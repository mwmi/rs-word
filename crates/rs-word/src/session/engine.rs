//! 引擎 adapter：分离 Word / WPS Writer 的连接与身份核验，共享其余领域操作。

use anyhow::Result;
use windows::core::{IUnknown, Interface, GUID, HSTRING, PCWSTR};
use windows::Win32::Foundation::S_OK;
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, GetRunningObjectTable, IDispatch, IMoniker,
    CLSCTX_LOCAL_SERVER,
};
use windows::Win32::System::Ole::GetActiveObject;

use crate::com::dispatch::IDispatchWrapper;
use crate::error::WordError;
use crate::types::{ConnectMode, EngineIdentity, WordEngine};

/// 引擎专属的小surface（ProgID 候选、进程名、身份核验）。领域操作共用，不重复实现。
pub trait WordEngineAdapter: Send + Sync {
    fn engine(&self) -> WordEngine;
    /// ProgID 候选（按顺序尝试；WPS 不写死单个值）。
    fn prog_id_candidates(&self) -> &'static [&'static str];
    /// 可接受的进程可执行文件名（小写）。
    fn accepted_exe_names(&self) -> &'static [&'static str];
    /// 附加身份关键词（`Application.Name` 应包含其一；空表示不校验）。
    fn name_keywords(&self) -> &'static [&'static str];
}

/// Microsoft Word。
pub struct MsWordAdapter;

impl WordEngineAdapter for MsWordAdapter {
    fn engine(&self) -> WordEngine {
        WordEngine::MsWord
    }
    fn prog_id_candidates(&self) -> &'static [&'static str] {
        &["Word.Application"]
    }
    fn accepted_exe_names(&self) -> &'static [&'static str] {
        &["winword.exe"]
    }
    fn name_keywords(&self) -> &'static [&'static str] {
        &["Word"]
    }
}

/// WPS Writer。ProgID 候选在目标机实测，命中项记录在身份结果里。
pub struct WpsWriterAdapter;

impl WordEngineAdapter for WpsWriterAdapter {
    fn engine(&self) -> WordEngine {
        WordEngine::WpsWriter
    }
    fn prog_id_candidates(&self) -> &'static [&'static str] {
        // 目标机实测：WPS Writer 注册的是 KWPS.Application（.9 为版本化别名）。
        // 不列 wps.Application（未注册，会产生 0x800401F3 掩盖真实原因）。
        &["KWPS.Application", "KWPS.Application.9"]
    }
    fn accepted_exe_names(&self) -> &'static [&'static str] {
        &["wps.exe"]
    }
    fn name_keywords(&self) -> &'static [&'static str] {
        &["WPS", "Kingsoft", "金山"]
    }
}

/// 按引擎取 adapter。
pub fn adapter_for(engine: WordEngine) -> Box<dyn WordEngineAdapter> {
    match engine {
        WordEngine::MsWord => Box::new(MsWordAdapter),
        WordEngine::WpsWriter => Box::new(WpsWriterAdapter),
    }
}

/// 连接结果。
pub struct Connected {
    pub root: IDispatchWrapper,
    pub identity: EngineIdentity,
    pub prog_id: String,
    pub owned: bool,
}

/// Application 对象的进程探测（`Hwnd` → PID → exe）。
///
/// **注意**：`Hwnd` 是 Excel 专有成员；Word 与 WPS Writer 的 `Application` **没有 Hwnd**，
/// 因此这里多数情况返回 None，需回退到 `Path`/`Name` 特征识别（见 [`detect_engine`]）。
fn exe_probe(root: &IDispatchWrapper) -> Option<(u32, String)> {
    let hwnd = root.get("Hwnd").ok()?.int().ok()?;
    let pid = crate::process::pid_of_hwnd(hwnd).ok()?;
    let exe = crate::process::exe_name_of(pid).ok()?;
    Some((pid, exe))
}

/// 依据 `Application.Path` / `Application.Name` 判定实际引擎。
///
/// WPS 会把自身伪装成 `Name="Microsoft Word"`，因此不能只看 Name；
/// Path 是可靠信号：WPS=`…\Kingsoft\WPS Office\…`，MS Office=`…\Microsoft Office\…`。
pub fn detect_engine(path: &str, name: &str) -> Option<WordEngine> {
    let p = path.to_lowercase();
    let n = name.to_uppercase();
    if p.contains("kingsoft")
        || p.contains("wps office")
        || p.contains("\\wps\\")
        || n.contains("WPS")
        || n.contains("KINGSOFT")
        || name.contains("金山")
    {
        return Some(WordEngine::WpsWriter);
    }
    if p.contains("microsoft office") || p.contains("\\office1") {
        return Some(WordEngine::MsWord);
    }
    None
}

/// COM 对象归属引擎的判定。
enum Probe {
    Unknown,
    Match,
    Other { who: String },
}

fn probe_of(root: &IDispatchWrapper, adapter: &dyn WordEngineAdapter) -> Probe {
    // 1) 进程证据：仅当 Application 暴露 Hwnd（Excel/WPS 表格）时可用。
    if let Some((pid, exe)) = exe_probe(root) {
        if adapter.accepted_exe_names().contains(&exe.as_str()) {
            return Probe::Match;
        }
        if !exe.is_empty() {
            return Probe::Other {
                who: format!("{exe}（PID {pid}）"),
            };
        }
    }
    // 2) Path/Name 特征（Word / WPS Writer 主要靠这里）。
    let path = root
        .get("Path")
        .and_then(|v| v.string())
        .unwrap_or_default();
    let name = root
        .get("Name")
        .and_then(|v| v.string())
        .unwrap_or_default();
    match detect_engine(&path, &name) {
        Some(e) if e == adapter.engine() => Probe::Match,
        Some(e) => Probe::Other {
            who: format!("{}（Path={path}）", e.label()),
        },
        None => Probe::Unknown,
    }
}

/// 读取身份信息并核验归属（不能只信 ProgID）。
pub fn verify_identity(
    adapter: &dyn WordEngineAdapter,
    root: &IDispatchWrapper,
    prog_id: &str,
) -> Result<EngineIdentity> {
    let engine = adapter.engine();
    let name = root
        .get("Name")
        .and_then(|v| v.string())
        .unwrap_or_default();
    let version = root.get("Version").and_then(|v| v.string()).ok();
    let path = root.get("Path").and_then(|v| v.string()).ok();
    let (pid, exe) = match exe_probe(root) {
        Some((pid, exe)) => (Some(pid), Some(exe)),
        None => (None, None),
    };

    if let Probe::Other { who } = probe_of(root, adapter) {
        return Err(WordError::new(
            "connect.verify",
            engine.label(),
            format!("ProgID '{prog_id}' 实际连接到 {who}"),
        )
        .with_hint(
            "该 ProgID 已被另一引擎接管（WPS 的 Office 兼容模式会把自身注册到 Word.Application 名下）；确认目标应用已安装并运行",
        )
        .into());
    }

    Ok(EngineIdentity {
        engine,
        prog_id: prog_id.to_string(),
        name,
        version,
        path,
        pid,
        exe,
    })
}

/// 统一连接：`Attach` 接管已运行实例，`New` 新建实例。
pub fn connect(adapter: &dyn WordEngineAdapter, mode: ConnectMode) -> Result<Connected> {
    let engine = adapter.engine();
    // 逐候选记录失败原因（不再只留最后一条，避免被低优先级候选掩盖真实原因）。
    let mut errors: Vec<String> = Vec::new();
    for prog_id in adapter.prog_id_candidates() {
        let prog = HSTRING::from(*prog_id);
        let clsid = match unsafe { CLSIDFromProgID(PCWSTR::from_raw(prog.as_ptr())) } {
            Ok(c) => c,
            Err(e) => {
                errors.push(format!("{prog_id}：ProgID 未注册（{}）", hr_of(&e)));
                continue;
            }
        };
        let disp = match mode {
            ConnectMode::New => unsafe {
                match CoCreateInstance::<_, IDispatch>(&clsid, None, CLSCTX_LOCAL_SERVER) {
                    Ok(d) => d,
                    Err(e) => {
                        errors.push(format!("{prog_id}：新建实例失败（{}）", hr_of(&e)));
                        continue;
                    }
                }
            },
            ConnectMode::Attach => match attach_engine(&clsid, adapter) {
                Ok(d) => d,
                Err(e) => {
                    errors.push(format!("{prog_id}：接管失败（{e}）"));
                    continue;
                }
            },
        };
        let root = IDispatchWrapper::root(disp, engine.label());
        let identity = match verify_identity(adapter, &root, prog_id) {
            Ok(id) => id,
            Err(e) => {
                errors.push(format!("{prog_id}：身份核验失败（{e}）"));
                continue;
            }
        };
        return Ok(Connected {
            root,
            identity,
            prog_id: (*prog_id).to_string(),
            owned: matches!(mode, ConnectMode::New),
        });
    }
    let mode_hint = if matches!(mode, ConnectMode::Attach) {
        "Attach 需目标应用已运行；若未打开请先打开文档，或改用 mode=new（新建实例）"
    } else {
        "确认目标应用已安装且可启动"
    };
    Err(WordError::new(
        "connect",
        engine.label(),
        format!(
            "无法连接 {}（候选 ProgID: {:?}）",
            engine.label(),
            adapter.prog_id_candidates()
        ),
    )
    .with_hint(format!(
        "{mode_hint}。逐候选结果：\n  {}",
        errors.join("\n  ")
    ))
    .into())
}

/// HRESULT 简要表示（如 `0x800401F3 无效的类字符串`）。
fn hr_of(e: &windows::core::Error) -> String {
    format!("0x{:08X} {}", e.code().0 as u32, e)
}

/// Attach：经 ROT 接管已运行的目标引擎实例（按进程可执行名校验，避免误连被接管的实例）。
fn attach_engine(clsid: &GUID, adapter: &dyn WordEngineAdapter) -> Result<IDispatch> {
    let mut unk: Option<IUnknown> = None;
    unsafe { GetActiveObject(clsid, None, &mut unk).ok() };
    if let Some(disp) = unk.and_then(|u| u.cast::<IDispatch>().ok()) {
        let root = IDispatchWrapper::root(disp.clone(), adapter.engine().label());
        if matches!(probe_of(&root, adapter), Probe::Match) {
            return Ok(disp);
        }
    }
    unsafe {
        let rot = GetRunningObjectTable(0)?;
        let enum_moniker = rot.EnumRunning()?;
        let mut mons = [None::<IMoniker>];
        let mut fetched = 0u32;
        while enum_moniker.Next(&mut mons, Some(&mut fetched)) == S_OK && fetched == 1 {
            let Some(m) = mons[0].take() else { break };
            let Ok(obj) = rot.GetObject(&m) else { continue };
            let Ok(disp) = obj.cast::<IDispatch>() else {
                continue;
            };
            let root = IDispatchWrapper::root(disp.clone(), adapter.engine().label());
            if matches!(probe_of(&root, adapter), Probe::Match) {
                return Ok(disp);
            }
        }
    }
    Err(WordError::new(
        "connect.attach",
        adapter.engine().label(),
        format!(
            "ROT 中未找到运行中的 {} 实例",
            adapter.accepted_exe_names().join("/")
        ),
    )
    .with_hint("确认目标应用已打开文档；如未运行，改用 connect new")
    .into())
}
