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

/// 上溯到 Application：ROT 里注册的可能是 Document 等子对象（真机实测 Word 就是
/// 只把 Document 注册进 ROT，Application 反而不在），此时 `Path/Name` 落在文档上，
/// 必须读它的 `Application` 才能核验引擎归属。
fn as_application(root: IDispatchWrapper) -> IDispatchWrapper {
    if root.has_member("Documents") {
        return root;
    }
    if root.has_member("Application") {
        if let Ok(app) = root.get_obj("Application", "Application") {
            if app.has_member("Documents") {
                return app;
            }
        }
    }
    root
}

/// 实例的文档数。0 表示空壳实例（真机实测 WPS 的 `…/Automation -Embedding`
/// 后台进程就注册在 ROT 里且没有任何文档）。
fn document_count(app: &IDispatchWrapper) -> i32 {
    app.get_obj("Documents", "Documents")
        .and_then(|d| d.get("Count"))
        .and_then(|c| c.int())
        .unwrap_or(0)
}

/// 归属核验通过时返回 **Application** 包装（可能由 Document 上溯而来）。
fn matched_application(disp: IDispatch, adapter: &dyn WordEngineAdapter) -> Option<IDispatchWrapper> {
    let app = as_application(IDispatchWrapper::root(disp, adapter.engine().label()));
    matches!(probe_of(&app, adapter), Probe::Match).then_some(app)
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

/// 统一连接入口：按引擎分派到各自的连接写法。
///
/// - **WPS Writer**：接管已运行实例（`GetActiveObject` → ROT 枚举），按 `Application.Path`
///   核验归属；`mode=New` 时才新建，绝不主动启动 WPS。
/// - **Microsoft Word**：同样只接管已运行实例（找不到即报错，不新建）；`mode=New` 才新建。
pub fn connect(adapter: &dyn WordEngineAdapter, mode: ConnectMode) -> Result<Connected> {
    match adapter.engine() {
        WordEngine::WpsWriter => connect_wps(adapter, mode),
        WordEngine::MsWord => connect_word(adapter, mode),
    }
}

/// WPS 写法：逐候选 ProgID 接管已运行实例（不猜测、不主动新建）。
fn connect_wps(adapter: &dyn WordEngineAdapter, mode: ConnectMode) -> Result<Connected> {
    // 逐候选记录失败原因（不再只留最后一条，避免被低优先级候选掩盖真实原因）。
    let mut errors: Vec<String> = Vec::new();
    for &prog_id in adapter.prog_id_candidates() {
        let clsid = match clsid_of(prog_id) {
            Ok(c) => c,
            Err(msg) => {
                errors.push(msg);
                continue;
            }
        };
        let disp = match mode {
            ConnectMode::New => match create_instance(&clsid, prog_id) {
                Ok(d) => d,
                Err(msg) => {
                    errors.push(msg);
                    continue;
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
        match finish(adapter, disp, prog_id, matches!(mode, ConnectMode::New)) {
            Ok(c) => return Ok(c),
            Err(msg) => {
                errors.push(format!("{prog_id}：身份核验失败（{msg}）"));
                continue;
            }
        }
    }
    Err(connect_failed(adapter, mode, errors))
}

/// MS Word 写法：`Attach` 只接管已运行实例（与 WPS 一致，找不到就报错，绝不新建）；
/// 只有显式 `mode=New` 才新建实例。
fn connect_word(adapter: &dyn WordEngineAdapter, mode: ConnectMode) -> Result<Connected> {
    let prog_id = *adapter
        .prog_id_candidates()
        .first()
        .expect("Microsoft Word 至少有一个 ProgID 候选");
    let mut errors: Vec<String> = Vec::new();
    let clsid = match clsid_of(prog_id) {
        Ok(c) => c,
        Err(msg) => return Err(connect_failed(adapter, mode, vec![msg])),
    };
    let (disp, owned) = match mode {
        ConnectMode::New => match create_instance(&clsid, prog_id) {
            Ok(d) => (d, true),
            Err(msg) => return Err(connect_failed(adapter, mode, vec![msg])),
        },
        ConnectMode::Attach => match attach_engine(&clsid, adapter) {
            Ok(d) => (d, false),
            Err(e) => {
                errors.push(format!("{prog_id}：接管失败（{e}）"));
                return Err(connect_failed(adapter, mode, errors));
            }
        },
    };
    match finish(adapter, disp, prog_id, owned) {
        Ok(c) => Ok(c),
        Err(msg) => {
            errors.push(format!("{prog_id}：身份核验失败（{msg}）"));
            Err(connect_failed(adapter, mode, errors))
        }
    }
}

/// 连接收尾：包装 root 并核验身份（不能只信 ProgID）。
fn finish(
    adapter: &dyn WordEngineAdapter,
    disp: IDispatch,
    prog_id: &str,
    owned: bool,
) -> std::result::Result<Connected, String> {
    let root = IDispatchWrapper::root(disp, adapter.engine().label());
    let identity = verify_identity(adapter, &root, prog_id).map_err(|e| e.to_string())?;
    Ok(Connected {
        root,
        identity,
        prog_id: prog_id.to_string(),
        owned,
    })
}

/// ProgID → CLSID（未注册返回可读错误）。
fn clsid_of(prog_id: &str) -> std::result::Result<GUID, String> {
    let prog = HSTRING::from(prog_id);
    unsafe { CLSIDFromProgID(PCWSTR::from_raw(prog.as_ptr())) }
        .map_err(|e| format!("{prog_id}：ProgID 未注册（{}）", hr_of(&e)))
}

/// 新建实例（`CoCreateInstance`，本地服务器）。
fn create_instance(clsid: &GUID, prog_id: &str) -> std::result::Result<IDispatch, String> {
    unsafe { CoCreateInstance::<_, IDispatch>(clsid, None, CLSCTX_LOCAL_SERVER) }
        .map_err(|e| format!("{prog_id}：新建实例失败（{}）", hr_of(&e)))
}

/// 连接失败的统一错误（含逐候选原因与按模式的提示）。
fn connect_failed(
    adapter: &dyn WordEngineAdapter,
    mode: ConnectMode,
    errors: Vec<String>,
) -> anyhow::Error {
    let mode_hint = if matches!(mode, ConnectMode::Attach) {
        "Attach 需目标应用已运行；若未打开请先打开文档，或改用 mode=new（新建实例）"
    } else {
        "确认目标应用已安装且可启动"
    };
    WordError::new(
        "connect",
        adapter.engine().label(),
        format!(
            "无法连接 {}（候选 ProgID: {:?}）",
            adapter.engine().label(),
            adapter.prog_id_candidates()
        ),
    )
    .with_hint(format!(
        "{mode_hint}。逐候选结果：\n  {}",
        errors.join("\n  ")
    ))
    .into()
}

/// HRESULT 简要表示（如 `0x800401F3 无效的类字符串`）。
fn hr_of(e: &windows::core::Error) -> String {
    format!("0x{:08X} {}", e.code().0 as u32, e)
}

/// Attach：接管已运行的目标引擎实例。
///
/// 先 `GetActiveObject`（可能被另一引擎接管，需身份核验），再枚举 ROT。
/// 同一引擎可能注册多个对象（真机实测 WPS 会把无文档的后台 `…/Automation -Embedding`
/// 进程注册进 ROT；MS Word 只把 Document 注册进 ROT）。因此按下列优先级取实例：
/// **可见且有文档 > 有文档 > 可见 > 其他**，避免绑到后台空壳或自动化实例。
fn attach_engine(clsid: &GUID, adapter: &dyn WordEngineAdapter) -> Result<IDispatch> {
    let mut best: Option<(u8, IDispatch)> = None;

    let mut consider = |app: IDispatchWrapper| -> Option<IDispatch> {
        let score = match (app_visible(&app), document_count(&app) > 0) {
            (true, true) => 3,
            (false, true) => 2,
            (true, false) => 1,
            (false, false) => 0,
        };
        let better = match &best {
            Some((s, _)) => score > *s,
            None => true,
        };
        if better {
            best = Some((score, app.disp));
        }
        (score == 3).then(|| best.as_ref().unwrap().1.clone())
    };

    let mut unk: Option<IUnknown> = None;
    unsafe { GetActiveObject(clsid, None, &mut unk).ok() };
    if let Some(disp) = unk.and_then(|u| u.cast::<IDispatch>().ok()) {
        if let Some(app) = matched_application(disp, adapter) {
            if let Some(found) = consider(app) {
                return Ok(found);
            }
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
            let Some(app) = matched_application(disp, adapter) else {
                continue;
            };
            if let Some(found) = consider(app) {
                return Ok(found);
            }
        }
    }
    if let Some((_, disp)) = best {
        return Ok(disp);
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

/// 实例是否有可见窗口（后台 `-Embedding` 实例通常不可见）。
fn app_visible(app: &IDispatchWrapper) -> bool {
    app.get("Visible")
        .and_then(|v| v.bool())
        .unwrap_or(false)
}
