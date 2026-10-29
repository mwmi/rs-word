//! VBA：只读 `status`、运行既有宏 `run_existing_macro`，以及**受授权**的临时注入
//! `execute` + `cleanup_stale_temp_modules`（10 手册 §12 / 08 章）。
//!
//! 三条路径严格区分：
//! 1. 运行既有宏 —— `Application.Run(MacroName, varg1..varg30)`，不触碰 VBProject；
//! 2. 临时代码 —— `Document.VBProject → VBComponents.Add(1) → CodeModule.AddFromString →
//!    Application.Run → 只删除本次自建并登记归属的模块`；
//! 3. 默认关闭：需宿主显式 `allow_vba_injection=true` + `confirm="yes"`，绝不改安全设置绕过。
//!
//! **错误不弹窗、不挂死、不假成功**（四层，与 rs-excel 同构 —— 同一套 WPS/Office VBA
//! 运行时，2026-10-14 起在表格侧逐条实测出来）：
//!   ① 运行时错误：入口包进 `On Error GoTo` 的包装 Function，错误写成返回值带出来（否则
//!      WPS/Word 弹模态框，整条 COM 链路卡到用户手动关窗为止）；
//!   ② 编译错误：`Application.Run` 卡在模态「编译错误」框上**永远不返回**（实测 60s+），
//!      由 [VbaGuard] 看门狗在 Run 期间按**窗口类名**处理：模态框露面等 400ms 点掉并带回
//!      框内原文，VBE 窗口最多藏两次（循环藏是在跟 VBA 抢窗口，实测会把进程搞挂）；
//!   ③ 编译错误被点掉后 Run **自己不当回事**（不报错、只回个空值）—— 当成成功就是
//!      「假执行成功」（宿主实报：回执 status=ok，VBE 里其实停在中断态）。两个判据都要用：
//!      Run 期间**出现过 VBA 错误框**，或返回值**不是文本**（包装 Function 声明 `As String`，
//!      真跑起来必然是 VT_BSTR，回空 = 压根没执行）；
//!   ④ 中断态救援：确定框被点掉后 WPS 直接进**中断态**（VBE 标题带「[中断]」，一切
//!      Application 级 COM 调用被拒、Run 永远阻塞）。Run 期间由旁路线程盯 VBE 标题，
//!      发现就从**第二条 COM 连接**执行 VBE「重置」，见 [spawn_break_rescuer]。

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::thread::sleep;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use anyhow::Result;
use windows::Win32::Foundation::{HWND, LPARAM};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, EnumWindows, GetClassNameW, GetWindowTextW, IsWindowVisible, SendMessageW,
    ShowWindow, BM_CLICK, SW_HIDE, SW_SHOW,
};

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::{ComValue, WordValue};
use crate::document::Document;
use crate::error::{SideEffectStatus, WordError};
use crate::types::{
    ConnectMode, VbaArg, VbaExecuteRequest, VbaExecuteResult, VbaResult, VbaRunRequest, VbaStatus,
    WordEngine,
};

/// 本库自建临时模块的前缀（只清理登记归属的模块，以及同前缀的残留）。
pub const TEMP_PREFIX: &str = "RsWordTmp_";

/// vbext_ct_StdModule。
const VBEXT_CT_STD_MODULE: i32 = 1;

/// 拒绝出现在注入代码里的构造（模态框会阻塞 STA 线程；外部进程调用超出文档沙箱保证）。
const BANNED: [&str; 3] = ["MsgBox", "Shell", "CreateObject(\"WScript.Shell\")"];

/// `VbaExecuteResult.status` 的三种取值（宿主据此转成功/失败回执）
const STATUS_OK: &str = "ok";
const STATUS_COMPILE: &str = "compile_error";
const STATUS_RUNTIME: &str = "runtime_error";

/// 代码守卫：返回命中的被拒构造（`None` = 通过）。纯函数，可离线单测。
pub fn banned_construct(code: &str) -> Option<&'static str> {
    BANNED.iter().find(|b| code.contains(**b)).copied()
}

// ===== Run 期间的外观守卫 =====

/// VBE 主窗口的窗口类名（WPS/Office 通用）。**只能按类名认窗口** —— 模态错误框的标题也含
/// "Visual Basic"，按标题认会把模态框一起藏了，WPS 卡死得更彻底（藏了的模态框还在跑模态
/// 循环，整个软件的 COM 调用全部被拒）
const VBE_CLASS: &str = "wndclass_desked_gsk";
/// VBA 模态错误框：标准对话框类 + 标题含 Visual Basic（MsgBox 标题不是它，不会误伤）
const DIALOG_CLASS: &str = "#32770";
/// 看门狗轮询间隔
const WATCH_TICK: Duration = Duration::from_millis(80);
/// 对话框露面多久才去点：露面瞬间就点 WPS 接不住（点击后进程挂死，实测），
/// 等它初始化完再点，点完还要等一会儿验它真关了
const CLICK_DELAY: Duration = Duration::from_millis(400);
/// 每个对话框最多点几次（前两次点「确定」按钮，最后一次对框发 WM_CLOSE 兜底）
const MAX_CLICKS: u8 = 3;
/// 同一个 VBE 窗口最多藏几次：露面藏一次 + 对话框点掉后补一次，收尾再补一刀。
/// **绝不能循环藏** —— 实测每轮都 SW_HIDE 露头的 VBE 是在跟 VBA 抢窗口，进程整个挂死，
/// 比断点状态糟得多
const MAX_HIDES_PER_WINDOW: usize = 2;

/// 扫描到的窗口（hwnd, 类名, 标题）：只收 VBE 与模态框两类，别的不进锁
static FOUND: Mutex<Vec<(isize, String, String)>> = Mutex::new(Vec::new());
/// 模态框子窗口里的静态文本（错误原文）
static STATIC_TEXT: Mutex<Vec<String>> = Mutex::new(Vec::new());
/// 模态框子窗口里的按钮（hwnd, 标题）
static BUTTONS: Mutex<Vec<(isize, String)>> = Mutex::new(Vec::new());

unsafe extern "system" fn scan_top(hwnd: HWND, _: LPARAM) -> windows::core::BOOL {
    let mut cls = [0u16; 64];
    let m = GetClassNameW(hwnd, &mut cls);
    let class = String::from_utf16_lossy(&cls[..m as usize]);
    if class == VBE_CLASS || class == DIALOG_CLASS {
        let mut title = [0u16; 256];
        let n = GetWindowTextW(hwnd, &mut title);
        let text = String::from_utf16_lossy(&title[..n as usize]);
        FOUND.lock().unwrap().push((hwnd.0 as isize, class, text));
    }
    windows::core::BOOL(1)
}

unsafe extern "system" fn scan_static(hwnd: HWND, _: LPARAM) -> windows::core::BOOL {
    let mut cls = [0u16; 64];
    let m = GetClassNameW(hwnd, &mut cls);
    if String::from_utf16_lossy(&cls[..m as usize]) == "Static" {
        let mut title = [0u16; 512];
        let n = GetWindowTextW(hwnd, &mut title);
        let text = String::from_utf16_lossy(&title[..n as usize]).replace(['\r', '\n'], " ");
        if !text.trim().is_empty() {
            STATIC_TEXT.lock().unwrap().push(text);
        }
    }
    windows::core::BOOL(1)
}

unsafe extern "system" fn scan_button(hwnd: HWND, _: LPARAM) -> windows::core::BOOL {
    let mut cls = [0u16; 64];
    let m = GetClassNameW(hwnd, &mut cls);
    if String::from_utf16_lossy(&cls[..m as usize]) == "Button" {
        let mut title = [0u16; 64];
        let n = GetWindowTextW(hwnd, &mut title);
        BUTTONS
            .lock()
            .unwrap()
            .push((hwnd.0 as isize, String::from_utf16_lossy(&title[..n as usize])));
    }
    windows::core::BOOL(1)
}

fn windows_now() -> Vec<(isize, String, String)> {
    FOUND.lock().unwrap().clear();
    unsafe {
        let _ = EnumWindows(Some(scan_top), LPARAM(0));
    }
    FOUND.lock().unwrap().drain(..).collect()
}

fn children(dialog: isize, cb: unsafe extern "system" fn(HWND, LPARAM) -> windows::core::BOOL) {
    unsafe {
        let _ = EnumChildWindows(
            Some(HWND(dialog as *mut core::ffi::c_void)),
            Some(cb),
            LPARAM(0),
        );
    }
}

fn visible(hwnd: isize) -> bool {
    unsafe { IsWindowVisible(HWND(hwnd as *mut core::ffi::c_void)) }.as_bool()
}

fn hide(hwnd: isize) {
    unsafe { let _ = ShowWindow(HWND(hwnd as *mut core::ffi::c_void), SW_HIDE); };
}

fn show(hwnd: isize) {
    unsafe { let _ = ShowWindow(HWND(hwnd as *mut core::ffi::c_void), SW_SHOW); };
}

/// 读模态框里的错误原文（静态文本，如「编译错误： 方法或数据成员未找到」）
fn dialog_text(dialog: isize) -> String {
    STATIC_TEXT.lock().unwrap().clear();
    children(dialog, scan_static);
    STATIC_TEXT
        .lock()
        .unwrap()
        .drain(..)
        .collect::<Vec<_>>()
        .join(" ")
}

/// 点掉模态框：非最后一次点「确定/OK/是」按钮，最后一次对框发 WM_CLOSE 兜底。
/// 不点它，`Application.Run` 永远不返回
fn dismiss_dialog(dialog: isize, last_try: bool) {
    if !last_try {
        BUTTONS.lock().unwrap().clear();
        children(dialog, scan_button);
        let buttons: Vec<(isize, String)> = BUTTONS.lock().unwrap().drain(..).collect();
        for (hwnd, text) in buttons {
            let label = text.trim();
            if label == "确定" || label == "OK" || label == "是" {
                unsafe {
                    SendMessageW(HWND(hwnd as *mut core::ffi::c_void), BM_CLICK, None, None)
                };
                return;
            }
        }
    }
    unsafe { SendMessageW(HWND(dialog as *mut core::ffi::c_void), 0x0010, None, None) }; // WM_CLOSE
}

/// Run 期间的外观守卫（RAII）：VBA 模态错误框露面 → 等 [CLICK_DELAY] → 点掉；
/// VBE 窗口露面 → 最多藏 [MAX_HIDES_PER_WINDOW] 次；`finish` 时补藏一次、恢复用户**原本**
/// 开着的 VBE 窗口，交回捕获到的错误原文。
///
/// 为什么必须看门狗：这些窗口是在 `Application.Run` **执行期间**弹出来的，主线程堵在 COM
/// 调用里出不来，只能从旁线程处理
pub struct VbaGuard {
    stop: Arc<AtomicBool>,
    caught: Arc<Mutex<String>>,
    handle: Option<std::thread::JoinHandle<Vec<isize>>>,
}

impl VbaGuard {
    pub fn start() -> Self {
        let stop = Arc::new(AtomicBool::new(false));
        let caught = Arc::new(Mutex::new(String::new()));
        let handle = {
            let stop = stop.clone();
            let caught = caught.clone();
            std::thread::spawn(move || {
                // 第一轮就可见的 VBE 窗口 = 用户自己开着的，finish 时还给人家
                let mut keep: Vec<isize> = Vec::new();
                // (hwnd, 首次露面时刻, 已点次数)
                let mut dialogs: Vec<(isize, Instant, u8)> = Vec::new();
                // 每个 VBE 窗口已藏过几次（上限 MAX_HIDES_PER_WINDOW）
                let mut hides: Vec<isize> = Vec::new();
                let mut first = true;
                while !stop.load(Ordering::SeqCst) {
                    let now = windows_now();
                    let present: Vec<isize> = now.iter().map(|(hwnd, _, _)| *hwnd).collect();
                    for (hwnd, class, title) in &now {
                        if *class == VBE_CLASS {
                            if first && visible(*hwnd) {
                                keep.push(*hwnd);
                            }
                        } else if *class == DIALOG_CLASS
                            && title.contains("Visual Basic")
                            && !dialogs.iter().any(|(seen, _, _)| seen == hwnd)
                        {
                            dialogs.push((*hwnd, Instant::now(), 0));
                        }
                    }
                    first = false;
                    for (hwnd, seen, tries) in dialogs.iter_mut() {
                        if *tries >= MAX_CLICKS || seen.elapsed() < CLICK_DELAY {
                            continue;
                        }
                        if !present.contains(hwnd) {
                            continue; // 已经关了
                        }
                        // 点之前把原文读走；读不到也留个占位 —— 「有框」本身就是编译错误的信号
                        let text = dialog_text(*hwnd);
                        let note = if text.is_empty() {
                            "（VBA 错误对话框，未读到原文）".to_string()
                        } else {
                            text
                        };
                        if let Ok(mut slot) = caught.lock() {
                            if slot.is_empty() {
                                *slot = note;
                            }
                        }
                        dismiss_dialog(*hwnd, *tries + 1 >= MAX_CLICKS);
                        *tries += 1;
                        sleep(Duration::from_millis(500)); // 等它关，顺带别重复点
                    }
                    dialogs.retain(|(hwnd, _, tries)| *tries < MAX_CLICKS && present.contains(hwnd));
                    for (hwnd, class, _) in &now {
                        if *class == VBE_CLASS && visible(*hwnd) {
                            let times = hides.iter().filter(|hidden| *hidden == hwnd).count();
                            if times < MAX_HIDES_PER_WINDOW {
                                hides.push(*hwnd);
                                hide(*hwnd);
                            }
                        }
                    }
                    sleep(WATCH_TICK);
                }
                keep
            })
        };
        Self { stop, caught, handle: Some(handle) }
    }

    /// 停看门狗；Run 已返回，此时补藏一次 VBE（没人跟它抢了）、恢复用户原本开着的 VBE 窗口，
    /// 返回捕获到的错误框原文（没有则空串）
    pub fn finish(mut self) -> String {
        self.stop.store(true, Ordering::SeqCst);
        let keep = self
            .handle
            .take()
            .and_then(|handle| handle.join().ok())
            .unwrap_or_default();
        for (hwnd, class, _) in windows_now() {
            if class == VBE_CLASS && !keep.contains(&hwnd) && visible(hwnd) {
                hide(hwnd);
            }
        }
        for hwnd in keep {
            show(hwnd);
        }
        self.caught.lock().map(|text| text.clone()).unwrap_or_default()
    }
}

/// 把 VBE 从「中断态」拉回停止态（等价于人手点 VBE 工具栏的「重置」方块按钮）。
///
/// 中断态下一切 Application 级 COM 调用都被拒，但 VBE 仍可交互（人能在 VBE 里点恢复），
/// 所以走 VBE 对象模型：CommandBars 菜单里找 ID=228 的「重置」并执行（各语言标题不同 ——
/// Reset / 重置 / 重新设置，标题匹配保留兜底）。工程本就停止时执行重置是无害的空操作。
/// **谁调它**：见 [spawn_break_rescuer] —— 主线程堵在 Run 里时只有旁路线程能走到这里
fn exit_break_mode(raw: &IDispatchWrapper) {
    let Ok(vbe) = raw.get("VBE").and_then(|v| v.idispatch()) else { return };
    let Ok(bars) = vbe.get_obj("CommandBars", "CommandBars") else { return };
    // **不能按名称取**（Item("Menu Bar") 在 WPS 的 VBE 上必抛 0x80020009，实测）：按序号枚举
    let Ok(count) = bars.get("Count").and_then(|v| v.int()) else { return };
    for at in 1..=count {
        let Ok(bar) = bars
            .call("Item", vec![ComValue::Int(at)])
            .and_then(|v| v.idispatch())
        else {
            continue;
        };
        // 菜单栏的名字是**本地化**的（WPS 中文版叫「菜单条」，英文版 Menu Bar —— 实测）
        let name = bar.get("Name").ok().and_then(|v| v.string().ok()).unwrap_or_default();
        if !(name.eq_ignore_ascii_case("menu bar")
            || name.contains("菜单条")
            || name.contains("菜单栏"))
        {
            continue;
        }
        let Ok(tops) = bar.get_obj("Controls", "Controls") else { continue };
        let Ok(count) = tops.get("Count").and_then(|v| v.int()) else { continue };
        for at in 1..=count {
            let Ok(top) = tops
                .call("Item", vec![ComValue::Int(at)])
                .and_then(|v| v.idispatch())
            else {
                continue;
            };
            // 顶层是弹出菜单；普通按钮没有 Controls 子集合，读失败正好跳过
            let Ok(items) = top.get_obj("Controls", "Controls") else { continue };
            let Ok(n) = items.get("Count").and_then(|v| v.int()) else { continue };
            for index in 1..=n {
                let Ok(item) = items
                    .call("Item", vec![ComValue::Int(index)])
                    .and_then(|v| v.idispatch())
                else {
                    continue;
                };
                let caption = item.get("Caption").ok().and_then(|v| v.string().ok()).unwrap_or_default();
                let id = item.get("Id").ok().and_then(|v| v.int().ok());
                if id == Some(228)
                    || caption.to_ascii_lowercase().contains("reset")
                    || caption.contains("重置")
                {
                    let _ = item.call("Execute", vec![]);
                    return;
                }
            }
        }
    }
}

/// 中断态救援线程：Run 期间旁路轮询 VBE 主窗口标题，发现「[中断]/[break]」就从
/// **第二条 COM 连接**执行 VBE「重置」。
///
/// 为什么要它：模态框被点掉后 WPS 并不返回 Run，而是直接进中断态（VBE 标题「…[中断]…」），
/// 主线程从此堵死在 Run 里。中断态下 VBE 仍可交互，所以旁路连接（COM 服务器在断点态照常
/// 处理新进调用）能替人执行重置；工程回到停止态后 Run 自然返回
fn spawn_break_rescuer(engine: WordEngine, stop: Arc<AtomicBool>) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        // 本线程自带套间：worker 的 STA 套间属于它自己，跨套间不能直接用它的连接
        unsafe { let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED); }
        // 兜底上限：60s 内救不回来就放弃 —— 别陪着一条死连接转圈
        for _ in 0..120 {
            if stop.load(Ordering::SeqCst) {
                break;
            }
            sleep(Duration::from_millis(500));
            let broken = windows_now().iter().any(|(_, class, title)| {
                class == VBE_CLASS
                    && (title.contains("[中断]") || title.to_ascii_lowercase().contains("[break]"))
            });
            if !broken {
                continue;
            }
            let adapter = crate::session::engine::adapter_for(engine);
            if let Ok(connected) =
                crate::session::engine::connect(adapter.as_ref(), ConnectMode::Attach)
            {
                exit_break_mode(&connected.root);
            }
        }
        unsafe { CoUninitialize() };
    })
}

// ===== 执行 =====

/// 只读探测 VBA 能力（不执行未知宏、不改安全设置、不创建临时模块）。
pub fn status(app: &App, doc: &Document) -> VbaStatus {
    let host_run = app.raw().has_member("Run");
    let project_access = doc.raw().get_obj("VBProject", "VBProject").is_ok();
    let notes = if !host_run {
        Some("Application.Run 不可用（VBA 运行时缺失或引擎不支持）".into())
    } else if !project_access {
        Some(
            "VBProject 不可访问（信任中心/工程保护）；run_existing_macro 仍可用，临时注入需先在目标机授权"
                .into(),
        )
    } else {
        Some("具备注入的尝试条件；实际可用性以真机实测为准".into())
    };
    VbaStatus {
        available: host_run,
        host_run,
        project_access,
        injection_ready: host_run && project_access,
        notes,
    }
}

/// 运行既有宏（带外观守卫与中断救援：用户的工程本身有编译错误时同样会弹 VBE / 卡模态框）
pub fn run(app: &App, doc: &Document, req: &VbaRunRequest) -> Result<VbaResult> {
    if doc.id() != req.document_id {
        return Err(WordError::new("vba.run", doc.label(), "document_id 不匹配").into());
    }
    if req.macro_name.trim().is_empty() {
        return Err(WordError::new("vba.run", doc.label(), "macro_name 为空").into());
    }
    if req.args.len() > 30 {
        return Err(
            WordError::new("vba.run", doc.label(), "Application.Run 最多 30 个参数").into(),
        );
    }
    let (outcome, dialog) = run_guarded(app, &req.macro_name, &req.args);
    // Run 期间弹过 VBA 错误框（看门狗已点掉）= 宏没跑起来：工程编译不过，或宏自己抛了
    // 未处理的错误。Sub 正常结束也回空值，失败只能靠这个信号识破
    if !dialog.is_empty() {
        return Err(WordError::new(
            "vba.run",
            doc.label(),
            format!("宏 {} 没有执行：{dialog}", req.macro_name),
        )
        .with_hint("工程存在编译错误，或宏抛出了未处理的错误：在 VBE 里打开该工程修正后重试")
        .into());
    }
    let returned = outcome
        .map_err(|e| WordError::hint(e, "宏不存在或运行/编译错误；确认宏名与参数"))?
        .and_then(to_returned);
    Ok(VbaResult {
        document_id: req.document_id.clone(),
        operation_id: crate::new_operation_id(),
        returned,
        side_effect_status: SideEffectStatus::Unknown,
        applied: vec![format!("run {}", req.macro_name)],
    })
}

/// 一次带守卫的 `Application.Run`：返回（Run 结果, 期间捕获到的错误框原文）。
fn run_guarded(
    app: &App,
    macro_name: &str,
    args: &[VbaArg],
) -> (Result<Option<WordValue>>, String) {
    let guard = VbaGuard::start();
    let stop = Arc::new(AtomicBool::new(false));
    let rescuer = spawn_break_rescuer(app.engine(), Arc::clone(&stop));
    let outcome = run_macro(app, macro_name, args);
    let dialog = guard.finish();
    stop.store(true, Ordering::SeqCst);
    let _ = rescuer.join();
    (outcome, dialog)
}

/// 调用 `Application.Run`；返回宏返回值（`Sub` 通常为 None，**不等于失败**）。
fn run_macro(app: &App, macro_name: &str, args: &[VbaArg]) -> Result<Option<WordValue>> {
    let mut com_args: Vec<ComValue> = Vec::with_capacity(args.len() + 1);
    com_args.push(ComValue::Str(macro_name.to_string()));
    for a in args {
        com_args.push(match a {
            VbaArg::Str { value } => ComValue::Str(value.clone()),
            VbaArg::Int { value } => ComValue::Int(*value),
            VbaArg::Float { value } => ComValue::Float(*value),
            VbaArg::Bool { value } => ComValue::Bool(*value),
        });
    }
    let result = app.raw().call("Run", com_args)?;
    Ok(match result.value() {
        WordValue::Empty => None,
        other => Some(other),
    })
}

/// 宏返回值按文本带进回执（`Sub` 没有返回值，见 [run_macro]）
fn to_returned(value: WordValue) -> Option<String> {
    Some(value.to_string())
}

/// 临时注入（高风险）。`owned` 记录本库创建且尚未清理的模块名。
pub fn execute(
    app: &App,
    doc: &Document,
    req: &VbaExecuteRequest,
    owned: &mut Vec<String>,
    seq: &mut u64,
) -> Result<VbaExecuteResult> {
    if doc.id() != req.document_id {
        return Err(WordError::new("vba.execute", doc.label(), "document_id 不匹配").into());
    }
    if !req.allow_vba_injection {
        return Err(
            WordError::new("vba.execute", doc.label(), "临时 VBA 注入默认关闭")
                .with_hint("宿主显式授权后传 allow_vba_injection=true（并建议先 checkpoint）")
                .into(),
        );
    }
    if req.confirm.as_deref() != Some("yes") {
        return Err(
            WordError::new("vba.execute", doc.label(), "临时 VBA 注入为高风险操作")
                .with_hint("确认后传 confirm=\"yes\"")
                .into(),
        );
    }
    if !app.raw().has_member("Run") {
        return Err(WordError::new(
            "vba.execute",
            doc.label(),
            "vba_runtime_unavailable：Application.Run 不可用",
        )
        .with_hint("该引擎/版本没有可用的 VBA 运行时；先 vba.status 确认")
        .into());
    }
    if req.entry.trim().is_empty() {
        return Err(WordError::new("vba.execute", doc.label(), "entry 为空").into());
    }
    if req.code.trim().is_empty() {
        return Err(WordError::new("vba.execute", doc.label(), "code 为空").into());
    }
    // 注入的入口由下面的包装 Function 无参调用：带参入口没法既保住 On Error 包装、
    // 又把参数转发进去（VBA 里 ParamArray 转发不到带类型的形参），索性不支持并一次说清。
    // 要带参跑现有宏走 [run]（Application.Run 支持 30 个参数）
    if !req.args.is_empty() {
        return Err(WordError::new(
            "vba.execute",
            doc.label(),
            "注入执行不支持 args（入口由包装函数无参调用）",
        )
        .with_hint("把参数直接写进代码（Const / 变量）；要带参调用已存在的宏改用 vba.run")
        .into());
    }
    if let Some(banned) = banned_construct(&req.code) {
        return Err(WordError::new(
            "vba.execute",
            doc.label(),
            format!("注入代码包含被拒绝的构造 `{banned}`"),
        )
        .with_hint("不能用 MsgBox 回传结果，也不能启动外部进程；请返回值/文档状态回读")
        .into());
    }

    let project = doc.raw().get_obj("VBProject", "VBProject").map_err(|e| {
        WordError::hint(
            e,
            "vba_project_access_denied：需在信任中心显式允许 VBA 工程对象模型访问；不得绕过安全策略",
        )
    })?;
    let components = project.get_obj("VBComponents", "VBComponents")?;
    // 先清掉残留的临时模块：上次被中断 / 强杀时模块会留在用户工程里，而 `Application.Run`
    // 会编译**整个工程**，之后每次调用都弹一次「编译错误」（实报：点了确定下次还弹）
    drop_stale_modules(&components);
    let module_name = next_module_name(seq);
    let component = components
        .call("Add", vec![ComValue::Int(VBEXT_CT_STD_MODULE)])
        .map_err(|e| {
            let msg = e.to_string();
            if msg.contains("lock") || msg.contains("protect") {
                WordError::hint(
                    e,
                    "vba_project_locked：工程受保护，请先在目标机解除工程保护",
                )
            } else {
                WordError::hint(e, "无法新增标准模块：确认项目访问权限与 VBA 组件可用")
            }
        })?
        .idispatch_child(&module_name)?;
    component.put("Name", vec![ComValue::Str(module_name.clone())])?;
    let source = req.kind.wrap(&req.entry, &req.code);
    // 包装 Function：抓住入口过程未处理的运行时错误，写成返回值带出来（第 ① 层）。
    // **用户代码原样注入，不加行号前缀** —— 曾为了回报出错行给每行加数字行标签，实测行标签
    // 加不得（过程声明行 / End 行 / 注释行 / 空行 / Case 子句上都非法），整个模块编译不过
    let wrapper = format!("{module_name}_Run");
    let trapped = format!(
        "{source}\n\nPublic Function {wrapper}() As String\n\
         \x20   On Error GoTo VbaErr\n\
         \x20   {entry}\n\
         \x20   Exit Function\n\
         VbaErr:\n\
         \x20   {wrapper} = \"VBA 运行时错误 \" & Err.Number & \"：\" & Err.Description\n\
         End Function",
        entry = req.entry
    );
    component
        .get_obj("CodeModule", "CodeModule")?
        .call("AddFromString", vec![ComValue::Str(trapped)])
        .map_err(|e| WordError::hint(e, "VBA 代码语法错误或包含不允许的语句"))?;
    owned.push(module_name.clone());

    // 运行：只跑包装函数（限定名前缀避免撞名），期间开着外观守卫与中断救援
    let (outcome, dialog) = run_guarded(app, &format!("{module_name}.{wrapper}"), &[]);
    // 清理（只删本次创建并登记归属的模块）：执行结论优先，清理失败只记 cleanup_pending
    let mut cleanup_pending = false;
    if req.cleanup {
        match remove_component(&components, &component) {
            Ok(()) => owned.retain(|m| m != &module_name),
            Err(_) => cleanup_pending = true,
        }
    }

    let (status, returned, error_number, error_description) = match outcome {
        Ok(value) => {
            let text = match value {
                Some(WordValue::Text(text)) => Some(text),
                _ => None,
            };
            if !dialog.is_empty() {
                (STATUS_COMPILE, text, None, Some(dialog))
            } else {
                match text {
                    // 包装声明的是 As String，真跑起来必然是文本（空串也算）—— 拿不到就是
                    // 没编译过、函数压根没执行（第 ③ 层）
                    None => (
                        STATUS_COMPILE,
                        None,
                        None,
                        Some("没拿到包装函数的返回值：模块没通过编译，代码未执行".into()),
                    ),
                    Some(text) => split_error(text),
                }
            }
        }
        Err(e) => {
            let text = e.to_string();
            if !dialog.is_empty() || text.contains("Compile error") {
                let detail = if dialog.is_empty() { text } else { dialog };
                (STATUS_COMPILE, None, None, Some(detail))
            } else {
                (STATUS_RUNTIME, None, None, Some(text))
            }
        }
    };

    Ok(VbaExecuteResult {
        document_id: req.document_id.clone(),
        operation_id: crate::new_operation_id(),
        module_name: module_name.clone(),
        entry: req.entry.clone(),
        returned,
        status: status.to_string(),
        error_number,
        error_description,
        cleaned_up: req.cleanup && !cleanup_pending,
        cleanup_pending,
        side_effect_status: if status == STATUS_OK {
            SideEffectStatus::Applied
        } else {
            SideEffectStatus::Unknown
        },
        applied: vec![format!("execute {module_name}")],
    })
}

/// 生成并推进临时模块名（`RsWordTmp_<seq>_<unix_ms>`）。
fn next_module_name(seq: &mut u64) -> String {
    *seq += 1;
    let ms = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0);
    format!("{TEMP_PREFIX}{seq}_{ms}")
}

/// 清掉工程里残留的临时模块（只删 `RsWordTmp_` 前缀的 —— 那是本库的，用户模块一概不碰）。
/// 倒着删：移除会让后面的序号前移
fn drop_stale_modules(comps: &IDispatchWrapper) {
    let Ok(count) = comps.get("Count").and_then(|v| v.int()) else { return };
    for index in (1..=count).rev() {
        let Ok(comp) = comps
            .call("Item", vec![ComValue::Int(index)])
            .and_then(|v| v.idispatch())
        else {
            continue;
        };
        let name = comp.get("Name").ok().and_then(|v| v.string().ok()).unwrap_or_default();
        if name.starts_with(TEMP_PREFIX) {
            let _ = comps.call("Remove", vec![ComValue::Obj(comp.disp.clone())]);
        }
    }
}

/// 清理登记归属的临时模块（只删 `RsWordTmp_` 前缀且存在于 `owned` 的）。
pub fn cleanup_stale(doc: &Document, owned: &mut Vec<String>) -> Result<Vec<String>> {
    if owned.is_empty() {
        return Ok(Vec::new());
    }
    let project = doc
        .raw()
        .get_obj("VBProject", "VBProject")
        .map_err(|e| WordError::hint(e, "VBProject 不可访问：无法枚举临时模块"))?;
    let components = project.get_obj("VBComponents", "VBComponents")?;
    let count = components.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut removed = Vec::new();
    for i in (1..=count).rev() {
        let Ok(component) = components
            .call("Item", vec![ComValue::Int(i)])
            .and_then(|r| r.idispatch_child("VBComponent"))
        else {
            continue;
        };
        let name = component
            .get("Name")
            .and_then(|v| v.string())
            .unwrap_or_default();
        if !name.starts_with(TEMP_PREFIX) || !owned.contains(&name) {
            continue;
        }
        remove_component(&components, &component)?;
        removed.push(name);
    }
    owned.retain(|m| !removed.contains(m));
    Ok(removed)
}

/// 移除模块（`VBComponents.Remove(VBComponent)`）。
fn remove_component(components: &IDispatchWrapper, component: &IDispatchWrapper) -> Result<()> {
    components
        .call("Remove", vec![ComValue::Obj(component.disp.clone())])
        .map(|_| ())
}

/// 解析宏返回值：`ERR|<number>|<description>` 与包装函数抓到的
/// `VBA 运行时错误 <number>：<description>` 都算运行期错误。
fn split_error(
    returned: String,
) -> (&'static str, Option<String>, Option<i64>, Option<String>) {
    const PREFIX: &str = "VBA 运行时错误 ";
    let body = returned
        .strip_prefix("ERR|")
        .map(|rest| format!("VBA 运行时错误 {rest}").replace('|', "："))
        .or_else(|| {
            returned
                .starts_with(PREFIX)
                .then(|| returned.clone())
        });
    let Some(body) = body else {
        return (STATUS_OK, Some(returned), None, None);
    };
    let rest = &body[PREFIX.len()..];
    let (number, description) = match rest.split_once('：') {
        Some((number, description)) => (
            number.trim().parse::<i64>().ok(),
            Some(description.trim().to_string()),
        ),
        None => (rest.trim().parse::<i64>().ok(), None),
    };
    (
        STATUS_RUNTIME,
        Some(returned),
        number,
        Some(description.unwrap_or_else(|| rest.trim().to_string())),
    )
}
