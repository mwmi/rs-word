//! 诊断探针：定位 Attach 为何匹配不到已运行的 WPS/Word 实例。
//!
//! 打印：ProgID→CLSID、GetActiveObject 结果、Application.Name/Version/Hwnd、
//! Hwnd→PID→exe，以及 ROT 枚举到的对象概览。只读，不修改、不关闭应用。
//!
//! 运行：`cargo run -p rs-word --example probe_progid`

use windows::core::{IUnknown, Interface, HSTRING, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::Foundation::S_OK;
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, GetRunningObjectTable, IDispatch, IMoniker,
    CLSCTX_LOCAL_SERVER,
};
use windows::Win32::System::Ole::GetActiveObject;
use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

fn hr(e: &windows::core::Error) -> String {
    format!("0x{:08X} {}", e.code().0 as u32, e)
}

fn describe(disp: &IDispatch) -> String {
    // 通过 IDispatchWrapper 读取 Name/Version/Hwnd（复用库底座）。
    use rs_word::com::dispatch::IDispatchWrapper;
    let w = IDispatchWrapper::root(disp.clone(), "probe");
    let name = w.get("Name").and_then(|v| v.string()).unwrap_or_default();
    let version = w
        .get("Version")
        .and_then(|v| v.string())
        .unwrap_or_default();
    let path = w.get("Path").and_then(|v| v.string()).unwrap_or_default();
    let hwnd = w.get("Hwnd").and_then(|v| v.int());
    let hwnd_s = match hwnd {
        Ok(h) => {
            let mut pid = 0u32;
            unsafe { GetWindowThreadProcessId(HWND(h as u32 as usize as *mut _), Some(&mut pid)) };
            let exe = rs_word::process::exe_name_of(pid).unwrap_or_default();
            format!("hwnd={h} pid={pid} exe={exe}")
        }
        Err(e) => format!("Hwnd 读取失败: {e}"),
    };
    format!("Name={name:?} Version={version:?} Path={path:?} {hwnd_s}")
}

fn probe_progid(prog_id: &str) {
    println!("\n--- ProgID {prog_id} ---");
    let p = HSTRING::from(prog_id);
    let clsid = match unsafe { CLSIDFromProgID(PCWSTR::from_raw(p.as_ptr())) } {
        Ok(c) => c,
        Err(e) => {
            println!("  CLSIDFromProgID 失败: {}", hr(&e));
            return;
        }
    };
    println!("  CLSID = {clsid:?}");

    // GetActiveObject
    let mut unk: Option<IUnknown> = None;
    match unsafe { GetActiveObject(&clsid, None, &mut unk) } {
        Ok(()) => match unk.and_then(|u| u.cast::<IDispatch>().ok()) {
            Some(d) => println!("  GetActiveObject: 成功 → {}", describe(&d)),
            None => println!("  GetActiveObject: 成功但非 IDispatch"),
        },
        Err(e) => println!("  GetActiveObject 失败: {}", hr(&e)),
    }

    // CoCreateInstance（本地服务器；Office 常返回已运行实例）
    match unsafe { CoCreateInstance::<_, IDispatch>(&clsid, None, CLSCTX_LOCAL_SERVER) } {
        Ok(d) => println!("  CoCreateInstance: 成功 → {}", describe(&d)),
        Err(e) => println!("  CoCreateInstance 失败: {}", hr(&e)),
    }
}

fn probe_rot() {
    println!("\n--- ROT 枚举 ---");
    let rot = match unsafe { GetRunningObjectTable(0) } {
        Ok(r) => r,
        Err(e) => {
            println!("  GetRunningObjectTable 失败: {}", hr(&e));
            return;
        }
    };
    let em = match unsafe { rot.EnumRunning() } {
        Ok(e) => e,
        Err(e) => {
            println!("  EnumRunning 失败: {}", hr(&e));
            return;
        }
    };
    let mut mons = [None::<IMoniker>];
    let mut fetched = 0u32;
    let mut n = 0;
    while unsafe { em.Next(&mut mons, Some(&mut fetched)) } == S_OK && fetched == 1 {
        let Some(m) = mons[0].take() else { break };
        n += 1;
        let obj = unsafe { rot.GetObject(&m) };
        let desc = match obj {
            Ok(o) => match o.cast::<IDispatch>() {
                Ok(d) => describe(&d),
                Err(_) => "<非 IDispatch>".to_string(),
            },
            Err(e) => format!("GetObject 失败: {}", hr(&e)),
        };
        println!("  [{n}] {desc}");
    }
    println!("  共 {n} 项");
}

fn main() {
    // 主线程需初始化 COM 套间，否则 GetActiveObject/ROT 会返回 0x800401F0。
    unsafe {
        let _ = windows::Win32::System::Com::CoInitialize(None);
    }
    println!("rs-word PROGID/ROT 诊断（只读）");
    probe_progid("KWPS.Application");
    probe_progid("Word.Application");
    probe_progid("KET.Application");
    probe_rot();
}
