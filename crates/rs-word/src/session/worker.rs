//! STA worker 线程：COM 套间初始化、消息泵、请求串行处理。
//!
//! 所有 `IDispatch`/`VARIANT` 只在本线程创建/使用/Drop；外部线程只发送 typed DTO。

use std::thread::JoinHandle;

use anyhow::Result;
use crossbeam_channel::{Receiver, Sender};
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, PeekMessageW, TranslateMessage, MSG, PM_REMOVE,
};

use crate::error::ToolError;
use crate::session::{handle, Request, Response, Session};

/// 一条命令：请求 + oneshot 回复通道。
pub struct Command {
    pub req: Request,
    pub reply: Sender<Result<Response, ToolError>>,
}

/// 启动 STA worker，返回发送端与线程句柄。
pub fn spawn() -> Result<(Sender<Command>, JoinHandle<()>)> {
    let (tx, rx) = crossbeam_channel::unbounded::<Command>();
    let handle = std::thread::Builder::new()
        .name("rs-word-sta".into())
        .spawn(move || run(rx))?;
    Ok((tx, handle))
}

/// worker 主循环（在专属线程内执行）。连接由 `Request::Connect` 惰性触发。
fn run(rx: Receiver<Command>) {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }
    let mut session: Option<Session> = None;

    while let Ok(cmd) = rx.recv() {
        pump_messages();
        let is_shutdown = matches!(cmd.req, Request::Shutdown);
        let resp = handle(&mut session, cmd.req).map_err(|e| ToolError::from_anyhow(&e));
        let _ = cmd.reply.send(resp);
        if is_shutdown {
            break;
        }
    }
    drop(session);
    unsafe { CoUninitialize() };
}

/// 处理 Windows 消息，避免连接点/COM 回调被饿死。
fn pump_messages() {
    let mut msg = MSG::default();
    unsafe {
        while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}
