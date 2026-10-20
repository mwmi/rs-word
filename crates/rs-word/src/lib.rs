//! # rs-word
//!
//! Rust + COM 自动化 **WPS Writer / Microsoft Word**（Windows 专用），基于 `IDispatch` 后期绑定。
//!
//! ## 架构
//! - 所有 COM 对象只在专用 **STA worker 线程**创建/使用/Drop；对外 [`WordClient`] 是
//!   `Send + Sync` 的 channel 代理，AI/Tauri 只发 typed DTO，不接触 `IDispatch`/`VARIANT`。
//! - 超时返回 `result_unknown`，**不**声称取消已进入的 COM 调用。
//!
//! ## 约定
//! - 单位统一 points；颜色 `#RRGGBB`（facade 转 OLE）；Word 常量全部封装为 enum。
//! - 格式补丁字段为 `Option<T>`：`None` = 不改，永不默认置 0/false。
//! - Range 用 Word `Start/End`（非 UTF-8 字节）；跨 story 位置禁止混用。
//! - 每个写操作带 `operation_id`、分步结果与 read-back 验证。
#![cfg(windows)]
// 工具 schema 用单个 `json!` 描述全部工具，宏展开递归较深。
#![recursion_limit = "512"]
// ToolError 是刻意做得字段完整（code/category/message/hint/hresult/retryable/side_effect/next_action）的
// 结构化错误，供 AI 一次读取定位；函数签名统一 `Result<T, ToolError>` 保持同构，故不拆箱。
#![allow(clippy::result_large_err)]

pub mod advanced;
pub mod anchors;
pub mod app;
pub mod client;
pub mod com;
pub mod document;
pub mod edit;
pub mod error;
pub mod find;
pub mod format;
pub mod history;
pub mod layout;
pub mod media;
pub mod process;
pub mod range;
pub mod read;
pub mod review;
pub mod session;
pub mod table;
pub mod target;
pub mod tools;
pub mod types;
pub mod vba;
pub mod workflow;

pub use client::WordClient;
pub use error::{hex_to_ole, ole_to_hex, rgb, Category, SideEffectStatus, ToolError, WordError};
pub use types::*;

use std::sync::atomic::{AtomicU64, Ordering};

static OP_SEQ: AtomicU64 = AtomicU64::new(1);

/// 生成操作 ID（`op_<n>`）。
pub fn new_operation_id() -> String {
    format!("op_{}", OP_SEQ.fetch_add(1, Ordering::Relaxed))
}

pub mod prelude {
    pub use crate::client::WordClient;
    pub use crate::error::{rgb, ToolError, WordError};
    pub use crate::process::{exe_name_of, kill, kill_all, pid_of_hwnd, processes, ProcessInfo};
    pub use crate::types::*;
    pub use crate::{new_operation_id, Category, SideEffectStatus};
}
