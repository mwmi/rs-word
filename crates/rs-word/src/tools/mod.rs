//! AI 工具门面：JSON Schema 定义 + 工具路由。

pub mod router;
pub mod schemas;

pub use router::dispatch;
pub use schemas::{tool_definitions, tool_names};
