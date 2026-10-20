//! 结构化错误：内部 [`WordError`]（单行四要素）+ 对外 [`ToolError`]（稳定错误码）。
//!
//! - 内部错误格式（镜像 rs-excel，一次读取即获全部定位要素）：
//!   `[操作] @ 对象路径 | 底层原因 | 提示: 修复建议`
//! - 对外错误：`code` + `category` + `retryable` + `side_effect_status` + `next_action` + `hresult`，
//!   调用方据此决定「改参数 / 换目标 / 重试 / 停手」，不解析英文文案。

use std::fmt;

use serde::{Deserialize, Serialize};

/// 单行结构化内部错误。
#[derive(Debug)]
pub struct WordError {
    /// 什么操作，如 `put('Text')`
    pub op: String,
    /// 对象路径，如 `WPS > report.docx > Range[120,128]`
    pub obj: String,
    /// 底层 COM 原因，如 `0x80020009 ...`
    pub source: String,
    /// 修复建议
    pub hint: String,
}

impl WordError {
    pub fn new(op: impl Into<String>, obj: impl Into<String>, source: impl Into<String>) -> Self {
        Self {
            op: op.into(),
            obj: obj.into(),
            source: source.into(),
            hint: String::new(),
        }
    }

    pub fn with_hint(mut self, hint: impl Into<String>) -> Self {
        self.hint = hint.into();
        self
    }

    /// 给已有错误补提示（尽量保持 WordError 类型不丢失上下文）。
    pub fn hint(e: anyhow::Error, hint: impl Into<String>) -> anyhow::Error {
        match e.downcast::<WordError>() {
            Ok(mut err) => {
                err.hint = hint.into();
                err.into()
            }
            Err(e) => WordError::new("error", "", e.to_string())
                .with_hint(hint)
                .into(),
        }
    }
}

impl fmt::Display for WordError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "[{}] @ {} | {}", self.op, self.obj, self.source)?;
        if !self.hint.is_empty() {
            write!(f, " | 提示: {}", self.hint)?;
        }
        Ok(())
    }
}

impl std::error::Error for WordError {}

/// 错误分类：调用方按它决定「改参数 / 换目标 / 重试 / 停手」。
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Category {
    /// 参数不合法（地址、形状、枚举值），未触碰 COM。
    Validation,
    /// 文档 / 段落 / 表格 / 书签不存在或已失效。
    TargetNotFound,
    /// 目标在读取后已变化（stale_target / version_conflict）。
    Conflict,
    /// 权限不足（覆盖文件、保护状态、只读）。
    Permission,
    /// 当前引擎/版本不支持该能力。
    Unsupported,
    /// COM 瞬时失败（WPS 常见的 0x80020009），可安全重试一次。
    ComTransient,
    /// 其他 COM / 应用层失败。
    ComFailure,
    /// 执行后恢复应用状态失败。
    Recovery,
}

impl Category {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Validation => "validation",
            Self::TargetNotFound => "target_not_found",
            Self::Conflict => "conflict",
            Self::Permission => "permission_denied",
            Self::Unsupported => "unsupported",
            Self::ComTransient => "com_transient",
            Self::ComFailure => "com_failure",
            Self::Recovery => "recovery_failure",
        }
    }

    pub fn default_hint(self) -> &'static str {
        match self {
            Self::Validation => "按 message 指出的字段/形状改一次再发",
            Self::TargetNotFound => {
                "先 list/read 拿到可用目标（document_id/table_id/书签名），不要猜"
            }
            Self::Conflict => "先重新 word_read 取最新锚点/版本，再生成修改",
            Self::Permission => "换路径，或关闭占用文件/请求解除保护，不能绕过安全策略",
            Self::Unsupported => "当前引擎不支持该能力，改用等价操作",
            Self::ComTransient => "可原样重试一次；重复失败说明参数或状态有问题，先读目标",
            Self::ComFailure => "按 message 里的对象路径与原因修正后重试，不要盲目重发",
            Self::Recovery => "应用状态可能没完全恢复，先看当前设置再继续",
        }
    }

    pub fn retryable(self) -> bool {
        matches!(self, Self::ComTransient)
    }
}

/// 失败时操作产生的副作用状态。
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SideEffectStatus {
    /// 可证明在产生副作用之前失败。
    NotApplied,
    /// 已生效。
    Applied,
    /// 部分生效（逐动作报告，不要让模型整组重跑）。
    PartiallyApplied,
    /// 无法确认是否已生效：先读目标再决定，不重发写操作。
    Unknown,
}

impl SideEffectStatus {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::NotApplied => "not_applied",
            Self::Applied => "applied",
            Self::PartiallyApplied => "partially_applied",
            Self::Unknown => "unknown",
        }
    }
}

/// 结构化对外错误（可序列化为工具回执里的 `error` 对象）。
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ToolError {
    /// 稳定错误码，如 `stale_target` / `com_transient`。
    pub code: String,
    pub category: Category,
    /// 出错的操作名。
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub operation: String,
    /// 出错的对象路径。
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub object_path: String,
    pub message: String,
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub hint: String,
    /// 底层 HRESULT（如 `0x80020009`）。
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hresult: Option<String>,
    pub retryable: bool,
    pub side_effect_status: SideEffectStatus,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next_action: Option<String>,
}

impl ToolError {
    pub fn new(code: impl Into<String>, category: Category, message: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            category,
            operation: String::new(),
            object_path: String::new(),
            message: message.into(),
            hint: category.default_hint().to_string(),
            hresult: None,
            retryable: category.retryable(),
            side_effect_status: SideEffectStatus::Unknown,
            next_action: None,
        }
    }

    /// 参数校验失败：确定没产生副作用。
    pub fn validation(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            retryable: false,
            side_effect_status: SideEffectStatus::NotApplied,
            ..Self::new(code, Category::Validation, message)
        }
    }

    pub fn unsupported(message: impl Into<String>) -> Self {
        Self {
            retryable: false,
            side_effect_status: SideEffectStatus::NotApplied,
            ..Self::new("unsupported_capability", Category::Unsupported, message)
        }
    }

    pub fn with_operation(mut self, operation: impl Into<String>) -> Self {
        self.operation = operation.into();
        self
    }

    pub fn with_object_path(mut self, path: impl Into<String>) -> Self {
        self.object_path = path.into();
        self
    }

    pub fn with_hint(mut self, hint: impl Into<String>) -> Self {
        self.hint = hint.into();
        self
    }

    pub fn with_hresult(mut self, hresult: impl Into<String>) -> Self {
        self.hresult = Some(hresult.into());
        self
    }

    pub fn with_side_effect(mut self, status: SideEffectStatus) -> Self {
        self.side_effect_status = status;
        self
    }

    pub fn with_next_action(mut self, action: impl Into<String>) -> Self {
        self.next_action = Some(action.into());
        self
    }

    /// 从 `anyhow` 错误转结构化错误（保留 `WordError` 的对象路径与 hint）。
    pub fn from_anyhow(e: &anyhow::Error) -> Self {
        if let Some(err) = e.downcast_ref::<ToolError>() {
            return err.clone();
        }
        let message = e.to_string();
        let (operation, object_path, hint, source) = match e.downcast_ref::<WordError>() {
            Some(err) => (
                err.op.clone(),
                err.obj.clone(),
                err.hint.clone(),
                err.source.clone(),
            ),
            None => (String::new(), String::new(), String::new(), message.clone()),
        };
        let hresult = extract_hresult(&source).or_else(|| extract_hresult(&message));
        let (code, category) = match hresult.as_deref() {
            Some("0x80020009") => ("com_transient", Category::ComTransient),
            Some("0x80020003") => ("unsupported_capability", Category::Unsupported),
            Some("0x80020005") => ("invalid_argument", Category::Validation),
            Some("0x800A03EC") => ("com_call_failed", Category::ComFailure),
            Some(_) => ("com_call_failed", Category::ComFailure),
            None => ("com_call_failed", Category::ComFailure),
        };
        Self {
            code: code.to_string(),
            category,
            operation,
            object_path,
            message,
            hint: if hint.is_empty() {
                category.default_hint().to_string()
            } else {
                hint
            },
            hresult,
            retryable: category.retryable(),
            side_effect_status: SideEffectStatus::Unknown,
            next_action: None,
        }
    }

    /// 消息里是否带 WPS 的瞬时错误码（调用方重试判据）。
    pub fn is_transient(message: &str) -> bool {
        message.contains("0x80020009")
    }
}

impl fmt::Display for ToolError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "[{}] {}", self.code, self.message)?;
        if !self.hint.is_empty() {
            write!(f, " | 提示: {}", self.hint)?;
        }
        Ok(())
    }
}

impl std::error::Error for ToolError {}

impl From<anyhow::Error> for ToolError {
    fn from(e: anyhow::Error) -> Self {
        Self::from_anyhow(&e)
    }
}

/// 从错误文本里取第一个 `0xXXXXXXXX`（统一大写）。
fn extract_hresult(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let mut i = 0;
    while i + 10 <= bytes.len() {
        if bytes[i] == b'0'
            && (bytes[i + 1] == b'x' || bytes[i + 1] == b'X')
            && bytes[i + 2..i + 10].iter().all(|b| b.is_ascii_hexdigit())
        {
            return Some(format!("0x{}", text[i + 2..i + 10].to_uppercase()));
        }
        i += 1;
    }
    None
}

/// RGB → OLE 颜色（COM 内部 `0x00BBGGRR`，与网页 RGB 相反）。
pub fn rgb(r: u8, g: u8, b: u8) -> u32 {
    (b as u32) << 16 | (g as u32) << 8 | r as u32
}

/// OLE 颜色 → `#RRGGBB` 文本。
pub fn ole_to_hex(c: u32) -> String {
    format!(
        "#{:02X}{:02X}{:02X}",
        c & 0xFF,
        (c >> 8) & 0xFF,
        (c >> 16) & 0xFF
    )
}

/// `#RRGGBB` / `RRGGBB` → OLE 颜色。
pub fn hex_to_ole(hex: &str) -> anyhow::Result<u32> {
    let s = hex.trim().trim_start_matches('#');
    if s.len() != 6 || !s.chars().all(|c| c.is_ascii_hexdigit()) {
        anyhow::bail!("颜色格式无效 '{hex}'，应为 #RRGGBB");
    }
    let r = u8::from_str_radix(&s[0..2], 16)?;
    let g = u8::from_str_radix(&s[2..4], 16)?;
    let b = u8::from_str_radix(&s[4..6], 16)?;
    Ok(rgb(r, g, b))
}
