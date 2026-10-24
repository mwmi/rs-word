//! 简易历史：journal（可逆转记录）+ 文件检查点；undo/redo 走原生 Document.Undo/Redo。
//!
//! COM 修改不保证都成为可靠的原生撤销步骤；不可精确逆转的动作标记 `reversible=false`。
//! 需要「一次调用 = 一步撤销」时用 [`begin_mark`] / [`end_mark`]（原生自定义撤销记录）。

use anyhow::Result;

use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::types::HistoryEntry;

/// 库级操作历史（后进先出）。
#[derive(Default)]
pub struct History {
    entries: Vec<HistoryEntry>,
}

impl History {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn record(&mut self, entry: HistoryEntry) {
        self.entries.push(entry);
    }

    /// 按文档过滤（新→旧）。
    pub fn list(&self, document_id: &str) -> Vec<HistoryEntry> {
        self.entries
            .iter()
            .rev()
            .filter(|e| e.document_id == document_id)
            .cloned()
            .collect()
    }

    pub fn clear(&mut self) {
        self.entries.clear();
    }

    /// 栈顶可逆转条目（用于 undo 前校验版本）。
    pub fn top(&self, document_id: &str) -> Option<&HistoryEntry> {
        self.entries
            .iter()
            .rev()
            .find(|e| e.document_id == document_id)
    }
}

/// 原生撤销（`Document.Undo`，Times 缺省 = 1）。
///
/// **必须在 Document 上调**：`Application.Undo` 只有 WPS 表格（ET）API 才有；
/// WPS Writer 与 MS Word 的撤销都挂在 Document 上（真机实测：在 Application 上
/// 调 `Undo` 报 0x80020006 未知名称，撤消整个不可用）。
pub fn undo(doc: &Document) -> Result<()> {
    doc.raw().call("Undo", vec![]).map(|_| ())
}

/// 原生重做（`Document.Redo`，Times 缺省 = 1）。同在 Document 上，与 undo 对称。
pub fn redo(doc: &Document) -> Result<()> {
    doc.raw().call("Redo", vec![]).map(|_| ())
}

/// 清空原生撤销栈（`Document.UndoClear`）：之后软件自己的 Ctrl+Z 也不能再撤。
pub fn undo_clear(doc: &Document) -> Result<()> {
    doc.raw().call("UndoClear", vec![]).map(|_| ())
}

/// 开始一段自定义撤销记录（`Application.UndoRecord.StartCustomRecord`）：
/// 从这一刻到 [`end_mark`] 之间的改动，在原生撤销栈里合成一步，名字就是 `label`。
///
/// 只在软件暴露 `UndoRecord` 时可用；WPS 不支持时这里报错，调用方自行降级。
pub fn begin_mark(app: &App, label: &str) -> Result<()> {
    app.raw()
        .get_obj("UndoRecord", "UndoRecord")?
        .call("StartCustomRecord", vec![ComValue::Str(label.to_string())])
        .map(|_| ())
}

/// 结束自定义撤销记录。与 [`begin_mark`] 成对调用；非自定义记录状态下调用会报错。
pub fn end_mark(app: &App) -> Result<()> {
    app.raw()
        .get_obj("UndoRecord", "UndoRecord")?
        .call("EndCustomRecord", vec![])
        .map(|_| ())
}

/// 创建文件检查点（Document.SaveCopyAs），返回绝对路径。
pub fn checkpoint(doc: &Document, dir: Option<&str>) -> Result<String> {
    let dir = match dir {
        Some(d) => crate::process::abs_path(d)?,
        None => std::env::temp_dir().to_string_lossy().into_owned(),
    };
    let stamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0);
    let name = doc.name().unwrap_or_else(|_| "document".into());
    let stem = std::path::Path::new(&name)
        .file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "document".into());
    let path = format!("{dir}\\rs_word_ckpt_{stem}_{stamp}.docx");
    doc.raw()
        .call(
            "SaveCopyAs",
            vec![crate::com::variant::ComValue::Str(path.clone())],
        )
        .map_err(|e| WordError::hint(e, "检查点写入失败：确认目录存在且可写"))?;
    Ok(path)
}
