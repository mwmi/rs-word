//! 简易历史：journal（可逆转记录）+ 文件检查点；undo/redo 走原生 Application.Undo/Redo。
//!
//! COM 修改不保证都成为可靠的原生撤销步骤；不可精确逆转的动作标记 `reversible=false`。

use anyhow::Result;

use crate::app::App;
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

/// 原生撤销（Application.Undo）。
pub fn undo(app: &App) -> Result<()> {
    app.raw().call("Undo", vec![]).map(|_| ())
}

/// 原生重做（Application.Redo）。
pub fn redo(app: &App) -> Result<()> {
    app.raw().call("Redo", vec![]).map(|_| ())
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
