//! 格式：字符、段落、样式、列表。

pub mod font;
pub mod lists;
pub mod paragraph;
pub mod styles;

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::document::Document;
use crate::target::resolve;
use crate::types::{FormatRequest, FormatResult, ListKind, Story};

pub use font::Font;
pub use paragraph::ParagraphFormat;
pub use styles::{StyleProfile, StyleRule};

/// 格式应用入口：先套样式，再应用明确字段覆盖，最后列表/重置；读取 Back 验证。
pub fn apply(
    app: &App,
    doc: &Document,
    req: &FormatRequest,
    doc_version: u64,
) -> Result<FormatResult> {
    let r = resolve(app, doc, &req.target)?.range;
    let mut applied = Vec::new();

    if let Some(style) = &req.style {
        styles::apply_style(&r, doc, style)?;
        applied.push(format!("style {style}"));
    }
    if let Some(font) = &req.font {
        applied.extend(r.font()?.apply(font)?);
    }
    if let Some(paragraph) = &req.paragraph {
        applied.extend(r.paragraph_format()?.apply(paragraph)?);
    }
    if let Some(list) = req.list {
        if list != ListKind::Custom {
            lists::apply_list(&r, list)?;
            applied.push(format!("list {list:?}"));
        }
    }
    if let Some(spec) = &req.list_template {
        lists::apply_template(app, &r, spec)?;
        applied.push(format!("list_template levels={}", spec.levels.len()));
    }
    if req.list == Some(ListKind::Custom) && req.list_template.is_none() {
        return Err(crate::error::WordError::new(
            "format",
            doc.label(),
            "list=custom 需要 list_template",
        )
        .with_hint("提供 list_template.levels（1..=9 级）")
        .into());
    }
    if let Some(reset) = &req.reset {
        if reset.font {
            r.font()?.reset()?;
            applied.push("reset.font".into());
        }
        if reset.paragraph {
            // 段落重置：套 Normal 并清零常见直接格式。
            styles::apply_style(&r, doc, "normal")?;
            applied.push("reset.paragraph".into());
        }
    }

    // read-back：重读目标文本作为验证证据。
    let anchor = anchors::anchor(doc.id(), Story::Main, &r, doc_version + 1, 80)?;
    Ok(FormatResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version: doc_version + 1,
        target: Some(anchor),
        applied,
        warnings: Vec::new(),
    })
}
