//! 格式：字符、段落、样式、列表。

pub mod font;
pub mod lists;
pub mod paragraph;
pub mod styles;

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::document::Document;
use crate::range::Range;
use crate::target::resolve_ranges;
use crate::types::{FormatRequest, FormatResult, ListKind, Story};

pub use font::Font;
pub use paragraph::ParagraphFormat;
pub use styles::{StyleProfile, StyleRule};

/// 格式应用入口：先套样式，再应用明确字段覆盖，最后列表/重置；读取 Back 验证。
///
/// 表格多选（`target.kind=table_cells`）会解析成多段 Range，一次调用把同一份补丁写到每段上
/// —— 模型不用（也没法）逐格调用。
pub fn apply(
    app: &App,
    doc: &Document,
    req: &FormatRequest,
    doc_version: u64,
) -> Result<FormatResult> {
    if req.list == Some(ListKind::Custom) && req.list_template.is_none() {
        return Err(crate::error::WordError::new(
            "format",
            doc.label(),
            "list=custom 需要 list_template",
        )
        .with_hint("提供 list_template.levels（1..=9 级）")
        .into());
    }
    let ranges = resolve_ranges(app, doc, &req.target)?;
    let mut applied = Vec::new();
    for resolved in &ranges {
        apply_range(app, doc, req, &resolved.range, &mut applied)?;
    }
    if ranges.len() > 1 {
        // 同一补丁写多段：回执里别把「bold」重复 N 次
        let mut unique: Vec<String> = Vec::with_capacity(applied.len() + 1);
        for item in applied {
            if !unique.contains(&item) {
                unique.push(item);
            }
        }
        unique.push(format!("ranges x{}", ranges.len()));
        applied = unique;
    }
    // read-back：重读目标文本作为验证证据（多段取外包范围）
    let anchor_range = if ranges.len() == 1 {
        ranges[0].range.clone()
    } else {
        let start = ranges.iter().map(|r| r.start).min().unwrap_or(0);
        let end = ranges.iter().map(|r| r.end).max().unwrap_or(start);
        doc.range(start, end)?
    };
    let anchor = anchors::anchor(doc.id(), Story::Main, &anchor_range, doc_version + 1, 80)?;
    Ok(FormatResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version: doc_version + 1,
        target: Some(anchor),
        applied,
        warnings: Vec::new(),
    })
}

/// 把一份格式补丁写到单段 Range 上（样式 → 字体 → 段落 → 列表 → 重置）。
fn apply_range(
    app: &App,
    doc: &Document,
    req: &FormatRequest,
    r: &Range,
    applied: &mut Vec<String>,
) -> Result<()> {
    if let Some(style) = &req.style {
        styles::apply_style(r, doc, style)?;
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
            lists::apply_list(r, list)?;
            applied.push(format!("list {list:?}"));
        }
    }
    if let Some(spec) = &req.list_template {
        lists::apply_template(app, r, spec)?;
        applied.push(format!("list_template levels={}", spec.levels.len()));
    }
    if let Some(reset) = &req.reset {
        if reset.font {
            r.font()?.reset()?;
            applied.push("reset.font".into());
        }
        if reset.paragraph {
            // 段落重置：套 Normal 并清零常见直接格式。
            styles::apply_style(r, doc, "normal")?;
            applied.push("reset.paragraph".into());
        }
    }
    Ok(())
}
