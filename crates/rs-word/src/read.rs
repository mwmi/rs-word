//! 读取：summary/outline/text/paragraphs/tables_summary/table/selection/metadata。
//! 默认只返回所需范围，带 limit/cursor/truncated/coverage，绝不静默伪装成全文。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::range::Range;
use crate::types::{Chunk, ReadRequest, ReadResult, ReadView, Story, TableSummary, Target};

/// 游标编码：`t:<char_offset>`（文本） / `p:<paragraph_index>`（段落）。
fn parse_cursor(c: &Option<String>) -> u32 {
    c.as_deref()
        .and_then(|s| s.split_once(':'))
        .and_then(|(_, n)| n.parse().ok())
        .unwrap_or(0)
}

/// 读取文档。
pub fn read(app: &App, doc: &Document, req: &ReadRequest, doc_version: u64) -> Result<ReadResult> {
    let mut result = ReadResult {
        document_id: req.document_id.clone(),
        doc_version,
        summary: None,
        outline: Vec::new(),
        chunks: Vec::new(),
        paragraphs: Vec::new(),
        tables: Vec::new(),
        selection: None,
        comments: Vec::new(),
        revisions: Vec::new(),
        truncated: false,
        next_cursor: None,
        coverage: vec![Story::Main],
        anchors: Vec::new(),
    };
    match req.view {
        ReadView::Summary => {
            let content = doc.content()?;
            let text = anchors::display_text(&content.text()?);
            result.summary = Some(format!(
                "字符 {} / 段落 {} / 表格 {} / 修订 {} / 批注 {}",
                text.chars().count(),
                doc.paragraph_count().unwrap_or(0),
                doc.table_count().unwrap_or(0),
                doc.revision_count().unwrap_or(0),
                doc.comment_count().unwrap_or(0),
            ));
            // 顺带给出前若干大纲项，帮助模型定位。
            result.outline = collect_outline(doc, 12)?;
        }
        ReadView::Outline => {
            result.outline = collect_outline(doc, req.max_items)?;
        }
        ReadView::Text => {
            let offset = parse_cursor(&req.cursor);
            let content = doc.content()?;
            let text = anchors::display_text(&content.text()?);
            let total = text.chars().count();
            let chunk: String = text
                .chars()
                .skip(offset as usize)
                .take(req.limit_chars)
                .collect();
            let end = offset as usize + chunk.chars().count();
            result.chunks.push(Chunk {
                start: content.start()? + offset as i32,
                end: content.start()? + end as i32,
                hash: anchors::hash_text(&chunk),
                text: chunk,
            });
            if end < total {
                result.truncated = true;
                result.next_cursor = Some(format!("t:{end}"));
            }
        }
        ReadView::Paragraphs => {
            let start = parse_cursor(&req.cursor).max(1);
            let paras = doc.raw().get_obj("Paragraphs", "Paragraphs")?;
            let count = paras.get("Count")?.int().unwrap_or(0) as u32;
            let end = (start + req.max_items as u32).min(count + 1);
            for i in start..end {
                let r = Range::new(
                    paras
                        .call("Item", vec![ComValue::Int(i as i32)])?
                        .idispatch_child("Paragraph")?
                        .get_obj("Range", "Paragraph.Range")?,
                );
                result.paragraphs.push(anchors::display_text(&r.text()?));
            }
            if end <= count {
                result.truncated = true;
                result.next_cursor = Some(format!("p:{end}"));
            }
        }
        ReadView::TablesSummary => {
            result.tables = collect_tables(doc, req.max_items)?;
        }
        ReadView::Table => {
            let (table_id, cells) = read_table(doc, &req.target, req.max_items)?;
            result.tables.push(TableSummary {
                table_id,
                index: 0,
                rows: cells.len() as u32,
                columns: cells.first().map(|r| r.len()).unwrap_or(0) as u32,
                preview: cells.first().map(|r| r.join(" | ")).unwrap_or_default(),
            });
        }
        ReadView::Selection => {
            let sel = app.raw().get_obj("Selection", "Selection")?;
            let r = Range::new(sel.get_obj("Range", "Selection.Range")?);
            let a = anchors::anchor(&req.document_id, Story::Main, &r, doc_version, 120)?;
            if r.start()? == r.end()? {
                result.selection = None;
            } else {
                result.selection = Some(a.clone());
                result.anchors.push(a);
            }
        }
        ReadView::Comments => {
            let comments = doc.raw().get_obj("Comments", "Comments")?;
            let count = comments.get("Count")?.int().unwrap_or(0);
            for i in 1..=count.min(req.max_items as i32) {
                let c = comments
                    .call("Item", vec![ComValue::Int(i)])?
                    .idispatch_child("Comment")?;
                result
                    .comments
                    .push(crate::review::comment_info(&c, i as u32));
            }
        }
        ReadView::Revisions => {
            let revisions = doc.raw().get_obj("Revisions", "Revisions")?;
            let count = revisions.get("Count")?.int().unwrap_or(0);
            for i in 1..=count.min(req.max_items as i32) {
                let rev = revisions
                    .call("Item", vec![ComValue::Int(i)])?
                    .idispatch_child("Revision")?;
                result
                    .revisions
                    .push(crate::review::revision_info(&rev, i as u32, doc_version)?);
            }
        }
        ReadView::Metadata => {
            result.summary = Some(format!(
                "name={} path={} saved={} read_only={}",
                doc.name().unwrap_or_default(),
                doc.path().unwrap_or_default(),
                doc.is_saved().unwrap_or(false),
                doc.read_only().unwrap_or(false),
            ));
        }
    }
    Ok(result)
}

/// 收集大纲（Heading 1..9 段落）。workflow prepare 复用此处，故对外可见。
pub fn collect_outline(doc: &Document, max_items: usize) -> Result<Vec<crate::types::OutlineItem>> {
    let paras = doc.raw().get_obj("Paragraphs", "Paragraphs")?;
    let count = paras.get("Count")?.int().unwrap_or(0);
    let mut out = Vec::new();
    for i in 1..=count {
        if out.len() >= max_items {
            break;
        }
        let para = paras
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child("Paragraph")?;
        let level = para.get("OutlineLevel").and_then(|v| v.int()).unwrap_or(10);
        if (1..=9).contains(&level) {
            let text = anchors::display_text(
                &para
                    .get_obj("Range", "Paragraph.Range")?
                    .get("Text")?
                    .string()?,
            );
            if !text.trim().is_empty() {
                out.push(crate::types::OutlineItem {
                    level: level as u32,
                    text: anchors::quote(&text, 80),
                    paragraph_index: i as u32,
                });
            }
        }
    }
    Ok(out)
}

/// 收集表格摘要。
fn collect_tables(doc: &Document, max_items: usize) -> Result<Vec<TableSummary>> {
    let tables = doc.tables()?;
    let count = tables.get("Count")?.int().unwrap_or(0);
    let mut out = Vec::new();
    for i in 1..=count.min(max_items as i32) {
        let t = tables
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child(&format!("tbl_{i}"))?;
        let rows = t
            .get_obj("Rows", "Rows")
            .and_then(|r| r.get("Count"))
            .and_then(|v| v.int())
            .unwrap_or(0);
        let cols = t
            .get_obj("Columns", "Columns")
            .and_then(|c| c.get("Count"))
            .and_then(|v| v.int())
            .unwrap_or(0);
        let preview = t
            .get_obj("Range", "Table.Range")
            .and_then(|r| r.get("Text"))
            .and_then(|v| v.string())
            .map(|s| anchors::quote(&anchors::display_text(&s), 60))
            .unwrap_or_default();
        out.push(TableSummary {
            table_id: format!("tbl_{i}"),
            index: i as u32,
            rows: rows.max(0) as u32,
            columns: cols.max(0) as u32,
            preview,
        });
    }
    Ok(out)
}

/// 读取单表单元格（按 table_id）。
fn read_table(
    doc: &Document,
    target: &Target,
    max_items: usize,
) -> Result<(String, Vec<Vec<String>>)> {
    let table_id = match target {
        Target::TableCell { table_id, .. } => table_id.clone(),
        _ => {
            return Err(WordError::new(
                "read.table",
                doc.label(),
                "view=table 需要 target.kind=table_cell 提供 table_id",
            )
            .into())
        }
    };
    let idx: i32 = table_id
        .strip_prefix("tbl_")
        .and_then(|s| s.parse().ok())
        .ok_or_else(|| anyhow::anyhow!("table_id '{table_id}' 非法"))?;
    let t = doc
        .tables()?
        .call("Item", vec![ComValue::Int(idx)])?
        .idispatch_child(&table_id)?;
    let rows = t
        .get_obj("Rows", "Rows")
        .and_then(|r| r.get("Count"))
        .and_then(|v| v.int())
        .unwrap_or(0);
    let cols = t
        .get_obj("Columns", "Columns")
        .and_then(|c| c.get("Count"))
        .and_then(|v| v.int())
        .unwrap_or(0);
    let mut cells = Vec::new();
    for r in 1..=rows.min(max_items as i32) {
        let mut row = Vec::new();
        for c in 1..=cols.min(max_items as i32) {
            let cell = t.call_named(
                "Cell",
                vec![],
                &[("Row", ComValue::Int(r)), ("Column", ComValue::Int(c))],
            );
            let text = match cell {
                Ok(cc) => cc
                    .idispatch_child("Cell")
                    .and_then(|c| c.get_obj("Range", "Cell.Range"))
                    .and_then(|rr| rr.get("Text"))
                    .and_then(|v| v.string())
                    .map(|s| anchors::display_text(&s))
                    .unwrap_or_default(),
                Err(_) => "<merged>".into(),
            };
            row.push(text);
        }
        cells.push(row);
    }
    Ok((table_id, cells))
}
