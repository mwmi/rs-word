//! 读取：summary/outline/text/paragraphs/tables_summary/table/selection/metadata。
//! 默认只返回所需范围，带 limit/cursor/truncated/coverage，绝不静默伪装成全文。
//!
//! 默认视图（summary）是**一次给全的概览**：统计 + 大纲 + 正文首块 + 表格清单 + 当前选区。
//! 模型为「先看看文档有什么」连读三次是常态（实报），一次给齐后这类往返就没了。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::range::Range;
use crate::types::{Chunk, ReadRequest, ReadResult, ReadView, Story, TableSummary, Target, TextAnchor};

/// 单次正文的硬上限：`limit_chars=0`（要全文）也只给这么多，剩下的用 cursor 续读
const BODY_HARD_MAX: usize = 40_000;

/// 游标编码：`t:<char_offset>`（文本） / `p:<paragraph_index>`（段落）。
fn parse_cursor(c: &Option<String>) -> u32 {
    c.as_deref()
        .and_then(|s| s.split_once(':'))
        .and_then(|(_, n)| n.parse().ok())
        .unwrap_or(0)
}

/// 正文可读化：段落标记 `\r` 显示成换行（JSON 里 `\r` 只是一串转义符，模型读着累），
/// 单元格标记 `\x07` 显示成空格。两者都是**一对一**替换，字符偏移不变，`start/end` 仍可直接当锚点。
fn readable(raw: &str) -> String {
    anchors::display_text(raw).replace('\r', "\n").replace('\x07', " ")
}

/// 把正文的一段塞进 `chunks`（`limit = 0` 表示要全文，仍受 [`BODY_HARD_MAX`] 约束）。
fn push_body(result: &mut ReadResult, text: &str, base: i32, offset: usize, limit: usize) -> usize {
    let total = text.chars().count();
    let take = if limit == 0 { BODY_HARD_MAX } else { limit.min(BODY_HARD_MAX) };
    let chunk: String = text.chars().skip(offset).take(take).collect();
    let end = offset + chunk.chars().count();
    result.chunks.push(Chunk {
        start: base + offset as i32,
        end: base + end as i32,
        hash: anchors::hash_text(&chunk),
        text: chunk,
    });
    if end < total {
        result.truncated = true;
        result.next_cursor = Some(format!("t:{end}"));
    }
    end
}

/// 当前选区（空选区返回 None）：summary 里顺带给，模型不必为「有没有选中」单独跑一趟。
fn current_selection(app: &App, document_id: &str, doc_version: u64) -> Result<Option<TextAnchor>> {
    let sel = app.raw().get_obj("Selection", "Selection")?;
    let range = Range::new(sel.get_obj("Range", "Selection.Range")?);
    if range.start()? == range.end()? {
        return Ok(None);
    }
    Ok(Some(anchors::anchor(document_id, Story::Main, &range, doc_version, 120)?))
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
        chars_total: None,
    };
    match req.view {
        // 概览：一次给齐「多大 / 什么结构 / 正文开头 / 有哪些表 / 有没有选中」
        ReadView::Summary => {
            let content = doc.content()?;
            let text = readable(&content.text()?);
            let total = text.chars().count();
            result.chars_total = Some(total);
            result.summary = Some(format!(
                "字符 {total} / 段落 {} / 表格 {} / 修订 {} / 批注 {}",
                doc.paragraph_count().unwrap_or(0),
                doc.table_count().unwrap_or(0),
                doc.revision_count().unwrap_or(0),
                doc.comment_count().unwrap_or(0),
            ));
            result.outline = collect_outline(doc, req.max_items.min(40))?;
            result.tables = collect_tables(doc, 8)?;
            push_body(&mut result, &text, content.start()?, 0, req.limit_chars);
            if let Some(anchor) = current_selection(app, &req.document_id, doc_version)? {
                result.selection = Some(anchor);
            }
        }
        ReadView::Outline => {
            result.outline = collect_outline(doc, req.max_items)?;
        }
        ReadView::Text => {
            let offset = parse_cursor(&req.cursor) as usize;
            let content = doc.content()?;
            let text = readable(&content.text()?);
            result.chars_total = Some(text.chars().count());
            push_body(&mut result, &text, content.start()?, offset, req.limit_chars);
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
        ReadView::Selection => match current_selection(app, &req.document_id, doc_version)? {
            Some(anchor) => {
                result.selection = Some(anchor.clone());
                result.anchors.push(anchor);
            }
            // 明说「没选中」：不然模型会以为读漏了，再读一次
            None => result.summary = Some("no selection (empty)".into()),
        },
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

/// 段落的大纲级别：先认 `OutlineLevel`，读不到再按样式名兜底。
///
/// 中文版 WPS 的标题样式叫「标题 1」，且 `OutlineLevel` 常常读回 10（正文），
/// 只认前者会让大纲整份是空的——模型就失去「一次看懂结构」的抓手。
fn heading_level(para: &crate::com::dispatch::IDispatchWrapper) -> Option<u32> {
    if let Some(level) = para.get("OutlineLevel").and_then(|v| v.int()).ok() {
        if (1..=9).contains(&level) {
            return Some(level as u32);
        }
    }
    let style = para.get("Style").and_then(|v| v.string()).ok()?;
    let name = style.trim();
    let lower = name.to_lowercase();
    if !(lower.starts_with("heading") || name.starts_with("标题")) {
        return None;
    }
    let level = name
        .chars()
        .filter(|c| c.is_ascii_digit())
        .collect::<String>()
        .parse::<u32>()
        .unwrap_or(1);
    Some(level.clamp(1, 9))
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
        let Some(level) = heading_level(&para) else {
            continue;
        };
        let text = anchors::display_text(
            &para
                .get_obj("Range", "Paragraph.Range")?
                .get("Text")?
                .string()?,
        );
        if !text.trim().is_empty() {
            out.push(crate::types::OutlineItem {
                level,
                text: anchors::quote(&text, 80),
                paragraph_index: i as u32,
            });
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
        Target::TableCell { table_id, .. }
        | Target::Table { table_id, .. }
        | Target::TableCells { table_id, .. } => table_id.clone(),
        _ => {
            return Err(WordError::new(
                "read.table",
                doc.label(),
                "view=table 需要 target.kind=table_cell / table_cells / table 提供 table_id",
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
