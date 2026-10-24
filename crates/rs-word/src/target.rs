//! Target 解析：tagged enum → 固定 Range（含 story 与边界校验）。
//!
//! 表格多选（`Target::TableCells`）会拆成多段：连续矩形一段一个 Range，
//! 由 [`resolve_ranges`] 交给格式一次批量写完（见 [`crate::format::apply`]）。

use std::collections::BTreeMap;

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::range::Range;
use crate::types::{FindOptions, HfKind, Story, Target};

/// 解析结果：固定 Range + 所属 story。
pub struct Resolved {
    pub story: Story,
    pub range: Range,
    pub start: i32,
    pub end: i32,
}

fn ensure_doc(doc: &Document, id: &str) -> Result<()> {
    if doc.id() == id {
        Ok(())
    } else {
        Err(WordError::new(
            "target",
            doc.label(),
            format!("document_id 不匹配：期望 {id}，当前 {}", doc.id()),
        )
        .with_hint("旧 document_id 在引擎切换/重新连接后失效，先 word_session.status 取新 ID")
        .into())
    }
}

/// 从 document_id 解析 `tbl_<n>` 里的 1 基序号。
fn table_index(table_id: &str) -> Result<i32> {
    table_id
        .strip_prefix("tbl_")
        .and_then(|s| s.parse::<i32>().ok())
        .filter(|n| *n >= 1)
        .ok_or_else(|| anyhow::anyhow!("table_id '{table_id}' 非法，应为 tbl_<1基序号>"))
}

fn resolved(range: Range, story: Story) -> Result<Resolved> {
    let start = range.start()?;
    let end = range.end()?;
    Ok(Resolved {
        story,
        range,
        start,
        end,
    })
}

/// 1 基表格序号 → Table 对象。
fn table_of(doc: &Document, table_id: &str) -> Result<IDispatchWrapper> {
    let idx = table_index(table_id)?;
    doc.tables()?
        .call("Item", vec![ComValue::Int(idx)])?
        .idispatch_child(table_id)
}

/// 表格行列数；读不到按 0（= 不做越界校验）。
fn table_dims(t: &IDispatchWrapper) -> (u32, u32) {
    let count = |name: &str| -> u32 {
        t.get_obj(name, name)
            .and_then(|c| c.get("Count"))
            .and_then(|v| v.int())
            .map(|n| n.max(0) as u32)
            .unwrap_or(0)
    };
    (count("Rows"), count("Columns"))
}

/// 单元格矩形（1 基，两端含）。
struct CellRect {
    r1: u32,
    c1: u32,
    r2: u32,
    c2: u32,
}

/// `Table.Cell(Row, Column)`：位置参数（WPS 的具名参数支持不全）。
fn cell_of(t: &IDispatchWrapper, row: u32, column: u32) -> Result<IDispatchWrapper> {
    t.call(
        "Cell",
        vec![ComValue::Int(row as i32), ComValue::Int(column as i32)],
    )?
    .idispatch_child("Cell")
}

fn cell_range(t: &IDispatchWrapper, row: u32, column: u32) -> Result<Range> {
    cell_of(t, row, column)
        .and_then(|cell| cell.get_obj("Range", "Cell.Range"))
        .map(Range::new)
        .map_err(|e| {
            WordError::hint(
                e,
                format!("第 {row} 行第 {column} 列取不到（合并单元格或表格已变）：先 word_read view=table 看结构"),
            )
        })
}

/// 矩形选区 → 一个连续 Range（左上格起点 … 右下格终点，含单元格结束标记）。
///
/// 整表/整行/整列优先走 Table/Rows/Columns 的 Range：合并单元格也算在内，且少一次取格。
fn rect_range(doc: &Document, t: &IDispatchWrapper, rect: &CellRect) -> Result<Range> {
    let (rows, cols) = table_dims(t);
    let full_width = cols > 0 && rect.c1 <= 1 && rect.c2 >= cols;
    let full_height = rows > 0 && rect.r1 <= 1 && rect.r2 >= rows;
    if full_width && full_height {
        if let Ok(range) = t.get_obj("Range", "Table.Range") {
            return Ok(Range::new(range));
        }
    }
    if full_width && rect.r1 == rect.r2 {
        if let Ok(range) = t
            .get_obj("Rows", "Table.Rows")
            .and_then(|rows| rows.call("Item", vec![ComValue::Int(rect.r1 as i32)]))
            .and_then(|row| row.idispatch_child("Row"))
            .and_then(|row| row.get_obj("Range", "Row.Range"))
        {
            return Ok(Range::new(range));
        }
    }
    if full_height && rect.c1 == rect.c2 {
        if let Ok(range) = t
            .get_obj("Columns", "Table.Columns")
            .and_then(|columns| columns.call("Item", vec![ComValue::Int(rect.c1 as i32)]))
            .and_then(|column| column.idispatch_child("Column"))
            .and_then(|column| column.get_obj("Range", "Column.Range"))
        {
            return Ok(Range::new(range));
        }
    }
    let start = cell_range(t, rect.r1, rect.c1)?.start()?;
    let end = cell_range(t, rect.r2, rect.c2)?.end()?;
    if end < start {
        return Err(WordError::new(
            "target",
            doc.label(),
            format!(
                "单元格范围 {}-{} 行 × {}-{} 列取到的位置反了（{start}>{end}），表格结构可能已变",
                rect.r1, rect.r2, rect.c1, rect.c2
            ),
        )
        .with_hint("重新 word_read view=tables_summary 再定位")
        .into());
    }
    doc.range(start, end)
}

/// 一维起止：都对 0 = 取满；只给一端 = 单行/单列；给反了自动换过来。
fn axis(start: u32, end: u32, count: u32) -> (u32, u32) {
    let (s, e) = match (start, end) {
        (0, 0) => (1, count.max(1)),
        (0, e) => (e, e),
        (s, 0) => (s, s),
        (s, e) => (s.min(e), s.max(e)),
    };
    (s.max(1), e.max(s).max(1))
}

fn check_bounds(doc: &Document, r1: u32, r2: u32, c1: u32, c2: u32, rows: u32, cols: u32) -> Result<()> {
    if r1 > 0 && c1 > 0 && (rows == 0 || r2 <= rows) && (cols == 0 || c2 <= cols) {
        return Ok(());
    }
    Err(WordError::new(
        "target",
        doc.label(),
        format!("单元格范围 {r1}-{r2} 行 × {c1}-{c2} 列越界（表格共 {rows} 行 {cols} 列）"),
    )
    .with_hint("行列都从 1 起；尺寸先 word_read view=tables_summary 或 word_table action=inspect")
    .into())
}

/// 多选目标展开成若干矩形段：同一行相邻列合成一段，一段只发一次 COM。
fn cell_rects(doc: &Document, t: &IDispatchWrapper, target: &Target) -> Result<Vec<CellRect>> {
    let Target::TableCells {
        row_start,
        row_end,
        col_start,
        col_end,
        cells,
        ..
    } = target
    else {
        return Ok(Vec::new());
    };
    let (rows, cols) = table_dims(t);
    let mut rects: Vec<CellRect> = Vec::new();
    if cells.is_empty() {
        let (r1, r2) = axis(*row_start, *row_end, rows);
        let (c1, c2) = axis(*col_start, *col_end, cols);
        check_bounds(doc, r1, r2, c1, c2, rows, cols)?;
        rects.push(CellRect { r1, c1, r2, c2 });
        return Ok(rects);
    }
    let mut by_row: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    for cell in cells {
        let (row, column) = cell.rc();
        if row == 0 || column == 0 {
            return Err(WordError::new(
                "target",
                doc.label(),
                format!("cells 里有 0：({row},{column})"),
            )
            .with_hint("行列都从 1 起，例 cells:[{\"row\":1,\"column\":2}]")
            .into());
        }
        by_row.entry(row).or_default().push(column);
    }
    for (row, mut columns) in by_row {
        columns.sort_unstable();
        columns.dedup();
        check_bounds(
            doc,
            row,
            row,
            columns[0],
            *columns.last().unwrap_or(&columns[0]),
            rows,
            cols,
        )?;
        let (mut run_start, mut prev) = (columns[0], columns[0]);
        for &column in &columns[1..] {
            if column == prev + 1 {
                prev = column;
                continue;
            }
            rects.push(CellRect { r1: row, c1: run_start, r2: row, c2: prev });
            run_start = column;
            prev = column;
        }
        rects.push(CellRect { r1: row, c1: run_start, r2: row, c2: prev });
    }
    Ok(rects)
}

/// 解析 Target 为固定 Range。
pub fn resolve(app: &App, doc: &Document, target: &Target) -> Result<Resolved> {
    match target {
        Target::Document { document_id } => {
            ensure_doc(doc, document_id)?;
            resolved(doc.content()?, Story::Main)
        }
        Target::Selection { document_id } => {
            ensure_doc(doc, document_id)?;
            let sel = app.raw().get_obj("Selection", "Selection")?;
            resolved(
                Range::new(sel.get_obj("Range", "Selection.Range")?),
                Story::Main,
            )
        }
        Target::Range {
            document_id,
            story,
            start,
            end,
            expected_text,
            text_hash,
            ..
        } => {
            ensure_doc(doc, document_id)?;
            if *story != Story::Main {
                return Err(WordError::new(
                    "target",
                    doc.label(),
                    format!(
                        "MVP 暂不支持在 {story} story 上用 Range 定位（请用 header_footer target）"
                    ),
                )
                .with_hint("改用 kind=header_footer，或先限定 main story")
                .into());
            }
            let content = doc.content()?;
            let (c_start, c_end) = (content.start()?, content.end()?);
            if *start < c_start || *end > c_end || end < start {
                return Err(WordError::new(
                    "target",
                    doc.label(),
                    format!("Range[{start},{end}] 越界（正文 [{c_start},{c_end}]）"),
                )
                .with_hint("重新 word_read 取最新锚点")
                .into());
            }
            let range = doc.range(*start, *end)?;
            anchors::verify(&range, expected_text.as_deref(), text_hash.as_deref())?;
            resolved(range, Story::Main)
        }
        Target::Paragraph {
            document_id,
            paragraph_index,
            expected_text,
        } => {
            ensure_doc(doc, document_id)?;
            let paras = doc.raw().get_obj("Paragraphs", "Paragraphs")?;
            let count = paras.get("Count")?.int()?;
            if *paragraph_index < 1 || *paragraph_index as i32 > count {
                return Err(WordError::new(
                    "target",
                    doc.label(),
                    format!("段落号 {paragraph_index} 越界（共 {count} 段）"),
                )
                .with_hint("段落号为 1 基；重新 word_read view=outline 获取")
                .into());
            }
            let range = Range::new(
                paras
                    .call("Item", vec![ComValue::Int(*paragraph_index as i32)])?
                    .idispatch_child("Paragraph")?
                    .get_obj("Range", "Paragraph.Range")?,
            );
            if let Some(exp) = expected_text {
                anchors::verify(&range, Some(exp), None)?;
            }
            resolved(range, Story::Main)
        }
        Target::Find {
            document_id,
            query,
            occurrence,
        } => {
            ensure_doc(doc, document_id)?;
            resolved(
                crate::find::resolve_occurrence(doc, query, *occurrence, &FindOptions::default())?,
                Story::Main,
            )
        }
        Target::Bookmark {
            document_id,
            bookmark,
        } => {
            ensure_doc(doc, document_id)?;
            let range = Range::new(
                doc.raw()
                    .get_obj("Bookmarks", "Bookmarks")?
                    .call("Item", vec![ComValue::Str(bookmark.clone())])
                    .map_err(|e| {
                        WordError::hint(
                            e,
                            format!("书签 '{bookmark}' 不存在：先 word_read view=metadata"),
                        )
                    })?
                    .idispatch_child(bookmark)?
                    .get_obj("Range", "Bookmark.Range")?,
            );
            resolved(range, Story::Main)
        }
        Target::TableCell {
            document_id,
            table_id,
            row,
            column,
        } => {
            ensure_doc(doc, document_id)?;
            let idx = table_index(table_id)?;
            let range = Range::new(
                doc.tables()?
                    .call("Item", vec![ComValue::Int(idx)])?
                    .idispatch_child(table_id)?
                    .call_named(
                        "Cell",
                        vec![],
                        &[
                            ("Row", ComValue::Int(*row as i32)),
                            ("Column", ComValue::Int(*column as i32)),
                        ],
                    )?
                    .idispatch_child("Cell")?
                    .get_obj("Range", "Cell.Range")?,
            );
            resolved(range, Story::Main)
        }
        Target::Table {
            document_id,
            table_id,
        } => {
            ensure_doc(doc, document_id)?;
            resolved(
                Range::new(table_of(doc, table_id)?.get_obj("Range", "Table.Range")?),
                Story::Main,
            )
        }
        // 单 Range 的语义只在连续矩形上成立：零散多选由 resolve_ranges 拆段（格式专用）
        Target::TableCells {
            document_id,
            table_id,
            ..
        } => {
            ensure_doc(doc, document_id)?;
            let t = table_of(doc, table_id)?;
            let rects = cell_rects(doc, &t, target)?;
            if rects.len() != 1 {
                return Err(WordError::new(
                    "target",
                    doc.label(),
                    format!(
                        "target=table_cells 是 {} 段不连续选区，只有格式（word_format）能一次写完",
                        rects.len()
                    ),
                )
                .with_hint("改文字/删内容请给连续矩形，或逐个单元格定位")
                .into());
            }
            resolved(rect_range(doc, &t, &rects[0])?, Story::Main)
        }
        Target::HeaderFooter {
            document_id,
            section,
            hf_kind,
            variant,
        } => {
            ensure_doc(doc, document_id)?;
            let secs = doc.sections()?;
            let count = secs.get("Count")?.int()?;
            if *section < 1 || *section as i32 > count {
                return Err(WordError::new(
                    "target",
                    doc.label(),
                    format!("节号 {section} 越界（共 {count} 节）"),
                )
                .into());
            }
            let sec = secs
                .call("Item", vec![ComValue::Int(*section as i32)])?
                .idispatch_child(&format!("Section[{section}]"))?;
            let coll_name = match hf_kind {
                HfKind::Header => "Headers",
                HfKind::Footer => "Footers",
            };
            let range = Range::new(
                sec.get_obj(coll_name, coll_name)?
                    .call("Item", vec![ComValue::Int(variant.code())])?
                    .idispatch_child("HeaderFooter")?
                    .get_obj("Range", "HeaderFooter.Range")?,
            );
            let story = match hf_kind {
                HfKind::Header => Story::Header,
                HfKind::Footer => Story::Footer,
            };
            resolved(range, story)
        }
    }
}

/// 解析成多段 Range：表格多选按连续矩形拆段，其余目标只有一段（= [`resolve`]）。
pub fn resolve_ranges(app: &App, doc: &Document, target: &Target) -> Result<Vec<Resolved>> {
    let Target::TableCells {
        document_id,
        table_id,
        ..
    } = target
    else {
        return Ok(vec![resolve(app, doc, target)?]);
    };
    ensure_doc(doc, document_id)?;
    let t = table_of(doc, table_id)?;
    let rects = cell_rects(doc, &t, target)?;
    let mut out = Vec::with_capacity(rects.len());
    for rect in &rects {
        out.push(resolved(rect_range(doc, &t, rect)?, Story::Main)?);
    }
    Ok(out)
}
