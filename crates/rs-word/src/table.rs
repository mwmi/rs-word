//! 表格：创建/批量填充/读写/增删行列/合并拆分/底纹对齐。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::{hex_to_ole, WordError};
use crate::range::Range;
use crate::target::resolve;
use crate::types::{CellRange, TableAction, TableRequest, TableResult, TableSummary};

/// 表格动作执行入口。
pub fn exec(
    app: &App,
    doc: &Document,
    req: &TableRequest,
    doc_version: u64,
) -> Result<TableResult> {
    let mut out = TableResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        table_id: req.table_id.clone(),
        info: None,
        cells: Vec::new(),
        applied: Vec::new(),
    };
    match req.action {
        TableAction::Create => {
            let (id, shape) = create(app, doc, req)?;
            out.table_id = Some(id);
            out.info = Some(shape);
            out.applied.push("create".into());
        }
        TableAction::Inspect => {
            let (id, shape) = inspect(doc, req)?;
            out.table_id = Some(id);
            out.info = Some(shape);
        }
        TableAction::ReadCells => {
            out.cells = read_cells(doc, req)?;
        }
        TableAction::SetCells => {
            set_cells(doc, req)?;
            out.applied.push("set_cells".into());
        }
        TableAction::InsertRow => {
            let t = table(doc, req)?;
            if let Some(r) = req.range.as_ref().map(|x| x.row_start) {
                t.get_obj("Rows", "Rows")?
                    .call("Add", vec![ComValue::Int(r as i32)])?;
            } else {
                t.get_obj("Rows", "Rows")?.call("Add", vec![])?;
            }
            out.applied.push("insert_row".into());
        }
        TableAction::DeleteRow => {
            delete_lines(doc, req, true)?;
            out.applied.push("delete_row".into());
        }
        TableAction::InsertColumn => {
            let t = table(doc, req)?;
            let c = req.range.as_ref().map(|x| x.col_start).unwrap_or(0);
            if c > 0 {
                t.get_obj("Columns", "Columns")?
                    .call("Add", vec![ComValue::Int(c as i32)])?;
            } else {
                t.get_obj("Columns", "Columns")?.call("Add", vec![])?;
            }
            out.applied.push("insert_column".into());
        }
        TableAction::DeleteColumn => {
            delete_lines(doc, req, false)?;
            out.applied.push("delete_column".into());
        }
        TableAction::MergeCells => {
            merge_cells(doc, req)?;
            out.applied.push("merge_cells".into());
        }
        TableAction::SplitCell => {
            let r = req.range.clone().unwrap_or_default();
            let cell = cell_of(doc, req, r.row_start.max(1), r.col_start.max(1))?;
            cell.call("Split", vec![ComValue::Int(1), ComValue::Int(1)])?;
            out.applied.push("split_cell".into());
        }
        TableAction::SetShading => {
            set_shading(doc, req)?;
            out.applied.push("set_shading".into());
        }
        TableAction::SetAlignment => {
            set_alignment(doc, req)?;
            out.applied.push("set_alignment".into());
        }
        TableAction::SetRepeatHeader => {
            let t = table(doc, req)?;
            let n = req.header_rows.max(1);
            let rows = t.get_obj("Rows", "Rows")?;
            for i in 1..=n {
                rows.call("Item", vec![ComValue::Int(i as i32)])?
                    .idispatch_child("Row")?
                    .put("HeadingFormat", vec![ComValue::Bool(true)])?;
            }
            out.applied.push(format!("repeat_header x{n}"));
        }
        TableAction::Autofit => {
            let t = table(doc, req)?;
            t.call("AutoFitBehavior", vec![ComValue::Int(2)])?; // wdAutoFitWindow
            out.applied.push("autofit".into());
        }
        TableAction::DeleteTable => {
            let t = table(doc, req)?;
            t.call("Delete", vec![])?;
            out.applied.push("delete_table".into());
        }
        TableAction::SetWidths => {
            out.applied.extend(apply_widths(doc, req)?);
        }
        TableAction::SetHeights => {
            out.applied.extend(apply_heights(doc, req)?);
        }
        TableAction::SetBorders => {
            out.applied.extend(apply_borders(doc, req)?);
        }
        TableAction::SetAllowRowBreak => {
            out.applied.extend(apply_allow_row_break(doc, req)?);
        }
    }
    Ok(out)
}

/// 取出并校验 `format`（V1 表格格式）。
fn spec_of<'a>(doc: &Document, req: &'a TableRequest) -> Result<&'a crate::types::TableFormatSpec> {
    let spec = req.format.as_ref().ok_or_else(|| {
        WordError::new(
            "table",
            doc.label(),
            format!("{:?} 需要 format（TableFormatSpec）", req.action),
        )
        .with_hint("提供 format.widths_pt / heights / borders / allow_row_break")
    })?;
    spec.validate()
        .map_err(|e| WordError::new("table", doc.label(), e).into())
        .map(|_| spec)
}

/// 列宽：按 `range.col_start`（缺省第 1 列）起逐列设置并回读。
fn apply_widths(doc: &Document, req: &TableRequest) -> Result<Vec<String>> {
    let spec = spec_of(doc, req)?;
    let widths = spec
        .widths_pt
        .as_ref()
        .ok_or_else(|| WordError::new("table.set_widths", doc.label(), "format.widths_pt 为空"))?;
    let t = table(doc, req)?;
    let cols = t.get_obj("Columns", "Columns")?;
    let count = cols.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let start = req.range.as_ref().map(|r| r.col_start).unwrap_or(1).max(1);
    let mut applied = Vec::new();
    for (i, w) in widths.iter().enumerate() {
        let idx = start + i as u32;
        if idx as i32 > count {
            break;
        }
        let col = cols
            .call("Item", vec![ComValue::Int(idx as i32)])?
            .idispatch_child(&format!("Column[{idx}]"))?;
        col.put("Width", vec![ComValue::Float(*w as f64)])?;
        let got = col.get("Width").and_then(|v| v.float()).unwrap_or(0.0);
        let ok = (got - *w as f64).abs() <= 0.5;
        applied.push(format!(
            "col {idx} width {w:.1}pt (read-back {got:.1}, ok={ok})"
        ));
    }
    Ok(applied)
}

/// 行高：`HeightRule` 先于 `Height`（避免自动行高覆盖），逐行回读。
fn apply_heights(doc: &Document, req: &TableRequest) -> Result<Vec<String>> {
    let spec = spec_of(doc, req)?;
    let heights = spec
        .heights
        .as_ref()
        .ok_or_else(|| WordError::new("table.set_heights", doc.label(), "format.heights 为空"))?;
    let t = table(doc, req)?;
    let rows = t.get_obj("Rows", "Rows")?;
    let count = rows.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut applied = Vec::new();
    for h in heights {
        if h.row as i32 > count {
            continue;
        }
        let row = rows
            .call("Item", vec![ComValue::Int(h.row as i32)])?
            .idispatch_child(&format!("Row[{}]", h.row))?;
        row.put("HeightRule", vec![ComValue::Int(h.rule.code())])?;
        row.put("Height", vec![ComValue::Float(h.height_pt as f64)])?;
        let got = row.get("Height").and_then(|v| v.float()).unwrap_or(0.0);
        let rule = row.get("HeightRule").and_then(|v| v.int()).unwrap_or(-1);
        applied.push(format!(
            "row {} height {:.1}pt rule {} (read-back {got:.1}/rule {rule})",
            h.row,
            h.height_pt,
            h.rule.code()
        ));
    }
    Ok(applied)
}

/// `Border.LineWidth` 只接受 `WdLineWidth` 枚举，把 pt 映射到最近的合法值。
fn line_width_code(pt: f32) -> i32 {
    const TABLE: [(f32, i32); 9] = [
        (0.25, 2),
        (0.5, 4),
        (0.75, 6),
        (1.0, 8),
        (1.5, 12),
        (2.25, 18),
        (3.0, 24),
        (4.5, 36),
        (6.0, 48),
    ];
    TABLE
        .iter()
        .min_by(|a, b| {
            (a.0 - pt)
                .abs()
                .partial_cmp(&(b.0 - pt).abs())
                .unwrap_or(std::cmp::Ordering::Equal)
        })
        .map(|(_, code)| *code)
        .unwrap_or(8)
}

/// 边框：`All` 展开为 6 条边；逐条设置并回读 `LineStyle`。
fn apply_borders(doc: &Document, req: &TableRequest) -> Result<Vec<String>> {
    let spec = spec_of(doc, req)?;
    let borders = spec
        .borders
        .as_ref()
        .ok_or_else(|| WordError::new("table.set_borders", doc.label(), "format.borders 为空"))?;
    let t = table(doc, req)?;
    let coll = t.get_obj("Borders", "Borders")?;
    let mut applied = Vec::new();
    for b in borders {
        for code in b.edge.codes() {
            let edge = coll
                .call("Item", vec![ComValue::Int(code)])?
                .idispatch_child(&format!("Border[{code}]"))?;
            edge.put("LineStyle", vec![ComValue::Int(b.style.code())])?;
            if b.style != crate::types::BorderStyle::None {
                if let Some(w) = b.width_pt {
                    edge.put("LineWidth", vec![ComValue::Int(line_width_code(w))])?;
                }
            }
            if let Some(c) = &b.color {
                edge.put("Color", vec![ComValue::Int(hex_to_ole(c)? as i32)])?;
            }
            let got = edge.get("LineStyle").and_then(|v| v.int()).unwrap_or(-1);
            applied.push(format!(
                "{:?}/edge {code} -> style {} (read-back {got})",
                b.edge,
                b.style.code()
            ));
        }
    }
    Ok(applied)
}

/// 行跨页：`Rows.AllowBreakAcrossPages` 并回读。
fn apply_allow_row_break(doc: &Document, req: &TableRequest) -> Result<Vec<String>> {
    let spec = spec_of(doc, req)?;
    let allow = spec.allow_row_break.ok_or_else(|| {
        WordError::new(
            "table.set_allow_row_break",
            doc.label(),
            "format.allow_row_break 为空",
        )
    })?;
    let rows = table(doc, req)?.get_obj("Rows", "Rows")?;
    rows.put("AllowBreakAcrossPages", vec![ComValue::Bool(allow)])?;
    let got = rows
        .get("AllowBreakAcrossPages")
        .and_then(|v| v.bool())
        .ok();
    Ok(vec![format!("allow_row_break={allow} (read-back {got:?})")])
}

fn table(doc: &Document, req: &TableRequest) -> Result<crate::com::dispatch::IDispatchWrapper> {
    let id = req.table_id.as_deref().ok_or_else(|| {
        anyhow::anyhow!("该动作需要 table_id（先 Inspect 或 word_read tables_summary）")
    })?;
    let idx: i32 = id
        .strip_prefix("tbl_")
        .and_then(|s| s.parse().ok())
        .ok_or_else(|| anyhow::anyhow!("table_id '{id}' 非法"))?;
    doc.tables()?
        .call("Item", vec![ComValue::Int(idx)])?
        .idispatch_child(id)
}

fn cell_of(
    doc: &Document,
    req: &TableRequest,
    row: u32,
    column: u32,
) -> Result<crate::com::dispatch::IDispatchWrapper> {
    table(doc, req)?
        .call_named(
            "Cell",
            vec![],
            &[
                ("Row", ComValue::Int(row as i32)),
                ("Column", ComValue::Int(column as i32)),
            ],
        )?
        .idispatch_child("Cell")
}

/// 创建表格并批量填充。
fn create(app: &App, doc: &Document, req: &TableRequest) -> Result<(String, TableSummary)> {
    if req.data.is_empty() || req.data[0].is_empty() {
        return Err(WordError::new("table.create", doc.label(), "data 为空").into());
    }
    let cols = req.data[0].len();
    if req.data.iter().any(|r| r.len() != cols) {
        return Err(
            WordError::new("table.create", doc.label(), "各行单元格数不一致（非矩形）").into(),
        );
    }
    let rows = req.data.len();
    let target = req
        .target
        .as_ref()
        .ok_or_else(|| anyhow::anyhow!("create 需要 target 定位插入点"))?;
    let at = resolve(app, doc, target)?.range.collapse_end()?;
    let t = doc.tables()?.call_named(
        "Add",
        vec![],
        &[
            ("Range", ComValue::Obj(at.raw().disp.clone())),
            ("NumRows", ComValue::Int(rows as i32)),
            ("NumColumns", ComValue::Int(cols as i32)),
            ("DefaultTableBehavior", ComValue::Int(1)),
            ("AutoFitBehavior", ComValue::Int(2)),
        ],
    )?;
    let t = t.idispatch_child("Table")?;
    // 一次填充：逐格写 body range（跳过单元格终止符）。
    for (ri, row) in req.data.iter().enumerate() {
        for (ci, value) in row.iter().enumerate() {
            let cell = t.call_named(
                "Cell",
                vec![],
                &[
                    ("Row", ComValue::Int((ri + 1) as i32)),
                    ("Column", ComValue::Int((ci + 1) as i32)),
                ],
            )?;
            let cr = Range::new(
                cell.idispatch_child("Cell")?
                    .get_obj("Range", "Cell.Range")?,
            );
            let (s, e) = (cr.start()?, cr.end()?);
            // 单元格 body 不含末尾 \r\x07 终止标记。
            let body = doc.range(s, (e - 1).max(s))?;
            body.set_text(value)?;
        }
    }
    if req.header_rows > 0 {
        let rows_coll = t.get_obj("Rows", "Rows")?;
        for i in 1..=req.header_rows.min(rows as u32) {
            rows_coll
                .call("Item", vec![ComValue::Int(i as i32)])?
                .idispatch_child("Row")?
                .put("HeadingFormat", vec![ComValue::Bool(true)])?;
        }
    }
    let count = doc.table_count().unwrap_or(0);
    let id = format!("tbl_{count}");
    let info = shape_summary(doc, &id, Some(&t))?;
    Ok((id, info))
}

fn inspect(doc: &Document, req: &TableRequest) -> Result<(String, TableSummary)> {
    let id = req.table_id.clone().unwrap_or_default();
    let info = shape_summary(doc, &id, None)?;
    Ok((id, info))
}

fn shape_summary(
    doc: &Document,
    id: &str,
    t: Option<&crate::com::dispatch::IDispatchWrapper>,
) -> Result<TableSummary> {
    let owned;
    let t = match t {
        Some(t) => t,
        None => {
            let idx: i32 = id
                .strip_prefix("tbl_")
                .and_then(|s| s.parse().ok())
                .unwrap_or(1);
            owned = doc
                .tables()?
                .call("Item", vec![ComValue::Int(idx)])?
                .idispatch_child(id)?;
            &owned
        }
    };
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
    Ok(TableSummary {
        table_id: id.to_string(),
        index: 0,
        rows: rows.max(0) as u32,
        columns: cols.max(0) as u32,
        preview,
    })
}

fn read_cells(doc: &Document, req: &TableRequest) -> Result<Vec<Vec<String>>> {
    let t = table(doc, req)?;
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
    let rng = req.range.clone().unwrap_or(CellRange {
        row_start: 1,
        row_end: rows.max(0) as u32,
        col_start: 1,
        col_end: cols.max(0) as u32,
    });
    let mut out = Vec::new();
    for r in rng.row_start..=rng.row_end {
        let mut row = Vec::new();
        for c in rng.col_start..=rng.col_end {
            let text = match t.call_named(
                "Cell",
                vec![],
                &[
                    ("Row", ComValue::Int(r as i32)),
                    ("Column", ComValue::Int(c as i32)),
                ],
            ) {
                Ok(cc) => cc
                    .idispatch_child("Cell")
                    .and_then(|cell| cell.get_obj("Range", "Cell.Range"))
                    .and_then(|rr| rr.get("Text"))
                    .and_then(|v| v.string())
                    .map(|s| anchors::display_text(&s))
                    .unwrap_or_default(),
                Err(_) => "<merged>".into(),
            };
            row.push(text);
        }
        out.push(row);
    }
    Ok(out)
}

fn set_cells(doc: &Document, req: &TableRequest) -> Result<()> {
    let t = table(doc, req)?;
    let rng = req.range.clone().unwrap_or_default();
    for (ri, row) in req.data.iter().enumerate() {
        for (ci, value) in row.iter().enumerate() {
            let r = rng.row_start + ri as u32;
            let c = rng.col_start + ci as u32;
            let cell = t.call_named(
                "Cell",
                vec![],
                &[
                    ("Row", ComValue::Int(r as i32)),
                    ("Column", ComValue::Int(c as i32)),
                ],
            )?;
            let cr = Range::new(
                cell.idispatch_child("Cell")?
                    .get_obj("Range", "Cell.Range")?,
            );
            let (s, e) = (cr.start()?, cr.end()?);
            doc.range(s, (e - 1).max(s))?.set_text(value)?;
        }
    }
    Ok(())
}

/// 删除行/列（按索引从大到小，避免前删后索引位移）。
fn delete_lines(doc: &Document, req: &TableRequest, is_row: bool) -> Result<()> {
    let t = table(doc, req)?;
    let coll_name = if is_row { "Rows" } else { "Columns" };
    let coll = t.get_obj(coll_name, coll_name)?;
    let rng = req.range.clone().unwrap_or_default();
    let (start, end) = if is_row {
        (rng.row_start, rng.row_end)
    } else {
        (rng.col_start, rng.col_end)
    };
    for i in (start..=end.max(start)).rev() {
        if i == 0 {
            continue;
        }
        coll.call("Item", vec![ComValue::Int(i as i32)])?
            .idispatch_child("Item")?
            .call("Delete", vec![])?;
    }
    Ok(())
}

fn merge_cells(doc: &Document, req: &TableRequest) -> Result<()> {
    let rng = req.range.clone().unwrap_or_default();
    if rng.row_end < rng.row_start || rng.col_end < rng.col_start {
        return Err(WordError::new("table.merge", doc.label(), "合并范围非法（非矩形）").into());
    }
    // 取矩形对角两格 Merge（左下/右上无差别，Word 以连续矩形合并）。
    let a = cell_of(doc, req, rng.row_start, rng.col_start)?;
    let b = cell_of(doc, req, rng.row_end, rng.col_end)?;
    a.call_named(
        "Merge",
        vec![],
        &[("MergeTo", ComValue::Obj(b.disp.clone()))],
    )?;
    // 读回结构（合并后 shape 变化由调用方决定是否重新枚举）。
    Ok(())
}

fn set_shading(doc: &Document, req: &TableRequest) -> Result<()> {
    let rng = req.range.clone().unwrap_or_default();
    let t = table(doc, req)?;
    for r in rng.row_start..=rng.row_end {
        for c in rng.col_start..=rng.col_end {
            if let Ok(cell) = t.call_named(
                "Cell",
                vec![],
                &[
                    ("Row", ComValue::Int(r as i32)),
                    ("Column", ComValue::Int(c as i32)),
                ],
            ) {
                let cell = cell.idispatch_child("Cell")?;
                let hex = req
                    .data
                    .first()
                    .and_then(|row| row.first())
                    .map(|s| hex_to_ole(s))
                    .transpose()?
                    .unwrap_or(0x00FF_FFFF);
                cell.get_obj("Shading", "Shading")?
                    .put("BackgroundPatternColor", vec![ComValue::Int(hex as i32)])?;
            }
        }
    }
    Ok(())
}

fn set_alignment(doc: &Document, req: &TableRequest) -> Result<()> {
    let rng = req.range.clone().unwrap_or_default();
    let t = table(doc, req)?;
    // 垂直对齐：req.data[0][0] 语义值 0/1/2 = top/center/bottom（wdCellVerticalAlignment）。
    let v = req
        .data
        .first()
        .and_then(|row| row.first())
        .and_then(|s| s.parse::<i32>().ok())
        .unwrap_or(1);
    for r in rng.row_start..=rng.row_end {
        for c in rng.col_start..=rng.col_end {
            if let Ok(cell) = t.call_named(
                "Cell",
                vec![],
                &[
                    ("Row", ComValue::Int(r as i32)),
                    ("Column", ComValue::Int(c as i32)),
                ],
            ) {
                cell.idispatch_child("Cell")?
                    .put("VerticalAlignment", vec![ComValue::Int(v)])?;
            }
        }
    }
    Ok(())
}
