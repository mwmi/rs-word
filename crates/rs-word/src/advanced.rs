//! 高级：脚注/尾注、交叉引用、文档结构快照与结构 diff（V2）。
//!
//! - 脚注/尾注属于独立 story，`Range` 位置不能与正文混用；这里只做 note 级读写与列示。
//! - 交叉引用用 `Range.InsertCrossReference`，只实现已验证的 `CrossReferenceType/Kind` 子集，
//!   未验证子集返回 `unsupported_capability`（不猜 Word 常量）。
//! - 结构 diff 由 `diff_snapshots` 纯函数完成，可离线单测；快照采集只读。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::range::Range;
use crate::target::resolve;
use crate::types::{
    AdvancedAction, AdvancedRequest, AdvancedResult, CrossRefInfo, NoteInfo, NoteKind,
    StructureChange, StructureParagraph, StructureSnapshot, TableShapeSnapshot,
};

/// wdFieldRef。
const FIELD_REF: i32 = 3;

/// 高级动作执行入口。
pub fn exec(
    app: &App,
    doc: &Document,
    req: &AdvancedRequest,
    doc_version: u64,
) -> Result<AdvancedResult> {
    let mut out = AdvancedResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        notes: Vec::new(),
        cross_refs: Vec::new(),
        snapshot: None,
        diff: Vec::new(),
        applied: Vec::new(),
    };
    match req.action {
        AdvancedAction::ListNotes => {
            out.notes = list_notes(doc, NoteKind::Footnote, req.max_items as usize)?;
            out.notes
                .extend(list_notes(doc, NoteKind::Endnote, req.max_items as usize)?);
        }
        AdvancedAction::AddFootnote | AdvancedAction::AddEndnote => {
            let kind = if matches!(req.action, AdvancedAction::AddFootnote) {
                NoteKind::Footnote
            } else {
                NoteKind::Endnote
            };
            let note = add_note(app, doc, req, kind)?;
            out.applied.push(format!("add {}", kind.label()));
            out.notes.push(note);
            out.doc_version = doc_version + 1;
        }
        AdvancedAction::InsertCrossReference => {
            let inserted = insert_cross_reference(app, doc, req)?;
            out.applied
                .push(format!("insert_cross_reference {inserted}"));
            out.doc_version = doc_version + 1;
        }
        AdvancedAction::ListCrossReferences => {
            out.cross_refs = list_cross_references(doc, req.max_items as usize)?;
            let fields = doc
                .content()
                .map(|r| {
                    r.raw()
                        .get_obj("Fields", "Fields")
                        .and_then(|f| f.get("Count"))
                        .and_then(|v| v.int())
                        .unwrap_or(0)
                })
                .unwrap_or(0);
            out.applied
                .push(format!("fields={fields} refs={}", out.cross_refs.len()));
        }
        AdvancedAction::StructureSnapshot => {
            let snap = snapshot(doc, doc_version, req.max_items as usize)?;
            out.applied
                .push(format!("structure_snapshot {} 段", snap.paragraphs.len()));
            out.snapshot = Some(snap);
        }
        AdvancedAction::StructureDiff => {
            let base = req.snapshot.as_ref().ok_or_else(|| {
                WordError::new(
                    "advanced",
                    doc.label(),
                    "structure_diff 需要 snapshot（来自 structure_snapshot）",
                )
                .with_hint("先执行 action=structure_snapshot 保存基线")
            })?;
            if base.document_id != doc.id() {
                return Err(WordError::new(
                    "advanced",
                    doc.label(),
                    format!(
                        "快照来自其它文档（{}），不能与当前文档比较",
                        base.document_id
                    ),
                )
                .with_hint("对当前文档重新取快照")
                .into());
            }
            let current = snapshot(doc, doc_version, req.max_items as usize)?;
            out.diff = diff_snapshots(base, &current);
            out.applied
                .push(format!("structure_diff 变更 {} 项", out.diff.len()));
            out.snapshot = Some(current);
        }
    }
    Ok(out)
}

/// 安全读取集合计数（成员缺失按 0）。
fn count_of(doc: &Document, member: &str) -> u32 {
    doc.raw()
        .get_obj(member, member)
        .and_then(|c| c.get("Count"))
        .and_then(|v| v.int())
        .ok()
        .unwrap_or(0)
        .max(0) as u32
}

/// 列示脚注/尾注。
fn list_notes(doc: &Document, kind: NoteKind, max_items: usize) -> Result<Vec<NoteInfo>> {
    let Some(coll) = doc.raw().get_obj(kind.collection(), kind.collection()).ok() else {
        return Ok(Vec::new());
    };
    let count = coll.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut out = Vec::new();
    for i in 1..=count.min(max_items as i32) {
        let note = coll
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child(&format!("{}[{}]", kind.label(), i))?;
        let range = Range::new(note.get_obj("Range", "Note.Range")?);
        let text = range.text().unwrap_or_default();
        out.push(NoteInfo {
            note_id: format!("{}_{i}", kind.prefix()),
            kind: kind.label().to_string(),
            reference: note
                .get("Reference")
                .and_then(|v| v.string())
                .ok()
                .filter(|s| !s.is_empty()),
            quote: anchors::quote(&text, 80),
            start: range.start().ok(),
            end: range.end().ok(),
        });
    }
    Ok(out)
}

/// 添加脚注/尾注（`Footnotes.Add(Range, Reference, Text)`）。
fn add_note(app: &App, doc: &Document, req: &AdvancedRequest, kind: NoteKind) -> Result<NoteInfo> {
    let target = req
        .target
        .as_ref()
        .ok_or_else(|| WordError::new("advanced", doc.label(), "add_note 需要 target"))?;
    let text = req
        .text
        .as_ref()
        .ok_or_else(|| WordError::new("advanced", doc.label(), "add_note 需要 text"))?;
    let at = resolve(app, doc, target)?.range;
    let coll = doc
        .raw()
        .get_obj(kind.collection(), kind.collection())
        .map_err(|e| {
            WordError::hint(
                e,
                format!("该引擎不支持 {}（{}）", kind.label(), kind.collection()),
            )
        })?;
    let note = coll
        .call_named(
            "Add",
            vec![],
            &[
                ("Range", ComValue::Obj(at.raw().disp.clone())),
                (
                    "Reference",
                    ComValue::Str(req.reference.clone().unwrap_or_default()),
                ),
                ("Text", ComValue::Str(text.clone())),
            ],
        )
        .map_err(|e| {
            WordError::hint(
                e,
                "检查插入点是否为有效正文位置（页眉/脚注内不能再次插入脚注）",
            )
        })?
        .idispatch_child("Note")?;
    // 回读：note 文本应包含写入内容。
    let range = Range::new(note.get_obj("Range", "Note.Range")?);
    let got = range.text().unwrap_or_default();
    let idx = coll.get("Count").and_then(|v| v.int()).unwrap_or(1).max(1);
    Ok(NoteInfo {
        note_id: format!("{}_{idx}", kind.prefix()),
        kind: kind.label().to_string(),
        reference: note
            .get("Reference")
            .and_then(|v| v.string())
            .ok()
            .filter(|s| !s.is_empty()),
        quote: anchors::quote(&got, 80),
        start: range.start().ok(),
        end: range.end().ok(),
    })
}

/// 插入交叉引用。
fn insert_cross_reference(app: &App, doc: &Document, req: &AdvancedRequest) -> Result<String> {
    let ty = req.xref_type.ok_or_else(|| {
        WordError::new(
            "advanced",
            doc.label(),
            "insert_cross_reference 需要 xref_type",
        )
    })?;
    let kind = req.xref_kind.ok_or_else(|| {
        WordError::new(
            "advanced",
            doc.label(),
            "insert_cross_reference 需要 xref_kind",
        )
    })?;
    if !ty.allowed_kinds().contains(&kind) {
        return Err(WordError::new(
            "advanced",
            doc.label(),
            format!("{ty:?} 不支持引用内容 {kind:?}（未验证的常量不做猜测）"),
        )
        .with_hint(format!("改用 {:?}", ty.allowed_kinds()))
        .into());
    }
    let item = req.reference_item.as_deref().ok_or_else(|| {
        WordError::new(
            "advanced",
            doc.label(),
            "insert_cross_reference 需要 reference_item（书签名/标题文字）",
        )
    })?;
    let target = req.target.as_ref().ok_or_else(|| {
        WordError::new(
            "advanced",
            doc.label(),
            "insert_cross_reference 需要 target",
        )
    })?;
    let at = resolve(app, doc, target)?.range.collapse_start()?;
    at.raw().call_named(
        "InsertCrossReference",
        vec![],
        &[
            ("ReferenceType", ComValue::Int(ty.code())),
            ("ReferenceKind", ComValue::Int(kind.code())),
            ("ReferenceItem", ComValue::Str(item.to_string())),
            ("InsertAsHyperlink", ComValue::Bool(req.insert_as_hyperlink)),
            ("IncludePosition", ComValue::Bool(false)),
            ("SeparateNumbers", ComValue::Bool(false)),
            ("SeparatorString", ComValue::Str(" ".to_string())),
        ],
    )
    .map_err(|e| {
        WordError::hint(
            e,
            format!(
                "确认目标项存在且名称匹配（{ty:?} 的 {item}）；该引擎可能不支持 InsertCrossReference"
            ),
        )
    })?;
    let got = at.text().unwrap_or_default();
    Ok(format!("{ty:?}/{kind:?} -> {}", anchors::quote(&got, 40)))
}

/// 列示现有交叉引用（`Fields` 中 `wdFieldRef`）。
fn list_cross_references(doc: &Document, max_items: usize) -> Result<Vec<CrossRefInfo>> {
    let content = doc.content()?;
    let Ok(fields) = content.raw().get_obj("Fields", "Fields") else {
        return Ok(Vec::new());
    };
    let count = fields.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut out = Vec::new();
    for i in 1..=count.min(max_items as i32) {
        let f = fields
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child(&format!("Field[{i}]"))?;
        let ty = f.get("Type").and_then(|v| v.int()).unwrap_or(-1);
        let code = f
            .get_obj("Range", "Field.Range")
            .and_then(|r| r.get("Text"))
            .and_then(|v| v.string())
            .unwrap_or_default();
        // 主判据是字段类型；个别引擎的 REF 域类型码不同，退化为按字段代码文本判断。
        let is_ref = ty == FIELD_REF || code.to_uppercase().contains("REF ");
        if !is_ref {
            continue;
        }
        let result_text = f
            .get_obj("Result", "Field.Result")
            .and_then(|r| r.get("Text"))
            .and_then(|v| v.string())
            .unwrap_or_default();
        // 字段代码形如 `REF 书签名 \h`：取第二段作为目标项。
        let item = code
            .split_whitespace()
            .nth(1)
            .map(str::to_string)
            .filter(|s| !s.is_empty());
        let range = Range::new(f.get_obj("Range", "Field.Range")?);
        out.push(CrossRefInfo {
            reference_type: "ref".into(),
            reference_item: item,
            quote: anchors::quote(&result_text, 60),
            start: range.start().ok(),
            end: range.end().ok(),
        });
    }
    Ok(out)
}

/// 采集文档结构快照（只读）。
pub fn snapshot(doc: &Document, doc_version: u64, max_items: usize) -> Result<StructureSnapshot> {
    let paras = doc.raw().get_obj("Paragraphs", "Paragraphs")?;
    let count = paras.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut paragraphs = Vec::new();
    for i in 1..=count.min(max_items as i32) {
        let p = paras
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child(&format!("Paragraph[{i}]"))?;
        let range = Range::new(p.get_obj("Range", "Paragraph.Range")?);
        let text = anchors::display_text(&range.text().unwrap_or_default());
        paragraphs.push(StructureParagraph {
            index: i as u32,
            style: style_name(&p),
            outline_level: p.get("OutlineLevel").and_then(|v| v.int()).unwrap_or(10),
            chars: text.chars().count() as u32,
            quote: anchors::quote(&text, 40),
            text_hash: anchors::hash_text(&text),
        });
    }

    let tables = collect_tables(doc, max_items)?;
    Ok(StructureSnapshot {
        document_id: doc.id().to_string(),
        doc_version,
        name: doc.name().unwrap_or_default(),
        section_count: count_of(doc, "Sections"),
        paragraphs,
        tables,
        note_count: count_of(doc, "Footnotes") + count_of(doc, "Endnotes"),
        media_count: count_of(doc, "Shapes") + count_of(doc, "InlineShapes"),
        toc_count: count_of(doc, "TablesOfContents") as i32,
        // wdStatisticPages = 2；大文档算一次要几十毫秒，取不到就是 0（界面不显示即可）
        page_count: doc
            .raw()
            .call("ComputeStatistics", vec![ComValue::Int(2)])
            .and_then(|v| v.int())
            .unwrap_or(0),
    })
}

/// 段落样式名（`Style` 可能是对象或字符串）。
fn style_name(p: &IDispatchWrapper) -> String {
    if let Ok(v) = p.get("Style") {
        if let Ok(name) = v.string() {
            if !name.is_empty() && !name.contains("<未支持") {
                return name;
            }
        }
    }
    p.get_obj("Style", "Style")
        .and_then(|s| s.get("NameLocal"))
        .and_then(|v| v.string())
        .unwrap_or_default()
}

fn collect_tables(doc: &Document, max_items: usize) -> Result<Vec<TableShapeSnapshot>> {
    let tables = doc.tables()?;
    let count = tables.get("Count").and_then(|v| v.int()).unwrap_or(0);
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
            .map(|s| anchors::quote(&anchors::display_text(&s), 40))
            .unwrap_or_default();
        out.push(TableShapeSnapshot {
            index: i as u32,
            rows: rows.max(0) as u32,
            columns: cols.max(0) as u32,
            preview,
        });
    }
    Ok(out)
}

/// 比较两份结构快照（纯函数，可离线单测）。
pub fn diff_snapshots(a: &StructureSnapshot, b: &StructureSnapshot) -> Vec<StructureChange> {
    let mut out = Vec::new();
    let common = a.paragraphs.len().min(b.paragraphs.len());
    for i in 0..common {
        let (pa, pb) = (&a.paragraphs[i], &b.paragraphs[i]);
        if pa.style != pb.style {
            out.push(StructureChange {
                kind: "paragraph_style_changed".into(),
                index: pb.index,
                detail: format!("{} → {}", pa.style, pb.style),
            });
        }
        if pa.text_hash != pb.text_hash {
            out.push(StructureChange {
                kind: "paragraph_text_changed".into(),
                index: pb.index,
                detail: format!(
                    "{} 字 → {} 字：{} → {}",
                    pa.chars, pb.chars, pa.quote, pb.quote
                ),
            });
        }
    }
    for p in a.paragraphs.iter().skip(b.paragraphs.len()) {
        out.push(StructureChange {
            kind: "paragraph_removed".into(),
            index: p.index,
            detail: p.quote.clone(),
        });
    }
    for p in b.paragraphs.iter().skip(a.paragraphs.len()) {
        out.push(StructureChange {
            kind: "paragraph_inserted".into(),
            index: p.index,
            detail: p.quote.clone(),
        });
    }

    let tcommon = a.tables.len().min(b.tables.len());
    for i in 0..tcommon {
        let (ta, tb) = (&a.tables[i], &b.tables[i]);
        if ta.rows != tb.rows || ta.columns != tb.columns || ta.preview != tb.preview {
            out.push(StructureChange {
                kind: "table_shape_changed".into(),
                index: tb.index,
                detail: format!(
                    "{}x{}:{} → {}x{}:{}",
                    ta.rows, ta.columns, ta.preview, tb.rows, tb.columns, tb.preview
                ),
            });
        }
    }
    for t in a.tables.iter().skip(b.tables.len()) {
        out.push(StructureChange {
            kind: "table_removed".into(),
            index: t.index,
            detail: format!("{}x{}", t.rows, t.columns),
        });
    }
    for t in b.tables.iter().skip(a.tables.len()) {
        out.push(StructureChange {
            kind: "table_added".into(),
            index: t.index,
            detail: format!("{}x{}", t.rows, t.columns),
        });
    }

    if a.section_count != b.section_count {
        out.push(StructureChange {
            kind: "section_count_changed".into(),
            index: 0,
            detail: format!("{} → {}", a.section_count, b.section_count),
        });
    }
    if a.note_count != b.note_count {
        out.push(StructureChange {
            kind: "note_count_changed".into(),
            index: 0,
            detail: format!("{} → {}", a.note_count, b.note_count),
        });
    }
    if a.media_count != b.media_count {
        out.push(StructureChange {
            kind: "media_count_changed".into(),
            index: 0,
            detail: format!("{} → {}", a.media_count, b.media_count),
        });
    }
    out
}
