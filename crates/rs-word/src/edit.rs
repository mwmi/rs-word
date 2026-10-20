//! 编辑：insert/replace/delete/段落/分页分节符/书签/超链接/字段/套样式。
//! 先全量预检 target，再有序应用，最后 read-back 验证；部分成功返回 partially_applied。

use anyhow::Result;

use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::{SideEffectStatus, WordError};
use crate::format::styles;
use crate::target::resolve;
use crate::types::{ActionResult, EditAction, EditRequest, EditResult, InsertPosition, Target};

/// 编辑入口。
pub fn edit(app: &App, doc: &Document, req: &EditRequest, doc_version: u64) -> Result<EditResult> {
    // 全量预检：解析全部 target，任一失败则不执行任何动作。
    if req.validate_all_first {
        for a in &req.actions {
            if let Some(t) = action_target(a) {
                resolve(app, doc, t).map_err(|e| {
                    WordError::hint(e, "全量预检失败：未应用任何动作（validate_all_first=true）")
                })?;
            }
        }
    }

    let prev_track = doc.track_revisions().unwrap_or(false);
    if req.track_changes && !prev_track {
        doc.set_track_revisions(true)?;
    }
    let outcome = run_actions(app, doc, req, doc_version);
    // 无论成败都恢复原跟踪修订状态（restore 不会撤销已生成的修订）。
    if req.track_changes {
        let _ = doc.set_track_revisions(prev_track);
    }
    let actions = outcome?;

    let applied = actions.iter().filter(|a| a.ok).count();
    let failed = actions.len() - applied;
    let side_effect_status = if failed == 0 {
        SideEffectStatus::Applied
    } else if applied == 0 {
        SideEffectStatus::NotApplied
    } else {
        SideEffectStatus::PartiallyApplied
    };
    Ok(EditResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version: doc_version + u64::from(applied > 0),
        applied,
        failed,
        side_effect_status,
        actions,
    })
}

fn action_target(a: &EditAction) -> Option<&Target> {
    match a {
        EditAction::InsertText { target, .. }
        | EditAction::ReplaceText { target, .. }
        | EditAction::DeleteRange { target, .. }
        | EditAction::InsertParagraph { target, .. }
        | EditAction::DeleteParagraph { target }
        | EditAction::InsertPageBreak { target }
        | EditAction::InsertSectionBreak { target }
        | EditAction::SetBookmark { target, .. }
        | EditAction::InsertHyperlink { target, .. }
        | EditAction::InsertField { target, .. }
        | EditAction::ApplyStyle { target, .. } => Some(target),
        EditAction::DeleteBookmark { .. } => None,
    }
}

fn run_actions(
    app: &App,
    doc: &Document,
    req: &EditRequest,
    doc_version: u64,
) -> Result<Vec<ActionResult>> {
    let mut results = Vec::with_capacity(req.actions.len());
    for (i, action) in req.actions.iter().enumerate() {
        let op = op_name(action);
        match apply_one(app, doc, action, doc_version) {
            Ok((summary, target)) => results.push(ActionResult {
                index: i,
                op: op.into(),
                ok: true,
                summary,
                target,
                reversible: true,
                error_code: None,
            }),
            Err(e) => {
                let terr = crate::error::ToolError::from_anyhow(&e);
                results.push(ActionResult {
                    index: i,
                    op: op.into(),
                    ok: false,
                    summary: terr.message.clone(),
                    target: None,
                    reversible: false,
                    error_code: Some(terr.code),
                });
            }
        }
    }
    Ok(results)
}

fn op_name(a: &EditAction) -> &'static str {
    match a {
        EditAction::InsertText { .. } => "insert_text",
        EditAction::ReplaceText { .. } => "replace_text",
        EditAction::DeleteRange { .. } => "delete_range",
        EditAction::InsertParagraph { .. } => "insert_paragraph",
        EditAction::DeleteParagraph { .. } => "delete_paragraph",
        EditAction::InsertPageBreak { .. } => "insert_page_break",
        EditAction::InsertSectionBreak { .. } => "insert_section_break",
        EditAction::SetBookmark { .. } => "set_bookmark",
        EditAction::DeleteBookmark { .. } => "delete_bookmark",
        EditAction::InsertHyperlink { .. } => "insert_hyperlink",
        EditAction::InsertField { .. } => "insert_field",
        EditAction::ApplyStyle { .. } => "apply_style",
    }
}

/// 执行单个动作，返回 (摘要, 结果锚点)。
fn apply_one(
    app: &App,
    doc: &Document,
    action: &EditAction,
    _doc_version: u64,
) -> Result<(String, Option<crate::types::TextAnchor>)> {
    match action {
        EditAction::InsertText {
            target,
            text,
            position,
        } => {
            let r = resolve(app, doc, target)?.range;
            match position {
                InsertPosition::Start => r.collapse_start()?.insert_after(text)?,
                InsertPosition::End => r.collapse_end()?.insert_after(text)?,
                InsertPosition::Before => r.insert_before(text)?,
                InsertPosition::After => r.insert_after(text)?,
            }
            Ok((
                format!("在 {position:?} 插入 {} 字", text.chars().count()),
                None,
            ))
        }
        EditAction::ReplaceText { target, text } => {
            let r = resolve(app, doc, target)?.range;
            r.set_text(text)?;
            let actual = r.text()?;
            if !actual.contains(text.as_str()) {
                return Err(
                    WordError::new("edit.replace_text", r.label(), "写入后回读不匹配").into(),
                );
            }
            Ok((format!("替换为 {} 字", text.chars().count()), None))
        }
        EditAction::DeleteRange {
            target,
            include_paragraph_mark,
        } => {
            let r = resolve(app, doc, target)?.range;
            if *include_paragraph_mark {
                let para = r
                    .raw()
                    .get_obj("Paragraphs", "Paragraphs")?
                    .call("Item", vec![ComValue::Int(1)])?
                    .idispatch_child("Paragraph")?
                    .get_obj("Range", "Paragraph.Range")?;
                para.call("Delete", vec![])?;
            } else {
                r.delete()?;
            }
            Ok(("删除范围".into(), None))
        }
        EditAction::InsertParagraph {
            target,
            text,
            position,
            ..
        } => {
            let r = resolve(app, doc, target)?.range;
            let payload = format!("{text}\r");
            match position {
                InsertPosition::Before | InsertPosition::Start => r.insert_before(&payload)?,
                InsertPosition::After | InsertPosition::End => r.insert_after(&payload)?,
            }
            Ok((format!("插入段落（{} 字）", text.chars().count()), None))
        }
        EditAction::DeleteParagraph { target } => {
            let r = resolve(app, doc, target)?.range;
            r.raw()
                .get_obj("Paragraphs", "Paragraphs")?
                .call("Item", vec![ComValue::Int(1)])?
                .idispatch_child("Paragraph")?
                .get_obj("Range", "Paragraph.Range")?
                .call("Delete", vec![])?;
            Ok(("删除段落".into(), None))
        }
        EditAction::InsertPageBreak { target } => {
            let r = resolve(app, doc, target)?.range;
            r.collapse_start()?.insert_page_break()?;
            Ok(("插入分页符".into(), None))
        }
        EditAction::InsertSectionBreak { target } => {
            let r = resolve(app, doc, target)?.range;
            r.collapse_start()?.insert_section_break(2)?; // next page
            Ok(("插入分节符（下一页）".into(), None))
        }
        EditAction::SetBookmark { target, bookmark } => {
            let r = resolve(app, doc, target)?.range;
            let bms = doc.raw().get_obj("Bookmarks", "Bookmarks")?;
            if bms
                .call("Item", vec![ComValue::Str(bookmark.clone())])
                .is_ok()
            {
                bms.call("Item", vec![ComValue::Str(bookmark.clone())])?
                    .idispatch_child(bookmark)?
                    .call("Delete", vec![])?;
            }
            bms.call_named(
                "Add",
                vec![],
                &[
                    ("Name", ComValue::Str(bookmark.clone())),
                    ("Range", ComValue::Obj(r.raw().disp.clone())),
                ],
            )?;
            Ok((format!("设置书签 {bookmark}"), None))
        }
        EditAction::DeleteBookmark {
            document_id,
            bookmark,
        } => {
            if doc.id() != document_id {
                return Err(WordError::new(
                    "edit.delete_bookmark",
                    doc.label(),
                    "document_id 不匹配",
                )
                .into());
            }
            doc.raw()
                .get_obj("Bookmarks", "Bookmarks")?
                .call("Item", vec![ComValue::Str(bookmark.clone())])
                .map_err(|e| WordError::hint(e, format!("书签 {bookmark} 不存在")))?
                .idispatch_child(bookmark)?
                .call("Delete", vec![])?;
            Ok((format!("删除书签 {bookmark}"), None))
        }
        EditAction::InsertHyperlink {
            target,
            address,
            text,
        } => {
            let r = resolve(app, doc, target)?.range;
            let mut named = vec![
                ("Anchor", ComValue::Obj(r.raw().disp.clone())),
                ("Address", ComValue::Str(address.clone())),
            ];
            if let Some(t) = text {
                named.push(("TextToDisplay", ComValue::Str(t.clone())));
            }
            r.hyperlinks()?.call_named("Add", vec![], &named)?;
            Ok((format!("插入超链接 → {address}"), None))
        }
        EditAction::InsertField { target, field } => {
            let r = resolve(app, doc, target)?.range;
            r.fields()?.call_named(
                "Add",
                vec![],
                &[
                    ("Range", ComValue::Obj(r.raw().disp.clone())),
                    ("Type", ComValue::Int(field.code())),
                ],
            )?;
            Ok((format!("插入字段 {:?}", field), None))
        }
        EditAction::ApplyStyle { target, style } => {
            let r = resolve(app, doc, target)?.range;
            styles::apply_style(&r, doc, style)?;
            Ok((format!("套用样式 {style}"), None))
        }
    }
}
