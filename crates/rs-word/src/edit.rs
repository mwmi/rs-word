//! 编辑：insert/replace/delete/段落/分页分节符/书签/超链接/字段/套样式/生公式（全文 `$…$` 转换）。
//! 先全量预检 target，再有序应用，最后 read-back 验证；部分成功返回 partially_applied。

use anyhow::Result;

use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::{SideEffectStatus, WordError};
use crate::format::styles;
use crate::range::Range;
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
        |         EditAction::ApplyStyle { target, .. }
        | EditAction::InsertMarkdown { target, .. }
        | EditAction::ConvertMath { target } => Some(target),
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
        let mut notes = Vec::new();
        match apply_one(app, doc, action, doc_version, &mut notes) {
            Ok((summary, target)) => results.push(ActionResult {
                index: i,
                op: op.into(),
                ok: true,
                summary,
                target,
                warnings: notes,
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
                    warnings: notes,
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
        EditAction::InsertMarkdown { .. } => "insert_markdown",
        EditAction::ConvertMath { .. } => "convert_math",
    }
}

/// 文档以表格开头时，`Content.Start` 落在第一个单元格内，`InsertBefore` / `collapse_start`
/// 会把内容直接写进格子（实测 WPS：标题成了 Cell(1,1) 的文字）。`InsertParagraphBefore`
/// 是唯一能让表格整体后移的写法，之后 `[Start, Start]` 就是表格前的空段落。
///
/// 只处理 `kind=document` 且表格顶在正文最前的落点；其余情况返回 `None`，维持原路径。
fn slot_before_leading_table(doc: &Document, target: &Target) -> Result<Option<Range>> {
    if !matches!(target, Target::Document { .. }) {
        return Ok(None);
    }
    let content = doc.content()?;
    let at = content.start()?;
    let leading = doc
        .first_table_start()
        .map_or(false, |table_start| table_start <= at);
    if !leading || content.insert_paragraph_before().is_err() {
        return Ok(None);
    }
    // 自检：表格必须整体后移到落点之后；否则这次插入仍在格内，撤掉多出的段落退回原路径。
    let shifted = doc
        .first_table_start()
        .map_or(false, |table_start| table_start > at);
    if shifted {
        return Ok(Some(doc.range(at, at)?));
    }
    let _ = doc.range(at, at + 1).and_then(|r| r.delete());
    Ok(None)
}

/// 删除范围内第一段的整段（连同段落标记）。
fn delete_first_paragraph(r: &Range) -> Result<()> {
    r.raw()
        .get_obj("Paragraphs", "Paragraphs")?
        .call("Item", vec![ComValue::Int(1)])?
        .idispatch_child("Paragraph")?
        .get_obj("Range", "Paragraph.Range")?
        .call("Delete", vec![])?;
    Ok(())
}

/// 执行单个动作，返回 (摘要, 结果锚点)。
fn apply_one(
    app: &App,
    doc: &Document,
    action: &EditAction,
    _doc_version: u64,
    notes: &mut Vec<String>,
) -> Result<(String, Option<crate::types::TextAnchor>)> {
    match action {
        EditAction::InsertText {
            target,
            text,
            position,
        } => {
            let r = resolve(app, doc, target)?.range;
            // 文档以表格开头且插在最前：落点在第一格里，先写进表格前的空段落
            let slot = match position {
                InsertPosition::Start | InsertPosition::Before => {
                    slot_before_leading_table(doc, target)?
                }
                _ => None,
            };
            match slot {
                Some(slot) => slot.insert_after(text)?,
                None => match position {
                    InsertPosition::Start => r.collapse_start()?.insert_after(text)?,
                    InsertPosition::End => r.collapse_end()?.insert_after(text)?,
                    InsertPosition::Before => r.insert_before(text)?,
                    InsertPosition::After => r.insert_after(text)?,
                },
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
                delete_first_paragraph(&r)?;
            } else {
                r.delete()?;
            }
            Ok(("删除范围".into(), None))
        }
        EditAction::InsertParagraph {
            target,
            text,
            position,
            style,
        } => {
            let r = resolve(app, doc, target)?.range;
            let payload = format!("{text}\r");
            let width = payload.encode_utf16().count() as i32;
            // 记下新段落的字符位置：样式要在插入后单独套（Range 插入不会把新段纳进来）
            let (start, end) = match position {
                InsertPosition::Before | InsertPosition::Start => {
                    match slot_before_leading_table(doc, target)? {
                        // 表前空段落已就位：正文写到它的段落标记之前（标记已存在，不再补 \r）
                        Some(slot) => {
                            let at = slot.start()?;
                            slot.insert_after(text)?;
                            (at, at + width)
                        }
                        None => {
                            let at = r.start()?;
                            r.insert_before(&payload)?;
                            (at, at + width)
                        }
                    }
                }
                InsertPosition::After | InsertPosition::End => {
                    let at = r.end()?;
                    r.insert_after(&payload)?;
                    (at, at + width)
                }
            };
            let mut note = String::new();
            if let Some(name) = style.as_deref().map(str::trim).filter(|name| !name.is_empty()) {
                // 样式名由模型给，写错不该让整条插入作废（内容已经在文档里了）
                if doc.range(start, end)?.set_style(name).is_err() {
                    note = format!("，样式「{name}」没套上");
                }
            }
            Ok((format!("插入段落（{} 字）{note}", text.chars().count()), None))
        }
        EditAction::DeleteParagraph { target } => {
            delete_first_paragraph(&resolve(app, doc, target)?.range)?;
            Ok(("删除段落".into(), None))
        }
        EditAction::InsertPageBreak { target } => {
            let r = resolve(app, doc, target)?.range;
            match slot_before_leading_table(doc, target)? {
                Some(slot) => slot.insert_page_break()?,
                None => r.collapse_start()?.insert_page_break()?,
            }
            Ok(("插入分页符".into(), None))
        }
        EditAction::InsertSectionBreak { target } => {
            let r = resolve(app, doc, target)?.range;
            match slot_before_leading_table(doc, target)? {
                Some(slot) => slot.insert_section_break(2)?, // next page
                None => r.collapse_start()?.insert_section_break(2)?, // next page
            }
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
        EditAction::InsertMarkdown {
            target,
            markdown,
            position,
            style_map,
            math,
        } => {
            let written = crate::write::insert_markdown(
                app,
                doc,
                target,
                markdown,
                *position,
                style_map.as_ref(),
                math.unwrap_or(true),
            )?;
            notes.extend(written.warnings);
            let mut summary = format!(
                "写入 {} 段 / {} 表 / {} 公式",
                written.paragraphs, written.tables, written.equations
            );
            if written.toc {
                summary.push_str(" / 目录");
            }
            Ok((summary, None))
        }
        EditAction::ConvertMath { target } => {
            let range = resolve(app, doc, target)?.range;
            let start = range.start()?;
            let text = range.text()?;
            let converted = crate::math::convert_text(app, doc, start, &text);
            notes.extend(converted.notes);
            let mut summary = format!(
                "转换 {} 个独立公式 / {} 个行内公式",
                converted.display, converted.inline
            );
            if converted.failed > 0 {
                summary.push_str(&format!(" / {} 处未自动转换", converted.failed));
            }
            Ok((summary, None))
        }
    }
}
