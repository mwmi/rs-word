//! AI 工作流：`prepare`（只读，产出带锚点的分块）与 `apply`（校验后批量写入）。
//!
//! 库本身不调用模型：`prepare` 只负责给模型准备好受限的输入包，`apply` 只执行
//! 模型给出的结构化 patch（`WorkflowChange`）。两次调用之间用 `doc_version` + `text_hash`
//! 保证锚点不过期；任何目标不符返回 `stale_target`，绝不猜新偏移。
//!
//! 分块语义：单次 `prepare` 返回**一个** chunk（≤ `max_chars_per_chunk` 字符）并在
//! `next_cursor` 给出下一块起点（= 当前终点 - `overlap_chars`），多次调用即可覆盖长文档。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::{SideEffectStatus, WordError};
use crate::target::resolve;
use crate::types::{
    ActionResult, Story, Target, TextAnchor, WorkflowApplyMode, WorkflowApplyRequest,
    WorkflowApplyResult, WorkflowChange, WorkflowChangeOp, WorkflowChunk, WorkflowPrepareRequest,
    WorkflowPrepareResult, WorkflowTask,
};

/// `max_changes` 硬上限（防止模型一次塞入超大批量）。
const HARD_MAX_CHANGES: u32 = 500;

/// prepare：只读产出锚点分块。
pub fn prepare(
    app: &App,
    doc: &Document,
    req: &WorkflowPrepareRequest,
    doc_version: u64,
) -> Result<WorkflowPrepareResult> {
    let range = resolve(app, doc, &req.target)?.range;
    let text = anchors::display_text(&range.text()?);
    let chars: Vec<char> = text.chars().collect();
    let total = chars.len();
    let base = range.start()?;

    // 游标 `t:<char_offset>`（与 word_read 同源编码）。
    let offset = req
        .cursor
        .as_deref()
        .and_then(|c| c.split_once(':'))
        .and_then(|(_, n)| n.parse::<usize>().ok())
        .unwrap_or(0)
        .min(total);

    let chunk_size = req.max_chars_per_chunk.max(1);
    let overlap = req.overlap_chars.min(chunk_size.saturating_sub(1));
    let end = (offset + chunk_size).min(total);
    let chunk_text: String = chars[offset..end].iter().collect();
    let step = end.saturating_sub(offset).saturating_sub(overlap);
    let next_offset = if end < total {
        Some(offset + step.max(1))
    } else {
        None
    };

    let mut notes = Vec::new();
    if !req.task.default_writes_document() {
        notes.push(format!(
            "{:?} 默认只返回聊天结果，不写入源文档；确需写回请改 task 或显式调用 apply",
            req.task
        ));
    }
    if req.task == WorkflowTask::Format {
        notes.push("format 任务：先审核 outline/样式清单，再用 word_format 批量套样式".into());
    }
    if req.include_tables {
        notes.push(format!("文档表格数 {}", doc.table_count().unwrap_or(0)));
    }
    if let Some(instruction) = &req.instruction {
        notes.push(format!("instruction: {}", anchors::quote(instruction, 80)));
    }
    if let Some(lang) = &req.language {
        notes.push(format!("language: {lang}"));
    }

    let chunk = WorkflowChunk {
        chunk_id: format!("c_{offset}"),
        start: base + offset as i32,
        end: base + end as i32,
        text_hash: anchors::hash_text(&chunk_text),
        quote: anchors::quote(&chunk_text, 80),
        truncated: end < total,
        text: chunk_text,
    };
    let anchor = TextAnchor {
        document_id: doc.id().to_string(),
        story: Story::Main,
        start: chunk.start,
        end: chunk.end,
        quote: chunk.quote.clone(),
        text_hash: chunk.text_hash.clone(),
        doc_version,
    };

    let outline = if req.include_outline {
        crate::read::collect_outline(doc, req.max_items)?
    } else {
        Vec::new()
    };

    Ok(WorkflowPrepareResult {
        document_id: doc.id().to_string(),
        task: req.task,
        doc_version,
        outline,
        chunks: vec![chunk],
        anchors: vec![anchor],
        next_cursor: next_offset.map(|o| format!("t:{o}")),
        coverage: vec![Story::Main],
        notes,
    })
}

/// apply：预检 → 逐条应用（低置信降级为批注）→ 回读验证。
pub fn apply(
    app: &App,
    doc: &Document,
    req: &WorkflowApplyRequest,
    doc_version: u64,
) -> Result<WorkflowApplyResult> {
    if req.expected_doc_version != doc_version {
        return Err(WordError::new(
            "workflow.apply",
            doc.label(),
            format!(
                "文档版本已变化（期望 {}，实际 {doc_version}）：重新 prepare",
                req.expected_doc_version
            ),
        )
        .with_hint("重新 word_workflow prepare 取新锚点与版本")
        .into());
    }
    if req.changes.len() as u32 > req.max_changes {
        return Err(WordError::new(
            "workflow.apply",
            doc.label(),
            format!(
                "变更数 {} 超过 max_changes {}：请分批",
                req.changes.len(),
                req.max_changes
            ),
        )
        .into());
    }
    if req.changes.len() as u32 > HARD_MAX_CHANGES {
        return Err(WordError::new(
            "workflow.apply",
            doc.label(),
            format!("变更数超过硬上限 {HARD_MAX_CHANGES}"),
        )
        .into());
    }
    if req.apply_mode == WorkflowApplyMode::Direct && req.confirm.as_deref() != Some("yes") {
        return Err(WordError::new(
            "workflow.apply",
            doc.label(),
            "apply_mode=direct 直接改原文（不产生修订/批注）为高风险操作",
        )
        .with_hint("确认后传 confirm=\"yes\"，或改用 track_changes/comment")
        .into());
    }

    if req.validate_first {
        for (i, ch) in req.changes.iter().enumerate() {
            if let Err(e) = resolve(app, doc, &ch.target) {
                let terr = crate::error::ToolError::from_anyhow(&e);
                return Err(WordError::new(
                    "workflow.apply",
                    doc.label(),
                    format!("第 {i} 条变更预检失败（{}）：{}", terr.code, terr.message),
                )
                .with_hint("全部变更未应用；重新 prepare 后重发")
                .into());
            }
        }
    }

    let mut results = Vec::with_capacity(req.changes.len());
    let mut warnings = Vec::new();
    for (i, ch) in req.changes.iter().enumerate() {
        let op = classify(ch, req);
        match apply_one(app, doc, ch, op, req.apply_mode) {
            Ok(summary) => {
                if op == EffectiveOp::Comment && ch.op != WorkflowChangeOp::Comment {
                    warnings.push(format!(
                        "第 {i} 条置信度 {:?} 低于阈值 {}：已降级为批注",
                        ch.confidence, req.comment_below_confidence
                    ));
                }
                results.push(ActionResult {
                    index: i,
                    op: op_name(op),
                    ok: true,
                    summary,
                    target: None,
                    reversible: true,
                    error_code: None,
                });
            }
            Err(e) => {
                let terr = crate::error::ToolError::from_anyhow(&e);
                results.push(ActionResult {
                    index: i,
                    op: op_name(op),
                    ok: false,
                    summary: terr.message.clone(),
                    target: None,
                    reversible: false,
                    error_code: Some(terr.code),
                });
            }
        }
    }

    let applied = results.iter().filter(|r| r.ok).count();
    let failed = results.len() - applied;
    let skipped = warnings.len();
    let side_effect_status = if failed == 0 && applied == 0 {
        SideEffectStatus::NotApplied
    } else if failed == 0 {
        SideEffectStatus::Applied
    } else if applied == 0 {
        SideEffectStatus::NotApplied
    } else {
        SideEffectStatus::PartiallyApplied
    };
    Ok(WorkflowApplyResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version: doc_version + u64::from(applied > 0),
        applied,
        failed,
        skipped,
        side_effect_status,
        results,
        warnings,
    })
}

/// validate：只做预检（解析 target + 校验锚点/版本/数量），**不写入任何内容**。
pub fn validate(
    app: &App,
    doc: &Document,
    req: &WorkflowApplyRequest,
    doc_version: u64,
) -> Result<WorkflowApplyResult> {
    if req.expected_doc_version != doc_version {
        return Err(WordError::new(
            "workflow.validate",
            doc.label(),
            format!(
                "文档版本已变化（期望 {}，实际 {doc_version}）：重新 prepare",
                req.expected_doc_version
            ),
        )
        .into());
    }
    if req.changes.len() as u32 > req.max_changes || req.changes.len() as u32 > HARD_MAX_CHANGES {
        return Err(WordError::new(
            "workflow.validate",
            doc.label(),
            format!("变更数 {} 超过上限", req.changes.len()),
        )
        .into());
    }
    let mut results = Vec::with_capacity(req.changes.len());
    for (i, ch) in req.changes.iter().enumerate() {
        let op = classify(ch, req);
        match resolve(app, doc, &ch.target) {
            Ok(_) => results.push(ActionResult {
                index: i,
                op: op_name(op),
                ok: true,
                summary: "预检通过（未写入）".into(),
                target: None,
                reversible: true,
                error_code: None,
            }),
            Err(e) => {
                let terr = crate::error::ToolError::from_anyhow(&e);
                results.push(ActionResult {
                    index: i,
                    op: op_name(op),
                    ok: false,
                    summary: terr.message.clone(),
                    target: None,
                    reversible: false,
                    error_code: Some(terr.code),
                });
            }
        }
    }
    let failed = results.iter().filter(|r| !r.ok).count();
    Ok(WorkflowApplyResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        applied: 0,
        failed,
        skipped: 0,
        side_effect_status: SideEffectStatus::NotApplied,
        results,
        warnings: vec!["validate 仅预检：文档未被修改".into()],
    })
}

/// 实际生效的操作（批注降级后与声明的 `op` 不同）。
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum EffectiveOp {
    Replace,
    InsertAfter,
    Comment,
}

/// 纯决策：按 `apply_mode`/置信度决定实际生效的操作（可离线单测）。
pub fn classify_change(
    ch: &WorkflowChange,
    apply_mode: WorkflowApplyMode,
    comment_below: f32,
) -> EffectiveOp {
    if ch.op == WorkflowChangeOp::Comment || apply_mode == WorkflowApplyMode::Comment {
        return EffectiveOp::Comment;
    }
    if apply_mode == WorkflowApplyMode::TrackChanges {
        if let Some(c) = ch.confidence {
            if c < comment_below {
                return EffectiveOp::Comment;
            }
        }
    }
    match ch.op {
        WorkflowChangeOp::Replace => EffectiveOp::Replace,
        WorkflowChangeOp::InsertAfter => EffectiveOp::InsertAfter,
        WorkflowChangeOp::Comment => EffectiveOp::Comment,
    }
}

/// 供内部调用：从请求取参数。
fn classify(ch: &WorkflowChange, req: &WorkflowApplyRequest) -> EffectiveOp {
    classify_change(ch, req.apply_mode, req.comment_below_confidence)
}

fn op_name(op: EffectiveOp) -> String {
    match op {
        EffectiveOp::Replace => "replace",
        EffectiveOp::InsertAfter => "insert_after",
        EffectiveOp::Comment => "comment",
    }
    .to_string()
}

fn apply_one(
    app: &App,
    doc: &Document,
    ch: &WorkflowChange,
    op: EffectiveOp,
    mode: WorkflowApplyMode,
) -> Result<String> {
    let resolved = resolve(app, doc, &ch.target)?;
    let range = resolved.range;
    // 逐条应用前重新核对锚点（前面变更可能已移动位置）。
    if let Target::Range {
        expected_text,
        text_hash,
        ..
    } = &ch.target
    {
        anchors::verify(&range, expected_text.as_deref(), text_hash.as_deref())
            .map_err(|e| crate::error::WordError::hint(e, "锚点已过期：重新 prepare 后再提交"))?;
    }
    match op {
        EffectiveOp::Comment => {
            let text = comment_text(ch);
            doc.raw().get_obj("Comments", "Comments")?.call_named(
                "Add",
                vec![],
                &[
                    ("Range", ComValue::Obj(range.raw().disp.clone())),
                    ("Text", ComValue::Str(text.clone())),
                ],
            )?;
            Ok(format!("批注 {}", anchors::quote(&text, 40)))
        }
        EffectiveOp::Replace => {
            let text = ch.text.as_ref().ok_or_else(|| {
                WordError::new("workflow.apply", doc.label(), "replace 需要 text")
            })?;
            let run = || -> Result<()> {
                range.set_text(text)?;
                let got = range.text().unwrap_or_default();
                if !got.contains(text.as_str()) {
                    return Err(WordError::new(
                        "workflow.apply",
                        range.label(),
                        "写入后回读不匹配",
                    )
                    .into());
                }
                Ok(())
            };
            if mode == WorkflowApplyMode::TrackChanges {
                crate::review::with_tracking(doc, true, run)?;
            } else {
                run()?;
            }
            Ok(format!("替换为 {} 字", text.chars().count()))
        }
        EffectiveOp::InsertAfter => {
            let text = ch.text.as_ref().ok_or_else(|| {
                WordError::new("workflow.apply", doc.label(), "insert_after 需要 text")
            })?;
            let run = || -> Result<()> {
                range.collapse_end()?.insert_after(text)?;
                Ok(())
            };
            if mode == WorkflowApplyMode::TrackChanges {
                crate::review::with_tracking(doc, true, run)?;
            } else {
                run()?;
            }
            Ok(format!("末尾插入 {} 字", text.chars().count()))
        }
    }
}

/// 批注文本：`reason` 在前、`text` 在后（模型可只给其一）。
fn comment_text(ch: &WorkflowChange) -> String {
    match (&ch.reason, &ch.text) {
        (Some(r), Some(t)) => format!("{r} {t}"),
        (Some(r), None) => r.clone(),
        (None, Some(t)) => t.clone(),
        (None, None) => "（模型未给出说明）".to_string(),
    }
}
