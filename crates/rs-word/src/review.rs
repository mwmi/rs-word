//! 审阅：Track Changes、Revisions（接受/拒绝）、Comments（批注）。
//!
//! 修订由 `TrackRevisions=true` 时的常规编辑产生；不要伪造 Revision 对象。
//! 单条接受/拒绝前重新枚举并校验 hash，避免旧序号过期。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::{SideEffectStatus, WordError};
use crate::range::Range;
use crate::target::resolve;
use crate::types::{
    CommentInfo, ReviewAction, ReviewRequest, ReviewResult, RevisionFilter, RevisionInfo,
};

/// 修订类型码 → 可读名。
fn revision_type_name(code: i32) -> String {
    match code {
        0 => "insert".into(),
        1 => "delete".into(),
        2 => "property".into(),
        3 => "paragraph_property".into(),
        5 => "format".into(),
        10 => "style".into(),
        11 => "comment".into(),
        12 => "field".into(),
        _ => format!("type_{code}"),
    }
}

/// 读取单条修订信息。
pub fn revision_info(rev: &IDispatchWrapper, index: u32, doc_version: u64) -> Result<RevisionInfo> {
    let kind = rev.get("Type").and_then(|v| v.int()).unwrap_or(-1);
    let author = rev.get("Author").and_then(|v| v.string()).ok();
    let date = rev.get("Date").and_then(|v| v.string()).ok();
    let range = Range::new(rev.get_obj("Range", "Revision.Range")?);
    let text = range.text().unwrap_or_default();
    let hash = anchors::hash_text(&text);
    let short = &hash[7..15.min(hash.len())];
    Ok(RevisionInfo {
        revision_ref: format!("rev_{index}_{short}"),
        index,
        kind: revision_type_name(kind),
        author,
        date,
        quote: anchors::quote(&text, 60),
        start: range.start().ok(),
        text_hash: hash,
        doc_version,
    })
}

/// 读取单条批注信息。
pub fn comment_info(comment: &IDispatchWrapper, index: u32) -> CommentInfo {
    let author = comment.get("Author").and_then(|v| v.string()).ok();
    let text = comment
        .get_obj("Range", "Comment.Range")
        .and_then(|r| r.get("Text"))
        .and_then(|v| v.string())
        .unwrap_or_default();
    CommentInfo {
        comment_id: format!("cm_{index}"),
        author,
        text: anchors::display_text(&text),
        quote: None,
    }
}

/// 审阅入口。
pub fn exec(
    app: &App,
    doc: &Document,
    req: &ReviewRequest,
    doc_version: u64,
) -> Result<ReviewResult> {
    let mut result = ReviewResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        tracking: doc.track_revisions().unwrap_or(false),
        revisions: Vec::new(),
        comments: Vec::new(),
        applied: Vec::new(),
        comment: None,
        skipped: 0,
    };
    match req.action {
        ReviewAction::Status => {}
        ReviewAction::SetTracking => {
            let on = req.enabled.unwrap_or(true);
            doc.set_track_revisions(on)?;
            result.tracking = on;
            result.applied.push(format!("track_revisions={on}"));
        }
        ReviewAction::ListRevisions => {
            result.revisions = list_revisions(doc, req.limit, doc_version)?;
        }
        ReviewAction::AcceptRevision | ReviewAction::RejectRevision => {
            let r = req.revision_ref.as_deref().ok_or_else(|| {
                anyhow::anyhow!("accept/reject 需要 revision_ref（来自 list_revisions）")
            })?;
            let (idx, hash) = parse_rev_ref(r)?;
            let revisions = doc.raw().get_obj("Revisions", "Revisions")?;
            let count = revisions.get("Count")?.int().unwrap_or(0);
            if idx < 1 || idx > count {
                return Err(WordError::new(
                    "review",
                    doc.label(),
                    format!("revision_ref {r} 已过期"),
                )
                .into());
            }
            let rev = revisions
                .call("Item", vec![ComValue::Int(idx)])?
                .idispatch_child("Revision")?;
            let current = revision_info(&rev, idx as u32, doc_version)?;
            if !current.text_hash.starts_with(&hash) {
                return Err(WordError::new(
                    "review",
                    doc.label(),
                    "revision_ref 与当前修订不符（已过期）",
                )
                .with_hint("重新 list_revisions 获取新引用")
                .into());
            }
            let method = if matches!(req.action, ReviewAction::AcceptRevision) {
                "Accept"
            } else {
                "Reject"
            };
            rev.call(method, vec![])?;
            result.applied.push(format!("{method} revision #{idx}"));
        }
        ReviewAction::AcceptAll | ReviewAction::RejectAll => {
            if req.confirm.as_deref() != Some("yes") {
                return Err(
                    WordError::new("review", doc.label(), "批量接受/拒绝修订为高风险操作")
                        .with_hint("确认后传 confirm=\"yes\"")
                        .into(),
                );
            }
            let method = if matches!(req.action, ReviewAction::AcceptAll) {
                "AcceptAllRevisions"
            } else {
                "RejectAllRevisions"
            };
            doc.raw().call(method, vec![])?;
            result.applied.push(method.to_string());
        }
        ReviewAction::AcceptFiltered | ReviewAction::RejectFiltered => {
            if req.confirm.as_deref() != Some("yes") {
                return Err(WordError::new(
                    "review",
                    doc.label(),
                    "按条件批量接受/拒绝修订为高风险操作",
                )
                .with_hint("确认后传 confirm=\"yes\"（建议先 checkpoint）")
                .into());
            }
            let accept = matches!(req.action, ReviewAction::AcceptFiltered);
            let (applied, skipped) = batch_filtered(
                doc,
                req.filter.as_ref().unwrap_or(&RevisionFilter::default()),
                accept,
                req.limit,
                doc_version,
            )?;
            result.applied = applied;
            result.skipped = skipped;
        }
        ReviewAction::AddComment => {
            let target = req
                .target
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("add_comment 需要 target"))?;
            let text = req
                .text
                .as_deref()
                .ok_or_else(|| anyhow::anyhow!("add_comment 需要 text"))?;
            let r = resolve(app, doc, target)?.range;
            let c = doc.raw().get_obj("Comments", "Comments")?.call_named(
                "Add",
                vec![],
                &[
                    ("Range", ComValue::Obj(r.raw().disp.clone())),
                    ("Text", ComValue::Str(text.to_string())),
                ],
            )?;
            let c = c.idispatch_child("Comment")?;
            let idx = doc.comment_count().unwrap_or(1);
            let info = comment_info(&c, idx);
            result
                .applied
                .push(format!("add_comment on {}", anchors::quote(text, 30)));
            result.comment = Some(info);
        }
        ReviewAction::ListComments => {
            result.comments = list_comments(doc, req.limit)?;
        }
        ReviewAction::DeleteComment => {
            let id = req
                .comment_id
                .as_deref()
                .ok_or_else(|| anyhow::anyhow!("delete_comment 需要 comment_id"))?;
            let idx: i32 = id
                .strip_prefix("cm_")
                .and_then(|s| s.parse().ok())
                .ok_or_else(|| anyhow::anyhow!("comment_id '{id}' 非法"))?;
            doc.raw()
                .get_obj("Comments", "Comments")?
                .call("Item", vec![ComValue::Int(idx)])
                .map_err(|e| WordError::hint(e, "批注不存在或已被删除"))?
                .idispatch_child("Comment")?
                .call("Delete", vec![])?;
            result.applied.push(format!("delete comment {id}"));
        }
    }
    Ok(result)
}

/// 修订是否匹配筛选条件（纯函数：`authors`/`kinds` 为空表示不限）。
pub fn matches_filter(info: &RevisionInfo, f: &RevisionFilter) -> bool {
    if let Some(authors) = &f.authors {
        let author = info.author.clone().unwrap_or_default();
        if !authors.iter().any(|a| a == &author) {
            return false;
        }
    }
    if let Some(kinds) = &f.kinds {
        if !kinds.iter().any(|k| k == &info.kind) {
            return false;
        }
    }
    true
}

/// 按条件批量接受/拒绝：**倒序**处理（接受/拒绝会移除条目，倒序保证剩余索引不移位），
/// 每条重新读取并核对类型/作者；返回 (逐条结果, 跳过数)。
fn batch_filtered(
    doc: &Document,
    filter: &RevisionFilter,
    accept: bool,
    limit: u32,
    doc_version: u64,
) -> Result<(Vec<String>, u32)> {
    let cap = filter.max.unwrap_or(limit);
    let method = if accept { "Accept" } else { "Reject" };
    let mut applied = Vec::new();
    let mut nonmatching = 0u32;
    let mut failed = 0u32;
    let mut idx = doc
        .raw()
        .get_obj("Revisions", "Revisions")?
        .get("Count")?
        .int()
        .unwrap_or(0);
    while idx >= 1 && applied.len() < cap as usize {
        let rev = doc
            .raw()
            .get_obj("Revisions", "Revisions")?
            .call("Item", vec![ComValue::Int(idx)])?
            .idispatch_child(&format!("Revision[{idx}]"))?;
        let info = revision_info(&rev, idx as u32, doc_version)?;
        if !matches_filter(&info, filter) {
            nonmatching += 1;
            idx -= 1;
            continue;
        }
        match rev.call(method, vec![]) {
            Ok(_) => applied.push(format!(
                "{method} #{idx} [{}] {}",
                info.kind,
                anchors::quote(&info.quote, 30)
            )),
            Err(e) => {
                let terr = crate::error::ToolError::from_anyhow(&e);
                failed += 1;
                applied.push(format!("{method} #{idx} 失败: {}", terr.message));
            }
        }
        idx -= 1;
    }
    Ok((applied, nonmatching + failed))
}

fn parse_rev_ref(r: &str) -> Result<(i32, String)> {
    // rev_<index>_<hash8>
    let mut parts = r.splitn(3, '_');
    let _ = parts.next();
    let idx: i32 = parts
        .next()
        .and_then(|s| s.parse().ok())
        .ok_or_else(|| anyhow::anyhow!("revision_ref '{r}' 非法"))?;
    let hash = parts.next().unwrap_or("").to_string();
    Ok((idx, hash))
}

fn list_revisions(doc: &Document, limit: u32, doc_version: u64) -> Result<Vec<RevisionInfo>> {
    let revisions = doc.raw().get_obj("Revisions", "Revisions")?;
    let count = revisions.get("Count")?.int().unwrap_or(0);
    let mut out = Vec::new();
    for i in 1..=count.min(limit as i32) {
        let rev = revisions
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child("Revision")?;
        out.push(revision_info(&rev, i as u32, doc_version)?);
    }
    Ok(out)
}

fn list_comments(doc: &Document, limit: u32) -> Result<Vec<CommentInfo>> {
    let comments = doc.raw().get_obj("Comments", "Comments")?;
    let count = comments.get("Count")?.int().unwrap_or(0);
    let mut out = Vec::new();
    for i in 1..=count.min(limit as i32) {
        let c = comments
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child("Comment")?;
        out.push(comment_info(&c, i as u32));
    }
    Ok(out)
}

/// 用 Track Changes 策略执行一次修改：保存旧状态→开→改→回读修订数→恢复。
pub fn with_tracking<T>(
    doc: &Document,
    enabled: bool,
    f: impl FnOnce() -> Result<T>,
) -> Result<(T, SideEffectStatus)> {
    let prev = doc.track_revisions().unwrap_or(false);
    let changed = enabled && !prev;
    if changed {
        doc.set_track_revisions(true)?;
    }
    let out = f();
    if changed {
        let _ = doc.set_track_revisions(prev);
    }
    let t = out?;
    let status = if doc.revision_count().unwrap_or(0) > 0 {
        SideEffectStatus::Applied
    } else {
        SideEffectStatus::Unknown
    };
    Ok((t, status))
}
