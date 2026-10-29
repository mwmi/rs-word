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
    ViewSettings,
};

/// 文档窗口的 `View` 对象。
///
/// 用 `Document.ActiveWindow` 而不是 `Application.ActiveWindow`：前者**按文档**取窗口，
/// 目标文档不在前台也能读写它的显示方式（右栏要能管后台那份文档）。
fn view_of(doc: &Document) -> Result<IDispatchWrapper> {
    doc.raw()
        .get_obj("ActiveWindow", "Window")?
        .get_obj("View", "View")
}

/// 读当前显示方式：个别成员读不到就按「开着」兜底，别因为一个属性让整块空着。
fn read_view(doc: &Document) -> Result<ViewSettings> {
    let view = view_of(doc)?;
    let on = |name: &str| view.get(name).and_then(|v| v.bool()).unwrap_or(true);
    // WdRevisionsMode：0 = 批注框（右侧气球），1 = 嵌入正文
    let mode = match view.get("MarkupMode").and_then(|v| v.int()).unwrap_or(1) {
        0 => "balloon",
        _ => "inline",
    };
    Ok(ViewSettings {
        markup_mode: Some(mode.to_string()),
        show_markup: Some(on("ShowRevisionsAndComments")),
        show_comments: Some(on("ShowComments")),
        show_insertions: Some(on("ShowInsertionsAndDeletions")),
        show_format: Some(on("ShowFormatChanges")),
    })
}

/// 写显示方式：只改带上的字段。
fn apply_view(doc: &Document, want: &ViewSettings) -> Result<()> {
    let view = view_of(doc)?;
    if let Some(mode) = want.markup_mode.as_deref() {
        // 不认识的取值直接报错：静默当成嵌入，用户只会看到「点了没反应」
        let code = match mode {
            "inline" => 1,
            "balloon" => 0,
            other => anyhow::bail!("markup_mode 只认 inline / balloon，收到 '{other}'"),
        };
        view.put("MarkupMode", vec![code.into()])?;
    }
    for (name, value) in [
        ("ShowRevisionsAndComments", want.show_markup),
        ("ShowComments", want.show_comments),
        ("ShowInsertionsAndDeletions", want.show_insertions),
        ("ShowFormatChanges", want.show_format),
    ] {
        if let Some(on) = value {
            view.put(name, vec![on.into()])?;
        }
    }
    Ok(())
}

/// 修订类型码 → 可读名（`WdRevisionType`：**1=insert、2=delete**）。
///
/// 这张表曾在 1/2 上整体错位一位：插入被报成 delete、删除被报成 property，
/// 于是「修改」在回执里全成了删除（实报：校对后一键看全是删除）。
fn revision_type_name(code: i32) -> String {
    match code {
        0 => "none".into(),
        1 => "insert".into(),
        2 => "delete".into(),
        3 => "property".into(),
        4 => "paragraph_number".into(),
        5 => "display_field".into(),
        6 => "reconcile".into(),
        7 => "conflict".into(),
        8 => "style".into(),
        9 => "replace".into(),
        10 => "paragraph_property".into(),
        11 => "table_property".into(),
        12 => "section_property".into(),
        13 => "style_definition".into(),
        14 => "moved_from".into(),
        15 => "moved_to".into(),
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
    // hash 自带 "sha256:" 前缀：ref 里只放十六进制段的前 8 位。
    // 旧写法按下标盲切（&hash[7..15]）看着等价，实际把前缀长度写死了；
    // 而校验处拿这 8 位去 starts_with 带前缀的全量 hash，**永远不成立**（见下方 accept/reject）。
    let short: String = hash.strip_prefix("sha256:").unwrap_or(&hash).chars().take(8).collect();
    Ok(RevisionInfo {
        revision_ref: format!("rev_{index}_{short}"),
        index,
        kind: revision_type_name(kind),
        author,
        date,
        quote: anchors::quote(&text, 60),
        start: range.start().ok(),
        end: range.end().ok(),
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
        view: None,
        revisions: Vec::new(),
        comments: Vec::new(),
        applied: Vec::new(),
        comment: None,
        skipped: 0,
    };
    match req.action {
        // 显示方式顺带带上：右栏的修订弹层要它，这里不多花一趟 COM
        ReviewAction::Status => result.view = read_view(doc).ok(),
        ReviewAction::ViewStatus => result.view = Some(read_view(doc)?),
        ReviewAction::SetView => {
            apply_view(doc, req.view.as_ref().unwrap_or(&ViewSettings::default()))?;
            result.applied.push("set_view".into());
            result.view = Some(read_view(doc)?);
        }
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
            // 比的是 hash 的十六进制段：text_hash 形如 "sha256:<hex>"，直接 starts_with 短 hash 恒不成立
            let current_hex = current.text_hash.strip_prefix("sha256:").unwrap_or(&current.text_hash);
            if !current_hex.starts_with(&hash) {
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
    // 集合取一次：原先每轮都 get_obj 一遍，几百条修订就是几百次多余 COM
    let revisions = doc.raw().get_obj("Revisions", "Revisions")?;
    let mut idx = revisions.get("Count")?.int().unwrap_or(0);
    while idx >= 1 && applied.len() < cap as usize {
        let rev = revisions
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
