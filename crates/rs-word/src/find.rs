//! Find/Replace：基于 `Range.Find`（不改动 UI 选区），支持预览与限量替换。

use anyhow::Result;

use crate::anchors;
use crate::document::Document;
use crate::error::WordError;
use crate::range::{Range, REPLACE_ONE};
use crate::types::{
    FindOptions, FindRequest, FindResult, MatchInfo, ReplaceApplyRequest, ReplacePreview,
    ReplacePreviewRequest, ReplaceResult, SideEffectStatus,
};

/// 预览/查找的硬上限，防止长文档失控。
const HARD_CAP: u32 = 5000;

/// 遍历匹配：返回前 `limit` 项与命中总数。
fn scan(
    doc: &Document,
    query: &str,
    opts: &FindOptions,
    limit: u32,
) -> Result<(Vec<MatchInfo>, u32)> {
    let content = doc.content()?;
    let content_end = content.end()?;
    let mut pos = content.start()?;
    let mut matches = Vec::new();
    let mut total = 0u32;
    while total < HARD_CAP {
        let search = doc.range(pos, content_end)?;
        let f = search.find()?;
        f.configure(query, opts)?;
        if !f.execute()? {
            break;
        }
        let s = search.start()?;
        let e = search.end()?;
        if e <= pos || e <= s {
            break; // 防死循环
        }
        total += 1;
        if (matches.len() as u32) < limit {
            let text = search.text()?;
            matches.push(MatchInfo {
                occurrence: total,
                start: s,
                end: e,
                quote: anchors::quote(&text, 60),
                text_hash: anchors::hash_text(&text),
                paragraph_index: None,
            });
        }
        pos = e;
    }
    Ok((matches, total))
}

/// 查找。
pub fn find(doc: &Document, req: &FindRequest, doc_version: u64) -> Result<FindResult> {
    if req.query.is_empty() {
        return Err(WordError::new("find", doc.label(), "查询文本为空").into());
    }
    let (matches, total) = scan(doc, &req.query, &req.options, req.limit.max(1))?;
    Ok(FindResult {
        document_id: doc.id().to_string(),
        doc_version,
        truncated: total > matches.len() as u32,
        matches,
        total,
    })
}

/// 解析第 `occurrence` 个匹配为 Range（供 Target::Find 定位）。
pub fn resolve_occurrence(
    doc: &Document,
    query: &str,
    occurrence: u32,
    opts: &FindOptions,
) -> Result<Range> {
    let content = doc.content()?;
    let content_end = content.end()?;
    let mut pos = content.start()?;
    let mut n = 0u32;
    while n < HARD_CAP {
        let search = doc.range(pos, content_end)?;
        let f = search.find()?;
        f.configure(query, opts)?;
        if !f.execute()? {
            break;
        }
        let s = search.start()?;
        let e = search.end()?;
        if e <= pos || e <= s {
            break;
        }
        n += 1;
        if n == occurrence.max(1) {
            return Ok(search);
        }
        pos = e;
    }
    Err(WordError::new(
        "find",
        doc.label(),
        format!("未找到第 {} 个匹配 '{query}'", occurrence.max(1)),
    )
    .with_hint("先 word_find_replace 查看实际命中数与位置")
    .into())
}

/// 替换预览（不改文档）。
pub fn preview(
    doc: &Document,
    req: &ReplacePreviewRequest,
    doc_version: u64,
) -> Result<ReplacePreview> {
    if req.find.is_empty() {
        return Err(WordError::new("replace.preview", doc.label(), "查找文本为空").into());
    }
    let (samples, total) = scan(doc, &req.find, &req.options, req.limit.max(1))?;
    let digest = anchors::hash_text(&format!("{}:{}:{}", req.find, req.replace, total));
    Ok(ReplacePreview {
        document_id: doc.id().to_string(),
        preview_id: format!("pv_{}", &digest[7..19]),
        doc_version,
        total,
        samples,
        max_replacements: Some(req.limit),
    })
}

/// 执行替换（限制最大替换数；默认从后往前由 Word 的 Replace=All 处理，这里用 One 循环以尊重上限）。
pub fn replace(doc: &Document, req: &ReplaceApplyRequest) -> Result<ReplaceResult> {
    if req.find.is_empty() {
        return Err(WordError::new("replace", doc.label(), "查找文本为空").into());
    }
    let content = doc.content()?;
    let content_end = content.end()?;
    let mut pos = content.start()?;
    let mut replaced = 0u32;
    while replaced < req.max_replacements && replaced < HARD_CAP {
        let search = doc.range(pos, content_end)?;
        let f = search.find()?;
        f.configure(&req.find, &req.options)?;
        f.set_replacement(&req.replace)?;
        if !f.execute_replace(REPLACE_ONE)? {
            break;
        }
        let s = search.start()?;
        let e = search.end()?;
        // 替换后命中范围长度变化，用替换文本长度估算新位置（保守：从 s 继续）。
        let new_end = s + req.replace.chars().count() as i32;
        if new_end <= s && e <= s {
            break;
        }
        pos = new_end.max(s + 1);
        replaced += 1;
    }
    Ok(ReplaceResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        replaced,
        doc_version: 0, // 由会话层填充
        side_effect_status: if replaced > 0 {
            SideEffectStatus::Applied
        } else {
            SideEffectStatus::NotApplied
        },
    })
}
