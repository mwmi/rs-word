//! 文本锚点：hash、短摘录、校验（stale_target）。

use anyhow::Result;
use sha2::{Digest, Sha256};

use crate::error::WordError;
use crate::range::Range;
use crate::types::{Story, TextAnchor};

/// 文本 sha256（`sha256:<hex>`）。
pub fn hash_text(s: &str) -> String {
    let mut h = Sha256::new();
    h.update(s.as_bytes());
    format!("sha256:{}", hex::encode(h.finalize()))
}

/// 展示层清理：去掉段落/单元格终止符 `\r`、`\x07`（不改底层位置）。
pub fn display_text(s: &str) -> String {
    s.trim_end_matches(['\r', '\x07']).to_string()
}

/// 短摘录（限制字符数，避免泄漏全文）。
pub fn quote(s: &str, max_chars: usize) -> String {
    let clean = display_text(s);
    if clean.chars().count() <= max_chars {
        clean
    } else {
        let mut q: String = clean.chars().take(max_chars).collect();
        q.push('…');
        q
    }
}

/// 从 Range 生成锚点。
pub fn anchor(
    document_id: &str,
    story: Story,
    range: &Range,
    doc_version: u64,
    quote_max: usize,
) -> Result<TextAnchor> {
    let text = range.text()?;
    Ok(TextAnchor {
        document_id: document_id.to_string(),
        story,
        start: range.start()?,
        end: range.end()?,
        quote: quote(&text, quote_max),
        text_hash: hash_text(&text),
        doc_version,
    })
}

/// 校验目标未过期：`expected_text` / `text_hash` 不符返回 `stale_target`。
pub fn verify(range: &Range, expected_text: Option<&str>, text_hash: Option<&str>) -> Result<()> {
    let actual = range.text()?;
    if let Some(exp) = expected_text {
        if actual != exp {
            return Err(
                WordError::new("verify", range.label(), "目标文本在读取后已变化")
                    .with_hint("重新 word_read 获取该段最新锚点，再重新生成修改")
                    .into(),
            );
        }
    }
    if let Some(h) = text_hash {
        let got = hash_text(&actual);
        if got != h {
            return Err(WordError::new(
                "verify",
                range.label(),
                format!("文本 hash 不匹配（期望 {h}，实际 {got}）"),
            )
            .with_hint("目标已变化，重新读取后再修改")
            .into());
        }
    }
    Ok(())
}
