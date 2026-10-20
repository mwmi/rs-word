//! Target 解析：tagged enum → 固定 Range（含 story 与边界校验）。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
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
