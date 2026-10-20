//! 样式：内置 StyleId → 常量映射、套用样式、`StyleProfile` 批量排版。

use anyhow::Result;
use serde::{Deserialize, Serialize};

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::format::font::apply_style_font;
use crate::format::paragraph::apply_paragraph_props;
use crate::range::Range;
use crate::types::{FontPatch, ParagraphPatch};

/// 语义 StyleId → Word 内置样式常量（集中管理，AI 永不传整数）。
pub fn builtin_style_constant(id: &str) -> Option<i32> {
    Some(match id {
        "normal" => -1,
        "heading_1" => -2,
        "heading_2" => -3,
        "heading_3" => -4,
        "heading_4" => -5,
        "heading_5" => -6,
        "heading_6" => -7,
        "heading_7" => -8,
        "heading_8" => -9,
        "heading_9" => -10,
        "caption" => -35,
        "title" => -63,
        "body_text" => -67,
        "subtitle" => -75,
        "list_paragraph" => -180,
        "quote" => -181,
        "intense_quote" => -182,
        _ => return None,
    })
}

/// `Document.Styles.Item(...)`；语义 ID 走内置常量，否则按显示名匹配。
fn style_object(doc: &Document, style: &str) -> Result<IDispatchWrapper> {
    let styles = doc.raw().get_obj("Styles", "Styles")?;
    let key = match builtin_style_constant(style) {
        Some(c) => ComValue::Int(c),
        None => ComValue::Str(style.to_string()),
    };
    styles
        .call("Item", vec![key])?
        .idispatch_child(&format!("Style[{style}]"))
}

/// 给 Range 套用样式（优先内置常量，避免本地化名称差异）。
pub fn apply_style(range: &Range, doc: &Document, style: &str) -> Result<()> {
    let obj = style_object(doc, style)?;
    range
        .raw()
        .put("Style", vec![ComValue::Obj(obj.disp.clone())])
}

/// 单条样式规则。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct StyleRule {
    pub style: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<FontPatch>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph: Option<ParagraphPatch>,
}

/// 文档排版配置（一次提交整套规则，比逐段硬编码更快）。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct StyleProfile {
    #[serde(default)]
    pub rules: Vec<StyleRule>,
}

/// 应用 StyleProfile（改样式定义，不逐段硬编码）。`dry_run=true` 只返回计划不修改。
pub fn apply_profile(doc: &Document, profile: &StyleProfile, dry_run: bool) -> Result<Vec<String>> {
    let mut applied = Vec::new();
    for rule in &profile.rules {
        if dry_run {
            applied.push(format!(
                "plan: style {} (font={} paragraph={})",
                rule.style,
                rule.font.is_some(),
                rule.paragraph.is_some()
            ));
            continue;
        }
        let s = style_object(doc, &rule.style)?;
        if let Some(f) = &rule.font {
            applied.extend(apply_style_font(&s.get_obj("Font", "Font")?, f)?);
        }
        if let Some(p) = &rule.paragraph {
            let fmt = s.get_obj("ParagraphFormat", "ParagraphFormat")?;
            let (a, w) = apply_paragraph_props(&fmt, p)?;
            applied.extend(a);
            applied.extend(w.iter().map(|x| format!("warning: {x}")));
        }
    }
    Ok(applied)
}
