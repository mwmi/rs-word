//! 字符格式：字体补丁（部分更新）、高亮、底纹。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::{hex_to_ole, WordError};
use crate::types::FontPatch;

/// 对任意 `Font` 对象应用字符属性（不含 Range 级高亮/底纹）。
/// 返回 (applied, warnings)：单项失败记 warning 不整体失败。
fn apply_font_props(
    font: &IDispatchWrapper,
    patch: &FontPatch,
) -> Result<(Vec<String>, Vec<String>)> {
    if patch.superscript == Some(true) && patch.subscript == Some(true) {
        return Err(
            WordError::new("format.font", &font.ctx, "上标与下标不能同时为 true")
                .with_hint("二者互斥，请只保留一个")
                .into(),
        );
    }
    if let Some(v) = patch.size_pt {
        if !(1.0..=200.0).contains(&v) {
            return Err(WordError::new(
                "format.font",
                &font.ctx,
                format!("字号 {v} 越界（1–200 pt）"),
            )
            .into());
        }
    }
    let mut applied = Vec::new();
    let mut warnings = Vec::new();
    let mut set = |name: &str, v: ComValue| match font.put(name, vec![v]) {
        Ok(()) => applied.push(format!("font.{name}")),
        Err(e) => warnings.push(format!("font.{name}: {e}")),
    };
    if let Some(v) = &patch.name {
        set("Name", ComValue::Str(v.clone()));
        set("NameAscii", ComValue::Str(v.clone()));
    }
    if let Some(v) = &patch.east_asia_name {
        set("NameFarEast", ComValue::Str(v.clone()));
    }
    if let Some(v) = patch.size_pt {
        set("Size", ComValue::Float(v as f64));
    }
    if let Some(v) = &patch.color {
        set("Color", ComValue::Int(hex_to_ole(v)? as i32));
    }
    if let Some(v) = patch.bold {
        set("Bold", ComValue::Bool(v));
    }
    if let Some(v) = patch.italic {
        set("Italic", ComValue::Bool(v));
    }
    if let Some(v) = patch.underline {
        set("Underline", ComValue::Int(v.code()));
    }
    if let Some(v) = patch.strike {
        set("StrikeThrough", ComValue::Bool(v));
    }
    if let Some(v) = patch.double_strike {
        set("DoubleStrikeThrough", ComValue::Bool(v));
    }
    if let Some(v) = patch.superscript {
        set("Superscript", ComValue::Bool(v));
    }
    if let Some(v) = patch.subscript {
        set("Subscript", ComValue::Bool(v));
    }
    if let Some(v) = patch.small_caps {
        set("SmallCaps", ComValue::Bool(v));
    }
    if let Some(v) = patch.all_caps {
        set("AllCaps", ComValue::Bool(v));
    }
    if let Some(v) = patch.hidden {
        set("Hidden", ComValue::Bool(v));
    }
    if let Some(v) = patch.char_spacing_pt {
        set("Spacing", ComValue::Float(v as f64));
    }
    Ok((applied, warnings))
}

/// `Range.Font` 入口；同时持有 Range 本体以设置 Range 级高亮/底纹。
pub struct Font {
    font: IDispatchWrapper,
    owner: IDispatchWrapper,
}

impl Font {
    pub(crate) fn new(font: IDispatchWrapper, owner: IDispatchWrapper) -> Self {
        Self { font, owner }
    }

    pub fn raw(&self) -> &IDispatchWrapper {
        &self.font
    }

    /// 应用字符格式补丁（`None` 字段不改）。
    pub fn apply(&self, patch: &FontPatch) -> Result<Vec<String>> {
        let (mut applied, mut warnings) = apply_font_props(&self.font, patch)?;
        if let Some(v) = patch.highlight {
            match self
                .owner
                .put("HighlightColorIndex", vec![ComValue::Int(v.code())])
            {
                Ok(()) => applied.push("font.highlight".into()),
                Err(e) => warnings.push(format!("font.highlight: {e}")),
            }
        }
        if let Some(v) = &patch.shading {
            let hex = hex_to_ole(v)?;
            match self
                .owner
                .get_obj("Shading", "Shading")
                .and_then(|s| s.put("BackgroundPatternColor", vec![ComValue::Int(hex as i32)]))
            {
                Ok(()) => applied.push("font.shading".into()),
                Err(e) => warnings.push(format!("font.shading: {e}")),
            }
        }
        applied.extend(warnings.iter().map(|w| format!("warning: {w}")));
        Ok(applied)
    }

    /// 恢复到样式继承格式（清除直接字符格式）。
    pub fn reset(&self) -> Result<()> {
        self.font.call("Reset", vec![]).map(|_| ())
    }
}

/// 供样式定义（`Style.Font`）等无 Range 上下文的场景应用字符属性。
pub(crate) fn apply_style_font(font: &IDispatchWrapper, patch: &FontPatch) -> Result<Vec<String>> {
    let (applied, warnings) = apply_font_props(font, patch)?;
    let mut out = applied;
    out.extend(warnings.iter().map(|w| format!("warning: {w}")));
    Ok(out)
}
