//! 段落格式：段落补丁（部分更新）。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::types::ParagraphPatch;

/// 对任意 `ParagraphFormat` 对象应用段落属性（`None` 字段不改）。
/// 返回 (applied, warnings)：单项失败记 warning 不整体失败。
pub(crate) fn apply_paragraph_props(
    fmt: &IDispatchWrapper,
    patch: &ParagraphPatch,
) -> Result<(Vec<String>, Vec<String>)> {
    let mut applied = Vec::new();
    let mut warnings = Vec::new();
    let mut set = |name: &str, v: ComValue| match fmt.put(name, vec![v]) {
        Ok(()) => applied.push(format!("paragraph.{name}")),
        Err(e) => warnings.push(format!("paragraph.{name}: {e}")),
    };
    if let Some(v) = patch.alignment {
        set("Alignment", ComValue::Int(v.code()));
    }
    if let Some(v) = patch.space_before_pt {
        set("SpaceBefore", ComValue::Float(v as f64));
    }
    if let Some(v) = patch.space_after_pt {
        set("SpaceAfter", ComValue::Float(v as f64));
    }
    if let Some(v) = patch.left_indent_pt {
        set("LeftIndent", ComValue::Float(v as f64));
    }
    if let Some(v) = patch.right_indent_pt {
        set("RightIndent", ComValue::Float(v as f64));
    }
    if let Some(v) = patch.first_line_indent_pt {
        set("FirstLineIndent", ComValue::Float(v as f64));
    }
    if let Some(v) = patch.keep_with_next {
        set("KeepWithNext", ComValue::Bool(v));
    }
    if let Some(v) = patch.keep_together {
        set("KeepTogether", ComValue::Bool(v));
    }
    if let Some(v) = patch.page_break_before {
        set("PageBreakBefore", ComValue::Bool(v));
    }
    if let Some(v) = patch.widow_control {
        set("WidowControl", ComValue::Bool(v));
    }
    if let Some(v) = patch.outline_level {
        set("OutlineLevel", ComValue::Int(v.code()));
    }
    // 行距：先写规则，再写数值（Multiple/Exactly/AtLeast 才有数值）。
    if let Some(ls) = patch.line_spacing {
        set("LineSpacingRule", ComValue::Int(ls.rule_code()));
        if let Some(n) = ls.value_for_com() {
            set("LineSpacing", ComValue::Float(n as f64));
        }
    }
    // 悬挂缩进：LeftIndent=h 且 FirstLineIndent=-h。
    if let Some(h) = patch.hanging_indent_pt {
        if h < 0.0 {
            warnings.push("hanging_indent: 负值已忽略".into());
        } else {
            set("LeftIndent", ComValue::Float(h as f64));
            set("FirstLineIndent", ComValue::Float(-(h as f64)));
        }
    }
    Ok((applied, warnings))
}

/// `Range.ParagraphFormat` 入口。
pub struct ParagraphFormat {
    raw: IDispatchWrapper,
}

impl ParagraphFormat {
    pub(crate) fn new(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 应用段落格式补丁（`None` 字段不改）。
    pub fn apply(&self, patch: &ParagraphPatch) -> Result<Vec<String>> {
        let (mut applied, warnings) = apply_paragraph_props(&self.raw, patch)?;
        applied.extend(warnings.iter().map(|w| format!("warning: {w}")));
        Ok(applied)
    }
}
