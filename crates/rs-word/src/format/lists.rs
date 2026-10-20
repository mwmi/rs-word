//! 列表：默认项目符号/编号/多级编号 + 自定义多级编号模板（`ListTemplateSpec`）。
//!
//! 自定义模板走 `Application.ListGalleries(wdOutlineNumberGallery).ListTemplates(7)`：
//! 该索引是画廊预留的自定义项，避免在文档里堆出多个重复模板（10 手册 §5.3）。

use anyhow::Result;

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::WordError;
use crate::range::Range;
use crate::types::{ListKind, ListTemplateSpec};

/// wdOutlineNumberGallery。
const OUTLINE_NUMBER_GALLERY: i32 = 3;
/// 画廊内预留的自定义模板索引。
const CUSTOM_TEMPLATE_INDEX: i32 = 7;

/// 对 Range 套用列表格式。
pub fn apply_list(range: &Range, kind: ListKind) -> Result<()> {
    let lf = range.list_format()?;
    match kind {
        ListKind::None => {
            lf.call("RemoveNumbers", vec![])?;
        }
        ListKind::Bullet => {
            lf.call("ApplyBulletDefault", vec![])?;
        }
        ListKind::Number => {
            lf.call("ApplyNumberDefault", vec![])?;
        }
        ListKind::OutlineNumber => {
            lf.call("ApplyOutlineNumberDefault", vec![])?;
        }
        ListKind::Custom => {
            return Err(WordError::new(
                "format.list",
                range.label(),
                "list=custom 需要同时提供 list_template",
            )
            .with_hint("提供 list_template.levels（1..=9 级），或使用 bullet/number/outline_number")
            .into());
        }
    }
    // 读回 ListType 作为验证（失败不致命）。
    let _ = lf.get("ListType").and_then(|v| v.int());
    Ok(())
}

/// 套用自定义多级编号模板（先写级别配置，再 ApplyListTemplate）。
pub fn apply_template(app: &App, range: &Range, spec: &ListTemplateSpec) -> Result<()> {
    spec.validate()
        .map_err(|e| WordError::new("format.list_template", range.label(), e))?;
    let lt = custom_template(app).map_err(|e| {
        WordError::hint(
            e,
            "该引擎不支持 ListGalleries/ListTemplates：自定义多级编号不可用",
        )
    })?;

    // Word 的成员名是 `ListLevels`；个别引擎只暴露 `Levels`，两者都试。
    let levels = lt
        .get_obj("ListLevels", "ListLevels")
        .or_else(|_| lt.get_obj("Levels", "Levels"))?;
    for l in &spec.levels {
        let level = levels
            .call("Item", vec![ComValue::Int(l.level as i32)])?
            .idispatch_child(&format!("Levels[{}]", l.level))?;
        let text = l
            .text
            .clone()
            .unwrap_or_else(|| l.format.default_text().into());
        level.put("NumberFormat", vec![ComValue::Str(text)])?;
        level.put("NumberStyle", vec![ComValue::Int(l.format.code())])?;
        level.put("StartAt", vec![ComValue::Int(l.start_at as i32)])?;
        level.put(
            "NumberPosition",
            vec![ComValue::Float(l.number_position_pt as f64)],
        )?;
        level.put(
            "TextPosition",
            vec![ComValue::Float(l.text_position_pt as f64)],
        )?;
        if let Some(tab) = l.tab_position_pt {
            level.put("TabPosition", vec![ComValue::Float(tab as f64)])?;
        }
        level.put("TrailingCharacter", vec![ComValue::Int(l.trailing.code())])?;
    }

    range.list_format()?.call_named(
        "ApplyListTemplate",
        vec![],
        &[
            ("ListTemplate", ComValue::Obj(lt.disp.clone())),
            (
                "ContinuePreviousList",
                ComValue::Bool(spec.continue_previous),
            ),
            ("ApplyTo", ComValue::Int(1)),
        ],
    )?;

    // 回读验证：ListType/ListLevelNumber 可读即视为已应用（值不强制断言）。
    let lf = range.list_format()?;
    let _ = lf.get("ListType").and_then(|v| v.int());
    let _ = lf.get("ListLevelNumber").and_then(|v| v.int());
    Ok(())
}

/// 取画廊预留的自定义模板。
fn custom_template(app: &App) -> Result<IDispatchWrapper> {
    app.raw()
        .get_obj("ListGalleries", "ListGalleries")?
        .call("Item", vec![ComValue::Int(OUTLINE_NUMBER_GALLERY)])?
        .idispatch_child("ListGallery")?
        .get_obj("ListTemplates", "ListTemplates")?
        .call("Item", vec![ComValue::Int(CUSTOM_TEMPLATE_INDEX)])?
        .idispatch_child("ListTemplate")
}

/// 当前列表级别（1 基）。
pub fn list_level(range: &Range) -> Result<i32> {
    range.list_format()?.get("ListLevelNumber")?.int()
}

/// 调整列表层级（缩进/升级）。
pub fn change_level(range: &Range, up: bool) -> Result<()> {
    let lf = range.list_format()?;
    let m = if up { "ListOutdent" } else { "ListIndent" };
    lf.call(m, vec![ComValue::Int(1)])?;
    Ok(())
}
