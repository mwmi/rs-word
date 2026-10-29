//! 标题编号：把 Heading 1..N 绑定到 Word 原生多级列表（编号不写进标题正文）。
//!
//! 走画廊预留的自定义模板（索引 7），与 `lists.rs` 同一个槽位：不往文档里堆重复模板，
//! 也不改内置画廊模板。绑定用 `Style.LinkToListTemplate`，之后插入/更新目录、导航窗格
//! 都能按标题级别工作。

use anyhow::Result;

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::format::styles::builtin_style_constant;
use crate::types::HeadingNumberingSpec;

/// wdOutlineNumberGallery。
const OUTLINE_NUMBER_GALLERY: i32 = 3;
/// 画廊内预留的自定义模板索引（与 lists.rs 一致）。
const CUSTOM_TEMPLATE_INDEX: i32 = 7;
/// wdListNumberStyleArabic。
const NUMBER_STYLE_ARABIC: i32 = 0;

/// 默认编号格式：1. / 1.1 / 1.1.1 ……
fn default_format(level: u8) -> String {
    let mut out = String::new();
    for i in 1..=level {
        if i > 1 {
            out.push('.');
        }
        out.push('%');
        out.push_str(&i.to_string());
    }
    if level == 1 {
        out.push('.');
    }
    out
}

/// 应用标题编号。单个级别失败不影响其它级别（返回 warning 由调用方展示）。
pub fn apply(app: &App, doc: &Document, spec: &HeadingNumberingSpec) -> Result<(Vec<String>, Vec<String>)> {
    let mut applied = Vec::new();
    let mut warnings = Vec::new();
    if spec.levels.is_empty() {
        return Ok((applied, warnings));
    }
    let lt = match custom_template(app) {
        Ok(lt) => lt,
        Err(e) => {
            warnings.push(format!("当前引擎不支持多级列表模板，标题编号未应用：{e}"));
            return Ok((applied, warnings));
        }
    };
    let levels = lt
        .get_obj("ListLevels", "ListLevels")
        .or_else(|_| lt.get_obj("Levels", "Levels"));
    let levels = match levels {
        Ok(levels) => levels,
        Err(e) => {
            warnings.push(format!("读不到多级列表级别，标题编号未应用：{e}"));
            return Ok((applied, warnings));
        }
    };
    let styles = doc.raw().get_obj("Styles", "Styles")?;

    for item in &spec.levels {
        let level = item.level.clamp(1, 9);
        let format = item
            .format
            .clone()
            .filter(|f| !f.trim().is_empty())
            .unwrap_or_else(|| default_format(level));
        let level_obj = match levels
            .call("Item", vec![ComValue::Int(level as i32)])
            .and_then(|v| v.idispatch_child("ListLevel"))
        {
            Ok(obj) => obj,
            Err(e) => {
                warnings.push(format!("第 {level} 级编号不可用：{e}"));
                continue;
            }
        };
        if let Err(e) = level_obj.put("NumberFormat", vec![ComValue::Str(format.clone())]) {
            warnings.push(format!("第 {level} 级编号格式未写入：{e}"));
            continue;
        }
        let _ = level_obj.put("NumberStyle", vec![ComValue::Int(NUMBER_STYLE_ARABIC)]);
        if let Some(start) = item.start_at {
            let _ = level_obj.put("StartAt", vec![ComValue::Int(start as i32)]);
        }
        let constant = match builtin_style_constant(&format!("heading_{level}")) {
            Some(code) => code,
            None => continue,
        };
        let style = match styles
            .call("Item", vec![ComValue::Int(constant)])
            .and_then(|v| v.idispatch_child("Style"))
        {
            Ok(style) => style,
            Err(e) => {
                warnings.push(format!("标题 {level} 样式不可用：{e}"));
                continue;
            }
        };
        match style.call(
            "LinkToListTemplate",
            vec![
                ComValue::Obj(lt.disp.clone()),
                ComValue::Int(level as i32),
            ],
        ) {
            Ok(_) => applied.push(format!("标题{level} → {format}")),
            Err(e) => warnings.push(format!("标题 {level} 未绑定编号：{e}")),
        }
    }
    Ok((applied, warnings))
}

fn custom_template(app: &App) -> Result<IDispatchWrapper> {
    app.raw()
        .get_obj("ListGalleries", "ListGalleries")?
        .call("Item", vec![ComValue::Int(OUTLINE_NUMBER_GALLERY)])?
        .idispatch_child("ListGallery")?
        .get_obj("ListTemplates", "ListTemplates")?
        .call("Item", vec![ComValue::Int(CUSTOM_TEMPLATE_INDEX)])?
        .idispatch_child("ListTemplate")
}
