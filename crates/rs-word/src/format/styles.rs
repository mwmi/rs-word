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
///
/// 入参做归一（忽略大小写/空格/下划线/连字符）：`Heading1` / `heading_1` / `heading 1`
/// 都是同一件事 —— 模板与模型两边的写法都能吃下。
pub fn builtin_style_constant(id: &str) -> Option<i32> {
    let key: String = id
        .trim()
        .chars()
        .filter(|c| !matches!(c, ' ' | '_' | '-'))
        .flat_map(|c| c.to_lowercase())
        .collect();
    Some(match key.as_str() {
        "normal" => -1,
        "heading1" => -2,
        "heading2" => -3,
        "heading3" => -4,
        "heading4" => -5,
        "heading5" => -6,
        "heading6" => -7,
        "heading7" => -8,
        "heading8" => -9,
        "heading9" => -10,
        "caption" => -35,
        "header" => -32,
        "footer" => -33,
        "title" => -63,
        "bodytext" => -67,
        "subtitle" => -75,
        "listparagraph" => -180,
        "quote" => -181,
        "intensequote" => -182,
        // 字符样式（只有字体属性，段落属性会被 apply_profile 跳过）
        "defaultparagraphfont" => -66,
        "hyperlink" => -86,
        "followedhyperlink" => -87,
        "emphasis" => -89,
        "strong" => -88,
        "footnotetext" => -34,
        "endnotetext" => -36,
        "footnotereference" => -39,
        "endnotereference" => -38,
        "commenttext" => -31,
        "commentreference" => -40,
        "pagenumber" => -69,
        "salutation" => -78,
        "date" => -77,
        "signature" => -80,
        "closing" => -79,
        _ => return None,
    })
}

/// 样式名归一：忽略大小写/空格/下划线/连字符（中文原样保留）。
fn style_key(id: &str) -> String {
    id.trim()
        .chars()
        .filter(|c| !matches!(c, ' ' | '_' | '-'))
        .flat_map(|c| c.to_lowercase())
        .collect()
}

/// 内置样式的本地化名：WPS 不认某个常量时，靠这些名字二次查找。
/// 例：`ListParagraph` 在中文 WPS 里叫「列表段落」，直接 `Styles.Add` 会报「样式名已使用」。
fn style_name_aliases(id: &str) -> Vec<String> {
    let key = style_key(id);
    if let Some(level) = key
        .strip_prefix("heading")
        .and_then(|rest| rest.parse::<u8>().ok())
    {
        return vec![format!("Heading {level}"), format!("标题 {level}"), format!("标题{level}")];
    }
    let fixed: &[&str] = match key.as_str() {
        "normal" => &["Normal", "正文"],
        "bodytext" => &["Body Text", "正文文本"],
        "listparagraph" => &["List Paragraph", "列表段落"],
        "title" => &["Title", "标题"],
        "subtitle" => &["Subtitle", "副标题"],
        "quote" => &["Quote", "引用"],
        "intensequote" => &["Intense Quote", "明显引用"],
        "caption" => &["Caption", "题注"],
        "code" => &["Code", "代码"],
        "header" => &["Header", "页眉"],
        "footer" => &["Footer", "页脚"],
        "footnotetext" => &["Footnote Text", "脚注文本"],
        "footnotereference" => &["Footnote Reference", "脚注引用"],
        "endnotetext" => &["Endnote Text", "尾注文本"],
        "endnotereference" => &["Endnote Reference", "尾注引用"],
        "commenttext" => &["Comment Text", "批注文字", "批注框文本"],
        "commentreference" => &["Comment Reference", "批注引用"],
        "defaultparagraphfont" => &["Default Paragraph Font", "默认段落字体"],
        "hyperlink" => &["Hyperlink", "超链接"],
        "followedhyperlink" => &["Followed Hyperlink", "已访问的超链接"],
        "emphasis" => &["Emphasis", "强调"],
        "strong" => &["Strong", "强烈强调"],
        "pagenumber" => &["Page Number", "页码"],
        "salutation" => &["Salutation", "称呼"],
        "date" => &["Date", "日期"],
        "signature" => &["Signature", "签名"],
        "closing" => &["Closing", "结束语"],
        _ => &[],
    };
    fixed.iter().map(|name| (*name).to_string()).collect()
}

/// `Document.Styles.Item(...)`；常量 → 原名 → 本地化别名，逐级兜底。
fn style_object(doc: &Document, style: &str) -> Result<IDispatchWrapper> {
    let styles = doc.raw().get_obj("Styles", "Styles")?;
    if let Some(constant) = builtin_style_constant(style) {
        if let Ok(obj) = styles.call("Item", vec![ComValue::Int(constant)]) {
            return obj.idispatch_child(&format!("Style[{style}]"));
        }
    }
    let mut names = vec![style.to_string()];
    names.extend(style_name_aliases(style));
    for name in names {
        if let Ok(obj) = styles.call("Item", vec![ComValue::Str(name)]) {
            return obj.idispatch_child(&format!("Style[{style}]"));
        }
    }
    Err(anyhow::anyhow!("样式「{style}」在文档里不存在"))
}

/// 全表按名称/本地化名找样式：`Styles.Add` 报「名称已使用」时的最后一条路。
fn find_style_named(styles: &IDispatchWrapper, id: &str) -> Option<IDispatchWrapper> {
    let mut targets: Vec<String> = vec![style_key(id)];
    targets.extend(style_name_aliases(id).iter().map(|name| style_key(name)));
    let count = styles.get("Count").ok()?.int().ok()?;
    for index in 1..=count {
        let Ok(style) = styles.call("Item", vec![ComValue::Int(index)]) else {
            continue;
        };
        let Ok(style) = style.idispatch_child("Style") else {
            continue;
        };
        for key in ["NameLocal", "Name"] {
            let Ok(value) = style.get(key) else { continue };
            let Ok(name) = value.string() else { continue };
            if targets.contains(&style_key(&name)) {
                return Some(style);
            }
        }
    }
    None
}

/// 给 Range 套用样式（优先内置常量，避免本地化名称差异）。
pub fn apply_style(range: &Range, doc: &Document, style: &str) -> Result<()> {
    let obj = style_object(doc, style)?;
    range
        .raw()
        .put("Style", vec![ComValue::Obj(obj.disp.clone())])
}

/// 同 [`apply_style`]，但样式不存在时先建一个（模板里的 `Code` / 自定义样式名）。
pub fn apply_style_or_create(range: &Range, doc: &Document, style: &str) -> Result<()> {
    let obj = style_object_or_create(doc, style)?;
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
    apply_profile_ext(doc, profile, dry_run).map(|(applied, _)| applied)
}

/// 同 [`apply_profile`]，但把单条样式的失败降级为 warning：
/// 一个模板里有 8 个样式，其中一个字体写错不该让整次排版作废。
pub fn apply_profile_ext(
    doc: &Document,
    profile: &StyleProfile,
    dry_run: bool,
) -> Result<(Vec<String>, Vec<String>)> {
    let mut applied = Vec::new();
    let mut warnings = Vec::new();
    for rule in &profile.rules {
        let name = rule.style.trim();
        if name.is_empty() {
            continue;
        }
        if dry_run {
            applied.push(format!(
                "计划：样式 {name}（字体 {} / 段落 {}）",
                rule.font.is_some(),
                rule.paragraph.is_some()
            ));
            continue;
        }
        let s = match style_object_or_create(doc, name) {
            Ok(s) => s,
            Err(e) => {
                warnings.push(format!("样式 {name} 不可用：{e}"));
                continue;
            }
        };
        if let Some(f) = &rule.font {
            match apply_style_font(&s.get_obj("Font", "Font")?, f) {
                Ok(items) => applied.extend(items.into_iter().map(|x| format!("{name}.{x}"))),
                Err(e) => warnings.push(format!("样式 {name} 字体未套上：{e}")),
            }
        }
        if let Some(p) = &rule.paragraph {
            // 字符样式（超链接/强调/批注引用…）没有 ParagraphFormat：跳过，不拖垮整套模板
            match s.get_obj("ParagraphFormat", "ParagraphFormat") {
                Ok(fmt) => match apply_paragraph_props(&fmt, p) {
                    Ok((a, w)) => {
                        applied.extend(a.into_iter().map(|x| format!("{name}.{x}")));
                        warnings.extend(w);
                    }
                    Err(e) => warnings.push(format!("样式 {name} 段落属性未套上：{e}")),
                },
                Err(_) => {}
            }
        }
    }
    Ok((applied, warnings))
}

/// 按段落序号批量套样式（一次 COM 循环）。
///
/// `expected_text_hash` 是分析阶段记下的段落指纹：不一致说明文档被改过，
/// 该段跳过并计入 warning —— 绝不把标题样式误套到别的段落上。
pub fn apply_assignments(
    doc: &Document,
    list: &[crate::types::StyleAssignment],
    dry_run: bool,
    reset_font: bool,
) -> Result<(Vec<String>, Vec<String>)> {
    let mut applied = Vec::new();
    let mut warnings = Vec::new();
    if list.is_empty() {
        return Ok((applied, warnings));
    }
    if dry_run {
        applied.push(format!("计划：按段套样式 {} 段", list.len()));
        return Ok((applied, warnings));
    }
    let paras = doc.raw().get_obj("Paragraphs", "Paragraphs")?;
    let count = paras.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut done = 0usize;
    let mut cache: std::collections::HashMap<String, IDispatchWrapper> =
        std::collections::HashMap::new();
    for item in list {
        let index = item.paragraph_index as i32;
        if index < 1 || index > count {
            warnings.push(format!("第 {index} 段不存在，已跳过"));
            continue;
        }
        let range = paras
            .call("Item", vec![ComValue::Int(index)])?
            .idispatch_child(&format!("Paragraph[{index}]"))?
            .get_obj("Range", "Paragraph.Range")?;
        let range = Range::new(range);
        if let Some(hash) = &item.expected_text_hash {
            let text = crate::anchors::display_text(&range.text().unwrap_or_default());
            if &crate::anchors::hash_text(&text) != hash {
                warnings.push(format!("第 {index} 段内容已变化，已跳过"));
                continue;
            }
        }
        // 同一样式在几百段里反复解析没必要：一次解析、整批复用（找不到时也只扫一遍全表）
        let style = match cache.get(&item.style) {
            Some(obj) => obj.clone(),
            None => match style_object_or_create(doc, &item.style) {
                Ok(obj) => {
                    cache.insert(item.style.clone(), obj.clone());
                    obj
                }
                Err(e) => {
                    warnings.push(format!("第 {index} 段套样式失败（{}）：{e}", item.style));
                    continue;
                }
            },
        };
        match range
            .raw()
            .put("Style", vec![ComValue::Obj(style.disp.clone())])
        {
            Ok(()) => {
                // 覆盖直接格式：只清字符格式（加粗/字号/颜色），段落属性仍由样式决定
                if reset_font {
                    if let Err(e) = range.font().and_then(|font| font.reset()) {
                        warnings.push(format!("第 {index} 段直接格式未清：{e}"));
                    }
                }
                // 行距自愈：固定行距比字号还小时，Word 会把字按行框裁掉（标题看着被遮住）
                if let Ok(true) = heal_line_spacing(&range) {
                    warnings.push(format!(
                        "第 {index} 段固定行距小于字号，已改为单倍行距（否则文字会被裁掉）"
                    ));
                }
                done += 1
            }
            Err(e) => warnings.push(format!("第 {index} 段套样式失败（{}）：{e}", item.style)),
        }
    }
    applied.push(format!("按段套样式 {done}/{} 段", list.len()));
    Ok((applied, warnings))
}

/// 行距自愈：`LineSpacingRule=4`（固定值）且磅数小于字号 → 改成单倍行距。
/// 只在确实会裁字时动手；正常文档（例如 16 磅字配 28 磅固定行距）不受影响。
fn heal_line_spacing(range: &Range) -> Result<bool> {
    let fmt = range.paragraph_format()?;
    let rule = fmt.raw().get("LineSpacingRule").and_then(|v| v.int()).unwrap_or(0);
    if rule != 4 && rule != 5 {
        return Ok(false);
    }
    let spacing = fmt.raw().get("LineSpacing").and_then(|v| v.float()).unwrap_or(0.0);
    let size = range
        .font()?
        .raw()
        .get("Size")
        .and_then(|v| v.float())
        .unwrap_or(0.0);
    if size <= 0.0 {
        return Ok(false);
    }
    // 固定值小于字号 = 裁字；多倍行距的磅数不到半行（< 6 磅）= 单位写错（1 磅被当成 1 倍）
    let broken = match rule {
        4 => spacing + 0.5 < size,
        5 => spacing < 6.0,
        _ => false,
    };
    if broken {
        fmt.raw().put("LineSpacingRule", vec![ComValue::Int(0)])?;
        return Ok(true);
    }
    Ok(false)
}

/// 样式对象：内置常量 / 名称都找不到时新建一个段落样式（模板里的自定义样式名）。
///
/// `Styles.Add` 失败多半不是「没这个样式」，而是「这个内置样式在本地叫别的名字」：
/// 再全表扫一遍本地化名，找不到才真正报错。
fn style_object_or_create(doc: &Document, name: &str) -> Result<IDispatchWrapper> {
    if let Ok(s) = style_object(doc, name) {
        return Ok(s);
    }
    let styles = doc.raw().get_obj("Styles", "Styles")?;
    match styles
        .call("Add", vec![ComValue::Str(name.to_string()), ComValue::Int(1)])
        .and_then(|obj| obj.idispatch_child(&format!("Style[{name}]")))
    {
        Ok(style) => Ok(style),
        Err(add_error) => {
            if let Some(found) = find_style_named(&styles, name) {
                return Ok(found);
            }
            if let Ok(obj) = styles.call("Item", vec![ComValue::Str(name.to_string())]) {
                return obj.idispatch_child(&format!("Style[{name}]"));
            }
            Err(add_error)
        }
    }
}
