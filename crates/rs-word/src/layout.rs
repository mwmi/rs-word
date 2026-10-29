//! 页面布局：PageSetup（按 Section）、页眉页脚、页码、目录、水印、字段更新、PDF 导出。

use anyhow::Result;

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::hex_to_ole;
use crate::range::Range;
use crate::types::{
    HfKind, LayoutAction, LayoutRequest, LayoutResult, PageNumberSpec, PageSetupPatch, TocSpec,
    WatermarkSpec,
};

/// mm → points（72 pt = 1 inch）。
fn mm_to_pt(mm: f32) -> f64 {
    mm as f64 * 72.0 / 25.4
}

/// 布局动作执行入口。
pub fn exec(
    app: &App,
    doc: &Document,
    req: &LayoutRequest,
    doc_version: u64,
) -> Result<LayoutResult> {
    let mut out = LayoutResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        applied: Vec::new(),
        warnings: Vec::new(),
        exported: None,
    };
    match req.action {
        LayoutAction::SetPageSetup => {
            let page = req
                .page
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("set_page_setup 需要 page 参数"))?;
            let (n, warnings) = apply_page_setup(doc, page)?;
            out.warnings.extend(warnings);
            out.applied.push(format!("page_setup on {n} section(s)"));
            out.doc_version = doc_version + 1;
        }
        LayoutAction::SetHeaderFooter => {
            let hf = req
                .header_footer
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("set_header_footer 需要 header_footer 参数"))?;
            apply_header_footer(doc, hf)?;
            out.applied
                .push(format!("{:?} section {}", hf.kind, hf.section));
            out.doc_version = doc_version + 1;
        }
        LayoutAction::InsertPageNumber => {
            let hf = req.header_footer.as_ref().ok_or_else(|| {
                anyhow::anyhow!("insert_page_number 需要 header_footer 指定 section")
            })?;
            insert_page_number(doc, hf.section, req.page_number.as_ref())?;
            out.applied.push("page_number".into());
            out.doc_version = doc_version + 1;
        }
        LayoutAction::UpdateFields => {
            update_fields(doc)?;
            out.applied.push("update_fields".into());
            out.doc_version = doc_version + 1;
        }
        LayoutAction::InsertToc => {
            let spec = req
                .toc
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("insert_toc 需要 toc 参数"))?;
            out.applied.push(insert_toc(doc, spec)?);
            out.doc_version = doc_version + 1;
        }
        LayoutAction::UpdateToc => {
            out.applied.push(update_toc(doc)?);
            out.doc_version = doc_version + 1;
        }
        LayoutAction::SetWatermark => {
            let spec = req
                .watermark
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("set_watermark 需要 watermark 参数"))?;
            out.applied.push(set_watermark(doc, spec)?);
            out.doc_version = doc_version + 1;
        }
        LayoutAction::SetHeaderFooterZones => {
            let zones = req
                .header_footer_zones
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("set_header_footer_zones 需要 header_footer_zones 参数"))?;
            out.applied.extend(apply_hf_zones(doc, zones)?);
            out.doc_version = doc_version + 1;
        }
        LayoutAction::HeadingNumbering => {
            let spec = req
                .numbering
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("heading_numbering 需要 numbering 参数"))?;
            let (applied, warnings) = crate::format::numbering::apply(app, doc, spec)?;
            out.applied.extend(applied);
            out.warnings.extend(warnings);
            out.doc_version = doc_version + 1;
        }
        LayoutAction::ExportPdf => {
            let path = req
                .pdf_path
                .as_deref()
                .ok_or_else(|| anyhow::anyhow!("export_pdf 需要 pdf_path"))?;
            let bytes = doc.export_pdf(path, req.overwrite)?;
            out.applied.push("export_pdf".into());
            out.exported = Some(crate::types::ExportResult {
                path: crate::process::abs_path(path)?,
                bytes,
            });
        }
    }
    Ok(out)
}

fn sections(doc: &Document) -> Result<crate::com::dispatch::IDispatchWrapper> {
    doc.sections()
}

/// 对每个 Section 应用页面设置（页面设置按 Section 生效）。
/// 扩展项（分栏/页眉页脚距/首页奇偶页）按尽力而为处理：失败记 warning，不中断整次应用。
fn apply_page_setup(doc: &Document, page: &PageSetupPatch) -> Result<(i32, Vec<String>)> {
    let secs = sections(doc)?;
    let count = secs.get("Count")?.int()?;
    let mut warnings = Vec::new();
    let mut align_warned = false;
    for i in 1..=count {
        let ps = secs
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child("Section")?
            .get_obj("PageSetup", "PageSetup")?;
        // 版式是「垂直居中」时，内容一旦比一页高，Word 会把上下溢出部分直接裁掉
        //（表现就是首行标题只剩半截、末行被截）—— 只提示，不擅自改用户的版式
        if !align_warned {
            if let Ok(align) = ps.get("VerticalAlignment").and_then(|v| v.int()) {
                if align != 0 {
                    warnings.push(
                        "该节版式为「垂直居中」：内容超过一页时上下会被裁掉（文字看着被遮住）；建议改为顶端对齐"
                            .into(),
                    );
                    align_warned = true;
                }
            }
        }
        if let Some(v) = page.size {
            ps.put("PaperSize", vec![ComValue::Int(v.code())])?;
        }
        if let Some(v) = page.orientation {
            ps.put("Orientation", vec![ComValue::Int(v.code())])?;
        }
        if let Some(v) = page.vertical_align {
            soft(
                &mut warnings,
                &ps,
                "VerticalAlignment",
                ComValue::Int(v.code()),
                "版式垂直对齐",
            );
        }
        if let Some(v) = page.margin_top_mm {
            ps.put("TopMargin", vec![ComValue::Float(mm_to_pt(v))])?;
        }
        if let Some(v) = page.margin_bottom_mm {
            ps.put("BottomMargin", vec![ComValue::Float(mm_to_pt(v))])?;
        }
        if let Some(v) = page.margin_left_mm {
            ps.put("LeftMargin", vec![ComValue::Float(mm_to_pt(v))])?;
        }
        if let Some(v) = page.margin_right_mm {
            ps.put("RightMargin", vec![ComValue::Float(mm_to_pt(v))])?;
        }
        if let Some(v) = page.gutter_mm {
            ps.put("Gutter", vec![ComValue::Float(mm_to_pt(v))])?;
        }
        if let Some(v) = page.header_distance_mm {
            soft(&mut warnings, &ps, "HeaderDistance", ComValue::Float(mm_to_pt(v)), "页眉距边界");
        }
        if let Some(v) = page.footer_distance_mm {
            soft(&mut warnings, &ps, "FooterDistance", ComValue::Float(mm_to_pt(v)), "页脚距边界");
        }
        if let Some(v) = page.different_first_page {
            soft(
                &mut warnings,
                &ps,
                "DifferentFirstPageHeaderFooter",
                ComValue::Bool(v),
                "首页页眉页脚不同",
            );
        }
        if let Some(v) = page.different_odd_even {
            soft(
                &mut warnings,
                &ps,
                "OddAndEvenPagesHeaderFooter",
                ComValue::Bool(v),
                "奇偶页页眉页脚不同",
            );
        }
        if let Some(columns) = page.columns {
            apply_columns(&ps, columns, page.column_spacing_mm, &mut warnings);
        }
    }
    Ok((count, warnings))
}

/// 单个属性写入：失败只记 warning（引擎差异不该让整页设置失败）。
fn soft(warnings: &mut Vec<String>, obj: &IDispatchWrapper, name: &str, value: ComValue, label: &str) {
    if let Err(e) = obj.put(name, vec![value]) {
        warnings.push(format!("{label} 未生效：{e}"));
    }
}

/// 分栏：TextColumns.SetCount + Spacing。
fn apply_columns(ps: &IDispatchWrapper, count: u16, spacing_mm: Option<f32>, warnings: &mut Vec<String>) {
    let result = ps.get_obj("TextColumns", "TextColumns").and_then(|tc| {
        tc.call("SetCount", vec![ComValue::Int(count.max(1) as i32)])?;
        if let Some(mm) = spacing_mm {
            tc.put("Spacing", vec![ComValue::Float(mm_to_pt(mm))])?;
        }
        Ok(())
    });
    if let Err(e) = result {
        warnings.push(format!("分栏未生效：{e}"));
    }
}

fn header_footer(
    doc: &Document,
    section: u32,
    kind: HfKind,
    variant: crate::types::HfVariant,
) -> Result<crate::com::dispatch::IDispatchWrapper> {
    let sec = sections(doc)?
        .call("Item", vec![ComValue::Int(section as i32)])?
        .idispatch_child(&format!("Section[{section}]"))?;
    let coll = match kind {
        HfKind::Header => "Headers",
        HfKind::Footer => "Footers",
    };
    sec.get_obj(coll, coll)?
        .call("Item", vec![ComValue::Int(variant.code())])?
        .idispatch_child("HeaderFooter")
}

fn apply_header_footer(doc: &Document, spec: &crate::types::HeaderFooterSpec) -> Result<()> {
    let hf = header_footer(doc, spec.section, spec.kind, spec.variant)?;
    hf.put(
        "LinkToPrevious",
        vec![ComValue::Bool(spec.link_to_previous)],
    )?;
    if let Some(text) = &spec.text {
        let range = hf.get_obj("Range", "HeaderFooter.Range")?;
        range.put("Text", vec![ComValue::Str(text.clone())])?;
        if let Some(align) = spec.alignment {
            range
                .get_obj("ParagraphFormat", "ParagraphFormat")?
                .put("Alignment", vec![ComValue::Int(align.code())])?;
        }
    }
    Ok(())
}

/// wdNumberStyle*：十进制/罗马/字母/中文。
fn number_style_code(format: &str) -> Option<i32> {
    let key = format.trim().to_ascii_lowercase().replace(['_', '-', ' '], "");
    Some(match key.as_str() {
        "decimal" | "arabic" | "1" => 0,
        "upperroman" => 1,
        "lowerroman" => 2,
        "upperletter" => 3,
        "lowerletter" => 4,
        "chinese" => 38,
        _ => return None,
    })
}

/// 页眉/页脚三区写入：`左\t中\t右`（默认页眉样式自带居中与右对齐制表位），
/// `{页码}` / `{总页数}` 换成 Word 字段（静态数字在分页变化后必错）。
fn apply_hf_zones(doc: &Document, zones: &crate::types::HeaderFooterZones) -> Result<Vec<String>> {
    let mut applied = Vec::new();
    let sec = sections(doc)?
        .call("Item", vec![ComValue::Int(1)])?
        .idispatch_child("Section[1]")?;
    let ps = sec.get_obj("PageSetup", "PageSetup")?;
    let _ = ps.put(
        "DifferentFirstPageHeaderFooter",
        vec![ComValue::Bool(zones.different_first_page)],
    );
    let _ = ps.put(
        "OddAndEvenPagesHeaderFooter",
        vec![ComValue::Bool(zones.different_odd_even)],
    );
    if let Some(format) = zones.page_number_format.as_deref().and_then(number_style_code) {
        let _ = ps.put("PageNumberStyle", vec![ComValue::Int(format)]);
    }
    if let Some(start) = zones.start_page_number {
        let _ = ps.put("StartingNumber", vec![ComValue::Int(start.max(1) as i32)]);
    }
    for (enabled, kind, left, center, right) in [
        (
            zones.header,
            HfKind::Header,
            zones.header_left.as_deref(),
            zones.header_center.as_deref(),
            zones.header_right.as_deref(),
        ),
        (
            zones.footer,
            HfKind::Footer,
            zones.footer_left.as_deref(),
            zones.footer_center.as_deref(),
            zones.footer_right.as_deref(),
        ),
    ] {
        if !enabled {
            continue;
        }
        let hf = header_footer(doc, 1, kind, crate::types::HfVariant::Primary)?;
        let story = hf.get_obj("Range", "HeaderFooter.Range")?;
        story.put("Text", vec![ComValue::Str(String::new())])?;
        let start = story.get("Start")?.int()?;
        let line = build_hf_line(left.unwrap_or(""), center.unwrap_or(""), right.unwrap_or(""));
        let mut pos = start;
        for part in line {
            match part {
                HfPart::Text(text) => {
                    if text.is_empty() {
                        continue;
                    }
                    let at = story_range(&story, pos)?;
                    at.insert_after(&text)?;
                    pos += text.encode_utf16().count() as i32;
                }
                HfPart::Field(code) => {
                    let at = story_range(&story, pos)?;
                    let fields = at.raw().get_obj("Fields", "Fields")?;
                    let field = fields
                        .call(
                            "Add",
                            vec![
                                ComValue::Obj(at.raw().disp.clone()),
                                ComValue::Int(code),
                            ],
                        )?
                        .idispatch_child("Field")?;
                    pos = field
                        .get_obj("Result", "Field.Result")
                        .and_then(|r| r.get("End"))
                        .and_then(|v| v.int())
                        .unwrap_or(pos + 1);
                }
            }
        }
        // 字体/字号/颜色：套到整段页眉页脚
        if let Ok(font) = story.get_obj("Font", "Font") {
            if let Some(name) = &zones.font {
                let _ = font.put("Name", vec![ComValue::Str(name.clone())]);
                let _ = font.put("NameFarEast", vec![ComValue::Str(name.clone())]);
            }
            if let Some(size) = zones.size_pt {
                let _ = font.put("Size", vec![ComValue::Float(size as f64)]);
            }
            if let Some(color) = zones.color.as_deref().and_then(|c| hex_to_ole(c).ok()) {
                let _ = font.put("Color", vec![ComValue::Int(color as i32)]);
            }
        }
        applied.push(format!("{:?} 三区已写入", kind));
    }
    Ok(applied)
}

enum HfPart {
    Text(String),
    Field(i32),
}

/// wdFieldPage / wdFieldNumPages。
const FIELD_PAGE: i32 = 33;
const FIELD_NUM_PAGES: i32 = 26;

fn build_hf_line(left: &str, center: &str, right: &str) -> Vec<HfPart> {
    let mut out = Vec::new();
    for (index, zone) in [left, center, right].into_iter().enumerate() {
        if index > 0 {
            out.push(HfPart::Text("\t".into()));
        }
        out.extend(zone_parts(zone));
    }
    out
}

fn zone_parts(zone: &str) -> Vec<HfPart> {
    let mut out = Vec::new();
    let mut rest = zone;
    while !rest.is_empty() {
        let page = rest.find("{页码}");
        let pages = rest.find("{总页数}");
        let next = match (page, pages) {
            (Some(a), Some(b)) => Some(a.min(b)),
            (Some(a), None) => Some(a),
            (None, Some(b)) => Some(b),
            (None, None) => None,
        };
        let Some(at) = next else {
            out.push(HfPart::Text(rest.to_string()));
            break;
        };
        if at > 0 {
            out.push(HfPart::Text(rest[..at].to_string()));
        }
        if rest[at..].starts_with("{页码}") {
            out.push(HfPart::Field(FIELD_PAGE));
            rest = &rest[at + "{页码}".len()..];
        } else {
            out.push(HfPart::Field(FIELD_NUM_PAGES));
            rest = &rest[at + "{总页数}".len()..];
        }
    }
    out
}

/// 在某个 story（页眉页脚）内的位置切片：`Document.Range` 只覆盖主 story。
fn story_range(story: &IDispatchWrapper, pos: i32) -> Result<Range> {
    let dup = story.call("Duplicate", vec![])?.idispatch_child("Range")?;
    dup.call("SetRange", vec![ComValue::Int(pos), ComValue::Int(pos)])?;
    Ok(Range::new(dup))
}

fn insert_page_number(doc: &Document, section: u32, spec: Option<&PageNumberSpec>) -> Result<()> {
    let sec = sections(doc)?
        .call("Item", vec![ComValue::Int(section as i32)])?
        .idispatch_child(&format!("Section[{section}]"))?;
    let ps = sec.get_obj("PageSetup", "PageSetup")?;
    if let Some(spec) = spec {
        if let Some(format) = spec.format.as_deref().and_then(number_style_code) {
            let _ = ps.put("PageNumberStyle", vec![ComValue::Int(format)]);
        }
        if let Some(start) = spec.start_at {
            let _ = ps.put("StartingNumber", vec![ComValue::Int(start.max(1) as i32)]);
        }
    }
    let hf = header_footer(
        doc,
        section,
        HfKind::Footer,
        crate::types::HfVariant::Primary,
    )?;
    hf.get_obj("PageNumbers", "PageNumbers")?.call_named(
        "Add",
        vec![],
        &[
            ("PageNumberAlignment", ComValue::Int(1)),
            ("FirstPage", ComValue::Bool(true)),
        ],
    )?;
    Ok(())
}

/// 目录：不存在则插入（可带标题段），已存在则原地更新。
pub(crate) fn insert_toc(doc: &Document, spec: &TocSpec) -> Result<String> {
    let max_level = spec.max_level.clamp(1, 9) as i32;
    let tocs = doc.raw().get_obj("TablesOfContents", "TablesOfContents")?;
    let count = tocs.get("Count").and_then(|v| v.int()).unwrap_or(0);
    if count > 0 {
        tocs.call("Item", vec![ComValue::Int(1)])?
            .idispatch_child("TablesOfContents[1]")?
            .call("Update", vec![])?;
        return Ok("目录已更新（不重复插入）".into());
    }
    // 落点：文档开头；首段是文档标题（Title）时排在标题之后，别把目录压在大标题上面。
    let mut at = doc.content()?.collapse_start()?.start()?;
    let mut title_width = 0;
    if let Ok(first) = doc
        .raw()
        .get_obj("Paragraphs", "Paragraphs")
        .and_then(|paras| paras.call("Item", vec![ComValue::Int(1)]))
        .and_then(|para| para.idispatch_child("Paragraph[1]"))
    {
        if is_title_style(&paragraph_style(&first)) {
            if let Ok(end) = first
                .get_obj("Range", "Paragraph.Range")
                .and_then(|range| range.get("End"))
                .and_then(|value| value.int())
            {
                at = end;
            }
        }
    }
    if let Some(title) = spec.title.as_deref().map(str::trim).filter(|t| !t.is_empty()) {
        let payload = format!("{title}\r");
        let width = payload.encode_utf16().count() as i32;
        doc.range(at, at)?.insert_after(&payload)?;
        if let Some(style) = spec.title_style.as_deref().filter(|s| !s.trim().is_empty()) {
            let _ = doc.range(at, at + width).and_then(|r| r.set_style(style));
        }
        title_width = width;
        at += width;
    }
    let target = doc.range(at, at)?;
    let toc = tocs.call(
        "Add",
        vec![
            ComValue::Obj(target.raw().disp.clone()),
            ComValue::Bool(true),
            ComValue::Int(1),
            ComValue::Int(max_level),
        ],
    )?;
    let toc = toc.idispatch_child("TablesOfContents[new]")?;
    let _ = toc.put("UseHyperlinks", vec![ComValue::Bool(true)]);
    let _ = toc.call("Update", vec![]);
    // 一个标题都没有时 Word 会塞一行「错误！未找到目录项」：清掉目录与标题段，
    // 宁可没有目录，也不要给用户留一行错误文字。
    let entries = toc
        .get_obj("Range", "TOC.Range")
        .and_then(|range| range.get("Text"))
        .and_then(|value| value.string())
        .unwrap_or_default();
    let lower = entries.to_lowercase();
    if entries.trim().is_empty()
        || entries.contains("错误")
        || entries.contains("未找到目录项")
        || lower.contains("no table of contents")
    {
        let _ = toc.call("Delete", vec![]);
        if spec.title.is_some() {
            if let Ok(range) = doc.range(at, at + title_width) {
                let _ = range.raw().call("Delete", vec![]);
            }
        }
        return Ok("没有可用标题（先套标题样式），目录已跳过".into());
    }
    // 目录之后补一个分页符：正文另起一页（拿不到 Range 就跳过，不影响目录本身）
    if let Ok(end) = toc
        .get_obj("Range", "TOC.Range")
        .and_then(|range| range.get("End"))
        .and_then(|value| value.int())
    {
        if let Ok(at) = doc.range(end, end) {
            let _ = at.raw().call("InsertBreak", vec![ComValue::Int(7)]);
        }
    }
    Ok(format!("已插入目录（{max_level} 级）"))
}

/// 段落样式名（`Style` 可能是对象或字符串，与快照同一读法）。
fn paragraph_style(para: &IDispatchWrapper) -> String {
    if let Ok(value) = para.get("Style") {
        if let Ok(name) = value.string() {
            if !name.is_empty() && !name.contains("<未支持") {
                return name;
            }
        }
    }
    para.get_obj("Style", "Style")
        .and_then(|style| style.get("NameLocal"))
        .and_then(|value| value.string())
        .unwrap_or_default()
}

fn is_title_style(name: &str) -> bool {
    let key = name.trim().to_lowercase().replace([' ', '_', '-'], "");
    key == "title" || key == "标题" || key == "文档标题"
}

fn update_toc(doc: &Document) -> Result<String> {
    let tocs = doc.raw().get_obj("TablesOfContents", "TablesOfContents")?;
    let count = tocs.get("Count").and_then(|v| v.int()).unwrap_or(0);
    if count == 0 {
        return Ok("没有目录可更新".into());
    }
    for i in 1..=count {
        if let Ok(toc) = tocs
            .call("Item", vec![ComValue::Int(i)])
            .and_then(|v| v.idispatch_child("TablesOfContents"))
        {
            let _ = toc.call("Update", vec![]);
        }
    }
    Ok(format!("目录已更新（{count} 个）"))
}

/// 文字水印：往第一节页眉塞一个衬于文字下方的艺术字形状。
/// 各引擎差异大，失败即报错由上层转成 warning。
fn set_watermark(doc: &Document, spec: &WatermarkSpec) -> Result<String> {
    let text = spec.text.trim();
    if text.is_empty() {
        return Ok("水印文本为空，已跳过".into());
    }
    let header = sections(doc)?
        .call("Item", vec![ComValue::Int(1)])?
        .idispatch_child("Section")?
        .get_obj("Headers", "Headers")?
        .call("Item", vec![ComValue::Int(1)])?
        .idispatch_child("HeaderFooter")?;
    let shapes = header.get_obj("Shapes", "Shapes")?;
    let shape = shapes.call_named(
        "AddTextEffect",
        vec![],
        &[
            ("PresetTextEffect", ComValue::Int(0)),
            ("Text", ComValue::Str(text.to_string())),
            (
                "FontName",
                ComValue::Str(spec.font.clone().unwrap_or_else(|| "宋体".into())),
            ),
            ("FontSize", ComValue::Float(spec.size_pt.unwrap_or(48.0) as f64)),
            ("FontBold", ComValue::Bool(false)),
            ("FontItalic", ComValue::Bool(false)),
            ("Left", ComValue::Float(0.0)),
            ("Top", ComValue::Float(0.0)),
        ],
    )?;
    let shape = shape.idispatch_child("Shape")?;
    // 浮于文字下方、相对页边距居中
    if let Ok(wrap) = shape.get_obj("WrapFormat", "WrapFormat") {
        let _ = wrap.put("Type", vec![ComValue::Int(3)]);
    }
    // 衬于文字下方（msoSendBehindText=5）：只设 WrapType=None 形状仍会浮在正文上方，把字盖住
    let _ = shape.call("ZOrder", vec![ComValue::Int(5)]);
    let _ = shape.put("RelativeHorizontalPosition", vec![ComValue::Int(0)]);
    let _ = shape.put("RelativeVerticalPosition", vec![ComValue::Int(0)]);
    let _ = shape.put("Left", vec![ComValue::Float(-999995.0)]);
    let _ = shape.put("Top", vec![ComValue::Float(-999995.0)]);
    if let Some(rotation) = spec.rotation_deg {
        let _ = shape.put("Rotation", vec![ComValue::Float(rotation as f64)]);
    }
    if let Some(fill) = shape.get_obj("Fill", "Fill").ok() {
        if let Some(color) = spec.color.as_deref().and_then(|c| hex_to_ole(c).ok()) {
            if let Ok(fore) = fill.get_obj("ForeColor", "ForeColor") {
                let _ = fore.put("RGB", vec![ComValue::Int(color as i32)]);
            }
        }
        let transparency = spec.transparency_percent.unwrap_or(70.0).clamp(0.0, 100.0) / 100.0;
        let _ = fill.put("Transparency", vec![ComValue::Float(transparency as f64)]);
    }
    Ok("已添加文字水印".into())
}

fn update_fields(doc: &Document) -> Result<()> {
    // 主 story 字段。
    doc.raw()
        .get_obj("Fields", "Fields")?
        .call("Update", vec![])?;
    // 各节页眉页脚字段（不同 story）。
    let secs = sections(doc)?;
    let count = secs.get("Count")?.int().unwrap_or(0);
    for i in 1..=count {
        let sec = secs
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child("Section")?;
        for coll in ["Headers", "Footers"] {
            for v in 1..=3 {
                if let Ok(hf) = sec
                    .get_obj(coll, coll)
                    .and_then(|c| c.call("Item", vec![ComValue::Int(v)]))
                    .and_then(|h| h.idispatch_child("HeaderFooter"))
                {
                    if let Ok(r) = hf.get_obj("Range", "HeaderFooter.Range") {
                        let _ = r
                            .get_obj("Fields", "Fields")
                            .and_then(|f| f.call("Update", vec![]));
                    }
                }
            }
        }
    }
    Ok(())
}
