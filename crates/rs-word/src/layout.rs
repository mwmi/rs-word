//! 页面布局：PageSetup（按 Section）、页眉页脚、页码、字段更新、PDF 导出。

use anyhow::Result;

use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::types::{HfKind, LayoutAction, LayoutRequest, LayoutResult, PageSetupPatch};

/// mm → points（72 pt = 1 inch）。
fn mm_to_pt(mm: f32) -> f64 {
    mm as f64 * 72.0 / 25.4
}

/// 布局动作执行入口。
pub fn exec(
    _app: &App,
    doc: &Document,
    req: &LayoutRequest,
    doc_version: u64,
) -> Result<LayoutResult> {
    let mut out = LayoutResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        applied: Vec::new(),
        exported: None,
    };
    match req.action {
        LayoutAction::SetPageSetup => {
            let page = req
                .page
                .as_ref()
                .ok_or_else(|| anyhow::anyhow!("set_page_setup 需要 page 参数"))?;
            let n = apply_page_setup(doc, page)?;
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
            insert_page_number(doc, hf.section)?;
            out.applied.push("page_number".into());
            out.doc_version = doc_version + 1;
        }
        LayoutAction::UpdateFields => {
            update_fields(doc)?;
            out.applied.push("update_fields".into());
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
fn apply_page_setup(doc: &Document, page: &PageSetupPatch) -> Result<i32> {
    let secs = sections(doc)?;
    let count = secs.get("Count")?.int()?;
    for i in 1..=count {
        let ps = secs
            .call("Item", vec![ComValue::Int(i)])?
            .idispatch_child("Section")?
            .get_obj("PageSetup", "PageSetup")?;
        if let Some(v) = page.size {
            ps.put("PaperSize", vec![ComValue::Int(v.code())])?;
        }
        if let Some(v) = page.orientation {
            ps.put("Orientation", vec![ComValue::Int(v.code())])?;
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
    }
    Ok(count)
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

fn insert_page_number(doc: &Document, section: u32) -> Result<()> {
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
