//! 媒体：行内/浮动图片、自动形状、文本框、WordArt（10 手册 §8）。
//!
//! 约定：
//! - 几何统一 points；`inline=false` 走 `Range.InlineShapes`，`inline=true` 走 `Document.Shapes`。
//! - 对象标识 `shape_<n>` / `inline_<n>`（无会话状态）；写入前核对 `Name`（`expect_name`）与 `Type`，
//!   不符返回 `stale_target`。
//! - WordArt（`Shapes.AddTextEffect`）缺失时不降级为普通文本/形状，直接 `unsupported_capability`。

use anyhow::Result;

use crate::anchors;
use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::target::resolve;
use crate::types::{AutoShapeKind, MediaAction, MediaItem, MediaRequest, MediaResult};

/// msoTextOrientationHorizontal。
const TEXT_ORIENTATION_HORIZONTAL: i32 = 1;

/// 媒体动作执行入口。
pub fn exec(
    app: &App,
    doc: &Document,
    req: &MediaRequest,
    doc_version: u64,
) -> Result<MediaResult> {
    let mut out = MediaResult {
        document_id: doc.id().to_string(),
        operation_id: crate::new_operation_id(),
        doc_version,
        media_id: None,
        items: Vec::new(),
        applied: Vec::new(),
    };
    match req.action {
        MediaAction::List => {
            out.items = list(doc, req.max_items as usize)?;
        }
        MediaAction::InsertPicture => {
            let (id, item) = insert_picture(app, doc, req)?;
            out.media_id = Some(id);
            out.items.push(item);
            out.applied.push("insert_picture".into());
            out.doc_version = doc_version + 1;
        }
        MediaAction::InsertShape => {
            let (id, item) = insert_shape(doc, req)?;
            out.media_id = Some(id);
            out.items.push(item);
            out.applied.push("insert_shape".into());
            out.doc_version = doc_version + 1;
        }
        MediaAction::InsertTextbox => {
            let (id, item) = insert_textbox(doc, req)?;
            out.media_id = Some(id);
            out.items.push(item);
            out.applied.push("insert_textbox".into());
            out.doc_version = doc_version + 1;
        }
        MediaAction::InsertWordart => {
            let (id, item) = insert_wordart(app, doc, req)?;
            out.media_id = Some(id);
            out.items.push(item);
            out.applied.push("insert_wordart".into());
            out.doc_version = doc_version + 1;
        }
        MediaAction::Resize => {
            let s = locate(doc, req)?;
            let w = req.width_pt;
            let h = req.height_pt;
            if w.is_none() && h.is_none() {
                return Err(validate_err(doc, "resize 需要 width_pt 或 height_pt"));
            }
            if let Some(w) = w {
                check_pt(doc, "width_pt", w, 1000.0)?;
                s.put("LockAspectRatio", vec![ComValue::Bool(false)])?;
                s.put("Width", vec![ComValue::Float(w as f64)])?;
            }
            if let Some(h) = h {
                check_pt(doc, "height_pt", h, 1000.0)?;
                s.put("LockAspectRatio", vec![ComValue::Bool(false)])?;
                s.put("Height", vec![ComValue::Float(h as f64)])?;
            }
            ok_read_back(&mut out, &s, "resize");
            out.doc_version = doc_version + 1;
        }
        MediaAction::SetPosition => {
            let s = locate(doc, req)?;
            if req.left_pt.is_none() && req.top_pt.is_none() {
                return Err(validate_err(doc, "set_position 需要 left_pt 或 top_pt"));
            }
            if let Some(v) = req.left_pt {
                s.put("Left", vec![ComValue::Float(v as f64)])?;
            }
            if let Some(v) = req.top_pt {
                s.put("Top", vec![ComValue::Float(v as f64)])?;
            }
            ok_read_back(&mut out, &s, "set_position");
            out.doc_version = doc_version + 1;
        }
        MediaAction::SetWrap => {
            let wrap = req
                .wrap
                .ok_or_else(|| validate_err(doc, "set_wrap 需要 wrap"))?;
            let s = locate(doc, req)?;
            s.get_obj("WrapFormat", "WrapFormat")?
                .put("Type", vec![ComValue::Int(wrap.code())])?;
            let got = s
                .get_obj("WrapFormat", "WrapFormat")?
                .get("Type")
                .and_then(|v| v.int())
                .unwrap_or(-1);
            out.applied.push(format!("set_wrap {:?} -> {got}", wrap));
            out.doc_version = doc_version + 1;
        }
        MediaAction::SetAltText => {
            let text = req
                .alt_text
                .as_ref()
                .ok_or_else(|| validate_err(doc, "set_alt_text 需要 alt_text"))?;
            let s = locate(doc, req)?;
            s.put("AlternativeText", vec![ComValue::Str(text.clone())])?;
            out.applied
                .push(format!("set_alt_text {}", anchors::quote(text, 40)));
            out.doc_version = doc_version + 1;
        }
        MediaAction::ZOrder => {
            let cmd = req
                .z_order
                .ok_or_else(|| validate_err(doc, "z_order 需要 z_order"))?;
            let s = locate(doc, req)?;
            s.call("ZOrder", vec![ComValue::Int(cmd.code())])?;
            out.applied.push(format!("z_order {:?}", cmd));
            out.doc_version = doc_version + 1;
        }
        MediaAction::Delete => {
            let s = locate(doc, req)?;
            s.call("Delete", vec![])?;
            out.applied.push("delete".into());
            out.doc_version = doc_version + 1;
        }
    }
    Ok(out)
}

fn validate_err(doc: &Document, msg: impl Into<String>) -> anyhow::Error {
    WordError::new("media", doc.label(), msg)
        .with_hint("检查 media 动作的必填字段")
        .into()
}

fn check_pt(doc: &Document, name: &str, v: f32, max: f32) -> Result<()> {
    if !v.is_finite() || !(1.0..=max).contains(&v) {
        return Err(validate_err(
            doc,
            format!("{name}={v} 越界（允许 1.0..={max} pt）"),
        ));
    }
    Ok(())
}

/// 顶部集合名（`Shapes` / `InlineShapes`）。
fn collection(doc: &Document, inline: bool) -> Result<IDispatchWrapper> {
    let name = if inline { "InlineShapes" } else { "Shapes" };
    doc.raw().get_obj(name, name)
}

/// 解析 `media_id` 为 `(inline, index)`。
fn parse_id(doc: &Document, id: &str) -> Result<(bool, i32)> {
    if let Some(n) = id.strip_prefix("shape_") {
        let i: i32 = n
            .parse()
            .map_err(|_| validate_err(doc, format!("media_id '{id}' 非法")))?;
        return Ok((false, i));
    }
    if let Some(n) = id.strip_prefix("inline_") {
        let i: i32 = n
            .parse()
            .map_err(|_| validate_err(doc, format!("media_id '{id}' 非法")))?;
        return Ok((true, i));
    }
    Err(validate_err(
        doc,
        format!("media_id '{id}' 非法，应为 shape_<n> / inline_<n>"),
    ))
}

/// 定位媒体对象并按 `expect_name` 核对。
fn locate(doc: &Document, req: &MediaRequest) -> Result<IDispatchWrapper> {
    let id = req
        .media_id
        .as_deref()
        .ok_or_else(|| validate_err(doc, "该动作需要 media_id（先 list）"))?;
    let (inline, idx) = parse_id(doc, id)?;
    let obj = collection(doc, inline)?
        .call("Item", vec![ComValue::Int(idx)])
        .map_err(|e| WordError::hint(e, "media_id 已过期（对象已删除或序号变化）：重新 list"))?
        .idispatch_child(id)?;
    if let Some(expect) = &req.expect_name {
        let name = obj.get("Name").and_then(|v| v.string()).unwrap_or_default();
        if &name != expect {
            return Err(WordError::new(
                "media",
                obj.ctx.clone(),
                format!("对象名不符（期望 {expect}，实际 {name}）"),
            )
            .with_hint("对象已变化：重新 media.list 取最新 media_id")
            .into());
        }
    }
    Ok(obj)
}

/// 读取对象几何与名称；对象类型码区分形状类别。
fn read_item(obj: &IDispatchWrapper, media_id: &str, inline: bool) -> Result<MediaItem> {
    let kind = classify(obj, inline);
    let name = if inline {
        media_id.to_string()
    } else {
        obj.get("Name").and_then(|v| v.string()).unwrap_or_default()
    };
    let getf = |p: &str| obj.get(p).and_then(|v| v.float()).unwrap_or(0.0) as f32;
    let alt = obj.get("AlternativeText").and_then(|v| v.string()).ok();
    let text = obj
        .get_obj("TextFrame", "TextFrame")
        .and_then(|tf| tf.get_obj("TextRange", "TextFrame.TextRange"))
        .and_then(|tr| tr.get("Text"))
        .and_then(|v| v.string())
        .ok()
        .filter(|s| !s.is_empty());
    Ok(MediaItem {
        media_id: media_id.to_string(),
        kind,
        name,
        left_pt: getf("Left"),
        top_pt: getf("Top"),
        width_pt: getf("Width"),
        height_pt: getf("Height"),
        alt_text: alt.filter(|s| !s.is_empty()),
        quote: text.map(|s| anchors::quote(&s, 60)),
    })
}

/// `Shape.Type` / `InlineShape.Type` → 可读类别（msoShapeType* / wdInlineShapeType*）。
fn classify(obj: &IDispatchWrapper, inline: bool) -> String {
    let t = obj.get("Type").and_then(|v| v.int()).unwrap_or(-1);
    if inline {
        return match t {
            3 | 4 => "inline_picture".into(),
            _ => format!("inline_shape_{t}"),
        };
    }
    match t {
        13 | 11 => "picture".into(),
        17 => "textbox".into(),
        15 => "wordart".into(),
        _ => format!("shape_{t}"),
    }
}

/// 枚举媒体列表。
pub fn list(doc: &Document, max_items: usize) -> Result<Vec<MediaItem>> {
    let mut out = Vec::new();
    for (inline, coll) in [(false, "Shapes"), (true, "InlineShapes")] {
        let Ok(c) = doc.raw().get_obj(coll, coll) else {
            continue;
        };
        let count = c.get("Count").and_then(|v| v.int()).unwrap_or(0);
        let prefix = if inline { "inline" } else { "shape" };
        for i in 1..=count.min(max_items as i32) {
            let obj = c
                .call("Item", vec![ComValue::Int(i)])?
                .idispatch_child(&format!("{coll}[{i}]"))?;
            let id = format!("{prefix}_{i}");
            out.push(read_item(&obj, &id, inline)?);
        }
    }
    Ok(out)
}

fn insert_picture(app: &App, doc: &Document, req: &MediaRequest) -> Result<(String, MediaItem)> {
    let path = req
        .path
        .as_deref()
        .ok_or_else(|| validate_err(doc, "insert_picture 需要 path"))?;
    if !std::path::Path::new(path).exists() {
        return Err(validate_err(doc, format!("图片文件不存在: {path}")));
    }
    let full = crate::process::abs_path(path)?;
    if let Some(w) = req.width_pt {
        check_pt(doc, "width_pt", w, 1000.0)?;
    }
    if let Some(h) = req.height_pt {
        check_pt(doc, "height_pt", h, 1000.0)?;
    }
    if req.inline {
        // 行内图：Range.InlineShapes.AddPicture(FileName, LinkToFile, SaveWithDocument, Range)。
        let target = req
            .target
            .as_ref()
            .ok_or_else(|| validate_err(doc, "行内图片需要 target 定位插入点"))?;
        let at = resolve(app, doc, target)?.range.collapse_start()?;
        let pic = at
            .raw()
            .get_obj("InlineShapes", "InlineShapes")?
            .call_named(
                "AddPicture",
                vec![],
                &[
                    ("FileName", ComValue::Str(full)),
                    ("LinkToFile", ComValue::Bool(false)),
                    ("SaveWithDocument", ComValue::Bool(true)),
                    ("Range", ComValue::Obj(at.raw().disp.clone())),
                ],
            )?
            .idispatch_child("InlineShape")?;
        pic.put("LockAspectRatio", vec![ComValue::Bool(true)])?;
        if let Some(w) = req.width_pt {
            pic.put("Width", vec![ComValue::Float(w as f64)])?;
        }
        let idx = doc
            .raw()
            .get_obj("InlineShapes", "InlineShapes")?
            .get("Count")
            .and_then(|v| v.int())
            .unwrap_or(1);
        let id = format!("inline_{idx}");
        let item = read_item(&pic, &id, true)?;
        return Ok((id, item));
    }
    // 浮动图：Shapes.AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height, Anchor)。
    let mut named: Vec<(&str, ComValue)> = vec![
        ("FileName", ComValue::Str(full)),
        ("LinkToFile", ComValue::Bool(false)),
        ("SaveWithDocument", ComValue::Bool(true)),
        ("Left", ComValue::Float(req.left_pt.unwrap_or(0.0) as f64)),
        ("Top", ComValue::Float(req.top_pt.unwrap_or(0.0) as f64)),
        (
            "Width",
            ComValue::Float(req.width_pt.unwrap_or(120.0) as f64),
        ),
        (
            "Height",
            ComValue::Float(req.height_pt.unwrap_or(90.0) as f64),
        ),
    ];
    if let Some(target) = &req.target {
        let at = resolve(app, doc, target)?.range;
        named.push(("Anchor", ComValue::Obj(at.raw().disp.clone())));
    }
    let shp = doc
        .raw()
        .get_obj("Shapes", "Shapes")?
        .call_named("AddPicture", vec![], &named)?
        .idispatch_child("Shape")?;
    finish_shape(doc, shp, req)
}

fn insert_shape(doc: &Document, req: &MediaRequest) -> Result<(String, MediaItem)> {
    let kind: AutoShapeKind = req
        .shape
        .ok_or_else(|| validate_err(doc, "insert_shape 需要 shape"))?;
    let shp = add_shape(doc, req, |shapes| {
        shapes.call_named(
            "AddShape",
            vec![],
            &[
                ("Type", ComValue::Int(kind.code())),
                ("Left", ComValue::Float(req.left_pt.unwrap_or(0.0) as f64)),
                ("Top", ComValue::Float(req.top_pt.unwrap_or(0.0) as f64)),
                (
                    "Width",
                    ComValue::Float(req.width_pt.unwrap_or(160.0) as f64),
                ),
                (
                    "Height",
                    ComValue::Float(req.height_pt.unwrap_or(60.0) as f64),
                ),
            ],
        )
    })?;
    apply_text(&shp, req)?;
    finish_shape(doc, shp, req)
}

fn insert_textbox(doc: &Document, req: &MediaRequest) -> Result<(String, MediaItem)> {
    let shp = add_shape(doc, req, |shapes| {
        shapes.call_named(
            "AddTextbox",
            vec![],
            &[
                ("Orientation", ComValue::Int(TEXT_ORIENTATION_HORIZONTAL)),
                ("Left", ComValue::Float(req.left_pt.unwrap_or(0.0) as f64)),
                ("Top", ComValue::Float(req.top_pt.unwrap_or(0.0) as f64)),
                (
                    "Width",
                    ComValue::Float(req.width_pt.unwrap_or(200.0) as f64),
                ),
                (
                    "Height",
                    ComValue::Float(req.height_pt.unwrap_or(80.0) as f64),
                ),
            ],
        )
    })?;
    apply_text(&shp, req)?;
    finish_shape(doc, shp, req)
}

fn insert_wordart(app: &App, doc: &Document, req: &MediaRequest) -> Result<(String, MediaItem)> {
    let text = req
        .text
        .as_ref()
        .ok_or_else(|| validate_err(doc, "insert_wordart 需要 text"))?;
    let font = req.font_name.clone().unwrap_or_else(|| "Arial".to_string());
    let size = req.font_size_pt.unwrap_or(36.0);
    let mut named: Vec<(&str, ComValue)> = vec![
        (
            "PresetTextEffect",
            ComValue::Int(req.preset_text_effect.unwrap_or(0) as i32),
        ),
        ("Text", ComValue::Str(text.clone())),
        ("FontName", ComValue::Str(font)),
        ("FontSize", ComValue::Float(size as f64)),
        ("FontBold", ComValue::Bool(req.bold)),
        ("FontItalic", ComValue::Bool(req.italic)),
        ("Left", ComValue::Float(req.left_pt.unwrap_or(0.0) as f64)),
        ("Top", ComValue::Float(req.top_pt.unwrap_or(0.0) as f64)),
    ];
    if let Some(target) = &req.target {
        let at = resolve(app, doc, target)?.range;
        named.push(("Anchor", ComValue::Obj(at.raw().disp.clone())));
    }
    let shapes = doc.raw().get_obj("Shapes", "Shapes")?;
    let shp = shapes
        .call_named("AddTextEffect", vec![], &named)
        .map_err(|e| {
            WordError::hint(
                e,
                "该引擎/版本不支持 WordArt（AddTextEffect）：改用普通形状+文本，不静默降级",
            )
        })?
        .idispatch_child("Shape")?;
    finish_shape(doc, shp, req)
}

/// `Document.Shapes` 上的创建（统一取出 Shapes 集合 + 错误美化）。
fn add_shape(
    doc: &Document,
    _req: &MediaRequest,
    f: impl FnOnce(&IDispatchWrapper) -> Result<crate::com::dispatch::DispatchResult>,
) -> Result<IDispatchWrapper> {
    let shapes = doc.raw().get_obj("Shapes", "Shapes")?;
    f(&shapes)?
        .idispatch_child("Shape")
        .map_err(|e| WordError::hint(e, "该引擎可能不支持该形状类型"))
}

/// 形状后面的收尾：文本、环绕、替代文本、编号与回读。
fn finish_shape(
    doc: &Document,
    shp: IDispatchWrapper,
    req: &MediaRequest,
) -> Result<(String, MediaItem)> {
    if let Some(wrap) = req.wrap {
        shp.get_obj("WrapFormat", "WrapFormat")?
            .put("Type", vec![ComValue::Int(wrap.code())])?;
    }
    if let Some(alt) = &req.alt_text {
        shp.put("AlternativeText", vec![ComValue::Str(alt.clone())])?;
    }
    let idx = doc
        .raw()
        .get_obj("Shapes", "Shapes")?
        .get("Count")
        .and_then(|v| v.int())
        .unwrap_or(1);
    let id = format!("shape_{idx}");
    let item = read_item(&shp, &id, false)?;
    Ok((id, item))
}

/// 写形状/文本框文本 + 字体。
fn apply_text(shp: &IDispatchWrapper, req: &MediaRequest) -> Result<()> {
    let Some(text) = &req.text else {
        return Ok(());
    };
    let tr = shp
        .get_obj("TextFrame", "TextFrame")?
        .get_obj("TextRange", "TextFrame.TextRange")?;
    tr.put("Text", vec![ComValue::Str(text.clone())])?;
    let font = tr.get_obj("Font", "Font")?;
    if let Some(name) = &req.font_name {
        font.put("Name", vec![ComValue::Str(name.clone())])?;
    }
    if let Some(size) = req.font_size_pt {
        font.put("Size", vec![ComValue::Float(size as f64)])?;
    }
    if req.bold {
        font.put("Bold", vec![ComValue::Int(-1)])?;
    }
    if req.italic {
        font.put("Italic", vec![ComValue::Int(-1)])?;
    }
    Ok(())
}

/// 写后回读几何/位置作为验证证据。
fn ok_read_back(out: &mut MediaResult, obj: &IDispatchWrapper, action: &str) {
    let w = obj.get("Width").and_then(|v| v.float()).unwrap_or(0.0);
    let h = obj.get("Height").and_then(|v| v.float()).unwrap_or(0.0);
    let l = obj.get("Left").and_then(|v| v.float()).unwrap_or(0.0);
    let t = obj.get("Top").and_then(|v| v.float()).unwrap_or(0.0);
    out.applied
        .push(format!("{action} -> {l:.1},{t:.1} {w:.1}x{h:.1}pt"));
}
