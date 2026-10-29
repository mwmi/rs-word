//! LaTeX → Word 原生公式（`$…$` 行内 / `$$…$$` 独立）。
//!
//! 只做「搬运 + 调引擎」，语义对齐两份 VBA 宏：
//! - **Microsoft Word**（excel.vbs）：去定界符 → 写入 LaTeX → `OMaths.Add` → `BuildUp` → 设 `Type`；
//! - **WPS Writer**（wps.vbs）：`OMaths.Add` 后选中公式发 `Ctrl+=`，触发 WPS 的「LaTeX 转专业格式」再设 `Type`
//!   （WPS 的 `BuildUp` 对 LaTeX 文本不生效）。
//!
//! Word 的 `BuildUp` 实测只认 UnicodeMath：首次建出来仍是 LaTeX 源码时，自动换 [`to_unicodemath`] 的语法重建一次。
//!
//! 转换不成时：**WPS 保留公式框**（LaTeX 留在里面，选中按 `Ctrl+=` 或点「转换」即可），
//! **Word 回落「去掉 `$` 的 LaTeX 原文」**；两种都不丢字、不中断整篇写入。
//! 建公式入口各走各的 VBA 原序（Word 文档级 / WPS 选区级），建完都会核对不是空公式。

use anyhow::Result;

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::range::Range;
use crate::types::WordEngine;

/// `WdOMathType`：独立公式 / 行内公式。
const DISPLAY: i32 = 0;
const INLINE: i32 = 1;
/// WPS 发完 `Ctrl+=` 后等它把 LaTeX 转成专业格式的上限（超时就按文本兜底，绝不无限等）。
const KEY_WAIT_MS: u64 = 2500;
/// 等公式对象真正建好的上限（WPS 的 `OMaths.Add` 不是同步完成，太早按键会被丢掉）。
const READY_WAIT_MS: u64 = 800;
/// 跨进程操作之间留的缓冲：选区/前台切换要 WPS 的消息循环处理完，按键才认。
const SETTLE_MS: u64 = 40;

/// 全文转换统计（`notes` 只留前几条，避免一大串重复提示刷屏）。
#[derive(Debug, Default, Clone)]
pub struct Converted {
    pub display: usize,
    pub inline: usize,
    pub failed: usize,
    pub notes: Vec<String>,
}

/// 去掉 `$`/`$$` 定界符与换行，得到可写进正文的 LaTeX 原文（兜底文本、`math=false` 都用它）。
pub fn plain_latex(src: &str) -> String {
    let raw = src.trim();
    let body = raw
        .strip_prefix("$$")
        .and_then(|text| text.strip_suffix("$$"))
        .or_else(|| raw.strip_prefix('$').and_then(|text| text.strip_suffix('$')))
        .unwrap_or(raw);
    let chars: Vec<char> = body.chars().collect();
    let mut out = String::with_capacity(body.len());
    let mut space = false;
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        // LaTeX 的换行命令 `\\` 在公式里就是空格
        if ch == '\\' && chars.get(i + 1) == Some(&'\\') {
            space = true;
            i += 2;
            continue;
        }
        // `\ neq` 是模型躲 JSON 转义的写法（`\n` 会被 JSON 吃掉）：反斜杠后隔着空格接命令名时还原成 `\neq`
        if ch == '\\' {
            let mut next = i + 1;
            while chars.get(next).is_some_and(|c| c.is_whitespace()) {
                next += 1;
            }
            if next > i + 1 && chars.get(next).is_some_and(|c| c.is_ascii_alphabetic()) {
                out.push('\\');
                i = next;
                continue;
            }
        }
        if ch.is_whitespace() {
            space = true;
            i += 1;
            continue;
        }
        if space && !out.is_empty() {
            out.push(' ');
        }
        space = false;
        out.push(ch);
        i += 1;
    }
    out.trim().to_string()
}

/// 把 `[start, start+len)` 里**已经写好**的 `$…$` / `$$…$$` 原文换成公式。
///
/// 写入阶段只落原文，等这段内容全部写完后按记录的位置**倒序**调用本函数：
/// 公式对象长度与字符数不同，先转后面的再转前面的，位置才不会漂。
/// 失败不返回 Err：Word 落「去掉 `$` 的 LaTeX 原文」、WPS 保留公式框，只给一条提示。
pub fn convert_at(
    app: &App,
    doc: &Document,
    start: i32,
    len: i32,
    src: &str,
    display: bool,
) -> (i32, Option<String>) {
    let latex = plain_latex(src);
    if latex.is_empty() {
        return (start + len.max(0), None);
    }
    let range = match doc.range(start, start + len.max(0)) {
        Ok(range) => range,
        Err(e) => return (start + len.max(0), Some(format!("公式位置已失效：{e}"))),
    };
    // 位置核对：这一段必须还是当初写进去的 `$…$` 原文，否则宁可不转，也不能改错地方
    if !is_delimited(&range.text().unwrap_or_default()) {
        return (
            start + len.max(0),
            Some("公式位置与正文对不上（前后有别的内容插进来了），已跳过".into()),
        );
    }
    if let Err(e) = range.set_text(&latex) {
        return (start + len.max(0), Some(format!("公式未能写入：{e}")));
    }
    let end = range
        .end()
        .unwrap_or(start + latex.encode_utf16().count() as i32);
    build(app, doc, start, end, &latex, display)
}

/// 把范围内**已有的** `$$…$$` / `$…$` 一次性转成公式（等价两个 VBS 宏，先独立后行内）。
///
/// 匹配以原始文本为准，**倒序**替换：公式对象长度与字符数不同，正序会让后面的位置全漂。
pub fn convert_text(app: &App, doc: &Document, start: i32, text: &str) -> Converted {
    let mut out = Converted::default();
    let hits = scan(text);
    if hits.is_empty() {
        return out;
    }
    let _screen = ScreenGuard::off(app);
    let equations = equation_spans(doc);
    let mut notes: Vec<String> = Vec::new();
    for hit in hits.into_iter().rev() {
        let from = start + hit.from as i32;
        let to = start + hit.to as i32;
        // 已有公式内部不算定界符（公式里可能就有 $）
        if equations.iter().any(|(s, e)| from < *e && to > *s) {
            continue;
        }
        let built = doc.range(from, to).and_then(|range| {
            // 位置核对：文档里已有公式/域时字符流与位置可能漂，宁可不转，也不能改错地方
            if !is_delimited(&range.text().unwrap_or_default()) {
                return Err(anyhow::anyhow!("位置核对未通过（文档里的公式/域会让字符流与位置错位），已跳过"));
            }
            range.set_text(&hit.latex)?;
            let end = range.end().unwrap_or(from + hit.latex.encode_utf16().count() as i32);
            Ok(build(app, doc, from, end, &hit.latex, hit.display))
        });
        match built {
            Ok((_, Some(warning))) => {
                out.failed += 1;
                if notes.len() < 3 {
                    notes.push(warning);
                }
            }
            Ok((_, None)) => {
                if hit.display {
                    out.display += 1;
                } else {
                    out.inline += 1;
                }
            }
            Err(e) => {
                out.failed += 1;
                if notes.len() < 3 {
                    notes.push(format!("公式未能写入：{e}"));
                }
            }
        }
    }
    out.notes.extend(notes);
    out
}

// ───────────────────── 建公式（两套引擎共用） ─────────────────────

/// 在 `[start,end)` 已是 `latex` 文本的前提下建公式，返回结束位置 + 失败提示。
///
/// 失败口径按引擎不同：**WPS 保留公式框**（LaTeX 留在里面，用户自己点「转换」）；
/// **Word 回落文本**（去掉 `$` 的 LaTeX 原文）。
fn build(
    app: &App,
    doc: &Document,
    start: i32,
    end: i32,
    latex: &str,
    display: bool,
) -> (i32, Option<String>) {
    // 旧版 .doc 不支持 OMML：正文已是纯 LaTeX，直接当文本留着，不白跑 COM
    if app.engine() == WordEngine::MsWord && compat_mode(doc) {
        return (end, None);
    }
    let mut current = latex.to_string();
    let mut end = end;
    let mut reason = String::new();
    let mut built = false;
    for attempt in 0..2 {
        match equation(app, doc, start, end, display, &current) {
            Ok((next, why)) => {
                end = next;
                built = true;
                match why {
                    None => return (end, None),
                    Some(why) => reason = why,
                }
            }
            Err(e) => reason = e.to_string(),
        }
        // Word：第一次按 VBA 直接喂 LaTeX；失败后换 UnicodeMath 语法（BuildUp 只认它）再试一次
        if app.engine() != WordEngine::MsWord || attempt == 1 {
            break;
        }
        let fixed = to_unicodemath(&current);
        if fixed.is_empty() || fixed == current {
            break;
        }
        match reset(doc, start, end, &fixed) {
            Ok(next) => {
                current = fixed;
                end = next;
            }
            Err(e) => {
                reason = e.to_string();
                break;
            }
        }
    }
    // WPS：转换不成也别动它 —— LaTeX 留在公式框里，人自己点「转换」/按 Ctrl+= 即可
    if app.engine() == WordEngine::WpsWriter {
        return (
            end,
            Some(if built {
                format!("公式已保留为 LaTeX（WPS Writer）：{reason}。选中该公式按 Ctrl+= 或点「转换」即可")
            } else {
                format!("公式未建成公式对象（WPS Writer）：{reason}。LaTeX 原文已留在正文")
            }),
        );
    }
    // Word：兜底留「去掉 $ 的 LaTeX 原文」，绝不把半成品公式或源码留在公式框里
    let end = reset(doc, start, end, latex).unwrap_or(end);
    (end, Some(format!("公式未能转换，已按文本写入：{reason}")))
}

/// 建一次公式对象；返回 (公式结束位置, 未转换原因)。`None` = 已转成专业格式。
fn equation(
    app: &App,
    doc: &Document,
    start: i32,
    end: i32,
    display: bool,
    source: &str,
) -> Result<(i32, Option<String>)> {
    let range = doc.range(start, end)?;
    let omath = add_omath(app, doc, &range)?;
    let mut why: Option<String> = None;
    if app.engine() == WordEngine::MsWord {
        if omath.has("BuildUp") {
            let _ = omath.call("BuildUp", vec![]);
        }
        let body = omath.get_obj("Range", "OMath.Range")?;
        if still_source(&text_of(&body), source) {
            why = Some("Word 的 BuildUp 未把 LaTeX 转成专业格式".into());
        }
    } else {
        // wps.vbs 第 76-77 行：等公式建好 → 切前台 → 选中公式 → 发一次 Ctrl+=
        let body = omath.get_obj("Range", "OMath.Range")?;
        wait_ready(&body);
        match focus(app) {
            Ok(()) => {
                let _ = body.call("Select", vec![]);
                std::thread::sleep(std::time::Duration::from_millis(SETTLE_MS));
                // 选区没落在公式上（WPS 忙的时候会吞掉 Select）就再选一次，否则按键作用在别处
                if !selection_holds_math(app) {
                    let _ = body.call("Select", vec![]);
                    std::thread::sleep(std::time::Duration::from_millis(SETTLE_MS));
                }
                press_ctrl_equals();
                wait_professional(&body, source, KEY_WAIT_MS);
                if still_source(&text_of(&body), source) {
                    why = Some("已发送 Ctrl+=，但公式仍是 LaTeX".into());
                }
            }
            Err(reason) => why = Some(format!("{reason}，未发送 Ctrl+=")),
        }
        // WPS 的 BuildUp 多半是空实现；按键没生效时再试一次（有该成员才调，不报错）
        if why.is_some() && omath.has("BuildUp") {
            let _ = omath.call("BuildUp", vec![]);
            wait_professional(&body, source, 300);
            if !still_source(&text_of(&body), source) {
                why = None;
            }
        }
    }
    let _ = omath.put(
        "Type",
        vec![ComValue::Int(if display { DISPLAY } else { INLINE })],
    );
    let body = omath.get_obj("Range", "OMath.Range")?;
    let end = body.get("End")?.int()?;
    Ok((end, why))
}

/// 建公式对象（返回 OMath）：按各引擎 VBA 的原生写法来。
/// - **Microsoft Word**（excel.vbs）：`Document.OMaths.Add(Range)`，不需要选区、不抢焦点；
/// - **WPS Writer**（wps.vbs）：`Range.Select()` → `Selection.OMaths.Add(Selection.Range)`。
///
/// 建完当场核对公式里确实有内容（WPS 的「在此处键入公式」占位不算），空壳删掉换下一条路径。
fn add_omath(app: &App, doc: &Document, range: &Range) -> Result<IDispatchWrapper> {
    let mut errors: Vec<String> = Vec::new();
    for path in engine_paths(app.engine()) {
        let added = match path {
            AddPath::Document => document_add(doc, range),
            AddPath::Selection => selection_add(app, doc, range),
        };
        match added.and_then(as_omath) {
            Ok(omath) => {
                let Ok(body) = omath.get_obj("Range", "OMath.Range") else {
                    errors.push("拿不到 OMath.Range".into());
                    continue;
                };
                if !is_blank(&text_of(&body)) {
                    return Ok(omath);
                }
                let _ = body.call("Delete", vec![]);
                errors.push("OMaths.Add 只建出空公式".into());
            }
            Err(e) => errors.push(e.to_string()),
        }
    }
    anyhow::bail!("OMaths.Add 未建出公式（{}）", errors.join("；"))
}

enum AddPath {
    /// `Document.OMaths.Add(Range)`（excel.vbs）
    Document,
    /// `Range.Select` → `Selection.OMaths.Add`（wps.vbs）
    Selection,
}

/// 建公式路径：**严格照各引擎自己的 VBA**。
/// WPS 只走选区（文档级 Add 在 WPS 上不认 range，会多出空公式）；
/// Word 走文档级，选区路径留给老版本兜底。
fn engine_paths(engine: WordEngine) -> &'static [AddPath] {
    match engine {
        WordEngine::WpsWriter => &[AddPath::Selection],
        WordEngine::MsWord => &[AddPath::Document, AddPath::Selection],
    }
}

fn document_add(doc: &Document, range: &Range) -> Result<IDispatchWrapper> {
    omaths_add(
        &doc.raw().get_obj("OMaths", "Document.OMaths")?,
        range.raw(),
    )
}

/// VBA 原序：`rng.Select` → `Selection.OMaths.Add(Selection.Range)`。
fn selection_add(app: &App, doc: &Document, range: &Range) -> Result<IDispatchWrapper> {
    let _ = doc.raw().call("Activate", vec![]);
    range.raw().call("Select", vec![])?;
    let selection = app.raw().get_obj("Selection", "Selection")?;
    let target = selection.get_obj("Range", "Selection.Range")?;
    omaths_add(&selection.get_obj("OMaths", "Selection.OMaths")?, &target)
}

/// 在某个 `OMaths` 集合上 `Add`（返回对象形态由 [`as_omath`] 识别）。
fn omaths_add(omaths: &IDispatchWrapper, target: &IDispatchWrapper) -> Result<IDispatchWrapper> {
    omaths
        .call("Add", vec![ComValue::Obj(target.disp.clone())])
        .and_then(|value| value.idispatch_child("OMaths.Add"))
}

/// `OMaths.Add` 各引擎返回值不一：Range（取 `Range.OMaths(1)`）/ OMaths 集合（`Item(1)`）/ 直接给 OMath。
fn as_omath(added: IDispatchWrapper) -> Result<IDispatchWrapper> {
    if added.has("OMaths") {
        return added
            .get_obj("OMaths", "Range.OMaths")?
            .call("Item", vec![ComValue::Int(1)])
            .and_then(|value| value.idispatch_child("OMath"));
    }
    if added.has("BuildUp") || added.has("Type") {
        return Ok(added);
    }
    if added.has("Item") {
        return added
            .call("Item", vec![ComValue::Int(1)])
            .and_then(|value| value.idispatch_child("OMath"));
    }
    anyhow::bail!("OMaths.Add 返回的对象无法识别（既不是 Range 也不是 OMath）")
}

/// 是否还停留在 LaTeX 源码形态：源码里有命令、公式里也还留着反斜杠。
fn still_source(text: &str, source: &str) -> bool {
    source.contains('\\') && text.contains('\\')
}

/// 公式框是不是空的：WPS 空公式会塞「在此处键入公式。」这类占位文案，不能当成建成功。
fn is_blank(text: &str) -> bool {
    text.replace("在此处键入公式。", "")
        .replace("在此处键入公式", "")
        .trim()
        .is_empty()
}

/// 这段文字是否真的以 `$` / `$$` 收尾（全文转换时核对位置，防止改错地方）。
fn is_delimited(text: &str) -> bool {
    let text = text.trim();
    text.len() > 1 && text.starts_with('$') && text.ends_with('$')
}

fn text_of(object: &IDispatchWrapper) -> String {
    object
        .get("Text")
        .and_then(|value| value.string())
        .unwrap_or_default()
}

/// 用 `text` 覆盖 `[start,end)`（删掉失败的公式对象再写回纯文本），返回新的结束位置。
fn reset(doc: &Document, start: i32, end: i32, text: &str) -> Result<i32> {
    let width = text.encode_utf16().count() as i32;
    if end <= start {
        let at = doc.range(start, start)?;
        at.insert_after(text)?;
        return Ok(at.end().unwrap_or(start + width));
    }
    if let Ok(range) = doc.range(start, end) {
        if range.delete().is_ok() {
            let at = doc.range(start, start)?;
            at.insert_after(text)?;
            return Ok(at.end().unwrap_or(start + width));
        }
        range.set_text(text)?;
        return Ok(range.end().unwrap_or(start + width));
    }
    Ok(end)
}

/// Word 是否处于 .doc / 2003 兼容模式（`CompatibilityMode < 14`）：这种文档不支持 OMML。
fn compat_mode(doc: &Document) -> bool {
    doc.raw()
        .get("CompatibilityMode")
        .and_then(|value| value.int())
        .map(|mode| mode < 14)
        .unwrap_or(false)
}

// ───────────────────── 全文定界符扫描 ─────────────────────

/// 一处 `$…$` / `$$…$$`（`from`/`to` 是相对 `text` 的 UTF-16 偏移，与 Word 位置一致）。
struct Hit {
    from: usize,
    to: usize,
    latex: String,
    display: bool,
}

/// 按段落扫描：Word 的 Find 通配同样不跨段，`$$` 优先于 `$`，`\$` 是转义不是定界符。
fn scan(text: &str) -> Vec<Hit> {
    let mut hits = Vec::new();
    let mut base = 0usize;
    for para in text.split('\r') {
        scan_paragraph(para, base, &mut hits);
        base += para.encode_utf16().count() + 1;
    }
    hits
}

fn scan_paragraph(para: &str, base: usize, hits: &mut Vec<Hit>) {
    let chars: Vec<char> = para.chars().collect();
    let mut at = 0usize;
    let mut offset = 0usize;
    while at < chars.len() {
        let ch = chars[at];
        if ch == '\\' {
            // 反斜杠命令（含 `\$`）整体跳过，别把命令里的字符当定界符
            at += 1;
            offset += ch.len_utf16();
            if let Some(next) = chars.get(at) {
                at += 1;
                offset += next.len_utf16();
            }
            continue;
        }
        if ch != '$' {
            at += 1;
            offset += ch.len_utf16();
            continue;
        }
        let double = chars.get(at + 1) == Some(&'$');
        let open = if double { 2 } else { 1 };
        match close_at(&chars, at + open, double) {
            Some((close, close_len)) if close > at + open => {
                let latex: String = chars[at + open..close].iter().collect();
                let inner: usize = latex.chars().map(char::len_utf16).sum();
                let width = open + inner + close_len;
                let latex = latex.trim().to_string();
                if !latex.is_empty() {
                    hits.push(Hit {
                        from: base + offset,
                        to: base + offset + width,
                        latex,
                        display: double,
                    });
                }
                at += width;
                offset += width;
            }
            _ => {
                at += open;
                offset += open;
            }
        }
    }
}

/// 从 `from` 起找闭合定界符（内容不含 `$`，即 VBA 的 `[!$]@`）；`$$` 必须成对。
fn close_at(chars: &[char], from: usize, double: bool) -> Option<(usize, usize)> {
    let mut at = from;
    while at < chars.len() {
        if chars[at] == '$' {
            if double {
                return (chars.get(at + 1) == Some(&'$')).then_some((at, 2));
            }
            return Some((at, 1));
        }
        at += 1;
    }
    None
}

/// 已有公式区间：落在里面的匹配要跳过。
fn equation_spans(doc: &Document) -> Vec<(i32, i32)> {
    let Ok(content) = doc.content() else {
        return Vec::new();
    };
    if !content.raw().has("OMaths") {
        return Vec::new();
    }
    let Ok(omaths) = content.raw().get_obj("OMaths", "Content.OMaths") else {
        return Vec::new();
    };
    let count = omaths.get("Count").and_then(|value| value.int()).unwrap_or(0);
    let mut out = Vec::new();
    for index in 1..=count {
        let Ok(item) = omaths
            .call("Item", vec![ComValue::Int(index)])
            .and_then(|value| value.idispatch_child("OMath"))
        else {
            continue;
        };
        let Ok(range) = item.get_obj("Range", "OMath.Range") else {
            continue;
        };
        if let (Ok(start), Ok(end)) = (
            range.get("Start").and_then(|value| value.int()),
            range.get("End").and_then(|value| value.int()),
        ) {
            out.push((start, end));
        }
    }
    out
}

// ───────────────────── WPS 按键 / 屏幕刷新 ─────────────────────

/// 把目标窗口置前（跨进程 `Activate` 常被前台锁定拦下），并确认前台窗口确实属于同一进程。
/// 确认不了就不发按键 —— 否则 `Ctrl+=` 会打进用户当前真正在用的窗口；原因回传给调用方写进提示。
fn focus(app: &App) -> Result<(), &'static str> {
    use windows::Win32::Foundation::HWND;
    use windows::Win32::UI::WindowsAndMessaging::{GetForegroundWindow, SetForegroundWindow};
    let handle = app
        .raw()
        .get_obj("ActiveWindow", "ActiveWindow")
        .and_then(|window| window.get("Hwnd"))
        .and_then(|value| value.int());
    let Ok(handle) = handle else {
        // 引擎没暴露窗口句柄：信一次 `Activate`，按前台已切换处理（老版本 WPS 常见）
        return Ok(());
    };
    unsafe {
        let _ = SetForegroundWindow(HWND(handle as u32 as usize as *mut core::ffi::c_void));
    }
    std::thread::sleep(std::time::Duration::from_millis(80));
    let front = unsafe { GetForegroundWindow() };
    if front.0.is_null() {
        return Err("没能把 WPS 切到前台");
    }
    let front = front.0 as usize as u32 as i32;
    match (
        crate::process::pid_of_hwnd(handle),
        crate::process::pid_of_hwnd(front),
    ) {
        (Ok(target), Ok(current)) if target == current => Ok(()),
        _ => Err("没能把 WPS 切到前台"),
    }
}

/// 发 `Ctrl+=`：WPS 里这是「把选中的 LaTeX 转成专业公式」的快捷键（wps.vbs 第 76-77 行）。
/// 带上真实扫描码，只认扫描码的应用（部分 WPS 版本）也能收到。
fn press_ctrl_equals() {
    use windows::Win32::UI::Input::KeyboardAndMouse::{
        keybd_event, MapVirtualKeyW, KEYBD_EVENT_FLAGS, KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC,
        VK_CONTROL, VK_OEM_PLUS,
    };
    const PRESS: KEYBD_EVENT_FLAGS = KEYBD_EVENT_FLAGS(0);
    let scan = |vk: u16| unsafe { MapVirtualKeyW(vk as u32, MAPVK_VK_TO_VSC) as u8 };
    let (ctrl, equals) = (VK_CONTROL.0, VK_OEM_PLUS.0);
    // 按键要「按住」一会儿：WPS 对极短的合成按键会漏掉
    unsafe {
        keybd_event(ctrl as u8, scan(ctrl), PRESS, 0);
        std::thread::sleep(std::time::Duration::from_millis(30));
        keybd_event(equals as u8, scan(equals), PRESS, 0);
        std::thread::sleep(std::time::Duration::from_millis(80));
        keybd_event(equals as u8, scan(equals), KEYEVENTF_KEYUP, 0);
        std::thread::sleep(std::time::Duration::from_millis(40));
        keybd_event(ctrl as u8, scan(ctrl), KEYEVENTF_KEYUP, 0);
    }
}

/// 当前选区里确实有公式（查不到 `Selection.OMaths` 的老版本不做拦截）。
fn selection_holds_math(app: &App) -> bool {
    app.raw()
        .get_obj("Selection", "Selection")
        .and_then(|selection| selection.get_obj("OMaths", "Selection.OMaths"))
        .and_then(|omaths| omaths.get("Count"))
        .and_then(|count| count.int())
        .map(|count| count > 0)
        .unwrap_or(true)
}

/// 等公式对象真正建好再发按键：`OMaths.Add` 返回时公式内容可能还没落进 Range。
fn wait_ready(body: &IDispatchWrapper) {
    let deadline = std::time::Instant::now() + std::time::Duration::from_millis(READY_WAIT_MS);
    loop {
        if !is_blank(&text_of(body)) {
            return;
        }
        if std::time::Instant::now() >= deadline {
            return;
        }
        std::thread::sleep(std::time::Duration::from_millis(25));
    }
}

/// 轮询等 WPS 转换完成；没有命令（无反斜杠）时无需等待。
fn wait_professional(body: &IDispatchWrapper, source: &str, budget_ms: u64) {
    if !source.contains('\\') {
        return;
    }
    let deadline = std::time::Instant::now() + std::time::Duration::from_millis(budget_ms);
    while std::time::Instant::now() < deadline {
        std::thread::sleep(std::time::Duration::from_millis(25));
        if !text_of(body).contains('\\') {
            return;
        }
    }
}

/// `Application.ScreenUpdating = False`（VBA 同款），Drop 时恢复原值，异常退出也不留副作用。
struct ScreenGuard<'a> {
    app: &'a App,
    previous: bool,
}

impl<'a> ScreenGuard<'a> {
    fn off(app: &'a App) -> Self {
        let previous = app.screen_updating().unwrap_or(true);
        if previous {
            let _ = app.set_screen_updating(false);
        }
        Self { app, previous }
    }
}

impl Drop for ScreenGuard<'_> {
    fn drop(&mut self) {
        if self.previous {
            let _ = self.app.set_screen_updating(true);
        }
    }
}

// ───────────────────── LaTeX → UnicodeMath（失败后的自动纠正） ─────────────────────

/// Word 的 `BuildUp` 只解析 UnicodeMath：`\frac{a}{b}` → `(a)/(b)`、`\sqrt{x}` → `√(x)`、
/// 上下标 → `^(…)`/`_(…)`、命令换成真字符。无法识别的命令只去反斜杠（不丢字）。
fn to_unicodemath(latex: &str) -> String {
    let mut parser = Parser::new(latex);
    let nodes = parser.seq(None);
    let mut out = String::new();
    linear(&nodes, &mut out);
    out.trim().to_string()
}

enum Node {
    Text(String),
    Frac(Vec<Node>, Vec<Node>),
    Rad(Vec<Node>),
    Sup(Vec<Node>, Vec<Node>),
    Sub(Vec<Node>, Vec<Node>),
}

fn linear(nodes: &[Node], out: &mut String) {
    for node in nodes {
        match node {
            Node::Text(text) => push_text(text, out),
            Node::Frac(num, den) => {
                out.push('(');
                linear(num, out);
                out.push_str(")/(");
                linear(den, out);
                out.push(')');
            }
            Node::Rad(body) => {
                out.push('√');
                out.push('(');
                linear(body, out);
                out.push(')');
            }
            Node::Sup(base, sup) => {
                power_base(base, out);
                out.push_str("^(");
                linear(sup, out);
                out.push(')');
            }
            Node::Sub(base, sub) => {
                power_base(base, out);
                out.push_str("_(");
                linear(sub, out);
                out.push(')');
            }
        }
    }
}

/// 单字符底数直接接 `^`/`_`（`x^2`）；多字符必须加括号，否则只作用到最后一个字符。
fn power_base(nodes: &[Node], out: &mut String) {
    if let [Node::Text(text)] = nodes {
        if text.chars().count() == 1 {
            push_text(text, out);
            return;
        }
    }
    out.push('(');
    linear(nodes, out);
    out.push(')');
}

/// `&` 在 UnicodeMath 里是矩阵对齐符，落在普通文本里换成空格；反斜杠一律不入公式。
fn push_text(text: &str, out: &mut String) {
    for ch in text.chars() {
        match ch {
            '\\' => {}
            '&' => out.push(' '),
            _ => out.push(ch),
        }
    }
}

struct Parser {
    chars: Vec<char>,
    at: usize,
}

impl Parser {
    fn new(text: &str) -> Self {
        Self {
            chars: text.chars().collect(),
            at: 0,
        }
    }

    fn peek(&self) -> Option<char> {
        self.chars.get(self.at).copied()
    }

    /// 读到 `}` 或结尾；`stop` 由调用方吃掉收尾括号。
    fn seq(&mut self, stop: Option<char>) -> Vec<Node> {
        let mut out = Vec::new();
        while let Some(ch) = self.peek() {
            if Some(ch) == stop || ch == '}' {
                break;
            }
            match ch {
                ' ' | '\t' | '\r' | '\n' => {
                    self.at += 1;
                }
                '^' | '_' => {
                    self.at += 1;
                    let arg = self.arg();
                    let base = out.pop().unwrap_or(Node::Text(String::new()));
                    out.push(if ch == '^' {
                        Node::Sup(vec![base], arg)
                    } else {
                        Node::Sub(vec![base], arg)
                    });
                }
                '{' => {
                    self.at += 1;
                    let inner = self.seq(Some('}'));
                    if self.peek() == Some('}') {
                        self.at += 1;
                    }
                    out.extend(inner);
                }
                '\\' => {
                    if let Some(node) = self.command() {
                        out.push(node);
                    }
                }
                _ => {
                    self.at += 1;
                    out.push(Node::Text(ch.to_string()));
                }
            }
        }
        out
    }

    /// 命令参数：`{…}` 或单个字符/命令。
    fn arg(&mut self) -> Vec<Node> {
        while self.peek() == Some(' ') {
            self.at += 1;
        }
        match self.peek() {
            Some('{') => {
                self.at += 1;
                let inner = self.seq(Some('}'));
                if self.peek() == Some('}') {
                    self.at += 1;
                }
                inner
            }
            Some('\\') => self.command().map(|node| vec![node]).unwrap_or_default(),
            Some(ch) => {
                self.at += 1;
                vec![Node::Text(ch.to_string())]
            }
            None => Vec::new(),
        }
    }

    fn command(&mut self) -> Option<Node> {
        if self.peek() != Some('\\') {
            return None;
        }
        self.at += 1;
        let mut name = String::new();
        while let Some(ch) = self.peek() {
            if ch.is_ascii_alphabetic() {
                name.push(ch);
                self.at += 1;
            } else {
                break;
            }
        }
        if name.is_empty() {
            let ch = self.peek()?;
            self.at += 1;
            // `\\` 是换行命令
            return Some(Node::Text(if ch == '\\' { " ".into() } else { ch.to_string() }));
        }
        Some(match name.as_str() {
            "frac" | "dfrac" | "tfrac" => Node::Frac(self.arg(), self.arg()),
            "sqrt" => {
                // 根指数（`[3]`）UnicodeMath 侧不参与，正文内容照旧
                if self.peek() == Some('[') {
                    self.at += 1;
                    let _ = self.seq(Some(']'));
                    if self.peek() == Some(']') {
                        self.at += 1;
                    }
                }
                Node::Rad(self.arg())
            }
            "text" | "mathrm" | "mathbf" | "mathit" | "operatorname" => {
                let mut out = String::new();
                linear(&self.arg(), &mut out);
                Node::Text(out)
            }
            "left" | "right" => {
                // 定界符可能隔着空格、也可能写成 `\{`；`\left.` 是不可见定界符，不出字符
                while self.peek() == Some(' ') {
                    self.at += 1;
                }
                if self.peek() == Some('\\') {
                    self.at += 1;
                }
                match self.peek() {
                    Some('.') => {
                        self.at += 1;
                        Node::Text(String::new())
                    }
                    Some(ch) => {
                        self.at += 1;
                        Node::Text(ch.to_string())
                    }
                    None => Node::Text(String::new()),
                }
            }
            "quad" | "qquad" | "," | ";" | ":" | "!" => Node::Text(" ".into()),
            other => Node::Text(symbol(other).unwrap_or_else(|| other.to_string())),
        })
    }
}

fn symbol(name: &str) -> Option<String> {
    Some(
        match name {
            "pm" => "±",
            "mp" => "∓",
            "times" => "×",
            "div" => "÷",
            "cdot" => "·",
            "le" | "leq" => "≤",
            "ge" | "geq" => "≥",
            "ne" | "neq" => "≠",
            "approx" => "≈",
            "equiv" => "≡",
            "infty" => "∞",
            "sum" => "∑",
            "prod" => "∏",
            "int" => "∫",
            "partial" => "∂",
            "nabla" => "∇",
            "to" | "rightarrow" => "→",
            "leftarrow" => "←",
            "Rightarrow" => "⇒",
            "in" => "∈",
            "notin" => "∉",
            "alpha" => "α",
            "beta" => "β",
            "gamma" => "γ",
            "delta" => "δ",
            "epsilon" => "ε",
            "theta" => "θ",
            "lambda" => "λ",
            "mu" => "μ",
            "pi" => "π",
            "rho" => "ρ",
            "sigma" => "σ",
            "tau" => "τ",
            "phi" => "φ",
            "omega" => "ω",
            "Delta" => "Δ",
            "Omega" => "Ω",
            "degree" => "°",
            _ => return None,
        }
        .to_string(),
    )
}
