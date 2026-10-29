//! Markdown → Word 原生块写入。
//!
//! 一次工具调用写完整篇：标题落成 Heading 样式（进入导航与目录），列表/表格建成原生对象，
//! 引用/代码用模板样式，`$$…$$` 走公式对象（见 [`crate::math`]），`[TOC]` 占位转目录。
//! 解析是**行级**的（标题/列表/表格/引用/代码/公式），不追求完整 CommonMark：
//! 模型给的是聊天里那种规整 Markdown，够用且没有正则回溯风险。

use std::collections::BTreeMap;

use anyhow::Result;

use crate::app::App;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::format::styles;
use crate::target::resolve;
use crate::types::{InsertPosition, ListKind, Target};

/// 写入结果摘要。
pub struct Written {
    pub paragraphs: usize,
    pub tables: usize,
    pub equations: usize,
    pub toc: bool,
    pub warnings: Vec<String>,
    /// 待转换的公式（写入阶段只落 `$…$` 原文，全部写完后倒序转换）。
    pending: Vec<Pending>,
}

/// 一处待转换的公式：写入时的起始位置 + 定界符原文长度。
struct Pending {
    at: i32,
    len: i32,
    latex: String,
    display: bool,
}

#[derive(Debug)]
enum Block {
    Heading { level: u8, text: String },
    Para(String),
    List { ordered: bool, items: Vec<(u8, String)> },
    Quote(String),
    Code(String),
    Rule,
    Table(Vec<Vec<String>>),
    Math(String),
    Toc,
}

/// 语义类别 → 样式 id 的缺省映射（模板没给时用）。
fn default_style(key: &str) -> String {
    match key {
        "title" => "title",
        "body" | "normal" => "normal",
        "quote" => "quote",
        "code" => "code",
        "caption" => "caption",
        "list" => "list_paragraph",
        k if k.starts_with('h') => "",
        _ => "normal",
    }
    .to_string()
}

fn style_for(map: Option<&BTreeMap<String, String>>, key: &str) -> String {
    if let Some(map) = map {
        if let Some(style) = map.get(key).filter(|s| !s.trim().is_empty()) {
            return style.clone();
        }
    }
    if let Some(level) = key.strip_prefix('h').and_then(|n| n.parse::<u8>().ok()) {
        return format!("heading_{level}");
    }
    let tag = default_style(key);
    if tag == "code" {
        // 文档里不一定有 Code 样式：apply_style 会按名字找，找不到才退回 Normal。
        return "code".into();
    }
    tag
}

/// 写入入口：返回摘要；单块失败记 warning，正文不整体作废。
pub fn insert_markdown(
    app: &App,
    doc: &Document,
    target: &Target,
    markdown: &str,
    position: InsertPosition,
    style_map: Option<&BTreeMap<String, String>>,
    math: bool,
) -> Result<Written> {
    let range = resolve(app, doc, target)?.range;
    let mut pos = match position {
        InsertPosition::Start | InsertPosition::Before => range.start()?,
        InsertPosition::End | InsertPosition::After => range.end()?,
    };
    let mut written = Written {
        paragraphs: 0,
        tables: 0,
        equations: 0,
        toc: false,
        warnings: Vec::new(),
        pending: Vec::new(),
    };
    let blocks = parse(markdown);
    if blocks.is_empty() {
        written.warnings.push("没有可写入的内容".into());
        return Ok(written);
    }
    for block in blocks {
        match write_block(doc, pos, &block, style_map, math, &mut written) {
            Ok(next) => pos = next,
            Err(e) => written.warnings.push(format!("{block:?} 写入失败：{e}")),
        }
    }
    // 正文全部落定后再转公式（`$…$` / `$$…$$` 都已经闭合、位置一次都没被公式对象扰动过）。
    // 倒序：先转靠后的公式，前面记录的位置才不会因为公式长度变化而失效。
    for pending in std::mem::take(&mut written.pending).into_iter().rev() {
        let (_, warning) = crate::math::convert_at(
            app,
            doc,
            pending.at,
            pending.len,
            &pending.latex,
            pending.display,
        );
        match warning {
            Some(warning) => written.warnings.push(warning),
            None => written.equations += 1,
        }
    }
    Ok(written)
}

fn write_block(
    doc: &Document,
    pos: i32,
    block: &Block,
    style_map: Option<&BTreeMap<String, String>>,
    math: bool,
    written: &mut Written,
) -> Result<i32> {
    match block {
        Block::Heading { level, text } => {
            let style = style_for(style_map, &format!("h{}", (*level).min(6)));
            let next = put_inline(doc, pos, text, math, written)?;
            style_range(doc, pos, next, &style, written);
            written.paragraphs += 1;
            Ok(next)
        }
        Block::Para(text) => {
            let next = put_inline(doc, pos, text, math, written)?;
            style_range(doc, pos, next, &style_for(style_map, "body"), written);
            written.paragraphs += 1;
            Ok(next)
        }
        Block::Quote(text) => {
            let next = put_inline(doc, pos, text, math, written)?;
            style_range(doc, pos, next, &style_for(style_map, "quote"), written);
            written.paragraphs += 1;
            Ok(next)
        }
        Block::Code(text) => {
            // 代码块：一个段落里的软换行（\v），保留缩进
            let payload = format!("{}\r", text.replace('\n', "\u{b}"));
            let next = put_raw(doc, pos, &payload)?;
            style_range(doc, pos, next, &style_for(style_map, "code"), written);
            written.paragraphs += 1;
            Ok(next)
        }
        Block::List { ordered, items } => {
            let list_style = style_for(style_map, "list");
            let mut next = pos;
            let mut ranges = Vec::new();
            for (level, item) in items {
                let start = next;
                next = put_inline(doc, start, item, math, written)?;
                style_range(doc, start, next, &list_style, written);
                ranges.push((start, next, *level));
            }
            if !ranges.is_empty() {
                let kind = if *ordered { ListKind::Number } else { ListKind::Bullet };
                let span = doc.range(pos, next)?;
                if let Err(e) = crate::format::lists::apply_list(&span, kind) {
                    written.warnings.push(format!("列表格式未套上：{e}"));
                }
                for (s, e, level) in ranges {
                    if level > 0 {
                        if let Ok(r) = doc.range(s, e) {
                            if let Ok(lf) = r.list_format() {
                                let _ = lf.put(
                                    "ListLevelNumber",
                                    vec![ComValue::Int((level + 1).min(9) as i32)],
                                );
                            }
                        }
                    }
                }
                written.paragraphs += 1;
            }
            Ok(next)
        }
        Block::Rule => {
            let next = put_raw(doc, pos, "\r")?;
            written.paragraphs += 1;
            Ok(next)
        }
        Block::Table(rows) => write_table(doc, pos, rows, style_map, written),
        Block::Math(src) => {
            // 模板关掉公式：按「去掉 $ 的 LaTeX 原文」落稿
            if !math {
                let text = format!("{}\r", crate::math::plain_latex(src));
                let next = put_raw(doc, pos, &text)?;
                style_range(doc, pos, next, &style_for(style_map, "body"), written);
                written.paragraphs += 1;
                return Ok(next);
            }
            // 开公式：先落带定界符的原文（`$$…$$` 已闭合），整篇写完再统一转
            let payload = format!("$${}$$", src.trim());
            let next = put_raw(doc, pos, &format!("{payload}\r"))?;
            style_range(doc, pos, next, &style_for(style_map, "body"), written);
            written.paragraphs += 1;
            written.pending.push(Pending {
                at: pos,
                len: payload.encode_utf16().count() as i32,
                latex: src.trim().to_string(),
                display: true,
            });
            Ok(next)
        }
        Block::Toc => {
            let spec = crate::types::TocSpec {
                max_level: 3,
                title: Some("目录".into()),
                title_style: None,
            };
            match crate::layout::insert_toc(doc, &spec) {
                Ok(note) => {
                    written.warnings.push(note);
                    written.toc = true;
                }
                Err(e) => written.warnings.push(format!("目录未插入：{e}")),
            }
            Ok(doc.range(pos, pos)?.end()?)
        }
    }
}

/// 插入一个段落：`$…$` / `$$…$$` 先按原文落稿（整篇写完统一转），其余按行内标记落稿。
/// 返回下一块起点。
fn put_inline(
    doc: &Document,
    pos: i32,
    text: &str,
    math: bool,
    written: &mut Written,
) -> Result<i32> {
    let mut cursor = pos;
    for piece in split_inline(text.trim_end()) {
        match piece {
            Inline::Text(text) => cursor = put_text_runs(doc, cursor, &text)?,
            // 模板关公式：落「去掉 $ 的 LaTeX 原文」
            Inline::Math { latex, .. } if !math => {
                cursor = put_raw(doc, cursor, &crate::math::plain_latex(&latex))?;
            }
            Inline::Math { latex, display } => {
                let payload = if display {
                    format!("$${latex}$$")
                } else {
                    format!("${latex}$")
                };
                let at = cursor;
                cursor = put_raw(doc, cursor, &payload)?;
                written.pending.push(Pending {
                    at,
                    len: payload.encode_utf16().count() as i32,
                    latex,
                    display,
                });
            }
        }
    }
    put_raw(doc, cursor, "\r")
}

/// 写普通文本（含行内标记），返回新位置。逐段写：公式对象的长度由 COM 决定，不能预先累加。
fn put_text_runs(doc: &Document, pos: i32, text: &str) -> Result<i32> {
    let mut cursor = pos;
    for run in scan_inline(text) {
        let len = run.text.encode_utf16().count() as i32;
        if len == 0 {
            continue;
        }
        let end = put_raw(doc, cursor, &run.text)?;
        if run.bold || run.italic || run.code {
            if let Ok(range) = doc.range(cursor, end) {
                if let Ok(font) = range.font() {
                    let patch = crate::types::FontPatch {
                        bold: run.bold.then_some(true),
                        italic: run.italic.then_some(true),
                        name: run.code.then(|| "Consolas".to_string()),
                        east_asia_name: run.code.then(|| "等线".to_string()),
                        ..Default::default()
                    };
                    let _ = font.apply(&patch);
                }
            }
        }
        cursor = end;
    }
    Ok(cursor)
}

/// 在 `pos` 处插入文本，返回**插入后 Range 的真实 End**。
/// 不能用 `pos + 字符数`：公式对象在 Word 里占的位置与字符数不同，手算会错位。
fn put_raw(doc: &Document, pos: i32, payload: &str) -> Result<i32> {
    let range = doc.range(pos, pos)?;
    range.insert_after(payload)?;
    range.end()
}

fn style_range(doc: &Document, start: i32, end: i32, style: &str, written: &mut Written) {
    if style.trim().is_empty() {
        return;
    }
    match doc
        .range(start, end)
        .and_then(|r| styles::apply_style_or_create(&r, doc, style))
    {
        Ok(()) => {}
        Err(e) => written.warnings.push(format!("样式「{style}」未套上：{e}")),
    }
}

fn write_table(
    doc: &Document,
    pos: i32,
    rows: &[Vec<String>],
    style_map: Option<&BTreeMap<String, String>>,
    written: &mut Written,
) -> Result<i32> {
    let row_count = rows.len().max(1) as i32;
    let col_count = rows.iter().map(Vec::len).max().unwrap_or(1).max(1) as i32;
    let at = doc.range(pos, pos)?;
    let table = doc
        .tables()?
        .call(
            "Add",
            vec![
                ComValue::Obj(at.raw().disp.clone()),
                ComValue::Int(row_count),
                ComValue::Int(col_count),
            ],
        )?
        .idispatch_child("Table")?;
    for (r, row) in rows.iter().enumerate() {
        for c in 0..col_count as usize {
            let value = row.get(c).map(String::as_str).unwrap_or("");
            if value.is_empty() {
                continue;
            }
            if let Ok(cell) = table
                .call("Cell", vec![ComValue::Int(r as i32 + 1), ComValue::Int(c as i32 + 1)])
                .and_then(|v| v.idispatch_child("Cell"))
            {
                if let Ok(range) = cell.get_obj("Range", "Cell.Range") {
                    let _ = range.put("Text", vec![ComValue::Str(plain_inline(value))]);
                }
            }
        }
    }
    // 表头加粗
    for c in 1..=col_count {
        if let Ok(cell) = table
            .call("Cell", vec![ComValue::Int(1), ComValue::Int(c)])
            .and_then(|v| v.idispatch_child("Cell"))
        {
            if let Ok(font) = cell
                .get_obj("Range", "Cell.Range")
                .and_then(|r| r.get_obj("Font", "Font"))
            {
                let _ = font.put("Bold", vec![ComValue::Bool(true)]);
            }
        }
    }
    let _ = style_map;
    written.tables += 1;
    // 表格末尾必须补一个段落标记：否则下一块内容会插进最后一个单元格里
    let end = table.get_obj("Range", "Table.Range")?.get("End")?.int()?;
    let _ = doc.range(end, end)?.insert_after("\r");
    Ok(end + 1)
}

/// 段落拆成「普通文本 / 公式」两段式：`$…$` 行内、`$$…$$` 独立。
enum Inline {
    Text(String),
    Math { latex: String, display: bool },
}

fn split_inline(text: &str) -> Vec<Inline> {
    let mut out = Vec::new();
    let mut buf = String::new();
    let mut rest = text;
    while let Some(start) = rest.find('$') {
        buf.push_str(&rest[..start]);
        let after = &rest[start + 1..];
        let double = after.starts_with('$');
        let inner = if double { &after[1..] } else { after };
        let delim = if double { "$$" } else { "$" };
        match inner.find(delim).filter(|end| *end > 0) {
            Some(end) => {
                if !buf.is_empty() {
                    out.push(Inline::Text(std::mem::take(&mut buf)));
                }
                out.push(Inline::Math {
                    latex: inner[..end].to_string(),
                    display: double,
                });
                rest = &inner[end + delim.len()..];
            }
            // 没配对（或空公式）：定界符按原文保留，别吞字
            None => {
                buf.push_str(delim);
                rest = inner;
            }
        }
    }
    buf.push_str(rest);
    if !buf.is_empty() {
        out.push(Inline::Text(buf));
    }
    out
}

/// 只留可读文本：去行内标记、`$…$` 落去定界符的 LaTeX 原文（表格单元格用）。
fn plain_inline(text: &str) -> String {
    split_inline(text)
        .into_iter()
        .map(|piece| match piece {
            Inline::Text(text) => strip_inline(&text),
            Inline::Math { latex, .. } => crate::math::plain_latex(&latex),
        })
        .collect()
}


#[derive(Debug, Clone)]
struct Run {
    text: String,
    bold: bool,
    italic: bool,
    code: bool,
}

/// 行内标记扫描：`**粗体**`、`*斜体*`、`` `代码` ``；不配对时按普通文字处理。
fn scan_inline(text: &str) -> Vec<Run> {
    let mut runs: Vec<Run> = Vec::new();
    let mut buf = String::new();
    let mut bold = false;
    let mut italic = false;
    let mut code = false;
    let chars: Vec<char> = text.chars().collect();
    let mut i = 0;
    let flush = |buf: &mut String, runs: &mut Vec<Run>, bold: bool, italic: bool, code: bool| {
        if !buf.is_empty() {
            runs.push(Run {
                text: std::mem::take(buf),
                bold,
                italic,
                code,
            });
        }
    };
    while i < chars.len() {
        let ch = chars[i];
        if ch == '`' {
            flush(&mut buf, &mut runs, bold, italic, code);
            code = !code;
            i += 1;
            continue;
        }
        if ch == '*' {
            let double = chars.get(i + 1) == Some(&'*');
            flush(&mut buf, &mut runs, bold, italic, code);
            if double {
                bold = !bold;
                i += 2;
            } else {
                italic = !italic;
                i += 1;
            }
            continue;
        }
        if ch == '\\' && chars.get(i + 1) == Some(&'*') {
            buf.push('*');
            i += 2;
            continue;
        }
        buf.push(ch);
        i += 1;
    }
    flush(&mut buf, &mut runs, bold, italic, code);
    if runs.is_empty() {
        runs.push(Run {
            text: String::new(),
            bold: false,
            italic: false,
            code: false,
        });
    }
    runs
}

fn strip_inline(text: &str) -> String {
    scan_inline(text).iter().map(|r| r.text.as_str()).collect()
}

// ───────────────────────── 解析 ─────────────────────────

fn parse(md: &str) -> Vec<Block> {
    let normalized = md.replace("\r\n", "\n").replace('\r', "\n");
    let lines: Vec<&str> = normalized.split('\n').collect();
    let mut blocks = Vec::new();
    let mut para: Vec<String> = Vec::new();
    let mut i = 0usize;
    while i < lines.len() {
        let raw = lines[i];
        let line = raw.trim_end();
        let trimmed = line.trim();
        if trimmed.is_empty() {
            flush_para(&mut para, &mut blocks);
            i += 1;
            continue;
        }
        // 围栏代码块
        if trimmed.starts_with("```") {
            flush_para(&mut para, &mut blocks);
            let mut body = Vec::new();
            i += 1;
            while i < lines.len() && !lines[i].trim_start().starts_with("```") {
                body.push(lines[i]);
                i += 1;
            }
            i += 1;
            blocks.push(Block::Code(body.join("\n")));
            continue;
        }
        // 目录占位
        let lower = trimmed.to_ascii_lowercase();
        if matches!(lower.as_str(), "[toc]" | "{{toc}}" | "<!-- toc -->" | "[[toc]]") {
            flush_para(&mut para, &mut blocks);
            blocks.push(Block::Toc);
            i += 1;
            continue;
        }
        // 标题
        if let Some((level, text)) = heading(trimmed) {
            flush_para(&mut para, &mut blocks);
            blocks.push(Block::Heading { level, text });
            i += 1;
            continue;
        }
        // 分隔线
        if is_rule(trimmed) {
            flush_para(&mut para, &mut blocks);
            blocks.push(Block::Rule);
            i += 1;
            continue;
        }
        // 块公式：`$$…$$` 单行或跨行（AI 常写成三行），由闭合的 `$$` 收尾。
        // `$$$…` 这类畸形写法不当开头，免得把后面整篇吞进公式。
        if let Some(body) = trimmed.strip_prefix("$$").filter(|body| !body.starts_with('$')) {
            if let Some(one) = body.strip_suffix("$$").map(str::trim).filter(|one| !one.is_empty()) {
                flush_para(&mut para, &mut blocks);
                blocks.push(Block::Math(one.to_string()));
                i += 1;
                continue;
            }
            let mut parts = vec![body.trim().to_string()];
            let mut next = i + 1;
            let mut closed = false;
            while next < lines.len() {
                let line = lines[next].trim();
                next += 1;
                match line.strip_suffix("$$") {
                    Some(head) => {
                        parts.push(head.trim().to_string());
                        closed = true;
                        break;
                    }
                    None => parts.push(line.to_string()),
                }
            }
            // 没闭合就按普通文字走（不能把后面整篇吞进公式）
            if closed {
                flush_para(&mut para, &mut blocks);
                blocks.push(Block::Math(parts.join(" ").trim().to_string()));
                i = next;
                continue;
            }
        }
        // 表格：本行含 |，下一行是分隔行
        if trimmed.contains('|') && i + 1 < lines.len() && is_table_sep(lines[i + 1]) {
            flush_para(&mut para, &mut blocks);
            let mut rows = vec![split_row(trimmed)];
            i += 2;
            while i < lines.len() && lines[i].trim().contains('|') {
                rows.push(split_row(lines[i].trim()));
                i += 1;
            }
            blocks.push(Block::Table(rows));
            continue;
        }
        // 列表
        if let Some((level, ordered, text)) = list_item(line) {
            flush_para(&mut para, &mut blocks);
            let mut items = vec![(level, text)];
            i += 1;
            while i < lines.len() {
                if lines[i].trim().is_empty() {
                    break;
                }
                match list_item(lines[i]) {
                    Some((l, o, t)) if o == ordered => {
                        items.push((l, t));
                        i += 1;
                    }
                    _ => break,
                }
            }
            blocks.push(Block::List { ordered, items });
            continue;
        }
        // 引用
        if let Some(body) = trimmed.strip_prefix('>') {
            flush_para(&mut para, &mut blocks);
            blocks.push(Block::Quote(body.trim().to_string()));
            i += 1;
            continue;
        }
        para.push(trimmed.to_string());
        i += 1;
    }
    flush_para(&mut para, &mut blocks);
    blocks
}

fn flush_para(para: &mut Vec<String>, blocks: &mut Vec<Block>) {
    if para.is_empty() {
        return;
    }
    let text = para.join(" ");
    para.clear();
    blocks.push(Block::Para(text));
}

fn heading(line: &str) -> Option<(u8, String)> {
    let hashes = line.chars().take_while(|c| *c == '#').count();
    if hashes == 0 || hashes > 6 {
        return None;
    }
    let rest = line[hashes..].trim_start();
    if rest.is_empty() {
        return None;
    }
    // `#标题` 也认（模型偶尔漏空格）
    Some((hashes as u8, rest.trim_end_matches('#').trim().to_string()))
}

fn is_rule(line: &str) -> bool {
    let t = line.trim();
    t.len() >= 3
        && (t.chars().all(|c| c == '-') || t.chars().all(|c| c == '*') || t.chars().all(|c| c == '_'))
}

fn is_table_sep(line: &str) -> bool {
    let t = line.trim();
    !t.is_empty()
        && t.contains('|')
        && t.chars().all(|c| matches!(c, '|' | '-' | ':' | ' '))
}

fn split_row(line: &str) -> Vec<String> {
    let t = line.trim().trim_start_matches('|').trim_end_matches('|');
    t.split('|').map(|c| c.trim().to_string()).collect()
}

fn list_item(line: &str) -> Option<(u8, bool, String)> {
    let indent = line.chars().take_while(|c| *c == ' ' || *c == '\t').count();
    let level = (indent / 2).min(3) as u8;
    let rest = line.trim_start();
    for marker in ["- ", "* ", "+ "] {
        if let Some(text) = rest.strip_prefix(marker) {
            return Some((level, false, text.trim().to_string()));
        }
    }
    let digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
    if digits > 0 && digits <= 3 {
        let after = &rest[digits..];
        if let Some(text) = after.strip_prefix(". ").or_else(|| after.strip_prefix(") ")) {
            return Some((level, true, text.trim().to_string()));
        }
    }
    None
}


