//! Range：文本读写、插入/删除、折叠、查找、字体/段落格式入口。
//!
//! `Range.Start/End` 是 Word 文档位置，**不是** Rust UTF-8 字节下标。
//! 修改范围前必须先核对 `expected_text`（见 [`crate::document::anchors`]）。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::format::{Font, ParagraphFormat};

/// 文档字符范围（含对象终止符时由其自身语义决定）。
#[derive(Clone)]
pub struct Range {
    raw: IDispatchWrapper,
}

impl Range {
    pub(crate) fn new(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    /// 底层 COM 包装。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    pub(crate) fn label(&self) -> &str {
        &self.raw.ctx
    }

    pub fn start(&self) -> Result<i32> {
        self.raw.get("Start")?.int()
    }

    pub fn end(&self) -> Result<i32> {
        self.raw.get("End")?.int()
    }

    /// 纯文本读取（含段落/单元格终止符，展示层需清理）。
    pub fn text(&self) -> Result<String> {
        self.raw.get("Text")?.string()
    }

    /// 替换范围内容（可能清除范围内局部格式/对象结构；调用方须先核对 expected_text）。
    pub fn set_text(&self, text: &str) -> Result<()> {
        self.raw.put("Text", vec![ComValue::Str(text.to_string())])
    }

    pub fn delete(&self) -> Result<()> {
        self.raw.call("Delete", vec![]).map(|_| ())
    }

    pub fn insert_before(&self, text: &str) -> Result<()> {
        self.raw
            .call("InsertBefore", vec![ComValue::Str(text.to_string())])
            .map(|_| ())
    }

    pub fn insert_after(&self, text: &str) -> Result<()> {
        self.raw
            .call("InsertAfter", vec![ComValue::Str(text.to_string())])
            .map(|_| ())
    }

    pub fn insert_paragraph_after(&self) -> Result<()> {
        self.raw.call("InsertParagraphAfter", vec![]).map(|_| ())
    }

    pub fn insert_paragraph_before(&self) -> Result<()> {
        self.raw.call("InsertParagraphBefore", vec![]).map(|_| ())
    }

    /// 插入分页符（wdPageBreak=7）。
    pub fn insert_page_break(&self) -> Result<()> {
        self.raw
            .call("InsertBreak", vec![ComValue::Int(7)])
            .map(|_| ())
    }

    /// 插入分节符（下一节/连续/偶数页/奇数页）。
    pub fn insert_section_break(&self, kind: i32) -> Result<()> {
        self.raw
            .call("InsertBreak", vec![ComValue::Int(kind)])
            .map(|_| ())
    }

    /// 折叠到终点（不修改原 Range：先 Duplicate 再 Collapse，wdCollapseEnd=0）。
    pub fn collapse_end(&self) -> Result<Range> {
        let dup = self
            .raw
            .call("Duplicate", vec![])?
            .idispatch_child("Duplicate")?;
        dup.call("Collapse", vec![ComValue::Int(0)])?;
        Ok(Range::new(dup))
    }

    /// 折叠到起点（wdCollapseStart=1）。
    pub fn collapse_start(&self) -> Result<Range> {
        let dup = self
            .raw
            .call("Duplicate", vec![])?
            .idispatch_child("Duplicate")?;
        dup.call("Collapse", vec![ComValue::Int(1)])?;
        Ok(Range::new(dup))
    }

    pub fn find(&self) -> Result<Find> {
        Ok(Find::new(self.raw.get_obj("Find", "Find")?))
    }

    /// 字体格式入口（同时可设置 Range 级高亮/底纹）。
    pub fn font(&self) -> Result<Font> {
        Ok(Font::new(
            self.raw.get_obj("Font", "Font")?,
            self.raw.clone(),
        ))
    }

    /// 段落格式入口。
    pub fn paragraph_format(&self) -> Result<ParagraphFormat> {
        Ok(ParagraphFormat::new(
            self.raw.get_obj("ParagraphFormat", "ParagraphFormat")?,
        ))
    }

    pub fn style(&self) -> Result<String> {
        self.raw.get("Style").and_then(|v| v.string())
    }

    pub fn set_style(&self, name: &str) -> Result<()> {
        self.raw.put("Style", vec![ComValue::Str(name.to_string())])
    }

    pub fn list_format(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("ListFormat", "ListFormat")
    }

    /// 书签集合（用于 Anchors.Add）。
    pub fn bookmarks(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Bookmarks", "Bookmarks")
    }

    /// 超链接集合。
    pub fn hyperlinks(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Hyperlinks", "Hyperlinks")
    }

    /// 字段集合（Fields.Add 插入 PAGE/NUMPAGES/DATE 等）。
    pub fn fields(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Fields", "Fields")
    }

    /// 该范围上的修订集合（用于回读验证生成的 Revision）。
    pub fn revisions(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Revisions", "Revisions")
    }
}

/// `Range.Find`（查找/替换）。执行后 Range 自身被更新为命中范围。
pub struct Find {
    raw: IDispatchWrapper,
}

/// 查找方向/换行常量。
pub const FIND_STOP: i32 = 0;
pub const REPLACE_ONE: i32 = 1;
pub const REPLACE_ALL: i32 = 2;

impl Find {
    pub(crate) fn new(raw: IDispatchWrapper) -> Self {
        Self { raw }
    }

    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 显式设置查找条件（不继承上次 UI 查找框状态）。
    pub fn configure(&self, query: &str, opts: &crate::types::FindOptions) -> Result<()> {
        self.raw.call("ClearFormatting", vec![])?;
        self.raw
            .put("Text", vec![ComValue::Str(query.to_string())])?;
        self.raw.put("Forward", vec![ComValue::Bool(true)])?;
        self.raw.put("Wrap", vec![ComValue::Int(FIND_STOP)])?;
        self.raw
            .put("MatchCase", vec![ComValue::Bool(opts.match_case)])?;
        self.raw
            .put("MatchWholeWord", vec![ComValue::Bool(opts.whole_word)])?;
        self.raw
            .put("MatchWildcards", vec![ComValue::Bool(opts.wildcards)])?;
        Ok(())
    }

    /// 查找下一处；命中返回 true（执行后持有 Find 的 Range 被更新）。
    pub fn execute(&self) -> Result<bool> {
        self.raw.call("Execute", vec![])?.bool()
    }

    /// 设置替换文本。
    pub fn set_replacement(&self, text: &str) -> Result<()> {
        let repl = self.raw.get_obj("Replacement", "Replacement")?;
        repl.call("ClearFormatting", vec![])?;
        repl.put("Text", vec![ComValue::Str(text.to_string())])
    }

    /// 执行替换（`Replace` 命名参数：wdReplaceOne=1 / wdReplaceAll=2）。
    pub fn execute_replace(&self, mode: i32) -> Result<bool> {
        self.raw
            .call_named("Execute", vec![], &[("Replace", ComValue::Int(mode))])?
            .bool()
    }
}
