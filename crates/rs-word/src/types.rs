//! 公共类型：引擎/模式、能力状态、Target、格式补丁与全部请求/响应 DTO。
//!
//! 约定：
//! - 单位统一 points；颜色用 `#RRGGBB`（facade 转 OLE）。
//! - 所有 Word 常量封装为 enum，AI/调用方永不传整数。
//! - 格式补丁字段为 `Option<T>`：`None` = 不改；重置由显式 `reset` 表达。

use serde::{Deserialize, Serialize};

pub use crate::error::SideEffectStatus;

// ───────────────────────── 引擎 / 连接 ─────────────────────────

/// 驱动引擎。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WordEngine {
    WpsWriter,
    MsWord,
}

impl WordEngine {
    pub fn label(self) -> &'static str {
        match self {
            Self::WpsWriter => "WPS Writer",
            Self::MsWord => "Microsoft Word",
        }
    }

    /// 默认（首选）ProgID。WPS 需候选探测，见 [`crate::session`]。
    pub fn primary_prog_id(self) -> &'static str {
        match self {
            Self::WpsWriter => "KWPS.Application",
            Self::MsWord => "Word.Application",
        }
    }

    /// 进程可执行文件名（小写，进程查看/救援匹配用）。
    pub fn exe_name(self) -> &'static str {
        match self {
            Self::WpsWriter => "wps.exe",
            Self::MsWord => "winword.exe",
        }
    }
}

/// 连接模式。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ConnectMode {
    /// 接管已运行实例（ROT）。默认，不抢占用户工作。
    Attach,
    /// 新建实例；仅关闭本库创建的实例。
    New,
}

/// 引擎选择策略。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum EnginePolicy {
    Explicit(WordEngine),
    PreferWps,
    PreferMsWord,
}

// ───────────────────────── 能力 ─────────────────────────

/// 能力状态（必须逐项真机记录，不得凭名称推断）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CapabilityStatus {
    Supported,
    Partial,
    Unsupported,
    Untested,
}

/// 单项能力信息。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CapabilityInfo {
    pub name: String,
    pub status: CapabilityStatus,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reason: Option<String>,
}

/// 引擎身份（连接后核验，不能只信 ProgID）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct EngineIdentity {
    pub engine: WordEngine,
    pub prog_id: String,
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub version: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pid: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub exe: Option<String>,
}

// ───────────────────────── Story / 坐标 ─────────────────────────

/// 文档 story（正文/页眉/页脚/脚注/尾注/批注/文本框）。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Story {
    #[default]
    Main,
    Header,
    Footer,
    Footnote,
    Endnote,
    Comment,
    TextBox,
}

impl std::fmt::Display for Story {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::Main => "main",
            Self::Header => "header",
            Self::Footer => "footer",
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Comment => "comment",
            Self::TextBox => "textbox",
        })
    }
}

/// 页眉/页脚类型。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HfKind {
    Header,
    Footer,
}

/// 页眉/页脚变体。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HfVariant {
    #[default]
    Primary,
    FirstPage,
    EvenPage,
}

impl HfVariant {
    /// wdHeaderFooterPrimary=1 / FirstPage=2 / EvenPages=3。
    pub fn code(self) -> i32 {
        match self {
            Self::Primary => 1,
            Self::FirstPage => 2,
            Self::EvenPage => 3,
        }
    }
}

// ───────────────────────── Target ─────────────────────────

/// 目标定位（tagged enum，互斥字段由 Rust 二次校验）。
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum Target {
    Document {
        document_id: String,
    },
    Selection {
        document_id: String,
    },
    Range {
        document_id: String,
        #[serde(default)]
        story: Story,
        start: i32,
        end: i32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expected_text: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        text_hash: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expected_doc_version: Option<u64>,
    },
    Paragraph {
        document_id: String,
        paragraph_index: u32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expected_text: Option<String>,
    },
    Find {
        document_id: String,
        query: String,
        #[serde(default = "default_one")]
        occurrence: u32,
    },
    Bookmark {
        document_id: String,
        bookmark: String,
    },
    TableCell {
        document_id: String,
        table_id: String,
        row: u32,
        column: u32,
    },
    /// 整张表格（含表尾标记）。
    Table {
        document_id: String,
        table_id: String,
    },
    /// 表格多选区：矩形（`row_start..=row_end` × `col_start..=col_end`）或零散单元格（`cells`）。
    /// 某维起止都为 0 = 该维不限定 → 整表取满；只给一端 = 单行/单列。
    TableCells {
        document_id: String,
        table_id: String,
        #[serde(default)]
        row_start: u32,
        #[serde(default)]
        row_end: u32,
        #[serde(default)]
        col_start: u32,
        #[serde(default)]
        col_end: u32,
        /// 零散多选（同一行相邻列会自动合并成一段）；给了它就不看上面的矩形。
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        cells: Vec<CellRef>,
    },
    HeaderFooter {
        document_id: String,
        #[serde(default = "default_one")]
        section: u32,
        hf_kind: HfKind,
        #[serde(default)]
        variant: HfVariant,
    },
}

/// 单元格引用（1 基）。`{"row":1,"column":2}` 与 `[1,2]` 都收。
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub enum CellRef {
    Obj {
        row: u32,
        #[serde(alias = "col", alias = "col_index")]
        column: u32,
    },
    Pair([u32; 2]),
}

impl CellRef {
    /// （行, 列），均从 1 起。
    pub fn rc(&self) -> (u32, u32) {
        match self {
            Self::Obj { row, column } => (*row, *column),
            Self::Pair([row, column]) => (*row, *column),
        }
    }
}

fn default_one() -> u32 {
    1
}

impl Target {
    pub fn document_id(&self) -> &str {
        match self {
            Self::Document { document_id }
            | Self::Selection { document_id }
            | Self::Range { document_id, .. }
            | Self::Paragraph { document_id, .. }
            | Self::Find { document_id, .. }
            | Self::Bookmark { document_id, .. }
            | Self::TableCell { document_id, .. }
            | Self::Table { document_id, .. }
            | Self::TableCells { document_id, .. }
            | Self::HeaderFooter { document_id, .. } => document_id,
        }
    }
}

/// 文本锚点（读取返回、修改回传校验）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TextAnchor {
    pub document_id: String,
    pub story: Story,
    pub start: i32,
    pub end: i32,
    pub quote: String,
    pub text_hash: String,
    pub doc_version: u64,
}

// ───────────────────────── 格式枚举 ─────────────────────────

/// 段落对齐。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ParagraphAlignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

impl ParagraphAlignment {
    /// wdAlignParagraph*。
    pub fn code(self) -> i32 {
        match self {
            Self::Left => 0,
            Self::Center => 1,
            Self::Right => 2,
            Self::Justify => 3,
            Self::Distribute => 4,
        }
    }
}

/// 下划线样式（子集，未验证枚举返回 unsupported 而非猜值）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum UnderlineStyle {
    None,
    Single,
    Words,
    Double,
    Dotted,
    Dash,
    DotDash,
    Thick,
    Wave,
}

impl UnderlineStyle {
    /// wdUnderline*。
    pub fn code(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Words => 2,
            Self::Double => 3,
            Self::Dotted => 4,
            Self::Dash => 5,
            Self::Thick => 6,
            Self::DotDash => 7,
            Self::Wave => 11,
        }
    }
}

/// 荧光笔颜色（WdColorIndex 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HighlightColor {
    None,
    Yellow,
    BrightGreen,
    Turquoise,
    Pink,
    Blue,
    Red,
    DarkBlue,
    Teal,
    Green,
    Violet,
    DarkRed,
    Gray,
}

impl HighlightColor {
    pub fn code(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Blue => 2,
            Self::Turquoise => 3,
            Self::BrightGreen => 4,
            Self::Pink => 5,
            Self::Red => 6,
            Self::Yellow => 7,
            Self::DarkBlue => 9,
            Self::Teal => 10,
            Self::Green => 11,
            Self::Violet => 12,
            Self::DarkRed => 13,
            Self::Gray => 15,
        }
    }
}

/// 行距。
#[derive(Clone, Copy, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum LineSpacing {
    Single,
    OneAndHalf,
    Double,
    /// 多倍行距。
    Multiple {
        value: f32,
    },
    /// 固定值（pt）。
    Exactly {
        value: f32,
    },
    /// 最小值（pt）。
    AtLeast {
        value: f32,
    },
}

impl LineSpacing {
    /// wdLineSpace*。
    pub fn rule_code(self) -> i32 {
        match self {
            Self::Single => 0,
            Self::OneAndHalf => 1,
            Self::Double => 2,
            Self::AtLeast { .. } => 3,
            Self::Exactly { .. } => 4,
            Self::Multiple { .. } => 5,
        }
    }

    /// Multiple/Exactly/AtLeast 需要写 LineSpacing 数值；其余返回 None。
    pub fn value_for_com(self) -> Option<f32> {
        match self {
            Self::Multiple { value } | Self::Exactly { value } | Self::AtLeast { value } => {
                Some(value)
            }
            _ => None,
        }
    }
}

/// 大纲级别（正文/1..9）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum OutlineLevel {
    BodyText,
    Level1,
    Level2,
    Level3,
    Level4,
    Level5,
    Level6,
    Level7,
    Level8,
    Level9,
}

impl OutlineLevel {
    pub fn code(self) -> i32 {
        match self {
            Self::Level1 => 1,
            Self::Level2 => 2,
            Self::Level3 => 3,
            Self::Level4 => 4,
            Self::Level5 => 5,
            Self::Level6 => 6,
            Self::Level7 => 7,
            Self::Level8 => 8,
            Self::Level9 => 9,
            Self::BodyText => 10,
        }
    }
}

// ───────────────────────── 格式补丁 ─────────────────────────

/// 字符格式补丁（None = 不改）。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct FontPatch {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_pt: Option<f32>,
    /// `#RRGGBB`。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineStyle>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub double_strike: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub superscript: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subscript: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub all_caps: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub char_spacing_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<HighlightColor>,
    /// 字符底纹背景色 `#RRGGBB`。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
}

/// 段落格式补丁（None = 不改）。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct ParagraphPatch {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<ParagraphAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_spacing: Option<LineSpacing>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space_before_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space_after_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left_indent_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right_indent_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_line_indent_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hanging_indent_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_with_next: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_together: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<OutlineLevel>,
}

/// 列表类型（`Custom` 需同时提供 `FormatRequest::list_template`）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ListKind {
    None,
    Bullet,
    Number,
    OutlineNumber,
    Custom,
}

/// 清除格式范围。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct ResetSpec {
    #[serde(default)]
    pub font: bool,
    #[serde(default)]
    pub paragraph: bool,
}

// ───────────────────────── 会话/文档 ─────────────────────────

/// 文档快照信息。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct DocumentInfo {
    pub document_id: String,
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    pub saved: bool,
    pub read_only: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table_count: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision_count: Option<u32>,
    pub is_active: bool,
}

/// 会话状态。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WordStatus {
    pub connected: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub engine: Option<WordEngine>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub identity: Option<EngineIdentity>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub active_document_id: Option<String>,
    pub documents: Vec<DocumentInfo>,
    pub capabilities: Vec<CapabilityInfo>,
    pub owns_instance: bool,
}

/// 保存文件格式。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SaveFormat {
    /// 保持原格式（Save）。
    Keep,
    Docx,
    Doc,
    Rtf,
    Txt,
    Html,
}

impl SaveFormat {
    /// wdFormat*。
    pub fn code(self) -> Option<i32> {
        Some(match self {
            Self::Keep => return None,
            Self::Docx => 16,
            Self::Doc => 0,
            Self::Rtf => 6,
            Self::Txt => 2,
            Self::Html => 8,
        })
    }
}

/// 关闭文档策略。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ClosePolicy {
    /// 有未保存修改则拒绝。
    RefuseIfDirty,
    /// 保存后关闭。
    Save,
    /// 丢弃修改后关闭（高风险，需授权）。
    Discard,
}

/// 引擎切换的脏文档策略。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DirtyPolicy {
    RefuseIfDirty,
    Save,
    Discard,
}

/// 保存结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct SaveResult {
    pub document_id: String,
    pub saved: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
}

/// 关闭结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CloseResult {
    pub document_id: String,
    pub closed: bool,
}

/// PDF 导出结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ExportResult {
    pub path: String,
    pub bytes: u64,
}

// ───────────────────────── 读取 ─────────────────────────

/// 读取视图。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ReadView {
    Summary,
    Outline,
    Text,
    Paragraphs,
    TablesSummary,
    Table,
    Selection,
    Comments,
    Revisions,
    Metadata,
}

/// 读取请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReadRequest {
    pub document_id: String,
    pub target: Target,
    #[serde(default = "default_view")]
    pub view: ReadView,
    #[serde(default = "default_limit_chars")]
    pub limit_chars: usize,
    #[serde(default = "default_max_items")]
    pub max_items: usize,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cursor: Option<String>,
}

fn default_view() -> ReadView {
    ReadView::Summary
}
fn default_limit_chars() -> usize {
    8000
}
fn default_max_items() -> usize {
    40
}

/// 文本块。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Chunk {
    pub start: i32,
    pub end: i32,
    pub text: String,
    pub hash: String,
}

/// 大纲项。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct OutlineItem {
    pub level: u32,
    pub text: String,
    pub paragraph_index: u32,
}

/// 表格摘要。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableSummary {
    pub table_id: String,
    pub index: u32,
    pub rows: u32,
    pub columns: u32,
    pub preview: String,
}

/// 读取结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReadResult {
    pub document_id: String,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub summary: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub outline: Vec<OutlineItem>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub chunks: Vec<Chunk>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub paragraphs: Vec<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<TableSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selection: Option<TextAnchor>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<CommentInfo>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub revisions: Vec<RevisionInfo>,
    pub truncated: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next_cursor: Option<String>,
    /// 覆盖范围（main/header/footer/...）。
    pub coverage: Vec<Story>,
    /// 本次读取涉及的锚点。
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub anchors: Vec<TextAnchor>,
}

// ───────────────────────── 编辑 ─────────────────────────

/// 插入位置。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum InsertPosition {
    #[default]
    Start,
    End,
    Before,
    After,
}

/// 编辑动作。
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum EditAction {
    InsertText {
        target: Target,
        text: String,
        #[serde(default)]
        position: InsertPosition,
    },
    ReplaceText {
        target: Target,
        text: String,
    },
    DeleteRange {
        target: Target,
        #[serde(default)]
        include_paragraph_mark: bool,
    },
    InsertParagraph {
        target: Target,
        text: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        style: Option<String>,
        #[serde(default)]
        position: InsertPosition,
    },
    DeleteParagraph {
        target: Target,
    },
    InsertPageBreak {
        target: Target,
    },
    InsertSectionBreak {
        target: Target,
    },
    SetBookmark {
        target: Target,
        bookmark: String,
    },
    DeleteBookmark {
        document_id: String,
        bookmark: String,
    },
    InsertHyperlink {
        target: Target,
        address: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        text: Option<String>,
    },
    InsertField {
        target: Target,
        field: FieldKind,
    },
    ApplyStyle {
        target: Target,
        style: String,
    },
}

/// 字段类型。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum FieldKind {
    Page,
    NumPages,
    Date,
    FileName,
    Time,
}

impl FieldKind {
    /// wdField* 常量。
    pub fn code(self) -> i32 {
        match self {
            Self::Page => 33,
            Self::NumPages => 26,
            Self::Date => 31,
            Self::FileName => 29,
            Self::Time => 32,
        }
    }
}

/// 编辑请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct EditRequest {
    pub document_id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub expected_doc_version: Option<u64>,
    #[serde(default = "default_true")]
    pub validate_all_first: bool,
    #[serde(default)]
    pub track_changes: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub idempotency_key: Option<String>,
    pub actions: Vec<EditAction>,
}

fn default_true() -> bool {
    true
}

/// 单动作结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ActionResult {
    pub index: usize,
    pub op: String,
    pub ok: bool,
    pub summary: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target: Option<TextAnchor>,
    pub reversible: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_code: Option<String>,
}

/// 编辑结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct EditResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    pub applied: usize,
    pub failed: usize,
    pub side_effect_status: SideEffectStatus,
    pub actions: Vec<ActionResult>,
}

// ───────────────────────── 查找 / 替换 ─────────────────────────

/// 查找选项。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct FindOptions {
    #[serde(default)]
    pub match_case: bool,
    #[serde(default)]
    pub whole_word: bool,
    #[serde(default)]
    pub wildcards: bool,
}

/// 单个匹配。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct MatchInfo {
    pub occurrence: u32,
    pub start: i32,
    pub end: i32,
    pub quote: String,
    pub text_hash: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_index: Option<u32>,
}

/// 查找请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FindRequest {
    pub document_id: String,
    pub query: String,
    #[serde(default)]
    pub options: FindOptions,
    #[serde(default = "default_limit_matches")]
    pub limit: u32,
}

fn default_limit_matches() -> u32 {
    50
}

/// 查找结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FindResult {
    pub document_id: String,
    pub doc_version: u64,
    pub matches: Vec<MatchInfo>,
    pub total: u32,
    pub truncated: bool,
}

/// 替换预览请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReplacePreviewRequest {
    pub document_id: String,
    pub find: String,
    pub replace: String,
    #[serde(default)]
    pub options: FindOptions,
    #[serde(default = "default_limit_matches")]
    pub limit: u32,
}

/// 替换预览结果（不改文档）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReplacePreview {
    pub document_id: String,
    pub preview_id: String,
    pub doc_version: u64,
    pub total: u32,
    pub samples: Vec<MatchInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub max_replacements: Option<u32>,
}

/// 替换执行请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReplaceApplyRequest {
    pub document_id: String,
    pub preview_id: String,
    pub find: String,
    pub replace: String,
    #[serde(default)]
    pub options: FindOptions,
    pub expected_doc_version: u64,
    #[serde(default = "default_max_replacements")]
    pub max_replacements: u32,
    #[serde(default)]
    pub track_changes: bool,
}

fn default_max_replacements() -> u32 {
    500
}

/// 替换执行结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReplaceResult {
    pub document_id: String,
    pub operation_id: String,
    pub replaced: u32,
    pub doc_version: u64,
    pub side_effect_status: SideEffectStatus,
}

// ───────────────────────── 格式 ─────────────────────────

/// 格式请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FormatRequest {
    pub document_id: String,
    pub target: Target,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub expected_doc_version: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<FontPatch>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph: Option<ParagraphPatch>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list: Option<ListKind>,
    /// `list=Custom` 时的自定义多级编号模板。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_template: Option<ListTemplateSpec>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reset: Option<ResetSpec>,
}

/// 格式结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct FormatResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target: Option<TextAnchor>,
    pub applied: Vec<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub warnings: Vec<String>,
}

// ───────────────────────── 表格 ─────────────────────────

/// 表格动作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TableAction {
    Create,
    Inspect,
    ReadCells,
    SetCells,
    InsertRow,
    DeleteRow,
    InsertColumn,
    DeleteColumn,
    MergeCells,
    SplitCell,
    SetWidths,
    SetHeights,
    SetAlignment,
    SetBorders,
    SetShading,
    SetRepeatHeader,
    SetAllowRowBreak,
    Autofit,
    DeleteTable,
}

/// 一维/二维数据区域。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct CellRange {
    pub row_start: u32,
    pub row_end: u32,
    pub col_start: u32,
    pub col_end: u32,
}

/// 表格请求（统一 envelope）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableRequest {
    pub document_id: String,
    pub action: TableAction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Target>,
    /// 创建表格的二维数据；或 set_cells 的填充数据。
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data: Vec<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<CellRange>,
    #[serde(default)]
    pub header_rows: u32,
    /// V1 表格格式（列宽/行高/边框/跨页行）；对应 `set_*` 动作。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<TableFormatSpec>,
}

/// 表格读取结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub info: Option<TableSummary>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub cells: Vec<Vec<String>>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
}

// ───────────────────────── 页面 / 页眉页脚 / PDF ─────────────────────────

/// 页面动作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum LayoutAction {
    SetPageSetup,
    SetHeaderFooter,
    InsertPageNumber,
    ExportPdf,
    UpdateFields,
}

/// 纸张大小。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PaperSize {
    A4,
    A3,
    A5,
    B4,
    B5,
    Letter,
    Legal,
}

impl PaperSize {
    /// wdPaper* 常量。
    pub fn code(self) -> i32 {
        match self {
            Self::A4 => 7,
            Self::A3 => 6,
            Self::A5 => 11,
            Self::B4 => 12,
            Self::B5 => 13,
            Self::Letter => 2,
            Self::Legal => 4,
        }
    }
}

/// 纸张方向。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Orientation {
    Portrait,
    Landscape,
}

impl Orientation {
    /// wdOrientPortrait=0 / wdOrientLandscape=1。
    pub fn code(self) -> i32 {
        match self {
            Self::Portrait => 0,
            Self::Landscape => 1,
        }
    }
}

/// 页面设置（mm 进入，内部转 pt）。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct PageSetupPatch {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<PaperSize>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<Orientation>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_top_mm: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_bottom_mm: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_left_mm: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_right_mm: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gutter_mm: Option<f32>,
}

/// 页眉页脚设置。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct HeaderFooterSpec {
    pub section: u32,
    pub kind: HfKind,
    #[serde(default)]
    pub variant: HfVariant,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<ParagraphAlignment>,
    #[serde(default)]
    pub link_to_previous: bool,
}

/// 布局请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct LayoutRequest {
    pub document_id: String,
    pub action: LayoutAction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page: Option<PageSetupPatch>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<HeaderFooterSpec>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pdf_path: Option<String>,
    #[serde(default)]
    pub overwrite: bool,
}

/// 布局结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct LayoutResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub exported: Option<ExportResult>,
}

// ───────────────────────── 审阅 ─────────────────────────

/// 修订信息。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct RevisionInfo {
    pub revision_ref: String,
    pub index: u32,
    #[serde(rename = "type")]
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub quote: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<i32>,
    pub text_hash: String,
    pub doc_version: u64,
}

/// 批注信息。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CommentInfo {
    pub comment_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub quote: Option<String>,
}

/// 审阅动作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ReviewAction {
    Status,
    SetTracking,
    ListRevisions,
    AcceptRevision,
    RejectRevision,
    AcceptAll,
    RejectAll,
    /// 按 `filter` 批量接受（高风险，需 confirm）。
    AcceptFiltered,
    /// 按 `filter` 批量拒绝（高风险，需 confirm）。
    RejectFiltered,
    AddComment,
    ListComments,
    DeleteComment,
}

/// 审阅请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReviewRequest {
    pub document_id: String,
    pub action: ReviewAction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub enabled: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision_ref: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Target>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment_id: Option<String>,
    /// 高风险动作的显式确认令牌（accept_all/reject_all/accept_filtered/reject_filtered）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub confirm: Option<String>,
    /// 修订筛选（accept_filtered/reject_filtered 用）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filter: Option<RevisionFilter>,
    #[serde(default = "default_limit_matches")]
    pub limit: u32,
}

/// 审阅结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ReviewResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    pub tracking: bool,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub revisions: Vec<RevisionInfo>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<CommentInfo>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comment: Option<CommentInfo>,
    /// 筛选批量时跳过（不匹配/已达上限/单条失败）的条数。
    #[serde(default, skip_serializing_if = "u32_is_zero")]
    pub skipped: u32,
}

fn u32_is_zero(v: &u32) -> bool {
    *v == 0
}

// ───────────────────────── 历史 ─────────────────────────

/// 历史条目。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct HistoryEntry {
    pub operation_id: String,
    pub kind: String,
    pub document_id: String,
    pub before_version: u64,
    pub after_version: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before_hash: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after_hash: Option<String>,
    pub reversible: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reason_not_reversible: Option<String>,
}

/// 历史结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct HistoryResult {
    pub document_id: String,
    pub operation_id: String,
    pub entries: Vec<HistoryEntry>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checkpoint_path: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
}

/// 历史动作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HistoryAction {
    List,
    Undo,
    Redo,
    Checkpoint,
    Clear,
    /// 开始一段自定义撤销记录（`Application.UndoRecord.StartCustomRecord`）：
    /// 之后到 `EndMark` 之前的所有改动在原生撤销栈里合成一步。
    BeginMark,
    /// 结束自定义撤销记录（`Application.UndoRecord.EndCustomRecord`）。
    EndMark,
}

/// 历史请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct HistoryRequest {
    pub document_id: String,
    pub action: HistoryAction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub checkpoint_dir: Option<String>,
    /// `BeginMark` 用的记录名（显示在软件自己的撤销下拉里）
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label: Option<String>,
}

// ───────────────────────── VBA ─────────────────────────

/// VBA 能力状态（只读探测）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct VbaStatus {
    pub available: bool,
    pub host_run: bool,
    pub project_access: bool,
    /// 具备临时注入的**尝试条件**（Run 可用 + VBProject 可访问）；不等于实测支持。
    pub injection_ready: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,
}

/// VBA 跑宏请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct VbaRunRequest {
    pub document_id: String,
    pub macro_name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub args: Vec<VbaArg>,
}

/// 宏标量参数。
#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum VbaArg {
    Str { value: String },
    Int { value: i32 },
    Float { value: f64 },
    Bool { value: bool },
}

/// VBA 结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct VbaResult {
    pub document_id: String,
    pub operation_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub returned: Option<String>,
    pub side_effect_status: SideEffectStatus,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
}

// ───────────────────────── 媒体（图片/形状/文本框/WordArt） ─────────────────────────

/// 媒体动作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum MediaAction {
    List,
    InsertPicture,
    InsertShape,
    InsertTextbox,
    InsertWordart,
    Resize,
    SetWrap,
    SetPosition,
    SetAltText,
    ZOrder,
    Delete,
}

/// 自动形状（`MsoAutoShapeType` 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum AutoShapeKind {
    Rectangle,
    RoundedRectangle,
    Diamond,
    Ellipse,
    Arrow,
}

impl AutoShapeKind {
    /// msoShape* 常量。
    pub fn code(self) -> i32 {
        match self {
            Self::Rectangle => 1,
            Self::RoundedRectangle => 5,
            Self::Diamond => 4,
            Self::Ellipse => 9,
            Self::Arrow => 33,
        }
    }
}

/// 文字环绕（`WdWrapType` 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WrapType {
    Inline,
    Square,
    Tight,
    Through,
    TopBottom,
    Behind,
    Front,
}

impl WrapType {
    /// wdWrap* 常量。
    pub fn code(self) -> i32 {
        match self {
            Self::Square => 0,
            Self::Tight => 1,
            Self::Through => 2,
            Self::TopBottom => 4,
            Self::Behind => 5,
            Self::Front => 6,
            Self::Inline => 7,
        }
    }
}

/// 图层顺序命令（`MsoZOrderCmd` 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ZOrderCommand {
    BringToFront,
    SendToBack,
    BringForward,
    SendBackward,
}

impl ZOrderCommand {
    pub fn code(self) -> i32 {
        match self {
            Self::BringToFront => 0,
            Self::SendToBack => 1,
            Self::BringForward => 2,
            Self::SendBackward => 3,
        }
    }
}

/// 媒体对象摘要。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct MediaItem {
    /// `shape_<n>` / `inline_<n>`。
    pub media_id: String,
    /// `inline_picture|picture|shape|textbox|wordart`。
    pub kind: String,
    pub name: String,
    pub left_pt: f32,
    pub top_pt: f32,
    pub width_pt: f32,
    pub height_pt: f32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub quote: Option<String>,
}

/// 媒体请求（统一 envelope）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct MediaRequest {
    pub document_id: String,
    pub action: MediaAction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Target>,
    /// 目标对象标识（`shape_<n>` / `inline_<n>`）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub media_id: Option<String>,
    /// 写前核对的对象名；不符返回 `stale_target`。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub expect_name: Option<String>,
    /// 图片路径（`insert_picture`，必须已存在）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    /// false（默认）= 行内图 `Range.InlineShapes`；true = 浮动图 `Document.Shapes`。
    #[serde(default)]
    pub inline: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height_pt: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape: Option<AutoShapeKind>,
    /// 形状/文本框/WordArt 的文本。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_pt: Option<f32>,
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    /// WordArt 预置效果（msoTextEffect*，默认 0）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset_text_effect: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<WrapType>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z_order: Option<ZOrderCommand>,
    #[serde(default = "default_limit_matches")]
    pub max_items: u32,
}

/// 媒体结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct MediaResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub media_id: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub items: Vec<MediaItem>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
}

// ───────────────────────── 表格格式（V1：列宽/行高/边框/跨页行） ─────────────────────────

/// 行高规则（`WdRowHeightRule`）。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum HeightRule {
    #[default]
    Auto,
    AtLeast,
    Exactly,
}

impl HeightRule {
    pub fn code(self) -> i32 {
        match self {
            Self::Auto => 0,
            Self::AtLeast => 1,
            Self::Exactly => 2,
        }
    }
}

/// 表格边线（`WdBorderType` 子集；`All` 展开为 6 条边）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TableEdge {
    Top,
    Left,
    Bottom,
    Right,
    InsideH,
    InsideV,
    All,
}

impl TableEdge {
    /// wdBorder* 常量；`All` 依次返回全部 6 条边。
    pub fn codes(self) -> Vec<i32> {
        match self {
            Self::Top => vec![-1],
            Self::Left => vec![-2],
            Self::Bottom => vec![-3],
            Self::Right => vec![-4],
            Self::InsideH => vec![-5],
            Self::InsideV => vec![-6],
            Self::All => vec![-1, -2, -3, -4, -5, -6],
        }
    }
}

/// 线型（`WdLineStyle` 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BorderStyle {
    None,
    Single,
    Double,
    Dotted,
    Dashed,
    DotDash,
    Wave,
}

impl BorderStyle {
    pub fn code(self) -> i32 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Double => 3,
            Self::Dotted => 4,
            Self::Dashed => 5,
            Self::DotDash => 6,
            Self::Wave => 11,
        }
    }
}

/// 单条边线规格。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct BorderSpec {
    pub edge: TableEdge,
    pub style: BorderStyle,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f32>,
    /// `#RRGGBB`。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// 单行行高。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct RowHeightSpec {
    /// 1 基行号。
    pub row: u32,
    pub height_pt: f32,
    #[serde(default)]
    pub rule: HeightRule,
}

/// 表格格式规格（`word_table` 的 `set_widths/set_heights/set_borders/set_allow_row_break`）。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct TableFormatSpec {
    /// 各列宽度（pt，按列序，从 `range.col_start` 或缺省第 1 列起）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub widths_pt: Option<Vec<f32>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub heights: Option<Vec<RowHeightSpec>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<Vec<BorderSpec>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_row_break: Option<bool>,
}

impl TableFormatSpec {
    /// 纯校验（不触碰 COM）。
    pub fn validate(&self) -> std::result::Result<(), String> {
        if let Some(ws) = &self.widths_pt {
            if ws.is_empty() {
                return Err("widths_pt 为空".into());
            }
            for w in ws {
                if !(1.0..=1000.0).contains(w) {
                    return Err(format!("列宽 {w} 越界（允许 1.0..=1000.0 pt）"));
                }
            }
        }
        if let Some(hs) = &self.heights {
            for h in hs {
                if h.row < 1 {
                    return Err("行号必须为 1 基".into());
                }
                if !(1.0..=2000.0).contains(&h.height_pt) {
                    return Err(format!("行高 {} 越界（允许 1.0..=2000.0 pt）", h.height_pt));
                }
            }
        }
        if let Some(bs) = &self.borders {
            for b in bs {
                if let Some(w) = b.width_pt {
                    if !(0.0..=96.0).contains(&w) {
                        return Err(format!("线宽 {w} 越界（允许 0.0..=96.0 pt）"));
                    }
                }
            }
        }
        if self.widths_pt.is_none()
            && self.heights.is_none()
            && self.borders.is_none()
            && self.allow_row_break.is_none()
        {
            return Err("format 未提供任何字段".into());
        }
        Ok(())
    }
}

// ───────────────────────── 修订筛选 ─────────────────────────

/// 修订筛选条件（全部为空 = 匹配全部）。
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct RevisionFilter {
    /// 作者白名单（大小写敏感）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub authors: Option<Vec<String>>,
    /// 类型白名单（`insert/delete/format/...`，见 `review::revision_type_name`）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kinds: Option<Vec<String>>,
    /// 最多处理条数。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<u32>,
}

// ───────────────────────── 高级（脚注尾注/交叉引用/结构 diff） ─────────────────────────

/// 高级动作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum AdvancedAction {
    ListNotes,
    AddFootnote,
    AddEndnote,
    InsertCrossReference,
    ListCrossReferences,
    StructureSnapshot,
    StructureDiff,
}

/// 注释类型。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NoteKind {
    Footnote,
    Endnote,
}

impl NoteKind {
    pub fn label(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }

    /// 集合成员名。
    pub fn collection(self) -> &'static str {
        match self {
            Self::Footnote => "Footnotes",
            Self::Endnote => "Endnotes",
        }
    }

    /// 标识前缀。
    pub fn prefix(self) -> &'static str {
        match self {
            Self::Footnote => "fn",
            Self::Endnote => "en",
        }
    }
}

/// 交叉引用类型（`WdReferenceType` 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CrossReferenceType {
    Heading,
    Bookmark,
    Footnote,
    Endnote,
    Figure,
    Table,
}

impl CrossReferenceType {
    pub fn code(self) -> i32 {
        match self {
            Self::Heading => 1,
            Self::Bookmark => 2,
            Self::Footnote => 3,
            Self::Endnote => 4,
            Self::Figure => 5,
            Self::Table => 6,
        }
    }

    /// 该类型允许的引用内容（只列已验证子集）。
    pub fn allowed_kinds(self) -> &'static [CrossReferenceKind] {
        match self {
            Self::Heading => &[CrossReferenceKind::ContentText],
            Self::Footnote | Self::Endnote => &[
                CrossReferenceKind::NumberFullContext,
                CrossReferenceKind::PageNumber,
            ],
            Self::Bookmark | Self::Figure | Self::Table => &[
                CrossReferenceKind::ContentText,
                CrossReferenceKind::PageNumber,
            ],
        }
    }
}

/// 交叉引用内容（`WdReferenceKind` 子集；未列出的不猜常量）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum CrossReferenceKind {
    ContentText,
    PageNumber,
    NumberFullContext,
}

impl CrossReferenceKind {
    pub fn code(self) -> i32 {
        match self {
            Self::ContentText => -1,
            Self::PageNumber => 7,
            Self::NumberFullContext => -4,
        }
    }
}

/// 脚注/尾注摘要。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct NoteInfo {
    /// `fn_<n>` / `en_<n>`。
    pub note_id: String,
    /// `footnote|endnote`。
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reference: Option<String>,
    pub quote: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end: Option<i32>,
}

/// 交叉引用摘要（列表中读出）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct CrossRefInfo {
    pub reference_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reference_item: Option<String>,
    pub quote: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end: Option<i32>,
}

/// 结构快照的段落项。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StructureParagraph {
    /// 1 基段落号。
    pub index: u32,
    /// 样式名（`Style.NameLocal`）。
    pub style: String,
    /// `OutlineLevel` 数值（10 = 正文）。
    pub outline_level: i32,
    pub chars: u32,
    pub quote: String,
    pub text_hash: String,
}

/// 结构快照的表格项。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableShapeSnapshot {
    pub index: u32,
    pub rows: u32,
    pub columns: u32,
    pub preview: String,
}

/// 文档结构快照（只读采集；用于 `structure_diff`）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StructureSnapshot {
    pub document_id: String,
    pub doc_version: u64,
    pub name: String,
    pub section_count: u32,
    pub paragraphs: Vec<StructureParagraph>,
    pub tables: Vec<TableShapeSnapshot>,
    pub note_count: u32,
    pub media_count: u32,
}

/// 结构差异项。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StructureChange {
    /// `paragraph_inserted|paragraph_removed|paragraph_style_changed|paragraph_text_changed|
    /// table_added|table_removed|table_shape_changed|section_count_changed|
    /// note_count_changed|media_count_changed`
    pub kind: String,
    pub index: u32,
    pub detail: String,
}

/// 高级请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct AdvancedRequest {
    pub document_id: String,
    pub action: AdvancedAction,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Target>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// 脚注/尾注的引用标记文本（缺省空串 = 自动编号）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xref_type: Option<CrossReferenceType>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xref_kind: Option<CrossReferenceKind>,
    /// 交叉引用目标项（如书签名、标题文字）。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference_item: Option<String>,
    #[serde(default)]
    pub insert_as_hyperlink: bool,
    /// `structure_diff` 的基线快照。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snapshot: Option<StructureSnapshot>,
    #[serde(default = "default_limit_matches")]
    pub max_items: u32,
}

/// 高级结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct AdvancedResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub notes: Vec<NoteInfo>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub cross_refs: Vec<CrossRefInfo>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub snapshot: Option<StructureSnapshot>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub diff: Vec<StructureChange>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
}

// ───────────────────────── AI 工作流（prepare/apply） ─────────────────────────

/// 工作流任务类型。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkflowTask {
    Proofread,
    Polish,
    Translate,
    Continue,
    Write,
    Format,
    Summarize,
    AmbiguityCheck,
    Custom,
}

impl WorkflowTask {
    /// 默认是否允许写入源文档（`summarize` 默认只回聊天）。
    pub fn default_writes_document(self) -> bool {
        !matches!(self, Self::Summarize)
    }
}

/// 写入模式。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkflowApplyMode {
    #[default]
    TrackChanges,
    Comment,
    Direct,
}

/// 单个 change 的操作。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkflowChangeOp {
    Replace,
    InsertAfter,
    Comment,
}

/// 模型生成的单个变更建议。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WorkflowChange {
    pub op: WorkflowChangeOp,
    pub target: Target,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reason: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub confidence: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub severity: Option<String>,
}

/// 工作流分块。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WorkflowChunk {
    pub chunk_id: String,
    pub start: i32,
    pub end: i32,
    pub text: String,
    pub text_hash: String,
    pub quote: String,
    pub truncated: bool,
}

/// prepare 请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WorkflowPrepareRequest {
    pub document_id: String,
    pub task: WorkflowTask,
    pub target: Target,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub instruction: Option<String>,
    #[serde(default = "default_chunk_chars")]
    pub max_chars_per_chunk: usize,
    #[serde(default = "default_overlap_chars")]
    pub overlap_chars: usize,
    #[serde(default = "default_true")]
    pub include_outline: bool,
    #[serde(default)]
    pub include_tables: bool,
    #[serde(default = "default_max_items")]
    pub max_items: usize,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cursor: Option<String>,
}

fn default_chunk_chars() -> usize {
    4000
}

fn default_overlap_chars() -> usize {
    200
}

/// prepare 结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WorkflowPrepareResult {
    pub document_id: String,
    pub task: WorkflowTask,
    pub doc_version: u64,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub outline: Vec<OutlineItem>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub chunks: Vec<WorkflowChunk>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub anchors: Vec<TextAnchor>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next_cursor: Option<String>,
    pub coverage: Vec<Story>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub notes: Vec<String>,
}

/// apply 请求。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WorkflowApplyRequest {
    pub document_id: String,
    pub task: WorkflowTask,
    pub expected_doc_version: u64,
    #[serde(default)]
    pub apply_mode: WorkflowApplyMode,
    pub changes: Vec<WorkflowChange>,
    #[serde(default = "default_true")]
    pub validate_first: bool,
    #[serde(default = "default_max_changes")]
    pub max_changes: u32,
    /// 低于该置信度的变更降级为批注（默认 0.75）。
    #[serde(default = "default_comment_confidence")]
    pub comment_below_confidence: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub confirm: Option<String>,
}

fn default_max_changes() -> u32 {
    100
}

fn default_comment_confidence() -> f32 {
    0.75
}

/// apply 结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct WorkflowApplyResult {
    pub document_id: String,
    pub operation_id: String,
    pub doc_version: u64,
    pub applied: usize,
    pub failed: usize,
    pub skipped: usize,
    pub side_effect_status: SideEffectStatus,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub results: Vec<ActionResult>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub warnings: Vec<String>,
}

// ───────────────────────── 自定义多级编号模板 ─────────────────────────

/// 编号格式（`WdListNumberStyle` 子集）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ListNumberFormat {
    Bullet,
    Decimal,
    LowerAlpha,
    UpperAlpha,
    LowerRoman,
    UpperRoman,
}

impl ListNumberFormat {
    /// wdListNumberStyle* 常量。
    pub fn code(self) -> i32 {
        match self {
            Self::Decimal => 0,
            Self::UpperRoman => 1,
            Self::LowerRoman => 2,
            Self::UpperAlpha => 3,
            Self::LowerAlpha => 4,
            Self::Bullet => 23,
        }
    }

    /// `NumberFormat` 默认文本（`%1` 为级别占位符）。
    pub fn default_text(self) -> &'static str {
        match self {
            Self::Bullet => "\u{F0B7}",
            _ => "%1.",
        }
    }
}

/// 编号后跟随符（`WdTrailingCharacter`）。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ListTrailing {
    #[default]
    Tab,
    Space,
    Nothing,
}

impl ListTrailing {
    pub fn code(self) -> i32 {
        match self {
            Self::Tab => 0,
            Self::Space => 1,
            Self::Nothing => 2,
        }
    }
}

/// 单个编号级别。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ListLevelSpec {
    /// 1..=9。
    pub level: u32,
    pub format: ListNumberFormat,
    /// `NumberFormat` 文本；缺省用 `format.default_text()`。
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default = "default_one_u32")]
    pub start_at: u32,
    #[serde(default)]
    pub number_position_pt: f32,
    #[serde(default)]
    pub text_position_pt: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_position_pt: Option<f32>,
    #[serde(default)]
    pub trailing: ListTrailing,
}

fn default_one_u32() -> u32 {
    1
}

/// 自定义多级编号模板。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct ListTemplateSpec {
    pub levels: Vec<ListLevelSpec>,
    /// 续接前一列表（false = 重新开始编号）。
    #[serde(default)]
    pub continue_previous: bool,
}

impl ListTemplateSpec {
    /// 纯校验（不触碰 COM）。
    pub fn validate(&self) -> std::result::Result<(), String> {
        if self.levels.is_empty() {
            return Err("levels 为空".into());
        }
        let mut seen = std::collections::BTreeSet::new();
        for l in &self.levels {
            if !(1..=9).contains(&l.level) {
                return Err(format!("级别 {} 越界（1..=9）", l.level));
            }
            if !seen.insert(l.level) {
                return Err(format!("级别 {} 重复", l.level));
            }
            if l.start_at < 1 {
                return Err(format!("级别 {} 的 start_at 必须 >= 1", l.level));
            }
            for (name, v) in [
                ("number_position_pt", l.number_position_pt),
                ("text_position_pt", l.text_position_pt),
            ] {
                if !v.is_finite() || v < 0.0 {
                    return Err(format!("级别 {} 的 {name} 必须为非负有限值", l.level));
                }
            }
            if let Some(t) = l.tab_position_pt {
                if !t.is_finite() || t < 0.0 {
                    return Err(format!(
                        "级别 {} 的 tab_position_pt 必须为非负有限值",
                        l.level
                    ));
                }
            }
            if l.format == ListNumberFormat::Bullet {
                let text = l.text.as_deref().unwrap_or(l.format.default_text());
                if text.chars().count() != 1 {
                    return Err(format!("级别 {} 的项目符号必须是单个字符", l.level));
                }
            }
        }
        Ok(())
    }
}

// ───────────────────────── VBA 注入 ─────────────────────────

/// VBA 动作（对外工具 action）。
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VbaAction {
    Status,
    RunExistingMacro,
    ExecuteBody,
    ExecuteModule,
    CleanupStaleTempModules,
}

/// 注入代码的形态。
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VbaCodeKind {
    /// 仅过程体，由库包成 `Public Sub <entry>()`（不允许 args）。
    #[default]
    Body,
    /// 完整标准模块源码（必须自带 `entry` 过程）。
    Module,
}

/// 临时注入请求（高风险：需 `allow_vba_injection` + `confirm="yes"`）。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct VbaExecuteRequest {
    pub document_id: String,
    #[serde(default)]
    pub kind: VbaCodeKind,
    pub entry: String,
    pub code: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub args: Vec<VbaArg>,
    #[serde(default)]
    pub allow_vba_injection: bool,
    #[serde(default = "default_true")]
    pub cleanup: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub expected_doc_version: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub confirm: Option<String>,
}

/// 临时注入结果。
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct VbaExecuteResult {
    pub document_id: String,
    pub operation_id: String,
    pub module_name: String,
    pub entry: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub returned: Option<String>,
    /// `ok|runtime_error|compile_error|unsupported` 等。
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_number: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_description: Option<String>,
    pub cleaned_up: bool,
    pub cleanup_pending: bool,
    pub side_effect_status: SideEffectStatus,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub applied: Vec<String>,
}

impl VbaCodeKind {
    /// 组装完整模块源码。
    pub fn wrap(self, entry: &str, code: &str) -> String {
        match self {
            Self::Body => format!("Option Explicit\nPublic Sub {entry}()\n{code}\nEnd Sub\n"),
            Self::Module => code.to_string(),
        }
    }
}
