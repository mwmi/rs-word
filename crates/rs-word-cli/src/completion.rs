//! 表格驱动的补全 + 候选菜单 + help 生成：[COMMANDS] 表是唯一事实来源。
//!
//! **扩展方式**（新命令两步完成，帮助与补全自动生效）：
//! ① 在 [COMMANDS] 表加一条 `CommandSpec`；
//! ② 在 commands.rs 的 `run()` 加一个 match 分支。
//!
//! 补全交互（基于 reedline ColumnarMenu）：输入行下方纵向列出候选，选中项反色高亮；
//! `Tab` 打开/循环候选，`Shift+Tab` 反向，方向键导航，回车接受。
//! 命令候选上下文感知：未连接 → 已连接 → 已选文档，隐藏无法执行的命令。

use std::borrow::Cow;
use std::cell::RefCell;
use std::rc::Rc;

use reedline::{
    default_emacs_keybindings, ColumnarMenu, Completer, CompletionResult, Emacs, KeyCode,
    KeyModifiers, MenuBuilder, Prompt, PromptEditMode, PromptHistorySearch, Reedline,
    ReedlineEvent, ReedlineMenu, Signal, Span, Suggestion,
};

use rs_word::process::processes;

use crate::commands::State;

// ── 候选常量 ──

const ENGINES: &[(&str, &str)] = &[("wps", "WPS Writer"), ("word", "Microsoft Word")];
const YES_NO: &[(&str, &str)] = &[("y", "是"), ("n", "否")];
const MODES: &[(&str, &str)] = &[("attach", "接管已打开实例"), ("new", "新建实例")];
const CLEAN: &[(&str, &str)] = &[
    ("refuse_if_dirty", "有改动则拒绝"),
    ("save", "保存"),
    ("discard", "丢弃"),
];
const FORMATS: &[(&str, &str)] = &[
    ("keep", "保持原格式"),
    ("docx", "DOCX"),
    ("doc", "DOC"),
    ("rtf", "RTF"),
    ("txt", "纯文本"),
    ("html", "HTML"),
];
const READS: &[(&str, &str)] = &[
    ("summary", "摘要"),
    ("outline", "大纲"),
    ("text", "正文分块"),
    ("paragraphs", "段落列表"),
    ("tables_summary", "表格摘要"),
    ("selection", "选区锚点"),
    ("comments", "批注"),
    ("revisions", "修订"),
    ("metadata", "元数据"),
];
const LISTS: &[(&str, &str)] = &[
    ("none", "无列表"),
    ("bullet", "项目符号"),
    ("number", "编号"),
    ("outline_number", "多级编号"),
];
const HISTORY_ACTS: &[(&str, &str)] = &[
    ("list", "历史列表"),
    ("undo", "撤销"),
    ("redo", "重做"),
    ("checkpoint", "检查点"),
    ("clear", "清空"),
];
const VBA_ACTS: &[(&str, &str)] = &[
    ("status", "能力探测"),
    ("run", "运行既有宏"),
    ("exec", "注入临时代码并运行（需 allow+confirm）"),
    ("cleanup", "清理本库临时模块"),
];
const MEDIA_ACTS: &[(&str, &str)] = &[
    ("list", "列出媒体对象"),
    ("pic", "插入图片"),
    ("shape", "插入形状"),
    ("textbox", "插入文本框"),
    ("wordart", "插入艺术字"),
    ("resize", "改尺寸"),
    ("wrap", "设置环绕"),
    ("pos", "设置位置"),
    ("alt", "替代文本"),
    ("zorder", "图层顺序"),
    ("delete", "删除"),
];
const NOTE_ACTS: &[(&str, &str)] = &[("list", "列出脚注/尾注"), ("add", "添加脚注/尾注")];
const XREF_ACTS: &[(&str, &str)] = &[("list", "列出交叉引用"), ("insert", "插入交叉引用")];
const STRUCT_ACTS: &[(&str, &str)] = &[("snapshot", "结构快照"), ("diff", "与快照比较")];
const WF_TASKS: &[(&str, &str)] = &[
    ("proofread", "校对"),
    ("polish", "润色"),
    ("translate", "翻译"),
    ("continue", "续写"),
    ("write", "写作"),
    ("format", "排版"),
    ("summarize", "总结"),
    ("ambiguity_check", "歧义检查"),
    ("custom", "自定义"),
];
const RAW_KINDS: &[(&str, &str)] = &[("get", "读属性"), ("call", "调方法"), ("put", "写属性")];

// ── 命令上下文分组 ──

const DOC_CMDS: &[&str] = &[
    "docs", "use", "open", "new", "save", "saveas", "pdf", "close",
];
const CTX_CMDS: &[&str] = &[
    "read", "find", "replace", "edit", "format", "table", "hdr", "ftr", "page", "review",
    "history", "vba", "media", "note", "xref", "struct", "wf",
];

/// 命令最低上下文要求：1 = 需已选文档，0 = 无要求。
fn required_ctx(name: &str) -> u8 {
    if CTX_CMDS.contains(&name) {
        1
    } else {
        0
    }
}

const G_HELP: &str = "帮助";
const G_CONN: &str = "连接与引擎";
const G_DOC: &str = "文档";
const G_READ: &str = "读取与查找";
const G_EDIT: &str = "编辑与格式";
const G_TABLE: &str = "表格";
const G_LAYOUT: &str = "页面与审阅";
const G_VBA: &str = "VBA";
const G_WIN: &str = "进程与窗口";
const G_EXT: &str = "扩展";

// ── 命令表（单一事实来源） ──

/// 参数位补全规格（下标 = 参数位置）。
pub enum ArgSpec {
    Fixed(&'static [(&'static str, &'static str)]),
    OptFixed(&'static [(&'static str, &'static str)]),
    Dynamic(fn(&mut State) -> Vec<String>, &'static str),
    Path,
    Any(&'static str),
    OptAny(&'static str),
    /// 子命令分支：(分支名, 说明, 分支参数)
    Sub(&'static [(&'static str, &'static str, &'static [ArgSpec])]),
}

pub struct CommandSpec {
    pub group: &'static str,
    pub name: &'static str,
    pub desc: &'static str,
    pub args: &'static [ArgSpec],
}

pub static COMMANDS: &[CommandSpec] = &[
    CommandSpec {
        group: G_HELP,
        name: "help",
        desc: "本帮助",
        args: &[],
    },
    CommandSpec {
        group: G_CONN,
        name: "engine",
        desc: "查看/切换引擎（切换后需重新 connect）",
        args: &[ArgSpec::OptFixed(ENGINES)],
    },
    CommandSpec {
        group: G_CONN,
        name: "connect",
        desc: "连接引擎（attach 接管 / new 新建）",
        args: &[ArgSpec::Fixed(MODES)],
    },
    CommandSpec {
        group: G_CONN,
        name: "switch",
        desc: "切换引擎（dirty 策略）",
        args: &[ArgSpec::Fixed(ENGINES), ArgSpec::OptFixed(CLEAN)],
    },
    CommandSpec {
        group: G_CONN,
        name: "status",
        desc: "会话状态",
        args: &[],
    },
    CommandSpec {
        group: G_DOC,
        name: "docs",
        desc: "列出文档",
        args: &[],
    },
    CommandSpec {
        group: G_DOC,
        name: "use",
        desc: "选定文档",
        args: &[ArgSpec::Dynamic(doc_candidates, "<序号|名称>")],
    },
    CommandSpec {
        group: G_DOC,
        name: "open",
        desc: "打开文档并选定",
        args: &[ArgSpec::Path],
    },
    CommandSpec {
        group: G_DOC,
        name: "new",
        desc: "新建文档并选定",
        args: &[],
    },
    CommandSpec {
        group: G_DOC,
        name: "save",
        desc: "保存当前文档",
        args: &[],
    },
    CommandSpec {
        group: G_DOC,
        name: "saveas",
        desc: "另存为",
        args: &[ArgSpec::Path, ArgSpec::OptFixed(FORMATS)],
    },
    CommandSpec {
        group: G_DOC,
        name: "pdf",
        desc: "导出 PDF",
        args: &[ArgSpec::Path, ArgSpec::OptFixed(YES_NO)],
    },
    CommandSpec {
        group: G_DOC,
        name: "close",
        desc: "关闭当前文档",
        args: &[ArgSpec::OptFixed(CLEAN)],
    },
    CommandSpec {
        group: G_READ,
        name: "read",
        desc: "读取（视图）",
        args: &[ArgSpec::Fixed(READS), ArgSpec::OptAny("限制")],
    },
    CommandSpec {
        group: G_READ,
        name: "selection",
        desc: "查看鼠标选区锚点",
        args: &[],
    },
    CommandSpec {
        group: G_READ,
        name: "find",
        desc: "查找文本",
        args: &[ArgSpec::Any("文本")],
    },
    CommandSpec {
        group: G_READ,
        name: "replace",
        desc: "替换：preview / apply",
        args: &[ArgSpec::Sub(&[
            (
                "preview",
                "预览替换",
                &[ArgSpec::Any("旧文本"), ArgSpec::Any("新文本")],
            ),
            (
                "apply",
                "执行替换（限量）",
                &[
                    ArgSpec::Any("旧文本"),
                    ArgSpec::Any("新文本"),
                    ArgSpec::OptAny("上限"),
                ],
            ),
        ])],
    },
    CommandSpec {
        group: G_EDIT,
        name: "edit",
        desc: "编辑：replace/insert/delete/paragraph",
        args: &[ArgSpec::Sub(&[
            (
                "replace",
                "替换 range（start:end）",
                &[
                    ArgSpec::Any("start"),
                    ArgSpec::Any("end"),
                    ArgSpec::Any("新文本"),
                ],
            ),
            (
                "insert",
                "在 range 起点插入",
                &[
                    ArgSpec::Any("start"),
                    ArgSpec::Any("end"),
                    ArgSpec::Any("文本"),
                ],
            ),
            (
                "delete",
                "删除 range",
                &[ArgSpec::Any("start"), ArgSpec::Any("end")],
            ),
            ("append", "末尾追加段落", &[ArgSpec::Any("文本")]),
        ])],
    },
    CommandSpec {
        group: G_EDIT,
        name: "format",
        desc: "套用样式 / 字号加粗",
        args: &[
            ArgSpec::Any("start"),
            ArgSpec::Any("end"),
            ArgSpec::Any("JSON 或 style=名称"),
        ],
    },
    CommandSpec {
        group: G_EDIT,
        name: "list",
        desc: "应用列表（range + 类型）",
        args: &[
            ArgSpec::Any("start"),
            ArgSpec::Any("end"),
            ArgSpec::Fixed(LISTS),
        ],
    },
    CommandSpec {
        group: G_TABLE,
        name: "table",
        desc: "表格操作",
        args: &[ArgSpec::Sub(&[
            ("list", "列出表格", &[]),
            (
                "show",
                "读取表格",
                &[ArgSpec::Dynamic(table_candidates, "<table_id>")],
            ),
            (
                "create",
                "创建表格（TSV 多行，单独 . 结束）",
                &[ArgSpec::Any("行"), ArgSpec::Any("列")],
            ),
            (
                "inspect",
                "查看结构",
                &[ArgSpec::Dynamic(table_candidates, "<table_id>")],
            ),
            (
                "delete",
                "删除表格",
                &[ArgSpec::Dynamic(table_candidates, "<table_id>")],
            ),
            (
                "widths",
                "列宽（pt，按列序）",
                &[
                    ArgSpec::Dynamic(table_candidates, "<table_id>"),
                    ArgSpec::OptAny("宽…"),
                ],
            ),
            (
                "heights",
                "行高（row:pt[:auto|atleast|exactly]）",
                &[
                    ArgSpec::Dynamic(table_candidates, "<table_id>"),
                    ArgSpec::OptAny("高…"),
                ],
            ),
            (
                "borders",
                "边框（edge:style[:pt[:#color]]）",
                &[
                    ArgSpec::Dynamic(table_candidates, "<table_id>"),
                    ArgSpec::OptAny("边线…"),
                ],
            ),
            (
                "rowbreak",
                "行跨页 on|off",
                &[
                    ArgSpec::Dynamic(table_candidates, "<table_id>"),
                    ArgSpec::OptFixed(YES_NO),
                ],
            ),
        ])],
    },
    CommandSpec {
        group: G_LAYOUT,
        name: "page",
        desc: "页面设置（A4/方向/边距）",
        args: &[ArgSpec::Sub(&[
            (
                "paper",
                "纸张大小",
                &[ArgSpec::OptFixed(&[
                    ("A4", "A4"),
                    ("A3", "A3"),
                    ("A5", "A5"),
                    ("Letter", "Letter"),
                ])],
            ),
            (
                "orient",
                "方向",
                &[ArgSpec::OptFixed(&[
                    ("portrait", "纵"),
                    ("landscape", "横"),
                ])],
            ),
            (
                "margins",
                "页边距（mm，上 下 左 右）",
                &[
                    ArgSpec::Any("上"),
                    ArgSpec::Any("下"),
                    ArgSpec::Any("左"),
                    ArgSpec::Any("右"),
                ],
            ),
        ])],
    },
    CommandSpec {
        group: G_LAYOUT,
        name: "hdr",
        desc: "页眉文本（section 1 起）",
        args: &[ArgSpec::OptAny("section"), ArgSpec::Any("文本")],
    },
    CommandSpec {
        group: G_LAYOUT,
        name: "ftr",
        desc: "页脚文本（section 1 起）",
        args: &[ArgSpec::OptAny("section"), ArgSpec::Any("文本")],
    },
    CommandSpec {
        group: G_LAYOUT,
        name: "review",
        desc: "修订/批注",
        args: &[ArgSpec::Sub(&[
            ("status", "修订状态", &[]),
            ("on", "开启跟踪修订", &[]),
            ("off", "关闭跟踪修订", &[]),
            ("list", "修订列表", &[]),
            ("accept", "接受修订", &[ArgSpec::Any("revision_ref")]),
            ("reject", "拒绝修订", &[ArgSpec::Any("revision_ref")]),
            (
                "acceptfilter",
                "按条件批量接受（confirm=yes）",
                &[ArgSpec::OptAny("author=/kind=/max=/confirm=")],
            ),
            (
                "rejectfilter",
                "按条件批量拒绝（confirm=yes）",
                &[ArgSpec::OptAny("author=/kind=/max=/confirm=")],
            ),
            (
                "comment",
                "添加批注",
                &[
                    ArgSpec::Any("start"),
                    ArgSpec::Any("end"),
                    ArgSpec::Any("文本"),
                ],
            ),
            ("comments", "批注列表", &[]),
        ])],
    },
    CommandSpec {
        group: G_EDIT,
        name: "media",
        desc: "图片/形状/文本框/WordArt",
        args: &[ArgSpec::Fixed(MEDIA_ACTS), ArgSpec::OptAny("参数…")],
    },
    CommandSpec {
        group: G_EXT,
        name: "note",
        desc: "脚注/尾注",
        args: &[
            ArgSpec::Fixed(NOTE_ACTS),
            ArgSpec::OptAny("footnote|endnote 文本…"),
        ],
    },
    CommandSpec {
        group: G_EXT,
        name: "xref",
        desc: "交叉引用",
        args: &[
            ArgSpec::Fixed(XREF_ACTS),
            ArgSpec::OptAny("类型 内容 目标项"),
        ],
    },
    CommandSpec {
        group: G_EXT,
        name: "struct",
        desc: "结构快照 / 结构 diff",
        args: &[ArgSpec::Fixed(STRUCT_ACTS), ArgSpec::OptAny("快照文件")],
    },
    CommandSpec {
        group: G_EXT,
        name: "wf",
        desc: "AI 工作流 prepare/validate/apply",
        args: &[ArgSpec::Sub(&[
            (
                "prepare",
                "取带锚点分块",
                &[ArgSpec::OptFixed(WF_TASKS), ArgSpec::OptAny("分块字符数")],
            ),
            ("validate", "只预检", &[ArgSpec::Any("request.json")]),
            ("apply", "应用变更", &[ArgSpec::Any("request.json")]),
        ])],
    },
    CommandSpec {
        group: G_LAYOUT,
        name: "fields",
        desc: "更新字段/目录",
        args: &[],
    },
    CommandSpec {
        group: G_LAYOUT,
        name: "history",
        desc: "历史/撤销/检查点",
        args: &[ArgSpec::OptFixed(HISTORY_ACTS)],
    },
    CommandSpec {
        group: G_VBA,
        name: "vba",
        desc: "VBA 能力检测 / 运行既有宏",
        args: &[
            ArgSpec::Fixed(VBA_ACTS),
            ArgSpec::OptAny("宏名"),
            ArgSpec::OptAny("参数"),
        ],
    },
    CommandSpec {
        group: G_WIN,
        name: "ps",
        desc: "查看引擎进程（wps.exe/winword.exe）",
        args: &[],
    },
    CommandSpec {
        group: G_WIN,
        name: "kill",
        desc: "强制结束指定 PID",
        args: &[ArgSpec::Dynamic(process_candidates, "<PID>")],
    },
    CommandSpec {
        group: G_WIN,
        name: "killall",
        desc: "结束当前引擎全部进程",
        args: &[],
    },
    CommandSpec {
        group: G_WIN,
        name: "appquit",
        desc: "关闭应用进程并断开",
        args: &[],
    },
    CommandSpec {
        group: G_WIN,
        name: "quit",
        desc: "退出本工具（不影响应用）",
        args: &[],
    },
    CommandSpec {
        group: G_EXT,
        name: "raw",
        desc: "直接调用 COM 成员（测试未封装 API）",
        args: &[
            ArgSpec::Fixed(RAW_KINDS),
            ArgSpec::Any("成员"),
            ArgSpec::OptAny("参数"),
        ],
    },
    CommandSpec {
        group: "",
        name: "h",
        desc: "同 help",
        args: &[],
    },
    CommandSpec {
        group: "",
        name: "?",
        desc: "同 help",
        args: &[],
    },
    CommandSpec {
        group: "",
        name: "exit",
        desc: "同 quit",
        args: &[],
    },
];

pub fn find_command(name: &str) -> Option<&'static CommandSpec> {
    COMMANDS.iter().find(|c| c.name == name)
}

// ── help 生成（与补全同源） ──

pub fn help_text() -> String {
    let mut out =
        String::from("命令列表（提示符 = 引擎|文档；Tab 打开候选菜单，方向键导航，回车确认）：\n");
    let mut group = "";
    for c in COMMANDS {
        if c.group.is_empty() {
            continue;
        }
        if c.group != group {
            out.push_str(&format!("\n── {} ──\n", c.group));
            group = c.group;
        }
        let mut usage = c.name.to_string();
        for a in c.args {
            usage.push(' ');
            usage.push_str(&a.usage());
        }
        let pad = 34usize.saturating_sub(display_width(&usage));
        out.push_str(&format!("  {}{} {}\n", usage, " ".repeat(pad), c.desc));
    }
    out
}

fn display_width(s: &str) -> usize {
    s.chars().map(|c| if c.is_ascii() { 1 } else { 2 }).sum()
}

impl ArgSpec {
    fn usage(&self) -> String {
        match self {
            Self::Fixed(f) => f.iter().map(|(v, _)| *v).collect::<Vec<_>>().join("|"),
            Self::OptFixed(f) => {
                let v = f.iter().map(|(v, _)| *v).collect::<Vec<_>>().join("|");
                format!("[{v}]")
            }
            Self::Dynamic(_, label) => label.to_string(),
            Self::Path => "<路径>".to_string(),
            Self::Any(p) => format!("<{p}>"),
            Self::OptAny(p) => format!("[{p}]"),
            Self::Sub(branches) => branches
                .iter()
                .map(|(n, _, args)| {
                    let mut u = n.to_string();
                    for a in *args {
                        u.push(' ');
                        u.push_str(&a.usage());
                    }
                    u
                })
                .collect::<Vec<_>>()
                .join(" | "),
        }
    }
}

// ── 动态候选 ──

/// 已打开文档候选：「序号. 名称」；未连接时先连接。
fn doc_candidates(st: &mut State) -> Vec<String> {
    let Some(c) = st.client() else { return vec![] };
    match c.list_documents() {
        Ok(docs) => docs
            .iter()
            .enumerate()
            .map(|(i, d)| format!("{}. {}", i + 1, d.name))
            .collect(),
        Err(_) => vec![],
    }
}

/// 表格候选：tbl_1..n（读当前文档 table_count）。
fn table_candidates(st: &mut State) -> Vec<String> {
    let Some(id) = st.doc_id().map(str::to_string) else {
        return vec![];
    };
    let Some(c) = st.client() else { return vec![] };
    match c.status() {
        Ok(s) => s
            .documents
            .iter()
            .find(|d| d.document_id == id)
            .and_then(|d| d.table_count)
            .map(|n| (1..=n).map(|i| format!("tbl_{i}")).collect())
            .unwrap_or_default(),
        Err(_) => vec![],
    }
}

/// 引擎进程候选：「PID. 名称」。
fn process_candidates(st: &mut State) -> Vec<String> {
    match processes(st.engine) {
        Ok(ps) => ps
            .iter()
            .map(|p| format!("{}. {}", p.pid, p.name))
            .collect(),
        Err(_) => vec![],
    }
}

fn list(c: &[(&'static str, &'static str)]) -> Vec<(String, String)> {
    c.iter()
        .map(|(v, d)| (v.to_string(), format!("{v:<10}{d}")))
        .collect()
}

// ── 补全引擎 ──

enum Source {
    Commands,
    Path,
    Free,
    List(Vec<(String, String)>),
}

fn matches_candidate(c: &str, lower: &str) -> bool {
    c.to_lowercase().starts_with(lower)
        || c.split_whitespace()
            .any(|t| t.to_lowercase().starts_with(lower))
}

fn insert_value(c: &str) -> &str {
    c.split_whitespace()
        .next()
        .unwrap_or(c)
        .trim_end_matches('.')
}

/// 文件路径候选：扫描当前 token 所在目录，目录自动带 `\`。
fn path_suggestions(partial: &str, span: Span) -> Vec<Suggestion> {
    let (dir, prefix, base) = match partial.rfind(['\\', '/']) {
        Some(i) => (&partial[..=i], &partial[i + 1..], partial[..=i].to_string()),
        None => (".", partial, String::new()),
    };
    let Ok(entries) = std::fs::read_dir(dir) else {
        return vec![];
    };
    let lower = prefix.to_lowercase();
    entries
        .flatten()
        .filter(|e| {
            e.file_name()
                .to_string_lossy()
                .to_lowercase()
                .starts_with(&lower)
        })
        .map(|e| {
            let name = e.file_name().to_string_lossy().to_string();
            let is_dir = e.file_type().map(|t| t.is_dir()).unwrap_or(false);
            Suggestion {
                value: if is_dir {
                    format!("{base}{name}\\")
                } else {
                    format!("{base}{name} ")
                },
                display_override: Some(name),
                span,
                ..Default::default()
            }
        })
        .collect()
}

pub struct MenuCompleter {
    state: Rc<RefCell<State>>,
}

// SAFETY：CLI 是单线程 REPL；reedline 的事件循环与补全始终在同一线程，
// state 中的 WordClient 是 Send+Sync 代理，指针不跨线程。
unsafe impl Send for MenuCompleter {}

impl MenuCompleter {
    pub fn new(state: Rc<RefCell<State>>) -> Self {
        Self { state }
    }

    fn source_at(&self, line: &str, pos: usize) -> Option<(usize, Source)> {
        let before = &line[..pos];
        let token_start = before.rfind(char::is_whitespace).map_or(0, |i| i + 1);
        let tokens: Vec<&str> = before[..token_start].split_whitespace().collect();
        let Some(cmd_name) = tokens.first() else {
            return Some((token_start, Source::Commands));
        };
        let spec = find_command(cmd_name)?;
        let arg_index = tokens.len() - 1;
        if let Some(ArgSpec::Sub(branches)) = spec.args.first() {
            if arg_index == 0 {
                let cands = branches
                    .iter()
                    .map(|(n, d, _)| (n.to_string(), format!("{n:<14}{d}")))
                    .collect();
                return Some((token_start, Source::List(cands)));
            }
            let branch = branches.iter().find(|(n, _, _)| *n == tokens[1]);
            let Some(arg) = branch.and_then(|(_, _, a)| a.get(arg_index - 1)) else {
                return Some((token_start, Source::Free));
            };
            return Some((token_start, self.classify(arg, &tokens)));
        }
        let Some(arg) = spec.args.get(arg_index) else {
            return Some((token_start, Source::Free));
        };
        Some((token_start, self.classify(arg, &tokens)))
    }

    fn classify(&self, arg: &ArgSpec, _tokens: &[&str]) -> Source {
        match arg {
            ArgSpec::Fixed(f) | ArgSpec::OptFixed(f) => Source::List(list(f)),
            ArgSpec::Dynamic(f, _) => match self.state.try_borrow_mut() {
                Ok(mut st) => Source::List(
                    f(&mut st)
                        .into_iter()
                        .map(|c| (insert_value(&c).to_string(), c))
                        .collect(),
                ),
                Err(_) => Source::Free,
            },
            ArgSpec::Path => Source::Path,
            _ => Source::Free,
        }
    }
}

impl Completer for MenuCompleter {
    fn complete(&mut self, line: &str, pos: usize) -> CompletionResult {
        let fallback = |token_start: usize| {
            CompletionResult::fresh(vec![Suggestion {
                value: line[token_start..pos].to_string(),
                display_override: Some("（无可用候选）".to_string()),
                span: Span::new(token_start, pos),
                ..Default::default()
            }])
        };
        let Some((token_start, src)) = self.source_at(line, pos) else {
            return fallback(line[..pos].rfind(char::is_whitespace).map_or(0, |i| i + 1));
        };
        let partial = &line[token_start..pos];
        let lower = partial.to_lowercase();
        let span = Span::new(token_start, pos);
        let suggestions: Vec<Suggestion> = match src {
            Source::Free => vec![],
            Source::Path => path_suggestions(partial, span),
            Source::Commands => {
                let (ctx, preferred) = match self.state.try_borrow() {
                    Ok(s) if s.doc_id().is_some() => (1u8, CTX_CMDS),
                    _ => (0, DOC_CMDS),
                };
                let mut specs: Vec<&CommandSpec> = COMMANDS
                    .iter()
                    .filter(|c| c.name.bytes().all(|b| b.is_ascii_alphanumeric()))
                    .filter(|c| required_ctx(c.name) <= ctx)
                    .collect();
                specs.sort_by_key(|c| !preferred.contains(&c.name));
                specs
                    .iter()
                    .filter(|c| c.name.starts_with(&lower))
                    .map(|c| Suggestion {
                        value: format!("{} ", c.name),
                        display_override: Some(format!("{:<10} {}", c.name, c.desc)),
                        span,
                        ..Default::default()
                    })
                    .collect()
            }
            Source::List(cands) => cands
                .into_iter()
                .filter(|(_, disp)| matches_candidate(disp, &lower))
                .map(|(value, disp)| Suggestion {
                    value: if value.ends_with(' ') {
                        value
                    } else {
                        format!("{value} ")
                    },
                    display_override: Some(disp),
                    span,
                    ..Default::default()
                })
                .collect(),
        };
        if suggestions.is_empty() {
            return fallback(token_start);
        }
        CompletionResult::fresh(suggestions)
    }
}

// ── 行编辑器组装 ──

pub type ReplEditor = Reedline;

pub struct DynPrompt {
    pub text: String,
}

impl Prompt for DynPrompt {
    fn render_prompt_left(&self) -> Cow<'_, str> {
        Cow::Borrowed(&self.text)
    }
    fn render_prompt_right(&self) -> Cow<'_, str> {
        Cow::Borrowed("")
    }
    fn render_prompt_indicator(&self, _m: PromptEditMode) -> Cow<'_, str> {
        Cow::Borrowed("")
    }
    fn render_prompt_multiline_indicator(&self) -> Cow<'_, str> {
        Cow::Borrowed(":: ")
    }
    fn render_prompt_history_search_indicator(&self, _s: PromptHistorySearch) -> Cow<'_, str> {
        Cow::Borrowed("历史搜索: ")
    }
}

pub fn editor(state: Rc<RefCell<State>>) -> ReplEditor {
    let mut keybindings = default_emacs_keybindings();
    keybindings.add_binding(
        KeyModifiers::NONE,
        KeyCode::Tab,
        ReedlineEvent::UntilFound(vec![
            ReedlineEvent::Menu("completion_menu".to_string()),
            ReedlineEvent::MenuNext,
        ]),
    );
    let menu = ColumnarMenu::default()
        .with_name("completion_menu")
        .with_columns(1);
    Reedline::create()
        .with_completer(Box::new(MenuCompleter::new(state)))
        .with_menu(ReedlineMenu::EngineCompleter(Box::new(menu)))
        .with_edit_mode(Box::new(Emacs::new(keybindings)))
}

/// 续行输入（TSV 多行等）：无补全菜单；中断视为结束。
pub fn read_continuation(rl: &mut Reedline, prompt: &str) -> anyhow::Result<String> {
    match rl.read_line(&DynPrompt {
        text: prompt.to_string(),
    })? {
        Signal::Success(buffer) => Ok(buffer),
        _ => Ok(".".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 命令表是补全 + help 的唯一事实来源：新命令必须同时出现在表与 help 里。
    #[test]
    fn commands_table_covers_new_features() {
        let names: Vec<&str> = COMMANDS.iter().map(|c| c.name).collect();
        for n in [
            "media", "note", "xref", "struct", "wf", "vba", "table", "review",
        ] {
            assert!(names.contains(&n), "命令表缺少 {n}");
        }
        let help = help_text();
        for n in ["media", "note", "xref", "struct", "wf"] {
            assert!(help.contains(n), "help 未列出 {n}");
        }
        // 命令表里每个命令都应能通过 find_command 取到（补全依赖）。
        for n in names {
            assert!(find_command(n).is_some(), "find_command 找不到 {n}");
        }
    }

    /// 表驱动的子命令枚举：表格 V1 与修订筛选批量都已暴露。
    #[test]
    fn subcommand_enums_expose_v1_actions() {
        let table = find_command("table").unwrap();
        let usage = table.args.first().unwrap().usage();
        for act in ["widths", "heights", "borders", "rowbreak"] {
            assert!(usage.contains(act), "table 子命令缺少 {act}");
        }
        let review = find_command("review").unwrap();
        let usage = review.args.first().unwrap().usage();
        for act in ["acceptfilter", "rejectfilter"] {
            assert!(usage.contains(act), "review 子命令缺少 {act}");
        }
        let vba = find_command("vba").unwrap();
        let usage = vba.args.first().unwrap().usage();
        assert!(usage.contains("exec") && usage.contains("cleanup"));
    }
}
