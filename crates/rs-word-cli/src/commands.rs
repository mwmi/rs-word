//! REPL 命令解析与执行。
//!
//! 命令的补全规格与帮助文本统一由 completion.rs 的 [COMMANDS] 表定义；
//! 新增命令时在 completion.rs 加一条表项，再在此处加一个 match 分支。

use std::cell::RefCell;

use anyhow::{bail, Context, Result};
use serde::Serialize;
use serde_json::json;

use rs_word::prelude::*;
use rs_word::process::processes;
use rs_word::{
    AdvancedRequest, ClosePolicy, ConnectMode, DirtyPolicy, EditAction, EditRequest, EnginePolicy,
    FindOptions, FindRequest, FormatRequest, HistoryAction, HistoryRequest, InsertPosition,
    LayoutAction, LayoutRequest, ListKind, MediaRequest, Orientation, PaperSize, ReadRequest,
    ReadView, ReplaceApplyRequest, ReplacePreviewRequest, ReviewAction, ReviewRequest, SaveFormat,
    Story, TableAction, TableRequest, Target, VbaCodeKind, VbaExecuteRequest, VbaRunRequest,
    WordClient, WordEngine, WordStatus, WorkflowApplyRequest, WorkflowPrepareRequest,
};

use crate::completion::{help_text, read_continuation, ReplEditor};

/// ToolError → anyhow（保留 Display 文案）。
fn ax(e: ToolError) -> anyhow::Error {
    anyhow::Error::new(e)
}

/// REPL 会话状态。
pub struct State {
    pub engine: WordEngine,
    client: Option<WordClient>,
    connected: bool,
    doc_id: Option<String>,
    doc_label: String,
}

impl State {
    pub fn new() -> Self {
        Self {
            engine: WordEngine::WpsWriter,
            client: None,
            connected: false,
            doc_id: None,
            doc_label: String::new(),
        }
    }

    pub fn prompt(&self) -> String {
        match &self.doc_id {
            None => format!("{}> ", engine_label(self.engine)),
            Some(_) => format!("{}|{}> ", engine_label(self.engine), self.doc_label),
        }
    }

    pub fn doc_id(&self) -> Option<&str> {
        self.doc_id.as_deref()
    }

    fn ensure_connected(&mut self) -> Result<()> {
        if self.client.is_none() {
            self.client = Some(WordClient::spawn().map_err(ax)?);
        }
        if !self.connected {
            let engine = self.engine;
            let c = self.client.as_ref().unwrap();
            c.connect(EnginePolicy::Explicit(engine), ConnectMode::Attach)
                .map_err(ax)
                .with_context(|| format!("无法连接 {}", engine_label(engine)))?;
            self.connected = true;
        }
        Ok(())
    }

    /// 惰性连接后的客户端（供补全动态候选调用）。
    pub fn client(&mut self) -> Option<&WordClient> {
        let _ = self.ensure_connected();
        self.client.as_ref()
    }

    /// 需已连接 + 已选文档的客户端。
    fn ctx(&mut self) -> Result<(&WordClient, String)> {
        self.ensure_connected()?;
        let id = self
            .doc_id
            .clone()
            .ok_or_else(|| anyhow::anyhow!("未选定文档：先 use <序号|名称> 或 open/new"))?;
        Ok((self.client.as_ref().unwrap(), id))
    }

    fn set_doc(&mut self, id: String, label: String) {
        self.doc_id = Some(id);
        self.doc_label = label;
    }
}

pub fn engine_label(e: WordEngine) -> &'static str {
    match e {
        WordEngine::WpsWriter => "wps",
        WordEngine::MsWord => "word",
    }
}

/// 执行一条命令。返回 false 表示退出 REPL。
pub fn execute(st: &RefCell<State>, line: &str, rl: &mut ReplEditor) -> bool {
    let cmd = first(line);
    if cmd == "quit" || cmd == "exit" {
        return false;
    }
    let args = rest(line);
    let mut state = st.borrow_mut();
    if let Err(e) = run(&mut state, cmd, args, rl) {
        println!("错误: {e:#}");
    }
    true
}

fn first(line: &str) -> &str {
    line.split_whitespace().next().unwrap_or("")
}
fn rest(line: &str) -> &str {
    let trimmed = line.trim_start();
    match trimmed.split_once(char::is_whitespace) {
        Some((_, r)) => r.trim_start(),
        None => "",
    }
}

fn show<T: Serialize>(r: Result<T, ToolError>) -> Result<()> {
    match r {
        Ok(v) => println!("{}", serde_json::to_string_pretty(&v)?),
        Err(e) => println!("失败: {e}"),
    }
    Ok(())
}

fn run(st: &mut State, cmd: &str, args: &str, rl: &mut ReplEditor) -> Result<()> {
    match cmd {
        "help" | "h" | "?" => print!("{}", help_text()),
        "engine" => {
            if args.is_empty() {
                println!("当前引擎: {}（connect 后生效）", engine_label(st.engine));
            } else {
                st.engine = match args {
                    "word" => WordEngine::MsWord,
                    "wps" => WordEngine::WpsWriter,
                    other => bail!("未知引擎 '{other}'，可选 wps|word"),
                };
                st.client = None;
                st.connected = false;
                st.doc_id = None;
                println!("已切换为 {}，重新 connect 生效", engine_label(st.engine));
            }
        }
        "connect" => {
            let mode = match args {
                "new" => ConnectMode::New,
                "" | "attach" => ConnectMode::Attach,
                other => bail!("用法: connect attach|new（实际 '{other}'）"),
            };
            if st.client.is_none() {
                st.client = Some(WordClient::spawn().map_err(ax)?);
            }
            let engine = st.engine;
            let c = st.client.as_ref().unwrap();
            let status = c
                .connect(EnginePolicy::Explicit(engine), mode)
                .map_err(ax)?;
            st.connected = true;
            let who = status
                .identity
                .as_ref()
                .map(|i| {
                    format!(
                        "（COM 自报名 {} {} / {}）",
                        i.name,
                        i.version.clone().unwrap_or_default(),
                        i.prog_id
                    )
                })
                .unwrap_or_default();
            println!("已连接 {}（{mode:?}）{who}", engine_label(engine));
        }
        "switch" => {
            let mut it = args.split_whitespace();
            let eng = match it.next() {
                Some("word") => WordEngine::MsWord,
                Some("wps") => WordEngine::WpsWriter,
                _ => bail!("用法: switch wps|word [refuse_if_dirty|save|discard]"),
            };
            let dirty = match it.next() {
                Some("save") => DirtyPolicy::Save,
                Some("discard") => DirtyPolicy::Discard,
                _ => DirtyPolicy::RefuseIfDirty,
            };
            st.ensure_connected()?;
            st.engine = eng;
            let c = st.client.as_ref().unwrap();
            let s = c.switch_engine(eng, dirty).map_err(ax)?;
            st.connected = true;
            st.doc_id = None;
            println!("已切换引擎: {}", engine_label(eng));
            print_status(&s);
        }
        "status" => {
            st.ensure_connected()?;
            let c = st.client.as_ref().unwrap();
            print_status(&c.status().map_err(ax)?);
        }
        "docs" => {
            let (c, _) = ctx_docs(st)?;
            for (i, d) in c.list_documents().map_err(ax)?.iter().enumerate() {
                println!(
                    "{}. {}  [{}] saved={} ro={} 段={} 表={} 修订={}",
                    i + 1,
                    d.name,
                    d.document_id,
                    d.saved,
                    d.read_only,
                    d.paragraph_count.unwrap_or(0),
                    d.table_count.unwrap_or(0),
                    d.revision_count.unwrap_or(0),
                );
            }
        }
        "use" => {
            let docs = ctx_docs(st)?.0.list_documents().map_err(ax)?;
            let (id, label) = pick_doc(&docs, args)?;
            st.set_doc(id, label);
            println!("已选定: {}", st.doc_label);
        }
        "open" => {
            let (c, _) = ctx_docs(st)?;
            if args.is_empty() {
                bail!("用法: open <路径>");
            }
            let d = c.open(args, false).map_err(ax)?;
            st.set_doc(d.document_id.clone(), d.name.clone());
            println!("已打开并选定: {}", d.name);
        }
        "new" => {
            let (c, _) = ctx_docs(st)?;
            let d = c.new_document(None).map_err(ax)?;
            st.set_doc(d.document_id.clone(), d.name.clone());
            println!("已新建并选定: {}", d.name);
        }
        "save" => {
            let (c, id) = st.ctx()?;
            show(c.save(&id, None, SaveFormat::Keep))?;
        }
        "saveas" => {
            let mut it = args.split_whitespace();
            let path = it
                .next()
                .ok_or_else(|| anyhow::anyhow!("用法: saveas <路径> [fmt]"))?;
            let fmt = parse_fmt(it.next());
            let (c, id) = st.ctx()?;
            show(c.save(&id, Some(path), fmt))?;
        }
        "pdf" => {
            let mut it = args.split_whitespace();
            let path = it
                .next()
                .ok_or_else(|| anyhow::anyhow!("用法: pdf <路径> [y|n]"))?;
            let overwrite = it.next() == Some("y");
            let (c, id) = st.ctx()?;
            show(c.layout(LayoutRequest {
                document_id: id,
                action: LayoutAction::ExportPdf,
                page: None,
                header_footer: None,
                toc: None,
                watermark: None,
                numbering: None,
                page_number: None,
                header_footer_zones: None,
                pdf_path: Some(path.to_string()),
                overwrite,
            }))?;
        }
        "close" => {
            let policy = match args {
                "save" => ClosePolicy::Save,
                "discard" => ClosePolicy::Discard,
                _ => ClosePolicy::RefuseIfDirty,
            };
            let (c, id) = st.ctx()?;
            show(c.close(&id, policy))?;
            st.doc_id = None;
            st.doc_label.clear();
        }
        "read" => {
            let view = match first(args) {
                "outline" => ReadView::Outline,
                "text" => ReadView::Text,
                "paragraphs" => ReadView::Paragraphs,
                "tables_summary" => ReadView::TablesSummary,
                "selection" => ReadView::Selection,
                "comments" => ReadView::Comments,
                "revisions" => ReadView::Revisions,
                "metadata" => ReadView::Metadata,
                _ => ReadView::Summary,
            };
            let limit: usize = rest(args)
                .split_whitespace()
                .next()
                .and_then(|s| s.parse().ok())
                .unwrap_or(8000);
            let (c, id) = st.ctx()?;
            show(c.read(ReadRequest {
                target: Target::Document {
                    document_id: id.clone(),
                },
                document_id: id,
                view,
                limit_chars: limit,
                max_items: 40,
                cursor: None,
            }))?;
        }
        "selection" => {
            let (c, id) = st.ctx()?;
            show(c.read(ReadRequest {
                target: Target::Selection {
                    document_id: id.clone(),
                },
                document_id: id,
                view: ReadView::Selection,
                limit_chars: 200,
                max_items: 5,
                cursor: None,
            }))?;
        }
        "find" => {
            let (c, id) = st.ctx()?;
            show(c.find(FindRequest {
                document_id: id,
                query: args.to_string(),
                options: FindOptions::default(),
                limit: 50,
            }))?;
        }
        "replace" => {
            let mut it = args.split_whitespace();
            let sub = it.next().unwrap_or("");
            let find = it.next().unwrap_or("").to_string();
            let replace = it.next().unwrap_or("").to_string();
            let (c, id) = st.ctx()?;
            match sub {
                "preview" => show(c.replace_preview(ReplacePreviewRequest {
                    document_id: id,
                    find,
                    replace,
                    options: FindOptions::default(),
                    limit: 200,
                }))?,
                "apply" => {
                    let pv = c
                        .replace_preview(ReplacePreviewRequest {
                            document_id: id.clone(),
                            find: find.clone(),
                            replace: replace.clone(),
                            options: FindOptions::default(),
                            limit: 500,
                        })
                        .map_err(ax)?;
                    show(c.replace_apply(ReplaceApplyRequest {
                        document_id: id,
                        preview_id: pv.preview_id,
                        find,
                        replace,
                        options: FindOptions::default(),
                        expected_doc_version: pv.doc_version,
                        max_replacements: 500,
                        track_changes: false,
                    }))?;
                }
                _ => bail!("用法: replace preview|apply <旧> <新>"),
            }
        }
        "edit" => edit(st, args)?,
        "format" => format_cmd(st, args)?,
        "list" => {
            let mut it = args.split_whitespace();
            let start: i32 = it.next().and_then(|s| s.parse().ok()).unwrap_or(0);
            let end: i32 = it.next().and_then(|s| s.parse().ok()).unwrap_or(start);
            let kind = match it.next() {
                Some("bullet") => ListKind::Bullet,
                Some("number") => ListKind::Number,
                Some("outline_number") => ListKind::OutlineNumber,
                _ => ListKind::None,
            };
            let (c, id) = st.ctx()?;
            show(c.format(FormatRequest {
                document_id: id.clone(),
                target: Target::Range {
                    document_id: id,
                    story: Story::Main,
                    start,
                    end,
                    expected_text: None,
                    text_hash: None,
                    expected_doc_version: None,
                },
                expected_doc_version: None,
                style: None,
                font: None,
                paragraph: None,
                list: Some(kind),
                list_template: None,
                reset: None,
                profile: None,
                assignments: None,
                dry_run: false,
                reset_font: false,
            }))?;
        }
        "table" => table(st, args, rl)?,
        "page" => page(st, args)?,
        "hdr" => header_footer(st, args, "header")?,
        "ftr" => header_footer(st, args, "footer")?,
        "review" => review(st, args)?,
        "media" => {
            let (c, id) = st.ctx()?;
            media(c, &id, args)?;
        }
        "note" => {
            let (c, id) = st.ctx()?;
            note(c, &id, args)?;
        }
        "xref" => {
            let (c, id) = st.ctx()?;
            xref(c, &id, args)?;
        }
        "struct" => {
            let (c, id) = st.ctx()?;
            structure(c, &id, args)?;
        }
        "wf" => {
            let (c, id) = st.ctx()?;
            workflow(c, &id, args)?;
        }
        "fields" => {
            let (c, id) = st.ctx()?;
            show(c.layout(LayoutRequest {
                document_id: id,
                action: LayoutAction::UpdateFields,
                page: None,
                header_footer: None,
                toc: None,
                watermark: None,
                numbering: None,
                page_number: None,
                header_footer_zones: None,
                pdf_path: None,
                overwrite: false,
            }))?;
        }
        "history" => {
            let action = match args {
                "undo" => HistoryAction::Undo,
                "redo" => HistoryAction::Redo,
                "checkpoint" => HistoryAction::Checkpoint,
                "clear" => HistoryAction::Clear,
                _ => HistoryAction::List,
            };
            let (c, id) = st.ctx()?;
            show(c.history(HistoryRequest {
                document_id: id,
                action,
                checkpoint_dir: None,
                label: None,
            }))?;
        }
        "vba" => {
            let (c, id) = st.ctx()?;
            match first(args) {
                "run" => {
                    let macro_name = rest(args)
                        .split_whitespace()
                        .next()
                        .unwrap_or("")
                        .to_string();
                    show(c.vba_run(VbaRunRequest {
                        document_id: id,
                        macro_name,
                        args: Vec::new(),
                    }))?;
                }
                "cleanup" => show(c.vba_cleanup(&id))?,
                "exec" => {
                    // vba exec <entry> <file> [allow] [module] [confirm=yes]
                    let mut parts = rest(args).split_whitespace();
                    let entry = parts.next().unwrap_or("").to_string();
                    let file = parts.next().unwrap_or("");
                    let code = std::fs::read_to_string(file)
                        .with_context(|| format!("读取代码文件 {file} 失败"))?;
                    let (mut allow, mut kind, mut confirm) = (false, VbaCodeKind::Body, None);
                    for tok in parts {
                        match tok {
                            "allow" => allow = true,
                            "module" => kind = VbaCodeKind::Module,
                            t if t.starts_with("confirm=") => {
                                confirm = Some(t.trim_start_matches("confirm=").to_string())
                            }
                            _ => {}
                        }
                    }
                    show(c.vba_execute(VbaExecuteRequest {
                        document_id: id,
                        kind,
                        entry,
                        code,
                        args: Vec::new(),
                        allow_vba_injection: allow,
                        cleanup: true,
                        expected_doc_version: None,
                        confirm,
                    }))?;
                }
                _ => show(c.vba_status(&id))?,
            }
        }
        "ps" => {
            for p in processes(st.engine)? {
                println!("{p}");
            }
        }
        "kill" => {
            let pid: u32 = args.trim().parse().context("用法: kill <PID>")?;
            rs_word::process::kill(pid)?;
            println!("已结束 PID {pid}");
        }
        "killall" => {
            let n = rs_word::process::kill_all(st.engine)?;
            println!("已结束 {n} 个进程");
        }
        "appquit" => {
            st.ensure_connected()?;
            let c = st.client.as_ref().unwrap();
            c.shutdown();
            st.client = None;
            st.connected = false;
            println!("已断开（应用进程请用 killall 或手动关闭）");
        }
        "raw" => {
            println!(
                "raw 属性/方法直调请使用库 API（App::raw()）；本 MVP CLI 仅做占位。参数: {args}"
            );
        }
        other => bail!("未知命令 '{other}'，输入 help 查看"),
    }
    Ok(())
}

fn ctx_docs(st: &mut State) -> Result<(&WordClient, ())> {
    st.ensure_connected()?;
    Ok((st.client.as_ref().unwrap(), ()))
}

fn pick_doc(docs: &[DocumentInfo], sel: &str) -> Result<(String, String)> {
    if sel.is_empty() {
        bail!("用法: use <序号|名称>");
    }
    if let Ok(n) = sel.parse::<usize>() {
        if n >= 1 && n <= docs.len() {
            let d = &docs[n - 1];
            return Ok((d.document_id.clone(), d.name.clone()));
        }
    }
    docs.iter()
        .find(|d| d.name == sel)
        .map(|d| (d.document_id.clone(), d.name.clone()))
        .ok_or_else(|| anyhow::anyhow!("未找到文档 '{sel}'"))
}

fn print_status(s: &WordStatus) {
    // 引擎名用我方选择结果：WPS 会自报 Name="Microsoft Word"，不能拿它当引擎名。
    let engine = s.engine.map(engine_label).unwrap_or("-");
    let self_report = s
        .identity
        .as_ref()
        .map(|i| {
            format!(
                "（COM 自报名: {:?} 版本: {} ProgID: {}）",
                i.name,
                i.version.clone().unwrap_or_default(),
                i.prog_id
            )
        })
        .unwrap_or_default();
    println!(
        "引擎: {} {}  已连接: {}  拥有实例: {}",
        engine, self_report, s.connected, s.owns_instance
    );
    println!(
        "活动文档: {}",
        s.active_document_id.clone().unwrap_or_else(|| "-".into())
    );
    for d in &s.documents {
        println!("  {}  [{}] saved={}", d.name, d.document_id, d.saved);
    }
    let supported = s
        .capabilities
        .iter()
        .filter(|c| matches!(c.status, CapabilityStatus::Supported))
        .count();
    println!("能力: {}/{} supported", supported, s.capabilities.len());
}

fn parse_fmt(s: Option<&str>) -> SaveFormat {
    match s {
        Some("docx") => SaveFormat::Docx,
        Some("doc") => SaveFormat::Doc,
        Some("rtf") => SaveFormat::Rtf,
        Some("txt") => SaveFormat::Txt,
        Some("html") => SaveFormat::Html,
        _ => SaveFormat::Keep,
    }
}

fn edit(st: &mut State, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("");
    let (c, id) = st.ctx()?;
    let actions = match sub {
        "replace" => {
            let s: i32 = it
                .next()
                .and_then(|x| x.parse().ok())
                .ok_or_else(|| anyhow::anyhow!("用法: edit replace <start> <end> <文本>"))?;
            let e: i32 = it
                .next()
                .and_then(|x| x.parse().ok())
                .ok_or_else(|| anyhow::anyhow!("缺 end"))?;
            let text = it.collect::<Vec<_>>().join(" ");
            vec![EditAction::ReplaceText {
                target: Target::Range {
                    document_id: id.clone(),
                    story: Story::Main,
                    start: s,
                    end: e,
                    expected_text: None,
                    text_hash: None,
                    expected_doc_version: None,
                },
                text,
            }]
        }
        "insert" => {
            let s: i32 = it
                .next()
                .and_then(|x| x.parse().ok())
                .ok_or_else(|| anyhow::anyhow!("用法: edit insert <start> <end> <文本>"))?;
            let e: i32 = it.next().and_then(|x| x.parse().ok()).unwrap_or(s);
            let text = it.collect::<Vec<_>>().join(" ");
            vec![EditAction::InsertText {
                target: Target::Range {
                    document_id: id.clone(),
                    story: Story::Main,
                    start: s,
                    end: e,
                    expected_text: None,
                    text_hash: None,
                    expected_doc_version: None,
                },
                text,
                position: InsertPosition::Start,
            }]
        }
        "delete" => {
            let s: i32 = it
                .next()
                .and_then(|x| x.parse().ok())
                .ok_or_else(|| anyhow::anyhow!("用法: edit delete <start> <end>"))?;
            let e: i32 = it.next().and_then(|x| x.parse().ok()).unwrap_or(s);
            vec![EditAction::DeleteRange {
                target: Target::Range {
                    document_id: id.clone(),
                    story: Story::Main,
                    start: s,
                    end: e,
                    expected_text: None,
                    text_hash: None,
                    expected_doc_version: None,
                },
                include_paragraph_mark: false,
            }]
        }
        "append" => {
            let text = it.collect::<Vec<_>>().join(" ");
            vec![EditAction::InsertParagraph {
                target: Target::Document {
                    document_id: id.clone(),
                },
                text,
                style: None,
                position: InsertPosition::End,
            }]
        }
        _ => bail!("用法: edit replace|insert|delete|append ..."),
    };
    show(c.edit(EditRequest {
        document_id: id,
        expected_doc_version: None,
        validate_all_first: true,
        track_changes: false,
        idempotency_key: None,
        actions,
    }))
}

fn format_cmd(st: &mut State, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let s: i32 = it
        .next()
        .and_then(|x| x.parse().ok())
        .ok_or_else(|| anyhow::anyhow!("用法: format <start> <end> style=名称 | font-json"))?;
    let e: i32 = it.next().and_then(|x| x.parse().ok()).unwrap_or(s);
    let spec = args.trim().splitn(3, ' ').nth(2).unwrap_or("");
    let (style, font, list, list_template) = if let Some(name) = spec.strip_prefix("style=") {
        (Some(name.trim().to_string()), None, None, None)
    } else if spec.starts_with('{') {
        let f: rs_word::FontPatch = serde_json::from_str(spec).context("font JSON 解析失败")?;
        (None, Some(f), None, None)
    } else if let Some(k) = spec.strip_prefix("list=") {
        (None, None, Some(list_kind(k.trim())), None)
    } else if let Some(json) = spec.strip_prefix("listtemplate=") {
        let t: rs_word::ListTemplateSpec =
            serde_json::from_str(json).context("list_template JSON 解析失败")?;
        (None, None, Some(ListKind::Custom), Some(t))
    } else {
        (None, None, None, None)
    };
    let (c, id) = st.ctx()?;
    show(c.format(FormatRequest {
        document_id: id.clone(),
        target: Target::Range {
            document_id: id,
            story: Story::Main,
            start: s,
            end: e,
            expected_text: None,
            text_hash: None,
            expected_doc_version: None,
        },
        expected_doc_version: None,
        style,
        font,
        paragraph: None,
        list,
        list_template,
        reset: None,
        profile: None,
        assignments: None,
        dry_run: false,
        reset_font: false,
    }))
}

/// 命令行列表类型名 → [`ListKind`]。
fn list_kind(name: &str) -> ListKind {
    match name {
        "bullet" => ListKind::Bullet,
        "number" => ListKind::Number,
        "outline_number" => ListKind::OutlineNumber,
        "custom" => ListKind::Custom,
        _ => ListKind::None,
    }
}

fn table(st: &mut State, args: &str, rl: &mut ReplEditor) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("");
    let (c, id) = st.ctx()?;
    match sub {
        "list" => show(c.read(ReadRequest {
            target: Target::Document {
                document_id: id.clone(),
            },
            document_id: id,
            view: ReadView::TablesSummary,
            limit_chars: 200,
            max_items: 40,
            cursor: None,
        }))?,
        "show" | "inspect" => {
            let tid = it.next().unwrap_or("");
            show(c.table(TableRequest {
                document_id: id.clone(),
                action: if sub == "show" {
                    TableAction::ReadCells
                } else {
                    TableAction::Inspect
                },
                table_id: Some(tid.to_string()),
                target: None,
                data: Vec::new(),
                range: None,
                header_rows: 0,
                format: None,
            }))?;
        }
        "create" => {
            let rows: usize = it.next().and_then(|x| x.parse().ok()).unwrap_or(2);
            let cols: usize = it.next().and_then(|x| x.parse().ok()).unwrap_or(3);
            println!("输入 {rows} 行，每行 {cols} 列（制表符分隔），单独 . 结束：");
            let mut data = Vec::new();
            for _ in 0..rows {
                let line = read_continuation(rl, "| ")?;
                if line == "." {
                    break;
                }
                let cells: Vec<String> = line.split('\t').map(|s| s.to_string()).collect();
                data.push(cells);
            }
            show(c.table(TableRequest {
                document_id: id.clone(),
                action: TableAction::Create,
                table_id: None,
                target: Some(Target::Document { document_id: id }),
                data,
                range: None,
                header_rows: 1,
                format: None,
            }))?;
        }
        "delete" => {
            let tid = it.next().unwrap_or("");
            show(c.table(TableRequest {
                document_id: id,
                action: TableAction::DeleteTable,
                table_id: Some(tid.to_string()),
                target: None,
                data: Vec::new(),
                range: None,
                header_rows: 0,
                format: None,
            }))?;
        }
        "widths" => {
            let tid = it.next().unwrap_or("").to_string();
            let widths: Vec<f32> = it.filter_map(|x| x.parse().ok()).collect();
            show(c.table(TableRequest {
                document_id: id,
                action: TableAction::SetWidths,
                table_id: Some(tid),
                target: None,
                data: Vec::new(),
                range: None,
                header_rows: 0,
                format: Some(rs_word::TableFormatSpec {
                    widths_pt: Some(widths),
                    ..Default::default()
                }),
            }))?;
        }
        "heights" => {
            let tid = it.next().unwrap_or("").to_string();
            let heights = it
                .filter_map(parse_height)
                .collect::<Vec<rs_word::RowHeightSpec>>();
            show(c.table(TableRequest {
                document_id: id,
                action: TableAction::SetHeights,
                table_id: Some(tid),
                target: None,
                data: Vec::new(),
                range: None,
                header_rows: 0,
                format: Some(rs_word::TableFormatSpec {
                    heights: Some(heights),
                    ..Default::default()
                }),
            }))?;
        }
        "borders" => {
            let tid = it.next().unwrap_or("").to_string();
            let borders = it.filter_map(parse_border).collect::<Vec<_>>();
            show(c.table(TableRequest {
                document_id: id,
                action: TableAction::SetBorders,
                table_id: Some(tid),
                target: None,
                data: Vec::new(),
                range: None,
                header_rows: 0,
                format: Some(rs_word::TableFormatSpec {
                    borders: Some(borders),
                    ..Default::default()
                }),
            }))?;
        }
        "rowbreak" => {
            let tid = it.next().unwrap_or("").to_string();
            let on = !matches!(it.next(), Some("off") | Some("false"));
            show(c.table(TableRequest {
                document_id: id,
                action: TableAction::SetAllowRowBreak,
                table_id: Some(tid),
                target: None,
                data: Vec::new(),
                range: None,
                header_rows: 0,
                format: Some(rs_word::TableFormatSpec {
                    allow_row_break: Some(on),
                    ..Default::default()
                }),
            }))?;
        }
        _ => {
            bail!("用法: table list|show|inspect|create|delete|widths|heights|borders|rowbreak ...")
        }
    }
    Ok(())
}

/// `row:height_pt[:rule]`（rule = auto|atleast|exactly）→ 行高规格。
fn parse_height(t: &str) -> Option<rs_word::RowHeightSpec> {
    let mut parts = t.split(':');
    let row: u32 = parts.next()?.parse().ok()?;
    let height_pt: f32 = parts.next()?.parse().ok()?;
    let rule = match parts.next() {
        Some("atleast") => rs_word::HeightRule::AtLeast,
        Some("exactly") => rs_word::HeightRule::Exactly,
        _ => rs_word::HeightRule::Auto,
    };
    Some(rs_word::RowHeightSpec {
        row,
        height_pt,
        rule,
    })
}

/// `edge:style[:width_pt[:#color]]` → 边线规格。
fn parse_border(t: &str) -> Option<rs_word::BorderSpec> {
    let mut parts = t.split(':');
    let edge = match parts.next()? {
        "top" => rs_word::TableEdge::Top,
        "left" => rs_word::TableEdge::Left,
        "bottom" => rs_word::TableEdge::Bottom,
        "right" => rs_word::TableEdge::Right,
        "inside_h" => rs_word::TableEdge::InsideH,
        "inside_v" => rs_word::TableEdge::InsideV,
        "all" => rs_word::TableEdge::All,
        _ => return None,
    };
    let style = match parts.next()? {
        "none" => rs_word::BorderStyle::None,
        "double" => rs_word::BorderStyle::Double,
        "dotted" => rs_word::BorderStyle::Dotted,
        "dashed" => rs_word::BorderStyle::Dashed,
        "dotdash" => rs_word::BorderStyle::DotDash,
        "wave" => rs_word::BorderStyle::Wave,
        _ => rs_word::BorderStyle::Single,
    };
    Some(rs_word::BorderSpec {
        edge,
        style,
        width_pt: parts.next().and_then(|w| w.parse().ok()),
        color: parts.next().map(str::to_string),
    })
}

fn page(st: &mut State, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("");
    let (c, id) = st.ctx()?;
    let mut page = rs_word::PageSetupPatch::default();
    match sub {
        "paper" => {
            page.size = Some(match it.next() {
                Some("A3") => PaperSize::A3,
                Some("A5") => PaperSize::A5,
                Some("Letter") => PaperSize::Letter,
                Some("Legal") => PaperSize::Legal,
                Some("B4") => PaperSize::B4,
                Some("B5") => PaperSize::B5,
                _ => PaperSize::A4,
            })
        }
        "orient" => {
            page.orientation = Some(match it.next() {
                Some("landscape") => Orientation::Landscape,
                _ => Orientation::Portrait,
            })
        }
        "margins" => {
            page.margin_top_mm = it.next().and_then(|x| x.parse().ok());
            page.margin_bottom_mm = it.next().and_then(|x| x.parse().ok());
            page.margin_left_mm = it.next().and_then(|x| x.parse().ok());
            page.margin_right_mm = it.next().and_then(|x| x.parse().ok());
        }
        _ => bail!("用法: page paper|orient|margins ..."),
    }
    show(c.layout(LayoutRequest {
        document_id: id,
        action: LayoutAction::SetPageSetup,
        page: Some(page),
        header_footer: None,
        toc: None,
        watermark: None,
        numbering: None,
        page_number: None,
        header_footer_zones: None,
        pdf_path: None,
        overwrite: false,
    }))
}

fn header_footer(st: &mut State, args: &str, kind: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let section: u32 = it.next().and_then(|x| x.parse().ok()).unwrap_or(1);
    let text = it.collect::<Vec<_>>().join(" ");
    let (c, id) = st.ctx()?;
    show(c.layout(LayoutRequest {
        document_id: id,
        action: LayoutAction::SetHeaderFooter,
        page: None,
        header_footer: Some(rs_word::HeaderFooterSpec {
            section,
            kind: if kind == "header" {
                rs_word::HfKind::Header
            } else {
                rs_word::HfKind::Footer
            },
            variant: rs_word::HfVariant::Primary,
            text: Some(text),
            alignment: None,
            link_to_previous: false,
        }),
        toc: None,
        watermark: None,
        numbering: None,
        page_number: None,
        header_footer_zones: None,
        pdf_path: None,
        overwrite: false,
    }))
}

fn review(st: &mut State, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("");
    let (c, id) = st.ctx()?;
    let action = match sub {
        "on" => ReviewAction::SetTracking,
        "off" => ReviewAction::SetTracking,
        "list" => ReviewAction::ListRevisions,
        "accept" => ReviewAction::AcceptRevision,
        "reject" => ReviewAction::RejectRevision,
        "acceptfilter" => ReviewAction::AcceptFiltered,
        "rejectfilter" => ReviewAction::RejectFiltered,
        "comment" => ReviewAction::AddComment,
        "comments" => ReviewAction::ListComments,
        _ => ReviewAction::Status,
    };
    let filtered = matches!(
        action,
        ReviewAction::AcceptFiltered | ReviewAction::RejectFiltered
    );
    // 筛选批量的参数形如 `author=张三 kind=insert max=10 confirm=yes`。
    let mut filter = rs_word::RevisionFilter::default();
    let mut confirm = None;
    let mut revision_ref = None;
    for tok in it {
        if let Some(v) = tok.strip_prefix("author=") {
            filter
                .authors
                .get_or_insert_with(Vec::new)
                .push(v.to_string());
        } else if let Some(v) = tok.strip_prefix("kind=") {
            filter
                .kinds
                .get_or_insert_with(Vec::new)
                .push(v.to_string());
        } else if let Some(v) = tok.strip_prefix("max=") {
            filter.max = v.parse().ok();
        } else if let Some(v) = tok.strip_prefix("confirm=") {
            confirm = Some(v.to_string());
        } else if revision_ref.is_none() {
            revision_ref = Some(tok.to_string());
        }
    }
    show(c.review(ReviewRequest {
        document_id: id.clone(),
        action,
        enabled: Some(sub == "on"),
        revision_ref,
        target: if sub == "comment" {
            Some(Target::Document { document_id: id })
        } else {
            None
        },
        text: Some("CLI 批注".into()),
        comment_id: None,
        confirm,
        filter: filtered.then_some(filter),
        view: None,
        limit: 50,
    }))
}

/// `media`：图片/形状/文本框/WordArt（几何单位 pt）。
fn media(c: &WordClient, id: &str, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("list");
    let doc = json!({ "kind": "document", "document_id": id });
    let mut spec = json!({ "document_id": id });
    match sub {
        "list" => spec["action"] = json!("list"),
        "pic" => {
            spec["action"] = json!("insert_picture");
            spec["path"] = json!(it.next().unwrap_or(""));
            spec["width_pt"] = json!(it.next().and_then(|x| x.parse::<f32>().ok()));
            spec["inline"] = json!(matches!(it.next(), Some("inline")));
            spec["target"] = doc;
        }
        "shape" => {
            spec["action"] = json!("insert_shape");
            spec["shape"] = json!(it.next().unwrap_or("rounded_rectangle"));
            let text = it.collect::<Vec<_>>().join(" ");
            if !text.is_empty() {
                spec["text"] = json!(text);
                spec["target"] = doc;
            }
        }
        "textbox" => {
            spec["action"] = json!("insert_textbox");
            spec["text"] = json!(it.collect::<Vec<_>>().join(" "));
        }
        "wordart" => {
            spec["action"] = json!("insert_wordart");
            spec["text"] = json!(it.collect::<Vec<_>>().join(" "));
        }
        "delete" => {
            spec["action"] = json!("delete");
            spec["media_id"] = json!(it.next().unwrap_or(""));
        }
        "resize" => {
            spec["action"] = json!("resize");
            spec["media_id"] = json!(it.next().unwrap_or(""));
            spec["width_pt"] = json!(it.next().and_then(|x| x.parse::<f32>().ok()));
            spec["height_pt"] = json!(it.next().and_then(|x| x.parse::<f32>().ok()));
        }
        "wrap" => {
            spec["action"] = json!("set_wrap");
            spec["media_id"] = json!(it.next().unwrap_or(""));
            spec["wrap"] = json!(it.next().unwrap_or("inline"));
        }
        "pos" => {
            spec["action"] = json!("set_position");
            spec["media_id"] = json!(it.next().unwrap_or(""));
            spec["left_pt"] = json!(it.next().and_then(|x| x.parse::<f32>().ok()));
            spec["top_pt"] = json!(it.next().and_then(|x| x.parse::<f32>().ok()));
        }
        "alt" => {
            spec["action"] = json!("set_alt_text");
            spec["media_id"] = json!(it.next().unwrap_or(""));
            spec["alt_text"] = json!(it.collect::<Vec<_>>().join(" "));
        }
        "zorder" => {
            spec["action"] = json!("z_order");
            spec["media_id"] = json!(it.next().unwrap_or(""));
            spec["z_order"] = json!(it.next().unwrap_or("bring_to_front"));
        }
        other => bail!("未知 media 子命令 '{other}'：list|pic|shape|textbox|wordart|delete|resize|wrap|pos|alt|zorder"),
    }
    let req: MediaRequest = serde_json::from_value(spec).context("media 参数不合法")?;
    show(c.media(req))
}

/// `note`：脚注/尾注。
fn note(c: &WordClient, id: &str, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("list");
    let mut spec = json!({ "document_id": id });
    match sub {
        "list" => spec["action"] = json!("list_notes"),
        "add" => {
            let kind = it.next().unwrap_or("footnote");
            spec["action"] = json!(if kind == "endnote" {
                "add_endnote"
            } else {
                "add_footnote"
            });
            spec["text"] = json!(it.collect::<Vec<_>>().join(" "));
            spec["target"] = json!({ "kind": "document", "document_id": id });
        }
        other => bail!("未知 note 子命令 '{other}'：list | add <footnote|endnote> <文本>"),
    }
    let req: AdvancedRequest = serde_json::from_value(spec).context("note 参数不合法")?;
    show(c.advanced(req))
}

/// `xref`：交叉引用。
fn xref(c: &WordClient, id: &str, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("list");
    let spec = match sub {
        "insert" => {
            let ty = it.next().unwrap_or("heading");
            let kind = it.next().unwrap_or("content_text");
            let item = it.collect::<Vec<_>>().join(" ");
            json!({
                "document_id": id,
                "action": "insert_cross_reference",
                "xref_type": ty,
                "xref_kind": kind,
                "reference_item": item,
                "target": { "kind": "document", "document_id": id }
            })
        }
        "list" => json!({ "document_id": id, "action": "list_cross_references" }),
        other => bail!("未知 xref 子命令 '{other}'：list | insert <type> <kind> <目标项>"),
    };
    let req: AdvancedRequest = serde_json::from_value(spec).context("xref 参数不合法")?;
    show(c.advanced(req))
}

/// `struct`：结构快照与 diff（`snapshot` 输出可直接喂给 `diff`）。
fn structure(c: &WordClient, id: &str, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("snapshot");
    match sub {
        "snapshot" => {
            let req: AdvancedRequest = serde_json::from_value(json!({
                "document_id": id,
                "action": "structure_snapshot"
            }))
            .context("simple 参数不合法")?;
            let r = c.advanced(req).map_err(ax)?;
            match r.snapshot {
                Some(s) => println!("{}", serde_json::to_string_pretty(&s)?),
                None => println!("{}", serde_json::to_string_pretty(&r)?),
            }
        }
        "diff" => {
            let file = it.next().unwrap_or("");
            let text = std::fs::read_to_string(file)
                .with_context(|| format!("读取快照文件 {file} 失败"))?;
            let base: serde_json::Value =
                serde_json::from_str(&text).context("快照 JSON 解析失败")?;
            let req: AdvancedRequest = serde_json::from_value(json!({
                "document_id": id,
                "action": "structure_diff",
                "snapshot": base
            }))
            .context("diff 参数不合法")?;
            show(c.advanced(req))?;
        }
        other => bail!("未知 struct 子命令 '{other}'：snapshot | diff <snapshot.json>"),
    }
    Ok(())
}

/// `wf`：AI 工作流 prepare / validate / apply。
fn workflow(c: &WordClient, id: &str, args: &str) -> Result<()> {
    let mut it = args.split_whitespace();
    let sub = it.next().unwrap_or("prepare");
    match sub {
        "prepare" => {
            let task = it.next().unwrap_or("proofread");
            let mut spec = json!({
                "document_id": id,
                "task": task,
                "target": { "kind": "document", "document_id": id }
            });
            if let Some(chars) = it.next().and_then(|x| x.parse::<usize>().ok()) {
                spec["max_chars_per_chunk"] = json!(chars);
            }
            let req: WorkflowPrepareRequest =
                serde_json::from_value(spec).context("wf prepare 参数不合法")?;
            show(c.workflow_prepare(req))?;
        }
        "validate" | "apply" => {
            let file = it.next().unwrap_or("");
            let text = std::fs::read_to_string(file)
                .with_context(|| format!("读取工作流请求 {file} 失败"))?;
            let mut req: WorkflowApplyRequest =
                serde_json::from_str(&text).context("工作流请求 JSON 解析失败")?;
            req.document_id = id.to_string();
            if sub == "validate" {
                show(c.workflow_validate(req))?;
            } else {
                show(c.workflow_apply(req))?;
            }
        }
        other => bail!(
            "未知 wf 子命令 '{other}'：prepare [task] [max_chars] | validate|apply <request.json>"
        ),
    }
    Ok(())
}
