//! 会话：STA 状态、引擎切换、文档生命周期与全部领域操作的编排。

pub mod capabilities;
pub mod engine;
pub mod worker;

use std::collections::HashMap;

use anyhow::Result;

use crate::app::App;
use crate::document::Document;
use crate::error::WordError;
use crate::history::History;
use crate::types::{
    AdvancedRequest, AdvancedResult, CapabilityInfo, ClosePolicy, CloseResult, ConnectMode,
    DirtyPolicy, DocumentInfo, EditRequest, EditResult, EngineIdentity, EnginePolicy, ExportResult,
    FindRequest, FindResult, FormatRequest, FormatResult, HistoryAction, HistoryEntry,
    HistoryRequest, HistoryResult, LayoutRequest, LayoutResult, MediaRequest, MediaResult,
    ReadRequest, ReadResult, ReplaceApplyRequest, ReplacePreview, ReplacePreviewRequest,
    ReplaceResult, ReviewRequest, ReviewResult, SaveFormat, SaveResult, TableRequest, TableResult,
    VbaExecuteRequest, VbaExecuteResult, VbaResult, VbaRunRequest, VbaStatus, WordEngine,
    WordStatus, WorkflowApplyRequest, WorkflowApplyResult, WorkflowPrepareRequest,
    WorkflowPrepareResult,
};

use engine::{adapter_for, WordEngineAdapter};

/// 文档槽位（库内稳定 ID + COM 对象）。
struct Slot {
    id: String,
    key: String,
    doc: Document,
}

/// 会话（仅在 STA worker 线程内使用）。
pub struct Session {
    pub engine: WordEngine,
    pub adapter: Box<dyn WordEngineAdapter>,
    pub prog_id: String,
    pub app: App,
    pub owned: bool,
    pub identity: Option<EngineIdentity>,
    pub caps: Vec<CapabilityInfo>,
    docs: Vec<Slot>,
    versions: HashMap<String, u64>,
    idem: HashMap<String, Response>,
    /// 本库创建且尚未清理的 VBA 临时模块（只清理这些）。
    owned_modules: Vec<String>,
    vba_seq: u64,
    pub history: History,
    next_seq: u64,
}

/// 客户端 → worker 的请求（typed，内部不序列化）。
pub enum Request {
    Connect {
        policy: EnginePolicy,
        mode: ConnectMode,
    },
    SwitchEngine {
        engine: WordEngine,
        dirty_policy: DirtyPolicy,
    },
    Status,
    ListDocuments,
    Open {
        path: String,
        read_only: bool,
    },
    NewDocument {
        template: Option<String>,
    },
    Activate {
        document_id: String,
    },
    Save {
        document_id: String,
        path: Option<String>,
        format: SaveFormat,
    },
    Close {
        document_id: String,
        policy: ClosePolicy,
    },
    Read(ReadRequest),
    Edit(EditRequest),
    Find(FindRequest),
    ReplacePreview(ReplacePreviewRequest),
    ReplaceApply(ReplaceApplyRequest),
    Format(FormatRequest),
    Table(TableRequest),
    Layout(LayoutRequest),
    Review(ReviewRequest),
    Media(MediaRequest),
    Advanced(AdvancedRequest),
    WorkflowPrepare(WorkflowPrepareRequest),
    WorkflowApply(WorkflowApplyRequest),
    WorkflowValidate(WorkflowApplyRequest),
    History(HistoryRequest),
    VbaStatus {
        document_id: String,
    },
    VbaRun(VbaRunRequest),
    VbaExecute(VbaExecuteRequest),
    VbaCleanup {
        document_id: String,
    },
    RawProbe {
        member: String,
    },
    Shutdown,
}

/// worker → 客户端的结果（typed）。
pub enum Response {
    Status(WordStatus),
    Documents(Vec<DocumentInfo>),
    Document(DocumentInfo),
    Save(SaveResult),
    Close(CloseResult),
    Export(ExportResult),
    Read(ReadResult),
    Edit(EditResult),
    Find(FindResult),
    ReplacePreview(ReplacePreview),
    Replace(ReplaceResult),
    Format(FormatResult),
    Table(TableResult),
    Layout(LayoutResult),
    Review(ReviewResult),
    Media(MediaResult),
    Advanced(AdvancedResult),
    WorkflowPrepare(WorkflowPrepareResult),
    WorkflowApply(WorkflowApplyResult),
    WorkflowValidate(WorkflowApplyResult),
    History(HistoryResult),
    VbaStatus(VbaStatus),
    Vba(VbaResult),
    VbaExecute(VbaExecuteResult),
    VbaCleanup(Vec<String>),
    Bool(bool),
    Unit,
}

impl Session {
    /// 建立会话（连接 + capability 探测）。
    pub fn connect(policy: EnginePolicy, mode: ConnectMode) -> Result<Self> {
        let engine = match policy {
            EnginePolicy::Explicit(e) => e,
            EnginePolicy::PreferMsWord => WordEngine::MsWord,
            EnginePolicy::PreferWps => WordEngine::WpsWriter,
        };
        let adapter = adapter_for(engine);
        let connected = engine::connect(adapter.as_ref(), mode)?;
        let app = App::new(connected.root, engine);
        let caps = capabilities::probe(&app);
        let mut s = Self {
            engine,
            adapter,
            prog_id: connected.prog_id,
            app,
            owned: connected.owned,
            identity: Some(connected.identity),
            caps,
            docs: Vec::new(),
            versions: HashMap::new(),
            idem: HashMap::new(),
            owned_modules: Vec::new(),
            vba_seq: 0,
            history: History::new(),
            next_seq: 0,
        };
        s.sync_documents()?;
        Ok(s)
    }

    /// 引擎切换：检查脏文档并按策略处理；旧 document_id 全部失效。
    pub fn switch_engine(
        &mut self,
        engine: WordEngine,
        dirty_policy: DirtyPolicy,
    ) -> Result<WordStatus> {
        if engine == self.engine {
            return self.status();
        }
        // 处理脏文档。
        let docs = self.app.documents()?;
        for d in &docs {
            if !d.is_saved().unwrap_or(true) {
                match dirty_policy {
                    DirtyPolicy::RefuseIfDirty => {
                        return Err(WordError::new(
                            "switch_engine",
                            self.app.engine().label(),
                            format!(
                                "文档 '{}' 有未保存修改，拒绝切换",
                                d.name().unwrap_or_default()
                            ),
                        )
                        .with_hint("先 save，或提供 dirty_policy=save/discard")
                        .into());
                    }
                    DirtyPolicy::Save => {
                        d.save()?;
                    }
                    DirtyPolicy::Discard => {}
                }
            }
        }
        // 重建会话（丢弃旧 COM 引用）。
        let adapter = adapter_for(engine);
        let connected = engine::connect(adapter.as_ref(), ConnectMode::Attach)?;
        self.engine = engine;
        self.adapter = adapter;
        self.prog_id = connected.prog_id;
        self.app = App::new(connected.root, engine);
        self.owned = connected.owned;
        self.identity = Some(connected.identity);
        self.caps = capabilities::probe(&self.app);
        self.docs.clear();
        self.versions.clear();
        self.idem.clear();
        self.owned_modules.clear();
        self.vba_seq = 0;
        self.next_seq = 0;
        self.sync_documents()?;
        self.status()
    }

    fn sync_documents(&mut self) -> Result<()> {
        let docs = self.app.documents()?;
        let mut slots = Vec::with_capacity(docs.len());
        for doc in docs {
            let full = doc.full_name().unwrap_or_default();
            let key = if full.is_empty() {
                doc.name().unwrap_or_default()
            } else {
                full
            };
            match self.docs.iter().find(|s| s.key == key) {
                Some(old) => {
                    let id = old.id.clone();
                    slots.push(Slot {
                        id: id.clone(),
                        key,
                        doc: doc.with_id(id),
                    });
                }
                None => {
                    self.next_seq += 1;
                    let id = format!("doc_{}", self.next_seq);
                    self.versions.insert(id.clone(), 0);
                    slots.push(Slot {
                        id: id.clone(),
                        key,
                        doc: doc.with_id(id),
                    });
                }
            }
        }
        self.docs = slots;
        Ok(())
    }

    fn doc(&self, id: &str) -> Result<Document> {
        self.docs
            .iter()
            .find(|s| s.id == id)
            .map(|s| s.doc.clone())
            .ok_or_else(|| {
                WordError::new("lookup", "", format!("document_id {id} 不存在或已失效"))
                    .with_hint("先 word_session.status 取最新文档列表")
                    .into()
            })
    }

    fn version(&self, id: &str) -> u64 {
        *self.versions.get(id).unwrap_or(&0)
    }

    fn bump(&mut self, id: &str) {
        *self.versions.entry(id.to_string()).or_insert(0) += 1;
    }

    fn info_of(&self, path: &str, active_key: &str) -> Option<DocumentInfo> {
        let s = self.docs.iter().find(|s| s.id == path)?;
        let d = &s.doc;
        Some(DocumentInfo {
            document_id: s.id.clone(),
            name: d.name().unwrap_or_default(),
            path: Some(d.path().unwrap_or_default()).filter(|p| !p.is_empty()),
            saved: d.is_saved().unwrap_or(true),
            read_only: d.read_only().unwrap_or(false),
            paragraph_count: d.paragraph_count(),
            table_count: d.table_count(),
            revision_count: d.revision_count(),
            is_active: s.key == active_key,
        })
    }

    pub fn status(&mut self) -> Result<WordStatus> {
        self.sync_documents()?;
        let active_key = self
            .app
            .active_document()
            .ok()
            .map(|d| {
                let full = d.full_name().unwrap_or_default();
                if full.is_empty() {
                    d.name().unwrap_or_default()
                } else {
                    full
                }
            })
            .unwrap_or_default();
        let mut documents = Vec::with_capacity(self.docs.len());
        for s in &self.docs {
            if let Some(info) = self.info_of(&s.id, &active_key) {
                documents.push(info);
            }
        }
        let active_document_id = self
            .docs
            .iter()
            .find(|s| s.key == active_key)
            .map(|s| s.id.clone());
        Ok(WordStatus {
            connected: true,
            engine: Some(self.engine),
            identity: self.identity.clone(),
            active_document_id,
            documents,
            capabilities: self.caps.clone(),
            owns_instance: self.owned,
        })
    }

    pub fn list_documents(&mut self) -> Result<Vec<DocumentInfo>> {
        Ok(self.status()?.documents)
    }

    pub fn open(&mut self, path: &str, read_only: bool) -> Result<DocumentInfo> {
        let d = self.app.open(path, read_only)?;
        self.sync_documents()?;
        let full = d.full_name().unwrap_or_default();
        let key = if full.is_empty() {
            d.name().unwrap_or_default()
        } else {
            full
        };
        self.docs
            .iter()
            .find(|s| s.key == key)
            .and_then(|s| self.info_of(&s.id, ""))
            .ok_or_else(|| WordError::new("open", key, "打开后未在文档列表中找到").into())
    }

    pub fn new_document(&mut self, template: Option<&str>) -> Result<DocumentInfo> {
        let d = self.app.add(template)?;
        self.sync_documents()?;
        let name = d.name().unwrap_or_default();
        self.docs
            .iter()
            .find(|s| s.key == name || s.doc.name().map(|n| n == name).unwrap_or(false))
            .and_then(|s| self.info_of(&s.id, ""))
            .ok_or_else(|| WordError::new("new", name, "新建后未在文档列表中找到").into())
    }

    pub fn activate(&mut self, document_id: &str) -> Result<DocumentInfo> {
        let d = self.doc(document_id)?;
        d.raw().call("Activate", vec![])?;
        self.info_of(document_id, "")
            .ok_or_else(|| WordError::new("activate", document_id, "文档已失效").into())
    }

    pub fn save(
        &mut self,
        document_id: &str,
        path: Option<&str>,
        format: SaveFormat,
    ) -> Result<SaveResult> {
        let d = self.doc(document_id)?;
        match path {
            Some(p) => d.save_as(p, format)?,
            None => d.save()?,
        }
        self.bump(document_id);
        Ok(SaveResult {
            document_id: document_id.to_string(),
            saved: true,
            path: Some(d.full_name().unwrap_or_default()).filter(|p| !p.is_empty()),
        })
    }

    pub fn close(&mut self, document_id: &str, policy: ClosePolicy) -> Result<CloseResult> {
        let d = self.doc(document_id)?;
        d.close(policy)?;
        self.docs.retain(|s| s.id != document_id);
        self.versions.remove(document_id);
        Ok(CloseResult {
            document_id: document_id.to_string(),
            closed: true,
        })
    }

    pub fn read(&mut self, req: &ReadRequest) -> Result<ReadResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        crate::read::read(&self.app, &d, req, v)
    }

    pub fn edit(&mut self, req: &EditRequest) -> Result<EditResult> {
        // 幂等：相同 key 返回首次结果。
        if let Some(key) = &req.idempotency_key {
            if let Some(Response::Edit(r)) = self.idem.get(key) {
                return Ok(r.clone());
            }
        }
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let before_hash = crate::anchors::hash_text(&d.content()?.text().unwrap_or_default());
        let result = crate::edit::edit(&self.app, &d, req, v)?;
        if result.applied > 0 {
            self.bump(&req.document_id);
            self.history.record(HistoryEntry {
                operation_id: result.operation_id.clone(),
                kind: "edit".into(),
                document_id: req.document_id.clone(),
                before_version: v,
                after_version: v + 1,
                before_hash: Some(before_hash),
                after_hash: Some(crate::anchors::hash_text(
                    &d.content()?.text().unwrap_or_default(),
                )),
                reversible: true,
                reason_not_reversible: None,
            });
        }
        if let Some(key) = &req.idempotency_key {
            self.idem
                .insert(key.clone(), Response::Edit(result.clone()));
        }
        Ok(result)
    }

    pub fn find(&mut self, req: &FindRequest) -> Result<FindResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        crate::find::find(&d, req, v)
    }

    pub fn replace_preview(&mut self, req: &ReplacePreviewRequest) -> Result<ReplacePreview> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        crate::find::preview(&d, req, v)
    }

    pub fn replace_apply(&mut self, req: &ReplaceApplyRequest) -> Result<ReplaceResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        if req.expected_doc_version != v {
            return Err(WordError::new(
                "replace",
                &req.document_id,
                format!(
                    "文档版本已变化（期望 {}，实际 {v}）：请重新 preview",
                    req.expected_doc_version
                ),
            )
            .into());
        }
        let mut r = crate::find::replace(&d, req)?;
        r.doc_version = v + u64::from(r.replaced > 0);
        if r.replaced > 0 {
            self.bump(&req.document_id);
        }
        Ok(r)
    }

    pub fn format(&mut self, req: &FormatRequest) -> Result<FormatResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        if let Some(exp) = req.expected_doc_version {
            if exp != v {
                return Err(WordError::new("format", &req.document_id, "文档版本已变化").into());
            }
        }
        let r = crate::format::apply(&self.app, &d, req, v)?;
        self.bump(&req.document_id);
        Ok(r)
    }

    pub fn table(&mut self, req: &TableRequest) -> Result<TableResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let r = crate::table::exec(&self.app, &d, req, v)?;
        if !r.applied.is_empty() {
            self.bump(&req.document_id);
        }
        Ok(r)
    }

    pub fn layout(&mut self, req: &LayoutRequest) -> Result<LayoutResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let r = crate::layout::exec(&self.app, &d, req, v)?;
        if !r.applied.is_empty() {
            self.bump(&req.document_id);
        }
        Ok(r)
    }

    pub fn review(&mut self, req: &ReviewRequest) -> Result<ReviewResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let r = crate::review::exec(&self.app, &d, req, v)?;
        if !r.applied.is_empty() {
            self.bump(&req.document_id);
        }
        Ok(r)
    }

    pub fn history(&mut self, req: &HistoryRequest) -> Result<HistoryResult> {
        let mut out = HistoryResult {
            document_id: req.document_id.clone(),
            operation_id: crate::new_operation_id(),
            entries: Vec::new(),
            checkpoint_path: None,
            applied: Vec::new(),
        };
        match req.action {
            HistoryAction::List => {
                out.entries = self.history.list(&req.document_id);
            }
            HistoryAction::Clear => {
                self.history.clear();
                out.applied.push("clear".into());
            }
            HistoryAction::Undo => {
                crate::history::undo(&self.app)?;
                self.bump(&req.document_id);
                out.applied.push("undo".into());
            }
            HistoryAction::Redo => {
                crate::history::redo(&self.app)?;
                self.bump(&req.document_id);
                out.applied.push("redo".into());
            }
            HistoryAction::Checkpoint => {
                let d = self.doc(&req.document_id)?;
                let path = crate::history::checkpoint(&d, req.checkpoint_dir.as_deref())?;
                out.applied.push("checkpoint".into());
                out.checkpoint_path = Some(path);
            }
        }
        Ok(out)
    }

    pub fn vba_status(&mut self, document_id: &str) -> Result<VbaStatus> {
        let d = self.doc(document_id)?;
        Ok(crate::vba::status(&self.app, &d))
    }

    pub fn vba_run(&mut self, req: &VbaRunRequest) -> Result<VbaResult> {
        let d = self.doc(&req.document_id)?;
        let r = crate::vba::run(&self.app, &d, req)?;
        self.bump(&req.document_id);
        Ok(r)
    }

    pub fn media(&mut self, req: &MediaRequest) -> Result<MediaResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let before = crate::anchors::hash_text(&d.content()?.text().unwrap_or_default());
        let r = crate::media::exec(&self.app, &d, req, v)?;
        if r.applied.is_empty() {
            return Ok(r);
        }
        self.bump(&req.document_id);
        self.history.record(HistoryEntry {
            operation_id: r.operation_id.clone(),
            kind: format!("media.{:?}", req.action),
            document_id: req.document_id.clone(),
            before_version: v,
            after_version: v + 1,
            before_hash: Some(before),
            after_hash: Some(crate::anchors::hash_text(
                &d.content()?.text().unwrap_or_default(),
            )),
            reversible: false,
            reason_not_reversible: Some("媒体对象需按 media_id 重新定位，不做自动撤销".into()),
        });
        Ok(r)
    }

    pub fn advanced(&mut self, req: &AdvancedRequest) -> Result<AdvancedResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let r = crate::advanced::exec(&self.app, &d, req, v)?;
        if !r.applied.is_empty() {
            self.bump(&req.document_id);
        }
        Ok(r)
    }

    pub fn workflow_prepare(
        &mut self,
        req: &WorkflowPrepareRequest,
    ) -> Result<WorkflowPrepareResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        crate::workflow::prepare(&self.app, &d, req, v)
    }

    pub fn workflow_apply(&mut self, req: &WorkflowApplyRequest) -> Result<WorkflowApplyResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        let before = crate::anchors::hash_text(&d.content()?.text().unwrap_or_default());
        let r = crate::workflow::apply(&self.app, &d, req, v)?;
        if r.applied > 0 {
            self.bump(&req.document_id);
            self.history.record(HistoryEntry {
                operation_id: r.operation_id.clone(),
                kind: format!("workflow.{:?}", req.task),
                document_id: req.document_id.clone(),
                before_version: v,
                after_version: v + 1,
                before_hash: Some(before),
                after_hash: Some(crate::anchors::hash_text(
                    &d.content()?.text().unwrap_or_default(),
                )),
                reversible: true,
                reason_not_reversible: None,
            });
        }
        Ok(r)
    }

    pub fn workflow_validate(&mut self, req: &WorkflowApplyRequest) -> Result<WorkflowApplyResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        crate::workflow::validate(&self.app, &d, req, v)
    }

    pub fn vba_execute(&mut self, req: &VbaExecuteRequest) -> Result<VbaExecuteResult> {
        let d = self.doc(&req.document_id)?;
        let v = self.version(&req.document_id);
        if let Some(exp) = req.expected_doc_version {
            if exp != v {
                return Err(
                    WordError::new("vba.execute", &req.document_id, "文档版本已变化").into(),
                );
            }
        }
        let mut owned = std::mem::take(&mut self.owned_modules);
        let mut seq = self.vba_seq;
        let result = crate::vba::execute(&self.app, &d, req, &mut owned, &mut seq);
        self.vba_seq = seq;
        self.owned_modules = owned;
        let r = result?;
        self.bump(&req.document_id);
        Ok(r)
    }

    pub fn vba_cleanup(&mut self, document_id: &str) -> Result<Vec<String>> {
        let d = self.doc(document_id)?;
        let mut owned = std::mem::take(&mut self.owned_modules);
        let removed = crate::vba::cleanup_stale(&d, &mut owned);
        self.owned_modules = owned;
        removed
    }
}

/// 处理一条请求（worker 调用）。
pub fn handle(session: &mut Option<Session>, req: Request) -> Result<Response> {
    match req {
        Request::Connect { policy, mode } => {
            *session = Some(Session::connect(policy, mode)?);
            Ok(Response::Status(session.as_mut().unwrap().status()?))
        }
        Request::Shutdown => Ok(Response::Unit),
        other => {
            let s = session.as_mut().ok_or_else(|| {
                anyhow::Error::new(
                    WordError::new(
                        "session",
                        "",
                        "尚未连接：先 connect 或直接调用其它工具（会自动 Attach）",
                    )
                    .with_hint("先 word_session.connect(engine, mode)"),
                )
            })?;
            match other {
                Request::SwitchEngine {
                    engine,
                    dirty_policy,
                } => Ok(Response::Status(s.switch_engine(engine, dirty_policy)?)),
                Request::Status => Ok(Response::Status(s.status()?)),
                Request::ListDocuments => Ok(Response::Documents(s.list_documents()?)),
                Request::Open { path, read_only } => {
                    Ok(Response::Document(s.open(&path, read_only)?))
                }
                Request::NewDocument { template } => {
                    Ok(Response::Document(s.new_document(template.as_deref())?))
                }
                Request::Activate { document_id } => {
                    Ok(Response::Document(s.activate(&document_id)?))
                }
                Request::Save {
                    document_id,
                    path,
                    format,
                } => Ok(Response::Save(s.save(
                    &document_id,
                    path.as_deref(),
                    format,
                )?)),
                Request::Close {
                    document_id,
                    policy,
                } => Ok(Response::Close(s.close(&document_id, policy)?)),
                Request::Read(r) => Ok(Response::Read(s.read(&r)?)),
                Request::Edit(r) => Ok(Response::Edit(s.edit(&r)?)),
                Request::Find(r) => Ok(Response::Find(s.find(&r)?)),
                Request::ReplacePreview(r) => Ok(Response::ReplacePreview(s.replace_preview(&r)?)),
                Request::ReplaceApply(r) => Ok(Response::Replace(s.replace_apply(&r)?)),
                Request::Format(r) => Ok(Response::Format(s.format(&r)?)),
                Request::Table(r) => Ok(Response::Table(s.table(&r)?)),
                Request::Layout(r) => Ok(Response::Layout(s.layout(&r)?)),
                Request::Review(r) => Ok(Response::Review(s.review(&r)?)),
                Request::Media(r) => Ok(Response::Media(s.media(&r)?)),
                Request::Advanced(r) => Ok(Response::Advanced(s.advanced(&r)?)),
                Request::WorkflowPrepare(r) => {
                    Ok(Response::WorkflowPrepare(s.workflow_prepare(&r)?))
                }
                Request::WorkflowApply(r) => Ok(Response::WorkflowApply(s.workflow_apply(&r)?)),
                Request::WorkflowValidate(r) => {
                    Ok(Response::WorkflowValidate(s.workflow_validate(&r)?))
                }
                Request::History(r) => Ok(Response::History(s.history(&r)?)),
                Request::VbaStatus { document_id } => {
                    Ok(Response::VbaStatus(s.vba_status(&document_id)?))
                }
                Request::VbaRun(r) => Ok(Response::Vba(s.vba_run(&r)?)),
                Request::VbaExecute(r) => Ok(Response::VbaExecute(s.vba_execute(&r)?)),
                Request::VbaCleanup { document_id } => {
                    Ok(Response::VbaCleanup(s.vba_cleanup(&document_id)?))
                }
                Request::RawProbe { member } => Ok(Response::Bool(s.app.raw().has_member(&member))),
                Request::Connect { .. } | Request::Shutdown => unreachable!(),
            }
        }
    }
}
