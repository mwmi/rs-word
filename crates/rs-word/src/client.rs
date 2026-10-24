//! `WordClient`：`Send + Sync` 的 STA worker 代理。外部线程只通过它发 typed DTO，
//! 永不接触 `IDispatch`/`VARIANT`。

use std::sync::Mutex;
use std::thread::JoinHandle;
use std::time::Duration;

use crossbeam_channel::Sender;

use crate::error::{Category, SideEffectStatus, ToolError};
use crate::session::worker::{spawn, Command};
use crate::session::{Request, Response};
use crate::types::{
    AdvancedRequest, AdvancedResult, ClosePolicy, CloseResult, ConnectMode, DocumentInfo,
    EditRequest, EditResult, EnginePolicy, FindRequest, FindResult, FormatRequest, FormatResult,
    HistoryAction, HistoryRequest, HistoryResult, LayoutRequest, LayoutResult, MediaRequest,
    MediaResult,
    ReadRequest, ReadResult, ReplaceApplyRequest, ReplacePreview, ReplacePreviewRequest,
    ReplaceResult, ReviewRequest, ReviewResult, SaveFormat, SaveResult, TableRequest, TableResult,
    VbaExecuteRequest, VbaExecuteResult, VbaResult, VbaRunRequest, VbaStatus, WordEngine,
    WordStatus, WorkflowApplyRequest, WorkflowApplyResult, WorkflowPrepareRequest,
    WorkflowPrepareResult,
};

const D_CONNECT: Duration = Duration::from_secs(20);
const D_QUICK: Duration = Duration::from_secs(10);
const D_READ: Duration = Duration::from_secs(30);
const D_WRITE: Duration = Duration::from_secs(60);

/// 面向宿主的客户端（可跨线程 clone；内部只有 channel sender）。
pub struct WordClient {
    tx: Sender<Command>,
    handle: Mutex<Option<JoinHandle<()>>>,
}

impl WordClient {
    /// 启动 STA worker（尚未连接引擎）。
    pub fn spawn() -> Result<Self, ToolError> {
        let (tx, handle) = spawn().map_err(|e| {
            ToolError::new("worker_spawn", Category::ComFailure, e.to_string())
                .with_hint("无法启动 STA 线程")
        })?;
        Ok(Self {
            tx,
            handle: Mutex::new(Some(handle)),
        })
    }

    /// 连接引擎（同一客户端可重复调用以切换/重建会话）。
    pub fn connect(
        &self,
        policy: EnginePolicy,
        mode: ConnectMode,
    ) -> Result<WordStatus, ToolError> {
        match self.call(Request::Connect { policy, mode }, D_CONNECT)? {
            Response::Status(s) => Ok(s),
            _ => Err(bad_variant("connect")),
        }
    }

    pub fn switch_engine(
        &self,
        engine: WordEngine,
        dirty_policy: crate::types::DirtyPolicy,
    ) -> Result<WordStatus, ToolError> {
        match self.call(
            Request::SwitchEngine {
                engine,
                dirty_policy,
            },
            D_CONNECT,
        )? {
            Response::Status(s) => Ok(s),
            _ => Err(bad_variant("switch_engine")),
        }
    }

    pub fn status(&self) -> Result<WordStatus, ToolError> {
        match self.call(Request::Status, D_QUICK)? {
            Response::Status(s) => Ok(s),
            _ => Err(bad_variant("status")),
        }
    }

    pub fn list_documents(&self) -> Result<Vec<DocumentInfo>, ToolError> {
        match self.call(Request::ListDocuments, D_QUICK)? {
            Response::Documents(v) => Ok(v),
            _ => Err(bad_variant("list_documents")),
        }
    }

    pub fn open(&self, path: &str, read_only: bool) -> Result<DocumentInfo, ToolError> {
        match self.call(
            Request::Open {
                path: path.to_string(),
                read_only,
            },
            D_CONNECT,
        )? {
            Response::Document(d) => Ok(d),
            _ => Err(bad_variant("open")),
        }
    }

    pub fn new_document(&self, template: Option<&str>) -> Result<DocumentInfo, ToolError> {
        match self.call(
            Request::NewDocument {
                template: template.map(str::to_string),
            },
            D_CONNECT,
        )? {
            Response::Document(d) => Ok(d),
            _ => Err(bad_variant("new_document")),
        }
    }

    /// 活动窗口句柄（`Application.ActiveWindow.Hwnd`）：供宿主做 Win32 置前（见 doc_focus）
    pub fn hwnd(&self) -> Result<i32, ToolError> {
        match self.call(Request::Hwnd, D_QUICK)? {
            Response::Hwnd(h) => Ok(h),
            _ => Err(bad_variant("hwnd")),
        }
    }

    pub fn activate(&self, document_id: &str) -> Result<DocumentInfo, ToolError> {
        match self.call(
            Request::Activate {
                document_id: document_id.to_string(),
            },
            D_QUICK,
        )? {
            Response::Document(d) => Ok(d),
            _ => Err(bad_variant("activate")),
        }
    }

    pub fn save(
        &self,
        document_id: &str,
        path: Option<&str>,
        format: SaveFormat,
    ) -> Result<SaveResult, ToolError> {
        match self.call(
            Request::Save {
                document_id: document_id.to_string(),
                path: path.map(str::to_string),
                format,
            },
            D_WRITE,
        )? {
            Response::Save(r) => Ok(r),
            _ => Err(bad_variant("save")),
        }
    }

    pub fn close(&self, document_id: &str, policy: ClosePolicy) -> Result<CloseResult, ToolError> {
        match self.call(
            Request::Close {
                document_id: document_id.to_string(),
                policy,
            },
            D_WRITE,
        )? {
            Response::Close(r) => Ok(r),
            _ => Err(bad_variant("close")),
        }
    }

    pub fn read(&self, req: ReadRequest) -> Result<ReadResult, ToolError> {
        match self.call(Request::Read(req), D_READ)? {
            Response::Read(r) => Ok(r),
            _ => Err(bad_variant("read")),
        }
    }

    pub fn edit(&self, req: EditRequest) -> Result<EditResult, ToolError> {
        match self.call(Request::Edit(req), D_WRITE)? {
            Response::Edit(r) => Ok(r),
            _ => Err(bad_variant("edit")),
        }
    }

    pub fn find(&self, req: FindRequest) -> Result<FindResult, ToolError> {
        match self.call(Request::Find(req), D_READ)? {
            Response::Find(r) => Ok(r),
            _ => Err(bad_variant("find")),
        }
    }

    pub fn replace_preview(&self, req: ReplacePreviewRequest) -> Result<ReplacePreview, ToolError> {
        match self.call(Request::ReplacePreview(req), D_READ)? {
            Response::ReplacePreview(r) => Ok(r),
            _ => Err(bad_variant("replace_preview")),
        }
    }

    pub fn replace_apply(&self, req: ReplaceApplyRequest) -> Result<ReplaceResult, ToolError> {
        match self.call(Request::ReplaceApply(req), D_WRITE)? {
            Response::Replace(r) => Ok(r),
            _ => Err(bad_variant("replace_apply")),
        }
    }

    pub fn format(&self, req: FormatRequest) -> Result<FormatResult, ToolError> {
        match self.call(Request::Format(req), D_WRITE)? {
            Response::Format(r) => Ok(r),
            _ => Err(bad_variant("format")),
        }
    }

    pub fn table(&self, req: TableRequest) -> Result<TableResult, ToolError> {
        match self.call(Request::Table(req), D_WRITE)? {
            Response::Table(r) => Ok(r),
            _ => Err(bad_variant("table")),
        }
    }

    pub fn layout(&self, req: LayoutRequest) -> Result<LayoutResult, ToolError> {
        match self.call(Request::Layout(req), D_WRITE)? {
            Response::Layout(r) => Ok(r),
            _ => Err(bad_variant("layout")),
        }
    }

    pub fn review(&self, req: ReviewRequest) -> Result<ReviewResult, ToolError> {
        match self.call(Request::Review(req), D_WRITE)? {
            Response::Review(r) => Ok(r),
            _ => Err(bad_variant("review")),
        }
    }

    pub fn history(&self, req: HistoryRequest) -> Result<HistoryResult, ToolError> {
        match self.call(Request::History(req), D_WRITE)? {
            Response::History(r) => Ok(r),
            _ => Err(bad_variant("history")),
        }
    }

    /// 原生历史动作（清空 / 自定义撤销记录开始与结束）：与 [`history`](Self::history) 同一通道，
    /// 只省掉调用方每次拼一份 `HistoryRequest`
    pub fn history_op(
        &self,
        document_id: &str,
        action: HistoryAction,
        label: Option<&str>,
    ) -> Result<HistoryResult, ToolError> {
        self.history(HistoryRequest {
            document_id: document_id.to_string(),
            action,
            checkpoint_dir: None,
            label: label.map(str::to_string),
        })
    }

    pub fn vba_status(&self, document_id: &str) -> Result<VbaStatus, ToolError> {
        match self.call(
            Request::VbaStatus {
                document_id: document_id.to_string(),
            },
            D_QUICK,
        )? {
            Response::VbaStatus(s) => Ok(s),
            _ => Err(bad_variant("vba_status")),
        }
    }

    pub fn vba_run(&self, req: VbaRunRequest) -> Result<VbaResult, ToolError> {
        match self.call(Request::VbaRun(req), D_WRITE)? {
            Response::Vba(r) => Ok(r),
            _ => Err(bad_variant("vba_run")),
        }
    }

    pub fn media(&self, req: MediaRequest) -> Result<MediaResult, ToolError> {
        match self.call(Request::Media(req), D_WRITE)? {
            Response::Media(r) => Ok(r),
            _ => Err(bad_variant("media")),
        }
    }

    pub fn advanced(&self, req: AdvancedRequest) -> Result<AdvancedResult, ToolError> {
        match self.call(Request::Advanced(req), D_WRITE)? {
            Response::Advanced(r) => Ok(r),
            _ => Err(bad_variant("advanced")),
        }
    }

    pub fn workflow_prepare(
        &self,
        req: WorkflowPrepareRequest,
    ) -> Result<WorkflowPrepareResult, ToolError> {
        match self.call(Request::WorkflowPrepare(req), D_READ)? {
            Response::WorkflowPrepare(r) => Ok(r),
            _ => Err(bad_variant("workflow_prepare")),
        }
    }

    pub fn workflow_apply(
        &self,
        req: WorkflowApplyRequest,
    ) -> Result<WorkflowApplyResult, ToolError> {
        match self.call(Request::WorkflowApply(req), D_WRITE)? {
            Response::WorkflowApply(r) => Ok(r),
            _ => Err(bad_variant("workflow_apply")),
        }
    }

    pub fn workflow_validate(
        &self,
        req: WorkflowApplyRequest,
    ) -> Result<WorkflowApplyResult, ToolError> {
        match self.call(Request::WorkflowValidate(req), D_READ)? {
            Response::WorkflowValidate(r) => Ok(r),
            _ => Err(bad_variant("workflow_validate")),
        }
    }

    pub fn vba_execute(&self, req: VbaExecuteRequest) -> Result<VbaExecuteResult, ToolError> {
        match self.call(Request::VbaExecute(req), D_WRITE)? {
            Response::VbaExecute(r) => Ok(r),
            _ => Err(bad_variant("vba_execute")),
        }
    }

    pub fn vba_cleanup(&self, document_id: &str) -> Result<Vec<String>, ToolError> {
        match self.call(
            Request::VbaCleanup {
                document_id: document_id.to_string(),
            },
            D_WRITE,
        )? {
            Response::VbaCleanup(v) => Ok(v),
            _ => Err(bad_variant("vba_cleanup")),
        }
    }

    /// 探测 Application 是否有某成员（只读）。
    pub fn has_member(&self, member: &str) -> Result<bool, ToolError> {
        match self.call(
            Request::RawProbe {
                member: member.to_string(),
            },
            D_QUICK,
        )? {
            Response::Bool(b) => Ok(b),
            _ => Err(bad_variant("has_member")),
        }
    }

    /// 关闭 worker（不影响已连接的 Office 应用进程）。
    pub fn shutdown(&self) {
        let (tx, rx) = crossbeam_channel::bounded(1);
        let _ = self.tx.send(Command {
            req: Request::Shutdown,
            reply: tx,
        });
        let _ = rx.recv_timeout(Duration::from_secs(5));
        if let Ok(mut guard) = self.handle.lock() {
            if let Some(h) = guard.take() {
                let _ = h.join();
            }
        }
    }

    /// 发送请求并等待回复；等待超时返回 `result_unknown`（不声称取消已进入的 COM 调用）。
    fn call(&self, req: Request, deadline: Duration) -> Result<Response, ToolError> {
        let (tx, rx) = crossbeam_channel::bounded(1);
        self.tx
            .send(Command { req, reply: tx })
            .map_err(|_| bad_variant("worker_gone"))?;
        match rx.recv_timeout(deadline) {
            Ok(res) => res,
            Err(crossbeam_channel::RecvTimeoutError::Timeout) => Err(ToolError::new(
                "result_unknown",
                Category::ComFailure,
                "worker 超时：COM 调用可能仍在执行，结果未知",
            )
            .with_hint("先 word_read/status 确认是否已生效，禁止直接重发写操作")
            .with_side_effect(SideEffectStatus::Unknown)
            .with_next_action("重新读取文档状态，不要盲目重试")),
            Err(crossbeam_channel::RecvTimeoutError::Disconnected) => {
                Err(bad_variant("worker_disconnected"))
            }
        }
    }
}

impl Drop for WordClient {
    fn drop(&mut self) {
        self.shutdown();
    }
}

fn bad_variant(op: &str) -> ToolError {
    ToolError::new(
        "internal_protocol",
        Category::ComFailure,
        format!("worker 返回了非预期的响应类型（{op}）"),
    )
}
