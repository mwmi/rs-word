//! 应用层（Application）：文档列表/新建/打开、可见性与弹窗告警、退出。
//!
//! `App` 不持有 COM 套间守卫 —— 套间由 STA worker 线程负责（见 [`crate::session`]）。
//! 所有方法只在 worker 线程内调用。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::WordError;
use crate::types::WordEngine;

/// Word / WPS Writer 应用（Application COM 对象）。
#[derive(Clone)]
pub struct App {
    raw: IDispatchWrapper,
    engine: WordEngine,
}

impl App {
    pub(crate) fn new(raw: IDispatchWrapper, engine: WordEngine) -> Self {
        Self { raw, engine }
    }

    pub fn engine(&self) -> WordEngine {
        self.engine
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    pub fn name(&self) -> Result<String> {
        self.raw.get("Name")?.string()
    }

    pub fn version(&self) -> Result<String> {
        self.raw.get("Version")?.string()
    }

    /// 文档数量。
    pub fn document_count(&self) -> Result<i32> {
        self.documents_raw()?.get("Count")?.int()
    }

    fn documents_raw(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Documents", "Documents")
    }

    /// 枚举全部文档（保持 COM 一基顺序）。
    pub fn documents(&self) -> Result<Vec<Document>> {
        let docs = self.documents_raw()?;
        let count = docs.get("Count")?.int()?;
        let mut out = Vec::with_capacity(count.max(0) as usize);
        for i in 1..=count {
            let d = docs
                .call("Item", vec![ComValue::Int(i)])?
                .idispatch_child(&format!("Document[{i}]"))?;
            out.push(Document::new(d, self.engine));
        }
        Ok(out)
    }

    /// 活动文档；无文档时返回 `no_active_document`。
    pub fn active_document(&self) -> Result<Document> {
        self.raw
            .get("ActiveDocument")
            .and_then(|v| v.idispatch_child("ActiveDocument"))
            .map(|d| Document::new(d, self.engine))
            .map_err(Error::no_active)
    }

    /// 新建文档（可选模板路径）。
    pub fn add(&self, template: Option<&str>) -> Result<Document> {
        let docs = self.documents_raw()?;
        let d = match template {
            Some(t) => {
                docs.call_named("Add", vec![], &[("Template", ComValue::Str(t.to_string()))])?
            }
            None => docs.call("Add", vec![])?,
        };
        Ok(Document::new(d.idispatch_child("Document")?, self.engine))
    }

    /// 打开文档（`FileName` + `ReadOnly` 命名参数；密码不入日志）。
    pub fn open(&self, path: &str, read_only: bool) -> Result<Document> {
        let full = crate::process::abs_path(path)?;
        let docs = self.documents_raw()?;
        let d = docs
            .call_named(
                "Open",
                vec![],
                &[
                    ("FileName", ComValue::Str(full.clone())),
                    ("ReadOnly", ComValue::Bool(read_only)),
                ],
            )
            .map_err(|e| {
                crate::error::WordError::hint(e, "确认文件存在、未被独占锁定，且格式受支持")
            })?;
        Ok(Document::new(d.idispatch_child(&full)?, self.engine))
    }

    pub fn visible(&self) -> Result<bool> {
        self.raw.get("Visible")?.bool()
    }

    pub fn set_visible(&self, on: bool) -> Result<()> {
        self.raw.put("Visible", vec![on.into()])
    }

    /// 弹窗告警开关（SaveAs 遇同名文件等不再卡模态框）。
    pub fn display_alerts(&self) -> Result<bool> {
        self.raw.get("DisplayAlerts")?.bool()
    }

    pub fn set_display_alerts(&self, on: bool) -> Result<()> {
        self.raw.put("DisplayAlerts", vec![on.into()])
    }

    /// 屏幕刷新开关。
    pub fn screen_updating(&self) -> Result<bool> {
        self.raw.get("ScreenUpdating")?.bool()
    }

    pub fn set_screen_updating(&self, on: bool) -> Result<()> {
        self.raw.put("ScreenUpdating", vec![on.into()])
    }

    /// 退出应用（高风险；Attach 实例默认禁止）。
    pub fn quit(&self) -> Result<()> {
        self.raw.call("Quit", vec![]).map(|_| ())
    }
}

/// 内部错误助手。
struct Error;

impl Error {
    fn no_active(e: anyhow::Error) -> anyhow::Error {
        WordError::hint(e, "当前没有打开的文档：先 open/new 或 list_documents")
    }
}
