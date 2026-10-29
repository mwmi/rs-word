//! Document：内容范围、统计、保存/另存/关闭/导出、修订与批注集合。

use anyhow::Result;

use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::error::WordError;
use crate::range::Range;
use crate::types::{ClosePolicy, SaveFormat, WordEngine};

/// 文档（Document COM 对象）。`Clone` 只是 AddRef 同一对象。
#[derive(Clone)]
pub struct Document {
    raw: IDispatchWrapper,
    engine: WordEngine,
    /// 库内生成的稳定 ID（由会话分配，避免只凭窗口标题定位）。
    id: String,
}

impl Document {
    pub(crate) fn new(raw: IDispatchWrapper, engine: WordEngine) -> Self {
        Self {
            raw,
            engine,
            id: String::new(),
        }
    }

    pub(crate) fn with_id(mut self, id: String) -> Self {
        self.id = id;
        self
    }

    pub fn id(&self) -> &str {
        &self.id
    }

    pub fn engine(&self) -> WordEngine {
        self.engine
    }

    /// 底层 COM 包装（扩展性核心）。
    pub fn raw(&self) -> &IDispatchWrapper {
        &self.raw
    }

    /// 对象路径标签（错误定位用）。
    pub(crate) fn label(&self) -> &str {
        &self.raw.ctx
    }

    pub fn name(&self) -> Result<String> {
        self.raw.get("Name")?.string()
    }

    pub fn full_name(&self) -> Result<String> {
        self.raw
            .get("FullName")
            .map(|v| v.string().unwrap_or_default())
    }

    pub fn path(&self) -> Result<String> {
        self.raw.get("Path").map(|v| v.string().unwrap_or_default())
    }

    /// 是否无未保存改动（`Saved=true` 只说明无未保存变更，不保证落盘）。
    pub fn is_saved(&self) -> Result<bool> {
        self.raw.get("Saved")?.bool()
    }

    pub fn read_only(&self) -> Result<bool> {
        self.raw.get("ReadOnly")?.bool()
    }

    fn count_of(&self, coll: &str) -> Option<u32> {
        self.raw
            .get_obj(coll, coll)
            .and_then(|c| c.get("Count"))
            .and_then(|v| v.int())
            .ok()
            .map(|n| n.max(0) as u32)
    }

    pub fn paragraph_count(&self) -> Option<u32> {
        self.count_of("Paragraphs")
    }
    pub fn table_count(&self) -> Option<u32> {
        self.count_of("Tables")
    }
    pub fn section_count(&self) -> Option<u32> {
        self.count_of("Sections")
    }
    pub fn revision_count(&self) -> Option<u32> {
        self.count_of("Revisions")
    }
    pub fn comment_count(&self) -> Option<u32> {
        self.count_of("Comments")
    }

    /// 主 Story 全文范围。
    pub fn content(&self) -> Result<Range> {
        Ok(Range::new(self.raw.get_obj("Content", "Content")?))
    }

    /// 指定字符范围（Word Start/End，非 UTF-8 字节）。
    pub fn range(&self, start: i32, end: i32) -> Result<Range> {
        Ok(Range::new(
            self.raw
                .call("Range", vec![ComValue::Int(start), ComValue::Int(end)])?
                .idispatch_child(&format!("Range[{start},{end}]"))?,
        ))
    }

    /// 末尾折叠点（文档终止符之前）。
    pub fn end_range(&self) -> Result<Range> {
        let content = self.raw.get_obj("Content", "Content")?;
        let end = content.get("End")?.int()?;
        // Word 的文档终止符位于 Content.End-1，插入点取该处。
        self.range((end - 1).max(0), (end - 1).max(0))
    }

    pub fn sections(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Sections", "Sections")
    }

    pub fn tables(&self) -> Result<IDispatchWrapper> {
        self.raw.get_obj("Tables", "Tables")
    }

    /// 第 `index` 张表的 `Range`；没有该表或探测失败时为 None（探测不该让主流程失败）。
    fn table_range(&self, index: i32) -> Option<Range> {
        self.tables()
            .ok()?
            .call("Item", vec![ComValue::Int(index)])
            .ok()?
            .idispatch_child("Table")
            .ok()?
            .get_obj("Range", "Table.Range")
            .ok()
            .map(Range::new)
    }

    /// 第一张表的 `Range.Start`；没有表格时为 None。
    pub fn first_table_start(&self) -> Option<i32> {
        self.table_range(1)?.start().ok()
    }

    /// 最后一张表的 `Range.End`；没有表格时为 None。
    pub fn last_table_end(&self) -> Option<i32> {
        let count = self.table_count()?;
        if count == 0 {
            return None;
        }
        self.table_range(count as i32)?.end().ok()
    }

    pub fn track_revisions(&self) -> Result<bool> {
        self.raw.get("TrackRevisions")?.bool()
    }

    /// 设置跟踪修订开关（调用方负责保存/恢复原值）。
    pub fn set_track_revisions(&self, on: bool) -> Result<()> {
        self.raw.put("TrackRevisions", vec![on.into()])
    }

    /// 保存到当前路径。
    pub fn save(&self) -> Result<()> {
        self.raw.call("Save", vec![]).map(|_| ())
    }

    /// 另存为；`SaveAs2` 不存在时回退 `SaveAs`。
    pub fn save_as(&self, path: &str, format: SaveFormat) -> Result<()> {
        let full = crate::process::abs_path(path)?;
        let mut named: Vec<(&str, ComValue)> = vec![("FileName", ComValue::Str(full))];
        if let Some(code) = format.code() {
            named.push(("FileFormat", ComValue::Int(code)));
        }
        match self.raw.call_named("SaveAs2", vec![], &named) {
            Ok(_) => Ok(()),
            Err(e) if crate::com::dispatch::is_member_notfound(&e) => {
                self.raw.call_named("SaveAs", vec![], &named).map(|_| ())
            }
            Err(e) => Err(e),
        }
    }

    /// 关闭（按策略保存/丢弃/拒绝脏文档）。
    pub fn close(&self, policy: ClosePolicy) -> Result<()> {
        let dirty = !self.is_saved().unwrap_or(true);
        let save_changes = match policy {
            ClosePolicy::RefuseIfDirty if dirty => {
                return Err(WordError::new(
                    "close",
                    self.label(),
                    "文档有未保存修改，已按 refuse_if_dirty 拒绝关闭",
                )
                .with_hint("先 save，或用 save_policy=save/discard")
                .into())
            }
            ClosePolicy::RefuseIfDirty | ClosePolicy::Save => -1, // wdSaveChanges
            ClosePolicy::Discard => 0,                            // wdDoNotSaveChanges
        };
        self.raw
            .call("Close", vec![ComValue::Int(save_changes)])
            .map(|_| ())
    }

    /// 导出 PDF（ExportAsFixedFormat；PDF 格式码 17）。
    pub fn export_pdf(&self, path: &str, overwrite: bool) -> Result<u64> {
        let full = crate::process::abs_path(path)?;
        if !overwrite && std::path::Path::new(&full).exists() {
            return Err(WordError::new(
                "export_pdf",
                self.label(),
                format!("目标已存在且未允许覆盖: {full}"),
            )
            .with_hint("设置 overwrite=true 或换一个路径")
            .into());
        }
        self.raw.call_named(
            "ExportAsFixedFormat",
            vec![],
            &[
                ("OutputFileName", ComValue::Str(full.clone())),
                ("ExportFormat", ComValue::Int(17)), // wdExportFormatPDF
                ("OpenAfterExport", ComValue::Bool(false)),
                ("OptimizeFor", ComValue::Int(0)), // wdExportOptimizeForPrint
            ],
        )?;
        let bytes = std::fs::metadata(&full)
            .map_err(|e| {
                WordError::new("export_pdf", self.label(), format!("导出后文件不可读: {e}"))
            })?
            .len();
        if bytes == 0 {
            return Err(WordError::new("export_pdf", self.label(), "导出文件为空").into());
        }
        Ok(bytes)
    }
}
