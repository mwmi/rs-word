//! VBA：只读 `status`、运行既有宏 `run_existing_macro`，以及**受授权**的临时注入
//! `execute_body`/`execute_module` + `cleanup_stale_temp_modules`（10 手册 §12 / 08 章）。
//!
//! 三条路径严格区分：
//! 1. 运行既有宏 —— `Application.Run(MacroName, varg1..varg30)`，不触碰 VBProject；
//! 2. 临时代码 —— `Document.VBProject → VBComponents.Add(1) → CodeModule.AddFromString →
//!    Application.Run → 只删除本次自建并登记归属的模块`；
//! 3. 默认关闭：需宿主显式 `allow_vba_injection=true` + `confirm="yes"`，绝不改安全设置绕过。

use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::Result;

use crate::app::App;
use crate::com::dispatch::IDispatchWrapper;
use crate::com::variant::ComValue;
use crate::document::Document;
use crate::error::{SideEffectStatus, WordError};
use crate::types::{
    VbaArg, VbaCodeKind, VbaExecuteRequest, VbaExecuteResult, VbaResult, VbaRunRequest, VbaStatus,
};

/// 本库自建临时模块的前缀（只清理登记归属的模块）。
pub const TEMP_PREFIX: &str = "RsWordTmp_";

/// vbext_ct_StdModule。
const VBEXT_CT_STD_MODULE: i32 = 1;

/// 拒绝出现在注入代码里的构造（模态框会阻塞 STA 线程；外部进程调用超出文档沙箱保证）。
const BANNED: [&str; 3] = ["MsgBox", "Shell", "CreateObject(\"WScript.Shell\")"];

/// 代码守卫：返回命中的被拒构造（`None` = 通过）。纯函数，可离线单测。
pub fn banned_construct(code: &str) -> Option<&'static str> {
    BANNED.iter().find(|b| code.contains(**b)).copied()
}

/// 只读探测 VBA 能力（不执行未知宏、不改安全设置、不创建临时模块）。
pub fn status(app: &App, doc: &Document) -> VbaStatus {
    let host_run = app.raw().has_member("Run");
    let project_access = doc.raw().get_obj("VBProject", "VBProject").is_ok();
    let notes = if !host_run {
        Some("Application.Run 不可用（VBA 运行时缺失或引擎不支持）".into())
    } else if !project_access {
        Some(
            "VBProject 不可访问（信任中心/工程保护）；run_existing_macro 仍可用，临时注入需先在目标机授权"
                .into(),
        )
    } else {
        Some("具备注入的尝试条件；实际可用性以真机实测为准".into())
    };
    VbaStatus {
        available: host_run,
        host_run,
        project_access,
        injection_ready: host_run && project_access,
        notes,
    }
}

/// 运行既有宏。
pub fn run(app: &App, doc: &Document, req: &VbaRunRequest) -> Result<VbaResult> {
    if doc.id() != req.document_id {
        return Err(WordError::new("vba.run", doc.label(), "document_id 不匹配").into());
    }
    if req.macro_name.trim().is_empty() {
        return Err(WordError::new("vba.run", doc.label(), "macro_name 为空").into());
    }
    if req.args.len() > 30 {
        return Err(
            WordError::new("vba.run", doc.label(), "Application.Run 最多 30 个参数").into(),
        );
    }
    let returned = run_macro(app, &req.macro_name, &req.args)
        .map_err(|e| WordError::hint(e, "宏不存在或运行/编译错误；确认宏名与参数"))?;
    Ok(VbaResult {
        document_id: req.document_id.clone(),
        operation_id: crate::new_operation_id(),
        returned,
        side_effect_status: SideEffectStatus::Unknown,
        applied: vec![format!("run {}", req.macro_name)],
    })
}

/// 调用 `Application.Run`；返回宏返回值文本（`Sub` 通常为 None，**不等于失败**）。
fn run_macro(app: &App, macro_name: &str, args: &[VbaArg]) -> Result<Option<String>> {
    let mut com_args: Vec<ComValue> = Vec::with_capacity(args.len() + 1);
    com_args.push(ComValue::Str(macro_name.to_string()));
    for a in args {
        com_args.push(match a {
            VbaArg::Str { value } => ComValue::Str(value.clone()),
            VbaArg::Int { value } => ComValue::Int(*value),
            VbaArg::Float { value } => ComValue::Float(*value),
            VbaArg::Bool { value } => ComValue::Bool(*value),
        });
    }
    let result = app.raw().call("Run", com_args)?;
    Ok(match result.value() {
        crate::com::variant::WordValue::Empty => None,
        other => Some(other.to_string()),
    })
}

/// 临时注入（高风险）。`owned` 记录本库创建且尚未清理的模块名。
pub fn execute(
    app: &App,
    doc: &Document,
    req: &VbaExecuteRequest,
    owned: &mut Vec<String>,
    seq: &mut u64,
) -> Result<VbaExecuteResult> {
    if doc.id() != req.document_id {
        return Err(WordError::new("vba.execute", doc.label(), "document_id 不匹配").into());
    }
    if !req.allow_vba_injection {
        return Err(
            WordError::new("vba.execute", doc.label(), "临时 VBA 注入默认关闭")
                .with_hint("宿主显式授权后传 allow_vba_injection=true（并建议先 checkpoint）")
                .into(),
        );
    }
    if req.confirm.as_deref() != Some("yes") {
        return Err(
            WordError::new("vba.execute", doc.label(), "临时 VBA 注入为高风险操作")
                .with_hint("确认后传 confirm=\"yes\"")
                .into(),
        );
    }
    if !app.raw().has_member("Run") {
        return Err(WordError::new(
            "vba.execute",
            doc.label(),
            "vba_runtime_unavailable：Application.Run 不可用",
        )
        .with_hint("该引擎/版本没有可用的 VBA 运行时；先 vba.status 确认")
        .into());
    }
    if req.entry.trim().is_empty() {
        return Err(WordError::new("vba.execute", doc.label(), "entry 为空").into());
    }
    if req.code.trim().is_empty() {
        return Err(WordError::new("vba.execute", doc.label(), "code 为空").into());
    }
    if matches!(req.kind, VbaCodeKind::Body) && !req.args.is_empty() {
        return Err(WordError::new(
            "vba.execute",
            doc.label(),
            "kind=body 不支持 args（包装成 Public Sub <entry>() 无参数）",
        )
        .with_hint("改用 kind=module 并自行声明带参入口")
        .into());
    }
    if req.args.len() > 30 {
        return Err(WordError::new("vba.execute", doc.label(), "最多 30 个参数").into());
    }
    if let Some(banned) = banned_construct(&req.code) {
        return Err(WordError::new(
            "vba.execute",
            doc.label(),
            format!("注入代码包含被拒绝的构造 `{banned}`"),
        )
        .with_hint("不能用 MsgBox 回传结果，也不能启动外部进程；请返回值/文档状态回读")
        .into());
    }

    let project = doc.raw().get_obj("VBProject", "VBProject").map_err(|e| {
        WordError::hint(
            e,
            "vba_project_access_denied：需在信任中心显式允许 VBA 工程对象模型访问；不得绕过安全策略",
        )
    })?;
    let components = project.get_obj("VBComponents", "VBComponents")?;
    let module_name = next_module_name(seq);
    let component = components
        .call("Add", vec![ComValue::Int(VBEXT_CT_STD_MODULE)])
        .map_err(|e| {
            let msg = e.to_string();
            if msg.contains("lock") || msg.contains("protect") {
                WordError::hint(
                    e,
                    "vba_project_locked：工程受保护，请先在目标机解除工程保护",
                )
            } else {
                WordError::hint(e, "无法新增标准模块：确认项目访问权限与 VBA 组件可用")
            }
        })?
        .idispatch_child(&module_name)?;
    component.put("Name", vec![ComValue::Str(module_name.clone())])?;
    let source = req.kind.wrap(&req.entry, &req.code);
    component
        .get_obj("CodeModule", "CodeModule")?
        .call("AddFromString", vec![ComValue::Str(source)])?;
    owned.push(module_name.clone());

    // 运行本次唯一入口（限定名前缀避免撞名）。
    let macro_name = format!("{module_name}.{}", req.entry);
    let outcome = run_macro(app, &macro_name, &req.args);

    // 清理（只删本次创建并登记归属的模块）。
    if req.cleanup {
        if let Err(e) = remove_component(&components, &component) {
            return Err(WordError::new(
                "vba.execute",
                doc.label(),
                format!("vba_cleanup_failed：临时模块 {module_name} 未能删除（{e}）"),
            )
            .with_hint("模块归属已保留；稍后执行 vba action=cleanup_stale_temp_modules")
            .into());
        }
        owned.retain(|m| m != &module_name);
    }

    let (status, returned, error_number, error_description) = match outcome {
        Ok(returned) => parse_returned(returned),
        Err(e) => {
            let text = e.to_string();
            if text.contains("Compile error") {
                return Err(WordError::new(
                    "vba.execute",
                    doc.label(),
                    format!("vba_compile_error：{text}"),
                )
                .with_hint("修正代码后重试；编译失败不重复运行")
                .into());
            }
            ("runtime_error".to_string(), None, None, Some(text))
        }
    };

    Ok(VbaExecuteResult {
        document_id: req.document_id.clone(),
        operation_id: crate::new_operation_id(),
        module_name: module_name.clone(),
        entry: req.entry.clone(),
        returned,
        status,
        error_number,
        error_description,
        cleaned_up: req.cleanup,
        cleanup_pending: false,
        side_effect_status: SideEffectStatus::Unknown,
        applied: vec![format!("execute {macro_name}")],
    })
}

/// 生成并推进临时模块名（`RsWordTmp_<seq>_<unix_ms>`）。
fn next_module_name(seq: &mut u64) -> String {
    *seq += 1;
    let ms = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0);
    format!("{TEMP_PREFIX}{seq}_{ms}")
}

/// 清理登记归属的临时模块（只删 `RsWordTmp_` 前缀且存在于 `owned` 的）。
pub fn cleanup_stale(doc: &Document, owned: &mut Vec<String>) -> Result<Vec<String>> {
    if owned.is_empty() {
        return Ok(Vec::new());
    }
    let project = doc
        .raw()
        .get_obj("VBProject", "VBProject")
        .map_err(|e| WordError::hint(e, "VBProject 不可访问：无法枚举临时模块"))?;
    let components = project.get_obj("VBComponents", "VBComponents")?;
    let count = components.get("Count").and_then(|v| v.int()).unwrap_or(0);
    let mut removed = Vec::new();
    for i in (1..=count).rev() {
        let Ok(component) = components
            .call("Item", vec![ComValue::Int(i)])
            .and_then(|r| r.idispatch_child("VBComponent"))
        else {
            continue;
        };
        let name = component
            .get("Name")
            .and_then(|v| v.string())
            .unwrap_or_default();
        if !name.starts_with(TEMP_PREFIX) || !owned.contains(&name) {
            continue;
        }
        remove_component(&components, &component)?;
        removed.push(name);
    }
    owned.retain(|m| !removed.contains(m));
    Ok(removed)
}

/// 移除模块（`VBComponents.Remove(VBComponent)`）。
fn remove_component(components: &IDispatchWrapper, component: &IDispatchWrapper) -> Result<()> {
    components
        .call("Remove", vec![ComValue::Obj(component.disp.clone())])
        .map(|_| ())
}

/// 解析宏返回值：`ERR|<number>|<description>` 视为运行期错误。
fn parse_returned(
    returned: Option<String>,
) -> (String, Option<String>, Option<i64>, Option<String>) {
    match returned {
        Some(s) if s.starts_with("ERR|") => {
            let mut parts = s.splitn(3, '|');
            let _ = parts.next();
            let num = parts.next().and_then(|n| n.parse::<i64>().ok());
            let desc = parts.next().map(str::to_string);
            ("runtime_error".to_string(), Some(s), num, desc)
        }
        other => ("ok".to_string(), other, None, None),
    }
}
