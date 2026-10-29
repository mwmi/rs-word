//! AI 工具门面：把工具名 + JSON 输入路由到 typed 调用，返回统一 JSON 回执。
//!
//! 参数校验：`target.kind` 与 `action` 的互斥字段在反序列化（tagged enum）时即被拒绝；
//! 其余语义校验在各领域模块内完成。

use serde_json::{json, Value};

use crate::client::WordClient;
use crate::error::{Category, ToolError};
use crate::types::*;

/// 分派一个工具调用，返回 `{ok, data|error}` 回执。
pub fn dispatch(client: &WordClient, name: &str, input: Value) -> Value {
    match run(client, name, input) {
        Ok(data) => json!({ "ok": true, "data": data }),
        Err(e) => json!({ "ok": false, "error": e }),
    }
}

fn run(client: &WordClient, name: &str, input: Value) -> Result<Value, ToolError> {
    match name {
        "word_session" => session(client, &input),
        "word_read" => ok(client.read(from_value(input, true)?)?),
        "word_edit" => ok(client.edit(from_value(input, false)?)?),
        "word_find_replace" => find_replace(client, &input),
        "word_format" => ok(client.format(from_value(input, false)?)?),
        "word_table" => ok(client.table(from_value(input, false)?)?),
        "word_layout" => ok(client.layout(from_value(input, false)?)?),
        "word_review" => ok(client.review(from_value(input, false)?)?),
        "word_media" => ok(client.media(from_value(input, false)?)?),
        "word_advanced" => ok(client.advanced(from_value(input, false)?)?),
        "word_workflow" => workflow(client, &input),
        "word_history" => ok(client.history(from_value(input, false)?)?),
        "word_vba" => vba(client, &input),
        other => Err(ToolError::validation(
            "unknown_tool",
            format!("未知工具 '{other}'"),
        )),
    }
}

fn ok(v: impl serde::Serialize) -> Result<Value, ToolError> {
    serde_json::to_value(v).map_err(|e| internal(e.to_string()))
}

fn internal(msg: String) -> ToolError {
    ToolError::new("internal", Category::ComFailure, msg)
}

fn parse_err(e: serde_json::Error) -> ToolError {
    ToolError::validation("invalid_argument", format!("参数不合法: {e}"))
}

/// 反序列化工具输入；`inject_target` 时若缺 `target` 则按 `document_id` 补 `document` 目标。
fn from_value<T: serde::de::DeserializeOwned>(
    mut input: Value,
    inject_target: bool,
) -> Result<T, ToolError> {
    if inject_target {
        if let Value::Object(map) = &mut input {
            if !map.contains_key("target") {
                let doc_id = map
                    .get("document_id")
                    .cloned()
                    .unwrap_or(Value::String(String::new()));
                map.insert(
                    "target".into(),
                    json!({ "kind": "document", "document_id": doc_id }),
                );
            }
        }
    }
    serde_json::from_value(input).map_err(parse_err)
}

fn req_str(input: &Value, key: &str) -> Result<String, ToolError> {
    input
        .get(key)
        .and_then(|v| v.as_str())
        .map(str::to_string)
        .ok_or_else(|| ToolError::validation("invalid_argument", format!("缺少必需参数 '{key}'")))
}

fn session(client: &WordClient, input: &Value) -> Result<Value, ToolError> {
    let action = input
        .get("action")
        .and_then(|v| v.as_str())
        .unwrap_or("status");
    match action {
        "status" => ok(client.status()?),
        "connect" => {
            let engine = parse_engine(input.get("engine").and_then(|v| v.as_str()));
            let mode = parse_mode(input.get("mode").and_then(|v| v.as_str()));
            ok(client.connect(EnginePolicy::Explicit(engine), mode)?)
        }
        "switch_engine" => {
            let engine = parse_engine(input.get("engine").and_then(|v| v.as_str()));
            let dp = input
                .get("dirty_policy")
                .and_then(|v| v.as_str())
                .unwrap_or("refuse_if_dirty");
            let dirty = match dp {
                "save" => DirtyPolicy::Save,
                "discard" => DirtyPolicy::Discard,
                _ => DirtyPolicy::RefuseIfDirty,
            };
            ok(client.switch_engine(engine, dirty)?)
        }
        "list_documents" => ok(client.list_documents()?),
        "open" => {
            let path = req_str(input, "path")?;
            let ro = input
                .get("read_only")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            ok(client.open(&path, ro)?)
        }
        "new" => {
            let t = input.get("template_path").and_then(|v| v.as_str());
            ok(client.new_document(t)?)
        }
        "activate" => {
            let id = req_str(input, "document_id")?;
            ok(client.activate(&id)?)
        }
        "save" => {
            let id = req_str(input, "document_id")?;
            let path = input.get("path").and_then(|v| v.as_str());
            let fmt = input
                .get("format")
                .and_then(|v| v.as_str())
                .map(parse_format)
                .unwrap_or(SaveFormat::Keep);
            ok(client.save(&id, path, fmt)?)
        }
        "close" => {
            let id = req_str(input, "document_id")?;
            let policy = match input
                .get("save_policy")
                .and_then(|v| v.as_str())
                .unwrap_or("refuse_if_dirty")
            {
                "save" => ClosePolicy::Save,
                "discard" => ClosePolicy::Discard,
                _ => ClosePolicy::RefuseIfDirty,
            };
            ok(client.close(&id, policy)?)
        }
        "export_pdf" => {
            let id = req_str(input, "document_id")?;
            let path = req_str(input, "path")?;
            let overwrite = input
                .get("overwrite")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            ok(client.layout(LayoutRequest {
                document_id: id,
                action: LayoutAction::ExportPdf,
                page: None,
                header_footer: None,
                toc: None,
                watermark: None,
                numbering: None,
                page_number: None,
                header_footer_zones: None,
                pdf_path: Some(path),
                overwrite,
            })?)
        }
        other => Err(ToolError::validation(
            "invalid_argument",
            format!("未知 session action '{other}'"),
        )),
    }
}

fn find_replace(client: &WordClient, input: &Value) -> Result<Value, ToolError> {
    let action = input
        .get("action")
        .and_then(|v| v.as_str())
        .unwrap_or("find");
    match action {
        "find" => ok(client.find(from_value(input.clone(), false)?)?),
        "preview_replace" => ok(client.replace_preview(from_value(input.clone(), false)?)?),
        "apply_replace" => ok(client.replace_apply(from_value(input.clone(), false)?)?),
        other => Err(ToolError::validation(
            "invalid_argument",
            format!("未知 find_replace action '{other}'"),
        )),
    }
}

fn workflow(client: &WordClient, input: &Value) -> Result<Value, ToolError> {
    let action = input
        .get("action")
        .and_then(|v| v.as_str())
        .unwrap_or("prepare");
    match action {
        "prepare" => ok(client.workflow_prepare(from_value(input.clone(), true)?)?),
        "apply" => ok(client.workflow_apply(from_value(input.clone(), false)?)?),
        "validate" => ok(client.workflow_validate(from_value(input.clone(), false)?)?),
        other => Err(ToolError::validation(
            "invalid_argument",
            format!("未知 workflow action '{other}'（可用 prepare/apply/validate）"),
        )),
    }
}

fn vba(client: &WordClient, input: &Value) -> Result<Value, ToolError> {
    let action = input
        .get("action")
        .and_then(|v| v.as_str())
        .unwrap_or("status");
    match action {
        "status" => {
            let id = req_str(input, "document_id")?;
            ok(client.vba_status(&id)?)
        }
        "run_existing_macro" => ok(client.vba_run(from_value(input.clone(), false)?)?),
        "execute_body" | "execute_module" => {
            let mut req: crate::types::VbaExecuteRequest =
                from_value(input.clone(), false)?;
            // action 决定代码形态（body 包装成无参 Sub，module 原样使用）。
            req.kind = if action == "execute_body" {
                crate::types::VbaCodeKind::Body
            } else {
                crate::types::VbaCodeKind::Module
            };
            ok(client.vba_execute(req)?)
        }
        "cleanup_stale_temp_modules" => {
            let id = req_str(input, "document_id")?;
            ok(client.vba_cleanup(&id)?)
        }
        "execute" => Err(ToolError::new(
            "invalid_argument",
            Category::Validation,
            "action 'execute' 已拆分为 execute_body / execute_module（需 allow_vba_injection=true）",
        )),
        other => Err(ToolError::validation(
            "invalid_argument",
            format!("未知 vba action '{other}'"),
        )),
    }
}

fn parse_engine(s: Option<&str>) -> WordEngine {
    match s {
        Some("ms_word") | Some("word") => WordEngine::MsWord,
        _ => WordEngine::WpsWriter,
    }
}

fn parse_mode(s: Option<&str>) -> ConnectMode {
    match s {
        Some("new") => ConnectMode::New,
        _ => ConnectMode::Attach,
    }
}

fn parse_format(s: &str) -> SaveFormat {
    match s {
        "docx" => SaveFormat::Docx,
        "doc" => SaveFormat::Doc,
        "rtf" => SaveFormat::Rtf,
        "txt" => SaveFormat::Txt,
        "html" => SaveFormat::Html,
        _ => SaveFormat::Keep,
    }
}
