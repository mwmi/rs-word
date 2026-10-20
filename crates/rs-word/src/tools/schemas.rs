//! 核心 AI 工具 JSON Schema（strict 风格：`additionalProperties:false` + 显式字段）。
//!
//! 后端仍需按 `target.kind` / `action` 做二次语义校验；本文件只描述形状。

use serde_json::{json, Value};

/// 工具名清单（默认注册的核心工具）。
pub fn tool_names() -> &'static [&'static str] {
    &[
        "word_session",
        "word_read",
        "word_edit",
        "word_find_replace",
        "word_format",
        "word_table",
        "word_layout",
        "word_review",
        "word_media",
        "word_advanced",
        "word_workflow",
        "word_history",
        "word_vba",
    ]
}

/// Target 对象的可复用 schema 片段。
fn target_schema() -> Value {
    json!({
        "type": "object",
        "additionalProperties": true,
        "properties": {
            "kind": {"type": "string", "enum": ["document","selection","range","paragraph","find","bookmark","table_cell","header_footer"]},
            "document_id": {"type": "string"},
            "story": {"type": ["string","null"]},
            "start": {"type": ["integer","null"]},
            "end": {"type": ["integer","null"]},
            "expected_text": {"type": ["string","null"]},
            "text_hash": {"type": ["string","null"]},
            "expected_doc_version": {"type": ["integer","null"]},
            "paragraph_index": {"type": ["integer","null"]},
            "query": {"type": ["string","null"]},
            "occurrence": {"type": ["integer","null"]},
            "bookmark": {"type": ["string","null"]},
            "table_id": {"type": ["string","null"]},
            "row": {"type": ["integer","null"]},
            "column": {"type": ["integer","null"]},
            "section": {"type": ["integer","null"]},
            "hf_kind": {"type": ["string","null"], "enum": ["header","footer", null]},
            "variant": {"type": ["string","null"]}
        },
        "required": ["kind", "document_id"]
    })
}

/// 全部核心工具定义。
pub fn tool_definitions() -> Value {
    json!([
        {
            "type": "function",
            "name": "word_session",
            "description": "连接/切换引擎、文档列表、打开/新建/保存/关闭、状态。action 枚举。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "action": {"type": "string", "enum": ["status","connect","switch_engine","list_documents","open","new","activate","save","close","export_pdf"]},
                    "engine": {"type": ["string","null"], "enum": ["wps_writer","ms_word", null]},
                    "mode": {"type": ["string","null"], "enum": ["attach","new", null]},
                    "dirty_policy": {"type": ["string","null"]},
                    "document_id": {"type": ["string","null"]},
                    "path": {"type": ["string","null"]},
                    "read_only": {"type": ["boolean","null"]},
                    "template_path": {"type": ["string","null"]},
                    "format": {"type": ["string","null"]},
                    "save_policy": {"type": ["string","null"]},
                    "overwrite": {"type": ["boolean","null"]}
                },
                "required": ["action"]
            }
        },
        {
            "type": "function",
            "name": "word_read",
            "description": "只读指定范围。默认 summary，不读全文；返回 anchor/doc_version/截断/coverage。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "target": target_schema(),
                    "view": {"type": ["string","null"], "enum": ["summary","outline","text","paragraphs","tables_summary","table","selection","comments","revisions","metadata", null]},
                    "limit_chars": {"type": ["integer","null"]},
                    "max_items": {"type": ["integer","null"]},
                    "cursor": {"type": ["string","null"]}
                },
                "required": ["document_id", "target"]
            }
        },
        {
            "type": "function",
            "name": "word_edit",
            "description": "文本插入/替换/删除、段落、分页分节符、书签、超链接、字段、套样式；一次提交多动作。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "expected_doc_version": {"type": ["integer","null"]},
                    "validate_all_first": {"type": ["boolean","null"]},
                    "track_changes": {"type": ["boolean","null"]},
                    "idempotency_key": {"type": ["string","null"]},
                    "actions": {"type": "array", "items": {"type": "object", "additionalProperties": true}}
                },
                "required": ["document_id", "actions"]
            }
        },
        {
            "type": "function",
            "name": "word_find_replace",
            "description": "查找 / 替换预览 / 执行替换。替换必须先 preview 再 apply（限量）。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "action": {"type": "string", "enum": ["find","preview_replace","apply_replace"]},
                    "document_id": {"type": "string"},
                    "query": {"type": ["string","null"]},
                    "find": {"type": ["string","null"]},
                    "replace": {"type": ["string","null"]},
                    "options": {"type": ["object","null"], "additionalProperties": true},
                    "limit": {"type": ["integer","null"]},
                    "preview_id": {"type": ["string","null"]},
                    "expected_doc_version": {"type": ["integer","null"]},
                    "max_replacements": {"type": ["integer","null"]},
                    "track_changes": {"type": ["boolean","null"]}
                },
                "required": ["action", "document_id"]
            }
        },
        {
            "type": "function",
            "name": "word_format",
            "description": "字符/段落/样式/列表格式（部分更新：None=不改）。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "target": target_schema(),
                    "expected_doc_version": {"type": ["integer","null"]},
                    "style": {"type": ["string","null"]},
                    "font": {"type": ["object","null"], "additionalProperties": true},
                    "paragraph": {"type": ["object","null"], "additionalProperties": true},
                    "list": {"type": ["string","null"], "enum": ["none","bullet","number","outline_number","custom", null]},
                    "list_template": {"type": ["object","null"], "additionalProperties": true},
                    "reset": {"type": ["object","null"], "additionalProperties": true}
                },
                "required": ["document_id", "target"]
            }
        },
        {
            "type": "function",
            "name": "word_table",
            "description": "表格创建/读写/增删行列/合并拆分/底纹对齐/列宽行高边框跨页行。action 枚举。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "action": {"type": "string", "enum": ["create","inspect","read_cells","set_cells","insert_row","delete_row","insert_column","delete_column","merge_cells","split_cell","set_widths","set_heights","set_borders","set_allow_row_break","set_shading","set_alignment","set_repeat_header","autofit","delete_table"]},
                    "table_id": {"type": ["string","null"]},
                    "target": {"type": ["object","null"], "additionalProperties": true},
                    "data": {"type": ["array","null"], "items": {"type": "array", "items": {"type": "string"}}},
                    "range": {"type": ["object","null"], "additionalProperties": true},
                    "header_rows": {"type": ["integer","null"]},
                    "format": {"type": ["object","null"], "additionalProperties": true,
                               "description": "{widths_pt, heights:[{row,height_pt,rule}], borders:[{edge,style,width_pt,color}], allow_row_break}"}
                },
                "required": ["document_id", "action"]
            }
        },
        {
            "type": "function",
            "name": "word_layout",
            "description": "页面设置/页眉页脚/页码/字段更新/PDF 导出。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "action": {"type": "string", "enum": ["set_page_setup","set_header_footer","insert_page_number","export_pdf","update_fields"]},
                    "page": {"type": ["object","null"], "additionalProperties": true},
                    "header_footer": {"type": ["object","null"], "additionalProperties": true},
                    "pdf_path": {"type": ["string","null"]},
                    "overwrite": {"type": ["boolean","null"]}
                },
                "required": ["document_id", "action"]
            }
        },
        {
            "type": "function",
            "name": "word_review",
            "description": "修订/批注：status/set_tracking/list/accept/reject/comment。accept_all/reject_all 需 confirm。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "action": {"type": "string", "enum": ["status","set_tracking","list_revisions","accept_revision","reject_revision","accept_all","reject_all","accept_filtered","reject_filtered","add_comment","list_comments","delete_comment"]},
                    "enabled": {"type": ["boolean","null"]},
                    "revision_ref": {"type": ["string","null"]},
                    "target": {"type": ["object","null"], "additionalProperties": true},
                    "text": {"type": ["string","null"]},
                    "comment_id": {"type": ["string","null"]},
                    "confirm": {"type": ["string","null"]},
                    "filter": {"type": ["object","null"], "additionalProperties": true,
                               "description": "{authors:[..], kinds:[insert|delete|format|...], max}"},
                    "limit": {"type": ["integer","null"]}
                },
                "required": ["document_id", "action"]
            }
        },
        {
            "type": "function",
            "name": "word_history",
            "description": "简易历史：list/undo/redo/checkpoint/clear。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "action": {"type": "string", "enum": ["list","undo","redo","checkpoint","clear"]},
                    "checkpoint_dir": {"type": ["string","null"]}
                },
                "required": ["document_id", "action"]
            }
        },
        {
            "type": "function",
            "name": "word_media",
            "description": "图片/形状/文本框/WordArt：list/insert_picture/insert_shape/insert_textbox/insert_wordart/resize/set_wrap/set_position/set_alt_text/z_order/delete。几何单位 pt。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "action": {"type": "string", "enum": ["list","insert_picture","insert_shape","insert_textbox","insert_wordart","resize","set_wrap","set_position","set_alt_text","z_order","delete"]},
                    "target": {"type": ["object","null"], "additionalProperties": true},
                    "media_id": {"type": ["string","null"]},
                    "expect_name": {"type": ["string","null"]},
                    "path": {"type": ["string","null"]},
                    "inline": {"type": ["boolean","null"]},
                    "left_pt": {"type": ["number","null"]},
                    "top_pt": {"type": ["number","null"]},
                    "width_pt": {"type": ["number","null"]},
                    "height_pt": {"type": ["number","null"]},
                    "shape": {"type": ["string","null"], "enum": ["rectangle","rounded_rectangle","diamond","ellipse","arrow", null]},
                    "text": {"type": ["string","null"]},
                    "font_name": {"type": ["string","null"]},
                    "font_size_pt": {"type": ["number","null"]},
                    "bold": {"type": ["boolean","null"]},
                    "italic": {"type": ["boolean","null"]},
                    "preset_text_effect": {"type": ["integer","null"]},
                    "wrap": {"type": ["string","null"], "enum": ["inline","square","tight","through","top_bottom","behind","front", null]},
                    "alt_text": {"type": ["string","null"]},
                    "z_order": {"type": ["string","null"], "enum": ["bring_to_front","send_to_back","bring_forward","send_backward", null]},
                    "max_items": {"type": ["integer","null"]}
                },
                "required": ["document_id", "action"]
            }
        },
        {
            "type": "function",
            "name": "word_advanced",
            "description": "脚注尾注、交叉引用、结构快照与 diff。只读操作为主；add_*/insert_* 会写入。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "document_id": {"type": "string"},
                    "action": {"type": "string", "enum": ["list_notes","add_footnote","add_endnote","insert_cross_reference","list_cross_references","structure_snapshot","structure_diff"]},
                    "target": {"type": ["object","null"], "additionalProperties": true},
                    "text": {"type": ["string","null"]},
                    "reference": {"type": ["string","null"]},
                    "xref_type": {"type": ["string","null"], "enum": ["heading","bookmark","footnote","endnote","figure","table", null]},
                    "xref_kind": {"type": ["string","null"], "enum": ["content_text","page_number","number_full_context", null]},
                    "reference_item": {"type": ["string","null"]},
                    "insert_as_hyperlink": {"type": ["boolean","null"]},
                    "snapshot": {"type": ["object","null"], "additionalProperties": true},
                    "max_items": {"type": ["integer","null"]}
                },
                "required": ["document_id", "action"]
            }
        },
        {
            "type": "function",
            "name": "word_workflow",
            "description": "AI 工作流两阶段：prepare 产出带锚点分块；validate 只预检；apply 校验后批量写入（低置信降级为批注）。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "action": {"type": "string", "enum": ["prepare","validate","apply"]},
                    "document_id": {"type": "string"},
                    "task": {"type": ["string","null"], "enum": ["proofread","polish","translate","continue","write","format","summarize","ambiguity_check","custom", null]},
                    "target": {"type": ["object","null"], "additionalProperties": true},
                    "language": {"type": ["string","null"]},
                    "instruction": {"type": ["string","null"]},
                    "max_chars_per_chunk": {"type": ["integer","null"]},
                    "overlap_chars": {"type": ["integer","null"]},
                    "include_outline": {"type": ["boolean","null"]},
                    "include_tables": {"type": ["boolean","null"]},
                    "max_items": {"type": ["integer","null"]},
                    "cursor": {"type": ["string","null"]},
                    "expected_doc_version": {"type": ["integer","null"]},
                    "apply_mode": {"type": ["string","null"], "enum": ["track_changes","comment","direct", null]},
                    "changes": {"type": ["array","null"], "items": {"type": "object", "additionalProperties": true}},
                    "validate_first": {"type": ["boolean","null"]},
                    "max_changes": {"type": ["integer","null"]},
                    "comment_below_confidence": {"type": ["number","null"]},
                    "confirm": {"type": ["string","null"]}
                },
                "required": ["action", "document_id"]
            }
        },
        {
            "type": "function",
            "name": "word_vba",
            "description": "VBA：status / run_existing_macro / execute_body / execute_module / cleanup_stale_temp_modules。注入默认关闭，需 allow_vba_injection=true + confirm=yes。",
            "parameters": {
                "type": "object",
                "additionalProperties": false,
                "properties": {
                    "action": {"type": "string", "enum": ["status","run_existing_macro","execute_body","execute_module","cleanup_stale_temp_modules"]},
                    "document_id": {"type": "string"},
                    "macro_name": {"type": ["string","null"]},
                    "args": {"type": ["array","null"], "items": {"type": "object", "additionalProperties": true}},
                    "entry": {"type": ["string","null"]},
                    "code": {"type": ["string","null"]},
                    "allow_vba_injection": {"type": ["boolean","null"]},
                    "cleanup": {"type": ["boolean","null"]},
                    "expected_doc_version": {"type": ["integer","null"]},
                    "confirm": {"type": ["string","null"]}
                },
                "required": ["action", "document_id"]
            }
        }
    ])
}
