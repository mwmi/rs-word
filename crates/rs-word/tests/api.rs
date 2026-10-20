//! 库级单元测试（无 COM，纯逻辑）：serde、锚点、错误映射、颜色、工具 schema。

use rs_word::types::{FindOptions, Story, Target};
use rs_word::{anchors, error, tools};

#[test]
fn target_range_serde_roundtrip() {
    let t = Target::Range {
        document_id: "doc_1".into(),
        story: Story::Main,
        start: 120,
        end: 168,
        expected_text: Some("旧标题".into()),
        text_hash: Some("sha256:abc".into()),
        expected_doc_version: Some(17),
    };
    let j = serde_json::to_value(&t).unwrap();
    assert_eq!(j["kind"], "range");
    assert_eq!(j["start"], 120);
    let back: Target = serde_json::from_value(j).unwrap();
    match back {
        Target::Range { start, end, .. } => {
            assert_eq!(start, 120);
            assert_eq!(end, 168);
        }
        _ => panic!("kind 丢失"),
    }
}

#[test]
fn target_kind_is_rejected_when_unknown() {
    let bad = serde_json::json!({"kind": "nonsense", "document_id": "d"});
    assert!(serde_json::from_value::<Target>(bad).is_err());
}

#[test]
fn target_document_id_accessor() {
    let t = Target::Find {
        document_id: "doc_9".into(),
        query: "x".into(),
        occurrence: 1,
    };
    assert_eq!(t.document_id(), "doc_9");
}

#[test]
fn anchors_hash_stable_and_quote_truncates() {
    let a = anchors::hash_text("这是一个测试");
    assert_eq!(a, anchors::hash_text("这是一个测试"));
    assert!(a.starts_with("sha256:"));
    let long = "字".repeat(50);
    let q = anchors::quote(&long, 10);
    assert_eq!(q.chars().count(), 11); // 10 字 + …
    assert!(q.ends_with('…'));
}

#[test]
fn display_text_strips_cell_markers() {
    assert_eq!(anchors::display_text("单元格内容\r\x07"), "单元格内容");
    assert_eq!(anchors::display_text("段落\r"), "段落");
}

#[test]
fn tool_error_maps_hresult_to_stable_code() {
    let e = anyhow::Error::new(error::WordError::new(
        "call('Range')",
        "WPS > doc > Range",
        "0x80020009 异常",
    ));
    let te = error::ToolError::from_anyhow(&e);
    assert_eq!(te.code, "com_transient");
    assert!(te.retryable);
    assert_eq!(te.object_path, "WPS > doc > Range");
    assert_eq!(te.hresult.as_deref(), Some("0x80020009"));
}

#[test]
fn color_helpers_roundtrip() {
    let ole = error::rgb(0x12, 0x34, 0x56);
    assert_eq!(error::ole_to_hex(ole), "#123456");
    assert_eq!(error::hex_to_ole("#123456").unwrap(), ole);
    assert!(error::hex_to_ole("nope").is_err());
}

#[test]
fn tool_schemas_are_valid_and_named() {
    let defs = tools::tool_definitions();
    let arr = defs.as_array().expect("应为数组");
    let names: Vec<&str> = arr.iter().filter_map(|d| d["name"].as_str()).collect();
    for n in tools::tool_names() {
        assert!(names.contains(n), "schema 缺少工具 {n}");
    }
    // 每个工具都应有 parameters 对象且 additionalProperties=false。
    for d in arr {
        assert!(d["parameters"].is_object(), "{} 缺 parameters", d["name"]);
    }
}

#[test]
fn find_options_default_is_off() {
    let o = FindOptions::default();
    assert!(!o.match_case && !o.whole_word && !o.wildcards);
}

#[test]
fn detect_engine_by_path_beats_masqueraded_name() {
    use rs_word::session::engine::detect_engine;
    use rs_word::WordEngine;
    // WPS Writer 伪装 Name="Microsoft Word"，只能靠 Path 识别。
    assert_eq!(
        detect_engine(
            "D:\\Kingsoft\\WPS Office\\12.1.0\\office6",
            "Microsoft Word"
        ),
        Some(WordEngine::WpsWriter)
    );
    assert_eq!(
        detect_engine(
            "C:\\Program Files\\Microsoft Office\\Root\\Office16",
            "Microsoft Word"
        ),
        Some(WordEngine::MsWord)
    );
    // WPS 表格进程名（无 Path 变更时按 name 兜底）
    assert_eq!(
        detect_engine("", "WPS Spreadsheets"),
        Some(WordEngine::WpsWriter)
    );
    assert_eq!(detect_engine("", ""), None);
}

// ── 媒体（V1）──

#[test]
fn media_enum_codes_are_stable() {
    use rs_word::{AutoShapeKind, WrapType, ZOrderCommand};
    assert_eq!(AutoShapeKind::RoundedRectangle.code(), 5);
    assert_eq!(AutoShapeKind::Rectangle.code(), 1);
    assert_eq!(AutoShapeKind::Ellipse.code(), 9);
    assert_eq!(WrapType::Inline.code(), 7);
    assert_eq!(WrapType::Square.code(), 0);
    assert_eq!(ZOrderCommand::BringToFront.code(), 0);
    assert_eq!(ZOrderCommand::SendBackward.code(), 3);
}

#[test]
fn media_request_serde_roundtrip() {
    use rs_word::MediaRequest;
    let req: MediaRequest = serde_json::from_value(serde_json::json!({
        "document_id": "doc_1",
        "action": "insert_shape",
        "shape": "rounded_rectangle",
        "text": "说明",
        "width_pt": 160.0
    }))
    .unwrap();
    assert_eq!(req.document_id, "doc_1");
    assert!(matches!(req.action, rs_word::MediaAction::InsertShape));
    assert_eq!(req.shape, Some(rs_word::AutoShapeKind::RoundedRectangle));
    assert!(!req.inline);
    let back = serde_json::to_value(&req).unwrap();
    assert_eq!(back["action"], "insert_shape");
}

// ── 表格 V1 格式 ──

#[test]
fn table_format_spec_validation() {
    use rs_word::{TableEdge, TableFormatSpec};
    let ok = TableFormatSpec {
        widths_pt: Some(vec![80.0, 120.0]),
        ..Default::default()
    };
    assert!(ok.validate().is_ok());
    let bad = TableFormatSpec {
        widths_pt: Some(vec![0.5]),
        ..Default::default()
    };
    assert!(bad.validate().is_err());
    let empty = TableFormatSpec::default();
    assert!(empty.validate().is_err());
    // All 展开为 6 条边。
    assert_eq!(TableEdge::All.codes().len(), 6);
    assert_eq!(rs_word::BorderStyle::None.code(), 0);
    let border = TableFormatSpec {
        borders: Some(vec![rs_word::BorderSpec {
            edge: TableEdge::Top,
            style: rs_word::BorderStyle::Double,
            width_pt: Some(1.0),
            color: Some("#FF0000".into()),
        }]),
        ..Default::default()
    };
    assert!(border.validate().is_ok());
    let heavy = TableFormatSpec {
        borders: Some(vec![rs_word::BorderSpec {
            edge: TableEdge::Top,
            style: rs_word::BorderStyle::Single,
            width_pt: Some(200.0),
            color: None,
        }]),
        ..Default::default()
    };
    assert!(heavy.validate().is_err());
}

// ── 修订筛选批量 ──

#[test]
fn revision_filter_matching() {
    use rs_word::{review, RevisionFilter, RevisionInfo};
    let info = RevisionInfo {
        revision_ref: "rev_1_abc".into(),
        index: 1,
        kind: "insert".into(),
        author: Some("张三".into()),
        date: None,
        quote: "新增".into(),
        start: Some(10),
        text_hash: "sha256:x".into(),
        doc_version: 1,
    };
    // 空筛选 = 全匹配。
    assert!(review::matches_filter(&info, &RevisionFilter::default()));
    let by_author = RevisionFilter {
        authors: Some(vec!["张三".into()]),
        ..Default::default()
    };
    assert!(review::matches_filter(&info, &by_author));
    let other_author = RevisionFilter {
        authors: Some(vec!["李四".into()]),
        ..Default::default()
    };
    assert!(!review::matches_filter(&info, &other_author));
    let by_kind = RevisionFilter {
        kinds: Some(vec!["delete".into()]),
        ..Default::default()
    };
    assert!(!review::matches_filter(&info, &by_kind));
}

// ── 结构 diff ──

#[test]
fn structure_diff_detects_changes() {
    use rs_word::advanced::diff_snapshots;
    use rs_word::{StructureParagraph, StructureSnapshot, TableShapeSnapshot};
    let para = |idx: u32, style: &str, hash: &str| StructureParagraph {
        index: idx,
        style: style.into(),
        outline_level: 10,
        chars: 4,
        quote: "正文".into(),
        text_hash: hash.into(),
    };
    let base = StructureSnapshot {
        document_id: "doc_1".into(),
        doc_version: 1,
        name: "a.docx".into(),
        section_count: 1,
        paragraphs: vec![
            para(1, "Normal", "sha256:aaa"),
            para(2, "Normal", "sha256:bbb"),
        ],
        tables: vec![TableShapeSnapshot {
            index: 1,
            rows: 2,
            columns: 3,
            preview: "字段 | 数量".into(),
        }],
        note_count: 0,
        media_count: 0,
    };
    // 同一快照 → 无变更。
    assert!(diff_snapshots(&base, &base).is_empty());

    let mut changed = base.clone();
    changed.paragraphs[0].style = "Heading 1".into();
    changed.paragraphs[1].text_hash = "sha256:ccc".into();
    changed.paragraphs.push(para(3, "Normal", "sha256:ddd"));
    changed.tables[0].rows = 3;
    let kinds: Vec<String> = diff_snapshots(&base, &changed)
        .into_iter()
        .map(|c| c.kind)
        .collect();
    assert!(kinds.contains(&"paragraph_style_changed".to_string()));
    assert!(kinds.contains(&"paragraph_text_changed".to_string()));
    assert!(kinds.contains(&"paragraph_inserted".to_string()));
    assert!(kinds.contains(&"table_shape_changed".to_string()));

    // 段落/表格减少。
    let removed = StructureSnapshot {
        paragraphs: vec![para(1, "Normal", "sha256:aaa")],
        tables: Vec::new(),
        ..base.clone()
    };
    let kinds: Vec<String> = diff_snapshots(&base, &removed)
        .into_iter()
        .map(|c| c.kind)
        .collect();
    assert!(kinds.contains(&"paragraph_removed".to_string()));
    assert!(kinds.contains(&"table_removed".to_string()));
}

// ── 工作流 ──

#[test]
fn workflow_apply_defaults_and_change_classification() {
    use rs_word::workflow::{classify_change, EffectiveOp};
    use rs_word::{
        Target, WorkflowApplyMode, WorkflowApplyRequest, WorkflowChange, WorkflowChangeOp,
        WorkflowTask,
    };
    let req: WorkflowApplyRequest = serde_json::from_value(serde_json::json!({
        "document_id": "doc_1",
        "task": "proofread",
        "expected_doc_version": 3,
        "changes": []
    }))
    .unwrap();
    assert_eq!(req.max_changes, 100);
    assert_eq!(req.comment_below_confidence, 0.75);
    assert!(req.validate_first);
    assert_eq!(req.apply_mode, WorkflowApplyMode::TrackChanges);
    assert!(matches!(req.task, WorkflowTask::Proofread));

    let change = |op, confidence| WorkflowChange {
        op,
        target: Target::Document {
            document_id: "doc_1".into(),
        },
        text: Some("新".into()),
        reason: None,
        confidence,
        severity: None,
    };
    // 高置信 track changes → 替换。
    assert_eq!(
        classify_change(
            &change(WorkflowChangeOp::Replace, Some(0.98)),
            WorkflowApplyMode::TrackChanges,
            0.75
        ),
        EffectiveOp::Replace
    );
    // 低置信 → 降级为批注。
    assert_eq!(
        classify_change(
            &change(WorkflowChangeOp::Replace, Some(0.5)),
            WorkflowApplyMode::TrackChanges,
            0.75
        ),
        EffectiveOp::Comment
    );
    // comment 模式一律批注；direct 模式不降级。
    assert_eq!(
        classify_change(
            &change(WorkflowChangeOp::Replace, Some(0.99)),
            WorkflowApplyMode::Comment,
            0.75
        ),
        EffectiveOp::Comment
    );
    assert_eq!(
        classify_change(
            &change(WorkflowChangeOp::InsertAfter, Some(0.1)),
            WorkflowApplyMode::Direct,
            0.75
        ),
        EffectiveOp::InsertAfter
    );
}

// ── VBA 注入 ──

#[test]
fn vba_body_wrapping_and_code_guard() {
    use rs_word::vba::banned_construct;
    use rs_word::VbaCodeKind;
    let wrapped = VbaCodeKind::Body.wrap("RsWordEntry", "Dim r As Range");
    assert!(wrapped.starts_with("Option Explicit"));
    assert!(wrapped.contains("Public Sub RsWordEntry()"));
    assert!(wrapped.trim_end().ends_with("End Sub"));
    // module 模式原样使用。
    let raw = "Public Function F() As String\nF = \"OK\"\nEnd Function\n";
    assert_eq!(VbaCodeKind::Module.wrap("F", raw), raw);
    // 代码守卫：模态框/外部进程被拒。
    assert_eq!(banned_construct("MsgBox \"hi\""), Some("MsgBox"));
    assert_eq!(banned_construct("Shell \"cmd\""), Some("Shell"));
    assert!(banned_construct("Dim r As Range\nr.InsertAfter \"x\"").is_none());
}

// ── 编号模板 / 工具 schema ──

#[test]
fn list_template_validation() {
    use rs_word::{ListLevelSpec, ListNumberFormat, ListTemplateSpec};
    let ok = ListTemplateSpec {
        levels: vec![
            ListLevelSpec {
                level: 1,
                format: ListNumberFormat::Decimal,
                text: Some("%1.".into()),
                start_at: 1,
                number_position_pt: 0.0,
                text_position_pt: 18.0,
                tab_position_pt: Some(18.0),
                trailing: rs_word::ListTrailing::Tab,
            },
            ListLevelSpec {
                level: 2,
                format: ListNumberFormat::LowerAlpha,
                text: None,
                start_at: 1,
                number_position_pt: 18.0,
                text_position_pt: 36.0,
                tab_position_pt: None,
                trailing: rs_word::ListTrailing::Space,
            },
        ],
        continue_previous: false,
    };
    assert!(ok.validate().is_ok());

    let dup = ListTemplateSpec {
        levels: vec![
            ok.levels[0].clone(),
            ListLevelSpec {
                level: 1,
                ..ok.levels[0].clone()
            },
        ],
        continue_previous: false,
    };
    assert!(dup.validate().is_err());

    let bad_bullet = ListTemplateSpec {
        levels: vec![ListLevelSpec {
            level: 1,
            format: ListNumberFormat::Bullet,
            text: Some("ab".into()),
            start_at: 1,
            number_position_pt: 0.0,
            text_position_pt: 10.0,
            tab_position_pt: None,
            trailing: rs_word::ListTrailing::Tab,
        }],
        continue_previous: false,
    };
    assert!(bad_bullet.validate().is_err());
    assert!(ListTemplateSpec {
        levels: Vec::new(),
        continue_previous: true,
    }
    .validate()
    .is_err());
}

#[test]
fn new_tools_are_registered_with_expected_actions() {
    let defs = tools::tool_definitions();
    let arr = defs.as_array().unwrap();
    for name in ["word_media", "word_advanced", "word_workflow"] {
        let d = arr
            .iter()
            .find(|d| d["name"] == name)
            .unwrap_or_else(|| panic!("schema 缺少工具 {name}"));
        let acts: Vec<String> = d["parameters"]["properties"]["action"]["enum"]
            .as_array()
            .map(|a| {
                a.iter()
                    .filter_map(|v| v.as_str().map(str::to_string))
                    .collect()
            })
            .unwrap_or_default();
        assert!(!acts.is_empty(), "{name} 缺少 action 枚举");
    }
    // word_table 必须暴露 V1 的四个格式动作。
    let table = arr.iter().find(|d| d["name"] == "word_table").unwrap();
    let acts: Vec<&str> = table["parameters"]["properties"]["action"]["enum"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|v| v.as_str())
        .collect();
    for a in [
        "set_widths",
        "set_heights",
        "set_borders",
        "set_allow_row_break",
    ] {
        assert!(acts.contains(&a), "word_table 缺少动作 {a}");
    }
}
