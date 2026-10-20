//! 基本用法示例：连接已打开的 WPS/Word → 在当前文档末尾追加演示文本
//! → 查找 → 替换预览 → 执行替换 → 回读核对。
//!
//! 运行：`cargo run -p rs-word --example basic`
//!
//! **Attach 模式：不新建 Word/WPS 进程，也不新建文档窗口**，直接写入你当前打开的那个文档。
//! 示例会向该文档末尾追加两段演示文字；替换前核对命中数，只处理本次追加的内容。
//! 示例不保存，避免改动你磁盘上的文件。

use rs_word::{
    ConnectMode, EditAction, EditRequest, EnginePolicy, FindOptions, FindRequest, InsertPosition,
    ReadRequest, ReadView, ReplaceApplyRequest, ReplacePreviewRequest, Target, WordClient,
    WordEngine,
};

/// 写入的演示文本（两段；`\r` = 段落标记，与库内 edit 的约定一致）。
const DEMO_TEXT: &str = "rs-word 演示第一段：苹果派与苹果汁。\r第二段：苹果很好吃。";
/// 查找/替换的关键词。
const FIND: &str = "苹果";
const REPLACE: &str = "橙子";
/// 追加文本中 `FIND` 出现的次数（用于核对替换范围，避免动到文档原有内容）。
const INSERTED_HITS: u32 = 3;

fn main() -> anyhow::Result<()> {
    // ── 1. 连接已运行的 WPS/Word（Attach，不新建进程）
    // WPS 会留下无文档的后台实例，因此逐个探测，取真正有文档的那个引擎。
    let (client, status) = [WordEngine::WpsWriter, WordEngine::MsWord]
        .into_iter()
        .find_map(|engine| {
            let c = WordClient::spawn().ok()?;
            match c.connect(EnginePolicy::Explicit(engine), ConnectMode::Attach) {
                Ok(s) if !s.documents.is_empty() => Some((c, s)),
                _ => {
                    c.shutdown();
                    None
                }
            }
        })
        .ok_or_else(|| anyhow::anyhow!("未找到已打开文档的 WPS/Word 实例，请先打开一个文档"))?;

    let engine = status.engine.map(|e| e.label()).unwrap_or("未知引擎");
    println!(
        "已连接 {engine}；当前已打开 {} 个文档（示例会向当前文档追加演示文字，不新建窗口）",
        status.documents.len()
    );

    // ── 2. 取当前窗口里打开的文档（不新建文档，避免 Word 再开一个窗口）
    let doc = status
        .active_document_id
        .clone()
        .or_else(|| status.documents.first().map(|d| d.document_id.clone()))
        .ok_or_else(|| anyhow::anyhow!("已连接但没有打开的文档"))?;
    let name = status
        .documents
        .iter()
        .find(|d| d.document_id == doc)
        .map(|d| d.name.clone())
        .unwrap_or_default();
    println!("\n写入目标: {name} [{doc}]（追加到文档末尾，不新建文档）");
    println!("  写入前: {}", read_summary(&client, &doc)?);

    // ── 3. 在文档末尾追加两段文本（validate_all_first=true，先全量预检再按序应用）
    let edit = client.edit(EditRequest {
        document_id: doc.clone(),
        expected_doc_version: None,
        validate_all_first: true,
        track_changes: false,
        idempotency_key: None,
        actions: vec![EditAction::InsertText {
            target: Target::Document {
                document_id: doc.clone(),
            },
            text: format!("\r{DEMO_TEXT}"),
            position: InsertPosition::End,
        }],
    })?;
    println!(
        "  写入完成: applied={} failed={} 副作用={:?}",
        edit.applied, edit.failed, edit.side_effect_status
    );
    for a in &edit.actions {
        if !a.ok {
            println!("    失败动作#{}: {}", a.index, a.summary);
        }
    }
    println!("  写入后: {}", read_summary(&client, &doc)?);

    // ── 4. 查找（只读）
    let found = client.find(FindRequest {
        document_id: doc.clone(),
        query: FIND.into(),
        options: FindOptions::default(),
        limit: 5,
    })?;
    println!("\n查找 '{FIND}': 命中 {} 处", found.total);
    for m in &found.matches {
        println!(
            "  #{} 位置[{}, {}] {:?}",
            m.occurrence, m.start, m.end, m.quote
        );
    }

    // ── 5. 替换预览（只读；返回 preview_id 与 doc_version）
    let preview = client.replace_preview(ReplacePreviewRequest {
        document_id: doc.clone(),
        find: FIND.into(),
        replace: REPLACE.into(),
        options: FindOptions::default(),
        limit: 20,
    })?;
    println!(
        "\n替换预览: '{FIND}' → '{REPLACE}' 共 {} 处（preview_id={}）",
        preview.total, preview.preview_id
    );

    // ── 6. 执行替换（回传 preview_id + expected_doc_version，目标失效会拒绝执行）
    // 只在命中数与本次追加文本一致时执行，避免改到文档原有内容。
    if preview.total == INSERTED_HITS {
        let replaced = client.replace_apply(ReplaceApplyRequest {
            document_id: doc.clone(),
            preview_id: preview.preview_id.clone(),
            find: FIND.into(),
            replace: REPLACE.into(),
            options: FindOptions::default(),
            expected_doc_version: preview.doc_version,
            max_replacements: 20,
            track_changes: false,
        })?;
        println!(
            "替换执行: {} 处（副作用={:?} operation_id={}）",
            replaced.replaced, replaced.side_effect_status, replaced.operation_id
        );
    } else {
        println!(
            "跳过替换：命中 {} 处 ≠ 本次追加的 {} 处，为避免改动文档原有内容不执行",
            preview.total, INSERTED_HITS
        );
    }

    // ── 7. 回读核对
    println!("\n替换后正文:\n{}", indent(&read_text(&client, &doc)?));
    println!("  替换后: {}", read_summary(&client, &doc)?);
    println!("\n示例不保存：Attach 模式不写你磁盘上的文件，改动可在编辑器里撤销。");

    // 只关闭本库 worker，不关闭 WPS/Word 实例（Attach 模式）。
    client.shutdown();
    Ok(())
}

/// 读摘要行（summary 视图）。
fn read_summary(client: &WordClient, id: &str) -> anyhow::Result<String> {
    let r = client.read(ReadRequest {
        document_id: id.to_string(),
        target: Target::Document {
            document_id: id.to_string(),
        },
        view: ReadView::Summary,
        limit_chars: 4000,
        max_items: 10,
        cursor: None,
    })?;
    Ok(r.summary.unwrap_or_default())
}

/// 读正文（text 视图；段落标记 `\r` 换成换行，便于终端查看）。
fn read_text(client: &WordClient, id: &str) -> anyhow::Result<String> {
    let r = client.read(ReadRequest {
        document_id: id.to_string(),
        target: Target::Document {
            document_id: id.to_string(),
        },
        view: ReadView::Text,
        limit_chars: 4000,
        max_items: 10,
        cursor: None,
    })?;
    Ok(r.chunks
        .iter()
        .map(|c| c.text.replace('\r', "\n"))
        .collect::<String>())
}

/// 多行文本整体缩进。
fn indent(s: &str) -> String {
    s.lines()
        .map(|l| format!("    {l}"))
        .collect::<Vec<_>>()
        .join("\n")
}
