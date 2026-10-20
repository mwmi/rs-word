//! 基本用法示例（只读为主）：Attach 当前文档，读摘要/大纲、查找并预览替换、查看修订。
//!
//! 运行：`cargo run -p rs-word --example basic`
//! 默认不写入文档（仅 find/replace preview 是只读）。

use rs_word::{
    ConnectMode, EnginePolicy, FindOptions, FindRequest, ReadRequest, ReadView,
    ReplacePreviewRequest, Target, WordClient, WordEngine,
};

fn main() -> anyhow::Result<()> {
    // 默认接 WPS；如需 Word 改成 WordEngine::MsWord。
    let engine = WordEngine::WpsWriter;
    let c = WordClient::spawn().map_err(anyhow::Error::new)?;
    let status = c
        .connect(EnginePolicy::Explicit(engine), ConnectMode::Attach)
        .map_err(anyhow::Error::new)?;

    let doc = status
        .documents
        .first()
        .ok_or_else(|| anyhow::anyhow!("当前没有打开的文档"))?;
    let id = doc.document_id.clone();
    println!("使用文档: {} [{}]", doc.name, id);

    // 读摘要（默认视图，不会返回全文）。
    let summary = c
        .read(ReadRequest {
            document_id: id.clone(),
            target: Target::Document {
                document_id: id.clone(),
            },
            view: ReadView::Summary,
            limit_chars: 4000,
            max_items: 30,
            cursor: None,
        })
        .map_err(anyhow::Error::new)?;
    println!("摘要: {:?}", summary.summary);
    for item in summary.outline.iter().take(10) {
        println!("  大纲 L{} {}", item.level, item.text);
    }

    // 查找（只读）。
    let found = c
        .find(FindRequest {
            document_id: id.clone(),
            query: "的".into(),
            options: FindOptions::default(),
            limit: 5,
        })
        .map_err(anyhow::Error::new)?;
    println!(
        "查找 '的' 命中 {} 处（展示 {} 条）",
        found.total,
        found.matches.len()
    );

    // 替换预览（只读，不改文档）。
    let preview = c
        .replace_preview(ReplacePreviewRequest {
            document_id: id.clone(),
            find: "的".into(),
            replace: "之".into(),
            options: FindOptions::default(),
            limit: 20,
        })
        .map_err(anyhow::Error::new)?;
    println!(
        "替换预览: {} 处（preview_id={}），执行替换需再调用 replace_apply",
        preview.total, preview.preview_id
    );

    c.shutdown();
    Ok(())
}
