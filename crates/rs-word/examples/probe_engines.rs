//! 引擎探测示例：只读连接 WPS Writer / Microsoft Word，打印身份与能力。
//!
//! 只做 Attach（接管已打开实例），不会关闭用户应用、不修改文档。
//! 运行：`cargo run -p rs-word --example probe_engines`

use rs_word::{ConnectMode, EnginePolicy, WordClient, WordEngine};

fn probe(label: &str, engine: WordEngine) {
    println!("\n=== {label} ===");
    let c = match WordClient::spawn() {
        Ok(c) => c,
        Err(e) => {
            println!("  无法启动 worker: {e}");
            return;
        }
    };
    match c.connect(EnginePolicy::Explicit(engine), ConnectMode::Attach) {
        Ok(status) => {
            if let Some(id) = &status.identity {
                println!(
                    "  已连接: name={:?} version={:?} prog_id={} pid={:?} exe={:?}",
                    id.name, id.version, id.prog_id, id.pid, id.exe
                );
            }
            println!("  文档 {} 个:", status.documents.len());
            for d in &status.documents {
                println!(
                    "    - {} [{}] saved={} ro={} 段={} 表={} 修订={}",
                    d.name,
                    d.document_id,
                    d.saved,
                    d.read_only,
                    d.paragraph_count.unwrap_or(0),
                    d.table_count.unwrap_or(0),
                    d.revision_count.unwrap_or(0),
                );
            }
            for cap in &status.capabilities {
                println!("    能力: {:?} {}", cap.status, cap.name);
            }
        }
        Err(e) => println!("  未连接（{label} 可能未运行）: {e}"),
    }
    c.shutdown();
}

fn main() {
    println!("rs-word 引擎探测（只读 Attach，不关闭任何应用）");
    probe("WPS Writer", WordEngine::WpsWriter);
    probe("Microsoft Word", WordEngine::MsWord);
    println!("\n完成。若两者均未连接，请先打开 WPS/Word 并新建或打开一个文档。");
}
