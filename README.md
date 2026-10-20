# rs-word

面向 AI 助手的 **Rust + COM 自动化 Word / WPS Writer 库**（Windows 专用），基于 `IDispatch` 后期绑定，配套交互式命令行工具 `rs-word-cli`。

> 本仓库覆盖 **MVP + V1/V2**：连接/切换、文档生命周期、Target/Range/锚点、文本读写、Find/Replace（含预览）、字符/段落格式、样式、列表（默认 + **自定义多级编号模板**）、表格基础 + **列宽/行高/边框/跨页行**、**媒体（图片/形状/文本框/WordArt）**、页面/页眉页脚/字段/PDF、修订与批注（含**按条件筛选批量**）、**脚注尾注/交叉引用/文档结构 diff**、简易历史、VBA `status` + `run_existing_macro` + **临时注入 `execute_body`/`execute_module`（默认关闭）**、**AI 工作流 prepare/validate/apply**、AI Tool 门面与能力矩阵。详见 [CAPABILITY_MATRIX.md](CAPABILITY_MATRIX.md) 与「实现边界」。

## 构建与运行

```bash
cargo build                       # 构建库与 CLI
cargo test --workspace            # 单元测试（无 COM）
cargo run -p rs-word-cli          # 交互式 REPL
cargo run -p rs-word --example probe_engines   # 引擎只读探测
cargo run -p rs-word --example basic           # 基本只读用法
```

要求：Windows + Rust 1.80+（本仓库用 `windows` 0.62）。仅 Windows 可编译（`#![cfg(windows)]`）。

## 集成到 Tauri / 宿主

```rust
use rs_word::prelude::*;
use rs_word::{ConnectMode, EnginePolicy, ReadRequest, ReadView, Target, WordClient};

let client = WordClient::spawn()?;                                    // 启动 STA worker
let status = client.connect(EnginePolicy::Explicit(WordEngine::WpsWriter), ConnectMode::Attach)?;
let doc = &status.documents[0];
let read = client.read(ReadRequest {
    document_id: doc.document_id.clone(),
    target: Target::Document { document_id: doc.document_id.clone() },
    view: ReadView::Summary,
    limit_chars: 4000,
    max_items: 30,
    cursor: None,
})?;
```

- `WordClient` 是 `Send + Sync` 的 channel 代理；**Tauri/async/AI 线程只发 typed DTO，绝不接触 `IDispatch`/`VARIANT`**。
- 所有 COM 对象只在专用 **STA worker 线程**创建/使用/Drop（含消息泵）。用 `client.shutdown()` 停止 worker（不影响已连接的 Office 进程）。

## 引擎与权限

- 引擎：`WordEngine::WpsWriter` / `WordEngine::MsWord`；模式 `Attach`（默认，接管已运行实例）/ `New`（新建）。
- WPS ProgID **不写死**：候选 `KWPS.Application` → `KWPS.Application.9`，连接后按 `Application.Path` 核验归属。
  ⚠ 真机实测：**Word / WPS Writer 的 `Application` 没有 `Hwnd` 成员**（那是 Excel 专有），且 WPS Writer 会伪装
  `Application.Name="Microsoft Word"`。因此引擎识别改用 `Path` 特征（WPS 含 `Kingsoft`/`WPS Office`，MS 含 `Microsoft Office`），
  避免把 WPS 误当 MS Word 或反之；详见 [CAPABILITY_MATRIX.md](CAPABILITY_MATRIX.md)。
- 安全：`Attach` 不关闭用户应用；接受/拒绝全部修订、按条件筛选批量、覆盖导出、关闭脏文档、任意 VBA 等高风险动作需显式确认或策略授权；**VBA 临时注入默认关闭**（需 `allow_vba_injection=true` + `confirm="yes"`，并先 checkpoint）。

## 能力说明

首次连接后 `status.capabilities` 逐项给出 `supported / partial / unsupported / untested`（按 `GetIDsOfNames` 只读探测）。**未在真机验证的能力一律为 `untested`**，不会伪报为 supported。详见 [CAPABILITY_MATRIX.md](CAPABILITY_MATRIX.md)。

## 目录

- [ARCHITECTURE.md](ARCHITECTURE.md)：STA worker、adapter、target、error、history 架构。
- [API_REFERENCE.md](API_REFERENCE.md)：Rust typed API 与 AI 工具 JSON Schema。
- [CAPABILITY_MATRIX.md](CAPABILITY_MATRIX.md)：Word/WPS 分列能力状态。

## 实现边界（诚实标注）

原「未实现项」的 6 个 V1 项 + 4 个 V2 项**均已落地为真实实现**（不再是硬编码拒绝）：

- **V1**：媒体（`word_media`：图片/形状/文本框/WordArt）、自定义多级编号模板（`list_template`）、
  表格 `set_widths/set_heights/set_borders/set_allow_row_break`、修订筛选批量（`accept_filtered`/`reject_filtered`）、
  AI 工作流（`word_workflow` prepare/validate/apply）。
- **V2**：脚注尾注、交叉引用、文档结构快照/diff（`word_advanced`）、VBA 临时注入（`execute_body`/`execute_module`）。

仍存在的边界（不谎报）：

- **真机验证范围**：WPS Writer 12.1.0.29161 下已实测通过 14 项新能力（证据见
  [CAPABILITY_MATRIX.md](CAPABILITY_MATRIX.md)）；**Microsoft Word 本机未运行 → Word 列全部 `untested`**。
  未实测项包括：尾注、`execute_module`、Word 侧全部行为、`连接 New` 是否独立于用户实例。
- **已实测为 partial**：`word_advanced.list_cross_references` —— WPS 下 `Content.Fields.Count=0`，
  插入的 REF 域枚举不到（插入本身可用）；结果中会回报 `fields=N refs=M`。
- **默认关闭 / 需显式授权**：VBA 临时注入（`allow_vba_injection=true` + `confirm="yes"`）；
  `apply_mode=direct`（不产生修订/批注直接改原文）；`accept_all/reject_all/accept_filtered/reject_filtered`（`confirm="yes"`）。
- **未验证的 Word 常量不猜**：交叉引用的 `type/kind` 只实现已验证子集，其余返回 `unsupported_capability`；
  WordArt 缺失时同样返回 `unsupported_capability`，**不**静默降级为普通形状。
- **结构 diff 的语义**：比较的是**段落样式/文本 hash/表格形状/计数**级快照（不是 OOXML 级 diff），
  基线快照由调用方保存并回传。
- **撤销诚实标注**：媒体、表格格式、VBA 注入等记入历史时 `reversible=false` 并写明原因，不做假撤销。
- **仍然不做**：段落边框/底纹本体实现、目录（TOC）生成、表格排序、文本↔表格互转、文档属性（V2 之外的更远项）。
