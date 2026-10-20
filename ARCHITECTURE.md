# rs-word 架构

## 分层

```text
AI Agent / Tauri UI / rs-word-cli
        │  typed DTO（serde JSON 仅在工具/CLI 边界）
        ▼
WordClient（Send + Sync channel 代理，不持有 COM）
        │  Request / Response（typed enum，经 crossbeam-channel）
        ▼
STA worker 线程（CoInitializeEx STA + 消息泵）
        │
        ▼
Session（会话状态：app / 文档槽位 / doc_version / 幂等缓存 / History / capabilities）
        │
        ├── App / Document / Range 包装（IDispatchWrapper）
        ├── 领域模块：target / anchors / read / edit / find / format / table / layout / review /
        │             media / advanced / workflow / history / vba
        └── Engine adapter：MsWordAdapter / WpsWriterAdapter（ProgID 候选 + 身份核验）
        ▼
com::dispatch（IDispatchWrapper：get/call/call_named/put）+ com::variant（ComValue/WordValue）
```

## STA worker 与线程模型

- `session::worker::spawn()` 起一个命名线程 `rs-word-sta`：`CoInitializeEx(COINIT_APARTMENTTHREADED)`，随后循环 `PeekMessage/TranslateMessage/DispatchMessage` + `recv`，串行处理每条 `Command`，退出时 `CoUninitialize`。
- 所有 `IDispatch`/`VARIANT` 只在该线程创建/使用/Drop；COM 资源只在这里释放。
- `WordClient` 内部只有一个 `Sender<Command>`，可跨线程 clone；不暴露任何 COM 指针。
- **超时语义**：调用方在 oneshot 上等待超时后返回 `ToolError{code:"result_unknown", side_effect_status:"unknown"}`，**不**声称取消了正在执行的 COM 调用。重新确认状态前禁止重发写操作。

## 会话与文档 ID

- `Session` 持有 `app: App`、`docs: Vec<Slot>`（`id/key/doc`）、`versions`、`idem`、`history`、`caps`、
  `owned_modules`（本库创建的 VBA 临时模块名）+ `vba_seq`。
- 文档稳定 ID `doc_<n>` 由库分配（按 FullName / Name 匹配，避免只凭窗口标题定位）。
- `doc_version`：库内自增计数，写操作成功即 +1；读取返回当前值；修改前用 `expected_doc_version` + `expected_text`/`text_hash` 共同判定 `stale_target`。
- 引擎切换清空 `docs/versions/idem/owned_modules`，旧 `document_id` 立即失效。

## 媒体 / 高级 / 工作流（V1+V2）

- **媒体**（`media.rs`）：`inline=false` 走 `Range.InlineShapes.AddPicture`，`inline=true` 走 `Document.Shapes.*`；
  形状/文本框/WordArt 统一 `Document.Shapes`。对象标识 `shape_<n>`/`inline_<n>`（无会话状态），写入前核对
  `Name`（可选 `expect_name`）与 `Type`，不符返回 `stale_target`；几何单位统一 pt；`AddTextEffect` 缺失时
  返回 `unsupported_capability`，**不**降级为普通形状。
- **高级**（`advanced.rs`）：脚注/尾注走 `Document.Footnotes|Endnotes.Add(...)`（独立 story，不回写正文 Range）；
  交叉引用走 `Range.InsertCrossReference`（只实现已验证的 type/kind 子集）；结构快照只读采集段落样式/表格形状/计数，
  `diff_snapshots` 是**纯函数**（可离线单测），`structure_diff` 用调用方回传的基线快照比较，无会话状态。
- **工作流**（`workflow.rs`）：`prepare` 只读产出**一个** chunk（≤ `max_chars_per_chunk`，`next_cursor` 指向
  下一块 = 终点 − overlap）+ 锚点/hash/大纲；`validate` 仅预检不写；`apply` 先校验 `expected_doc_version`，
  `validate_first` 时全量预检 target 后逐条应用（应用前逐条复核锚点），`track_changes`/`direct`/`comment` 三种模式，
  置信度低于 `comment_below_confidence` 自动降级为批注并记 `warnings`；`apply_mode=direct` 需 `confirm="yes"`。
  库不调用模型，只管理「输入包」与「结构化 patch」。

## 修订 / 历史 / VBA

- Track Changes：`Document.TrackRevisions` 保存旧值 → 开 → 改 → 回读 Revisions → **恢复旧值**（`review::with_tracking`）；修订由跟踪模式下的常规编辑产生，不伪造 Revision 对象。
- 修订筛选批量：`accept_filtered/reject_filtered` 需 `confirm="yes"`，**倒序**遍历（接受/拒绝会移除条目），
  逐条重读类型/作者并按 `RevisionFilter{authors,kinds,max}` 匹配，不匹配计入 `skipped`。
- History：库级 journal（operation_id / before-after version / hash / reversible）+ `SaveCopyAs` 检查点；`Undo/Redo` 走原生 `Application.Undo/Redo`；不可精确逆转标 `reversible=false`（媒体/注入等）。
- VBA：只读 `status`（`GetIDsOfNames` 探测 `Run`/`VBProject`，不执行、不改安全设置）+ `run_existing_macro`
  （`Application.Run`，≤30 参，`Sub` 返回 Empty 不等于失败）；临时注入 `execute_body`/`execute_module` 默认关闭，
  需 `allow_vba_injection=true` + `confirm="yes"`：模块名 `RsWordTmp_<seq>_<ms>` 并登记归属，
  结束只删除本次自建模块，失败返回 `vba_cleanup_failed`；代码守卫拒绝 `MsgBox`/`Shell`（模态框会阻塞 STA 线程）。
  `cleanup_stale_temp_modules` 只清理登记归属的 `RsWordTmp_*`。

## Target 与锚点

- `Target` 是 tagged enum：`document | selection | range | paragraph | find | bookmark | table_cell | header_footer`。
- `selection` 在命令开始时**一次性**捕获为固定 Range（Start/End/story/摘录/hash），后续只用该 Range。
- range 用 Word `Start/End`（**不是** Rust UTF-8 字节下标）；跨 story 位置禁止混用；MVP 仅 `main` story 的 range，「页眉/页脚」用 `header_footer` target。
- 修改前调用 `anchors::verify`（`expected_text` / `text_hash` 不符）返回 `stale_target`。

## 错误模型

- 内部 `WordError`：单行 `[操作] @ 对象路径 | 底层原因 | 提示: 建议`，由 `com::dispatch` 自动带上对象路径。
- 对外 `ToolError`：`code` / `category` / `message` / `hint` / `hresult` / `retryable` / `side_effect_status` / `next_action`。
  - 分类：`Validation / TargetNotFound / Conflict / Permission / Unsupported / ComTransient / ComFailure / Recovery`。
  - `retryable=true` 仅用于已确认无副作用的瞬时错误（WPS 常见 `0x80020009`）。
  - HRESULT 映射稳定错误码（见 `error.rs::from_anyhow`），调用方无需解析英文文案。

## 写操作可验证

- 每个写操作返回 `operation_id`、分步结果、`doc_version`、`side_effect_status`。
- `word_edit`：先全量预检 target（`validate_all_first`）→ 有序应用 → 逐动作 read-back；部分成功返回 `partially_applied` + 每步状态。
- Find/Replace：先 `preview_replace`（返回总数/样例/preview_id），执行需二次携带 `expected_doc_version` + `max_replacements`。
- 幂等：`EditRequest.idempotency_key` 命中缓存时返回首次结果，防止模型重试造成重复写入。

## 修订 / 历史 / VBA

- Track Changes：`Document.TrackRevisions` 保存旧值 → 开 → 改 → 回读 Revisions → **恢复旧值**（`review::with_tracking`）；修订由跟踪模式下的常规编辑产生，不伪造 Revision 对象。
- History：库级 journal（operation_id / before-after version / hash / reversible）+ `SaveCopyAs` 检查点；`Undo/Redo` 走原生 `Application.Undo/Redo`；不可精确逆转标 `reversible=false`。
- VBA：只读 `status`（`GetIDsOfNames` 探测 `Run`/`VBProject`，不执行、不改安全设置）+ `run_existing_macro`（`Application.Run`，≤30 参，`Sub` 返回 Empty 不等于失败）；临时注入返回 `unsupported_capability`（V1）。
