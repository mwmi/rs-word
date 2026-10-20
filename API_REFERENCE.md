# rs-word API 参考

所有 Rust API 由 `WordClient` 暴露（`Send + Sync`），调用方不接触 COM。返回值统一 `Result<T, ToolError>`。

## 1. Rust typed API（`WordClient`）

| 方法 | 说明 |
|---|---|
| `spawn() -> Result<WordClient, ToolError>` | 启动 STA worker（未连接引擎） |
| `connect(policy, mode) -> WordStatus` | 连接引擎；`policy=EnginePolicy::Explicit/PreferWps/PreferMsWord`，`mode=Attach/New` |
| `switch_engine(engine, dirty_policy) -> WordStatus` | 切换引擎（脏文档按 `RefuseIfDirty/Save/Discard`） |
| `status() -> WordStatus` | 会话状态（引擎身份、文档列表、能力、拥有实例） |
| `list_documents() -> Vec<DocumentInfo>` | 文档列表 |
| `open(path, read_only) -> DocumentInfo` / `new_document(template) -> DocumentInfo` | 打开/新建 |
| `activate(document_id) -> DocumentInfo` | 激活文档 |
| `save(document_id, path, format) -> SaveResult` | 保存/另存（`SaveFormat::{Keep,Docx,Doc,Rtf,Txt,Html}`） |
| `close(document_id, policy) -> CloseResult` | 关闭（`ClosePolicy::{RefuseIfDirty,Save,Discard}`） |
| `read(ReadRequest) -> ReadResult` | 读取（视图/分块/游标） |
| `edit(EditRequest) -> EditResult` | 批量文本编辑 |
| `find(FindRequest) -> FindResult` / `replace_preview(..) -> ReplacePreview` / `replace_apply(..) -> ReplaceResult` | 查找与替换（预览→执行） |
| `format(FormatRequest) -> FormatResult` | 字符/段落/样式/列表 |
| `table(TableRequest) -> TableResult` | 表格操作 |
| `layout(LayoutRequest) -> LayoutResult` | 页面/页眉页脚/PDF/字段 |
| `review(ReviewRequest) -> ReviewResult` | 修订/批注（含 `accept_filtered`/`reject_filtered` 筛选批量） |
| `media(MediaRequest) -> MediaResult` | 图片/形状/文本框/WordArt（list/insert/resize/wrap/pos/alt/z-order/delete） |
| `advanced(AdvancedRequest) -> AdvancedResult` | 脚注尾注/交叉引用/结构快照与 diff |
| `workflow_prepare(..) -> WorkflowPrepareResult` | 工作流第一阶段（只读分块 + 锚点） |
| `workflow_validate(..) -> WorkflowApplyResult` | 只预检，不写文档 |
| `workflow_apply(..) -> WorkflowApplyResult` | 校验后批量写入（低置信降级为批注） |
| `history(HistoryRequest) -> HistoryResult` | 历史/撤销/检查点 |
| `vba_status(document_id) -> VbaStatus` / `vba_run(VbaRunRequest) -> VbaResult` | VBA 只读探测 / 运行既有宏 |
| `vba_execute(VbaExecuteRequest) -> VbaExecuteResult` / `vba_cleanup(document_id) -> Vec<String>` | 临时注入（需授权+确认）/ 清理本库临时模块 |
| `has_member(member) -> bool` | 只读探测 Application 成员 |
| `shutdown()` | 停止 worker |

## 2. 关键 DTO

### Target
```json
{"kind":"document","document_id":"doc_1"}
{"kind":"selection","document_id":"doc_1"}
{"kind":"range","document_id":"doc_1","story":"main","start":120,"end":168,"expected_text":"旧标题","text_hash":"sha256:..","expected_doc_version":17}
{"kind":"paragraph","document_id":"doc_1","paragraph_index":4}
{"kind":"find","document_id":"doc_1","query":"待查文本","occurrence":1}
{"kind":"bookmark","document_id":"doc_1","bookmark":"ai_anchor"}
{"kind":"table_cell","document_id":"doc_1","table_id":"tbl_3","row":2,"column":3}
{"kind":"header_footer","document_id":"doc_1","section":1,"hf_kind":"header","variant":"primary"}
```

### 格式补丁（部分更新：`null`/缺省 = 不改）
```json
{"font":{"name":"等线","east_asia_name":"等线","size_pt":12,"bold":true,"color":"#222222","underline":"single","highlight":"yellow"},
 "paragraph":{"alignment":"justify","line_spacing":{"kind":"multiple","value":1.25},"space_after_pt":6,"first_line_indent_pt":24},
 "style":"heading_2","list":"bullet","reset":{"font":false,"paragraph":false}}
```

### 统一写入结果
```json
{"document_id":"doc_1","operation_id":"op_7","doc_version":18,"applied":3,"failed":0,
 "side_effect_status":"applied","actions":[{"index":0,"op":"replace_text","ok":true,"summary":"替换为 3 字","reversible":true}]}
```

### 表格 V1 格式（`TableRequest.format`，配合 `set_widths/set_heights/set_borders/set_allow_row_break`）
```json
{"action":"set_widths","table_id":"tbl_1","format":{"widths_pt":[80,60,120]}}
{"action":"set_heights","table_id":"tbl_1","format":{"heights":[{"row":1,"height_pt":24,"rule":"at_least"}]}}
{"action":"set_borders","table_id":"tbl_1","format":{"borders":[{"edge":"all","style":"single","width_pt":0.75,"color":"#4F81BD"}]}}
{"action":"set_allow_row_break","table_id":"tbl_1","format":{"allow_row_break":false}}
```

### 自定义多级编号（`FormatRequest.list` = `custom` + `list_template`）
```json
{"document_id":"doc_1","target":{"kind":"paragraph","document_id":"doc_1","paragraph_index":2},
 "list":"custom",
 "list_template":{"levels":[
   {"level":1,"format":"decimal","text":"%1.","number_position_pt":0,"text_position_pt":18,"tab_position_pt":18,"trailing":"tab"},
   {"level":2,"format":"lower_alpha","text":"%1.%2.","number_position_pt":18,"text_position_pt":36}],
  "continue_previous":false}}
```

### 修订筛选批量（`word_review.action = accept_filtered|reject_filtered`，需 `confirm:"yes"`）
```json
{"document_id":"doc_1","action":"accept_filtered","confirm":"yes",
 "filter":{"authors":["张三"],"kinds":["insert","delete"],"max":10}}
```

### 媒体（`word_media`）
```json
{"action":"insert_picture","document_id":"doc_1","target":{"kind":"document","document_id":"doc_1"},"path":"C:\\tmp\\a.png","inline":true,"width_pt":120}
{"action":"insert_shape","document_id":"doc_1","target":{"kind":"document","document_id":"doc_1"},"shape":"rounded_rectangle","text":"AI 说明","left_pt":60,"top_pt":60,"width_pt":180,"height_pt":60,"alt_text":"说明框"}
{"action":"insert_wordart","document_id":"doc_1","text":"艺术字","font_size_pt":30}
{"action":"list","document_id":"doc_1"}
{"action":"resize","document_id":"doc_1","media_id":"shape_1","width_pt":200}
```
返回 `items[]`：`{media_id,kind,name,left_pt,top_pt,width_pt,height_pt,alt_text,quote}`（`kind ∈ inline_picture|picture|shape|textbox|wordart`）。

### 高级（`word_advanced`）
```json
{"action":"add_footnote","document_id":"doc_1","target":{"kind":"document","document_id":"doc_1"},"text":"脚注内容"}
{"action":"insert_cross_reference","document_id":"doc_1","target":{"kind":"document","document_id":"doc_1"},
 "xref_type":"bookmark","xref_kind":"content_text","reference_item":"RsWordProbe","insert_as_hyperlink":true}
{"action":"structure_snapshot","document_id":"doc_1"}
{"action":"structure_diff","document_id":"doc_1","snapshot":{ /* 上一次 structure_snapshot 的返回 */ }}
```

### 工作流（`word_workflow`）
```json
// prepare：只读，返回 1 个 ≤max_chars_per_chunk 的 chunk + anchors + next_cursor
{"action":"prepare","document_id":"doc_1","task":"proofread","target":{"kind":"document","document_id":"doc_1"},
 "max_chars_per_chunk":4000,"overlap_chars":200,"include_outline":true}

// validate：只预检，不写文档（side_effect_status=not_applied）
{"action":"validate","document_id":"doc_1","task":"proofread","expected_doc_version":17,"changes":[ ... ]}

// apply：校验后批量写入；confidence < comment_below_confidence 自动降级为批注
{"action":"apply","document_id":"doc_1","task":"proofread","expected_doc_version":17,
 "apply_mode":"track_changes","comment_below_confidence":0.75,"max_changes":100,
 "changes":[{"op":"replace","target":{"kind":"find","document_id":"doc_1","query":"成本较高","occurrence":1},
             "text":"成本相对更高","reason":"口语化","confidence":0.9}]}
```
`apply_mode=direct`（直接改原文、不产生修订/批注）需 `confirm:"yes"`。

### VBA 临时注入（默认关闭）
```json
{"action":"execute_body","document_id":"doc_1","entry":"RsWordEntry",
 "code":"ThisDocument.Content.InsertAfter \"|VBA|\"\n",
 "allow_vba_injection":true,"confirm":"yes","cleanup":true}
{"action":"execute_module","document_id":"doc_1","entry":"RsWordEntry","code":"Public Function ...","allow_vba_injection":true,"confirm":"yes"}
{"action":"cleanup_stale_temp_modules","document_id":"doc_1"}
```
返回 `status ∈ ok|runtime_error`，含 `module_name/cleaned_up/cleanup_pending/error_number/error_description`；
注入失败码：`vba_runtime_unavailable` / `vba_project_access_denied` / `vba_project_locked` / `vba_compile_error` / `vba_cleanup_failed`。

### 错误回执
```json
{"code":"stale_target","category":"conflict","message":"目标文本在读取后已变化",
 "retryable":false,"side_effect_status":"not_applied","next_action":"重新读取最新锚点再修改"}
```

## 3. AI 工具（`tools::tool_definitions()`）

`word_session` · `word_read` · `word_edit` · `word_find_replace` · `word_format` · `word_table` · `word_layout` · `word_review` · `word_media` · `word_advanced` · `word_workflow` · `word_history` · `word_vba`。Schema 采用 `additionalProperties:false` + 显式必填字段；后端按 `target.kind`/`action` 二次语义校验。分派入口：

```rust
let reply = rs_word::tools::dispatch(&client, "word_read", serde_json::json!({ ... }));
// reply = {"ok":true,"data":{...}} 或 {"ok":false,"error":{...}}
```

### 示例
```json
// word_read
{"document_id":"doc_1","target":{"kind":"document","document_id":"doc_1"},"view":"outline","max_items":30}

// word_edit（校验全部目标后一次执行）
{"document_id":"doc_1","validate_all_first":true,"track_changes":true,"idempotency_key":"proof_1",
 "actions":[{"op":"replace_text","target":{"kind":"range","document_id":"doc_1","start":120,"end":126,"expected_text":"旧标题"},"text":"新标题"}]}

// word_find_replace
{"action":"preview_replace","document_id":"doc_1","find":"旧称","replace":"新称","limit":200}
{"action":"apply_replace","document_id":"doc_1","preview_id":"pv_ab12..","find":"旧称","replace":"新称","expected_doc_version":17,"max_replacements":200}

// word_review
{"document_id":"doc_1","action":"set_tracking","enabled":true}
{"document_id":"doc_1","action":"add_comment","target":{"kind":"range","document_id":"doc_1","start":500,"end":508},"text":"请补充依据"}
```

## 4. 最短 AI System Prompt 调用规则

```text
使用 Word 工具的规则：
1. 先 word_session.status 确认引擎与 document_id；不要用当前活动窗口猜目标文档。
2. 先 word_read 取必要范围（默认 summary/outline），不要默认读全文。
3. 修改必须带 target；优先使用 read 返回的 start/end + expected_text/text_hash。报告 stale_target 时重新读取，不要猜位置。
4. 多个修改合并到一次 word_edit / word_format / word_table 调用。
5. 全文替换先 preview_replace 检查数量与样例，再 apply_replace（带上 expected_doc_version 与 max_replacements）。
6. 校对/润色/翻译用 word_workflow：prepare 拿 chunk+锚点 → apply（默认 track_changes；低置信自动转批注）；summarize 默认只回聊天。
7. 图片/形状/文本框/艺术字用 word_media（几何单位 pt）；脚注/交叉引用/结构 diff 用 word_advanced。
8. 临时 VBA 注入默认关闭：只有宿主显式授权（allow_vba_injection=true + confirm=yes）才可执行，且先 checkpoint。
9. 遇到 result_unknown 不要直接重发写操作；先 word_read/status 验证是否已执行。
10. 接受全部修订、执行任意 VBA、覆盖/删除文件、关闭含未保存修改的文档，必须经宿主授权。
11. 输出只给任务摘要与少量例子，不要返回整份文档。
```
