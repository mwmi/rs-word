# rs-word 能力矩阵（WPS Writer / Microsoft Word）

> 状态定义：`supported` 实测可用；`partial` 部分可用/需变通；`unsupported` 确认不支持；`untested` 未在真机验证。
> 只读探测（`GetIDsOfNames`，不执行/不修改）实现的项，探测结果以「探测」标注。
> 请在目标机运行 `cargo run -p rs-word --example probe_engines`（只读 Attach，不关闭任何应用）复核。

## 真机实测记录（2026-10-20，Windows，本机）

- **WPS Writer 12.1.0.29161**（`D:\Kingsoft\WPS Office\12.1.0.29161\office6`，wps.exe 运行中）：
  - `KWPS.Application` 已注册，CLSID `{000209FF-0000-4B30-A977-D214852036FF}`；`GetActiveObject` 成功 → **Attach = supported**。
  - ⚠ **WPS 的 Application 没有 `Hwnd` 成员**（`0x80020006 未知名称`），且 `Application.Name = "Microsoft Word"`、
    `Version = "12.0"`（**伪装成 Word**）。引擎识别必须靠 `Application.Path`（含 `Kingsoft`/`WPS Office`）。
  - `word.Application`/`wps.Application` **未注册**（曾误列为候选，会产生 `0x800401F3` 掩盖真实原因，已移除）。
- **Microsoft Word**：本机已安装（`C:\Program Files\Microsoft Office\Root\Office16`）但**未运行** → 全部 `untested`。
  注意：WPS 会把自身注册到 `Word.Application` 名下，Attach `Word.Application` 可能连到 WPS；
  本库以 `Path` 核验归属，连到非目标引擎会返回 `connect.verify` 错误而非误用。

### V1/V2 新能力实测（2026-10-20，WPS Writer 12.1.0.29161，一次性测试文档后已丢弃）

取证方式：Attach 用户的 WPS → 新建一个测试文档 → 逐项写入并回读校验 → 丢弃并关闭该测试文档（未退出用户应用，未触碰用户文档）。
一次运行结果（全部 PASS，除下表中特别标注者）：

| 实测项 | 结果 | 回读证据 |
|---|---|---|
| 媒体：形状 / 文本框 / 行内图片 / WordArt / list | **supported** | `media.list` 得到 4 个对象 `["shape_1","textbox","wordart","inline_picture"]`，几何可回读 |
| 表格 set_widths | **supported** | 3 列 80/60/120 pt，回读 79.9/60.0/119.9（容差 ≤0.5 pt） |
| 表格 set_heights | **supported** | `row 1 height 24.0pt rule 1`，回读一致 |
| 表格 set_borders | **supported** | `edge=all` 展开 6 条边全部写为 `LineStyle=1` 且回读一致 |
| 表格 set_allow_row_break | **supported** | `AllowBreakAcrossPages=false`，回读 `Some(false)` |
| 自定义多级编号模板 | **supported** | `Application.ListGalleries(3).ListTemplates(7).ListLevels(i)` 配置成功并 `ApplyListTemplate` 返回 ok（注意：**WPS 的成员名是 `ListLevels`**，`Levels` 会 `0x80020006`） |
| 修订筛选批量 accept_filtered | **supported** | 按 `kinds=[insert,delete]` 处理 `applied=1 skipped=1` |
| 脚注 add_footnote / list_notes | **supported**（脚注） | 写入后回读 1 条脚注；**尾注未单独实测 → untested** |
| 交叉引用 insert_cross_reference | **supported** | 书签引用写入成功，插入点文本回读非空 |
| 交叉引用 list_cross_references | **partial** | 实测 `Content.Fields.Count = 0`：WPS 插入的 REF 域**无法从 `Content.Fields` 枚举**（插入可用、枚举不可用）；已在结果 `applied` 中回报 `fields=N refs=M` |
| 结构快照 / 结构 diff | **supported** | 快照 11 段；diff 正确报出 `paragraph_text_changed`、`note_count_changed` |
| 工作流 prepare / apply | **supported** | prepare 返回 1 个 chunk(v=20)+outline；apply `applied=1 failed=0`，低置信(0.4)按策略**降级为批注并写入 warnings** |
| VBA execute_body（注入） | **supported** | `status=ok`，临时模块 `RsWordTmp_1_<ms>` 创建后成功清理（`cleaned=true`）；`execute_module` 同路径未单独实测 → untested |
| 用户应用保护 | — | 全程只操作新建测试文档；结束时丢弃并关闭该文档，**未退出用户的 WPS 进程** |

## 能力矩阵

| Capability | MS Word 版本 | WPS Writer 版本 | Status | Evidence / notes |
|---|---|---|---|---|
| 连接 Attach（ProgID + 身份核验） | 未运行 | 12.1.0.29161 | **WPS supported / Word untested** | WPS：KWPS.Application + Path 含 Kingsoft；Word 未运行 |
| 连接 New（CoCreateInstance） | 待测 | 12.1.0.29161 | partial | WPS New 返回实例（Name=Microsoft Word, Path=Kingsoft）；**未在本库验证是否独立于用户实例** |
| 引擎切换（脏文档策略） | 待测 | 待测 | untested | 旧 document_id 立即失效 |
| Document 列表 / ActiveDocument | 待测 | 12.1.0.29161 | **WPS supported** | 实测列出「文字文稿1」 |
| 打开/新建/保存/另存(SaveAs2→SaveAs) | 待测 | 12.1.0.29161 | **WPS partial** | 新建/关闭(丢弃)实测通过；另存未实测 |
| Range 读（Word Start/End） | 待测 | 12.1.0.29161 | **WPS partial** | Content.Text 读取成功；写入待测 |
| Content.Start/End/Text 统计 | 待测 | 12.1.0.29161 | **WPS supported** | 段落/表格/修订/批注计数成功 |
| Find/Replace（Range.Find，preview+限量） | 待测 | 12.1.0.29161 | **WPS supported** | 实测按 occurrence 定位并替换成功 |
| 字符格式（字体/字号/颜色/高亮/上下标…） | 待测 | 待测 | untested | 部分更新；混合格式读值 |
| 段落格式（对齐/行距/缩进/分页控制） | 待测 | 待测 | untested | 内部统一 points |
| 段落底纹/边框 | 待测 | 待测 | untested | 段落边框未做本体实现 |
| 样式（内置常量映射 / 本地化） | 待测 | 待测 | untested | 优先内置常量，避免本地化名 |
| 列表（bullet/number/outline 默认） | 待测 | 待测 | untested | — |
| 自定义多级编号模板 | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | 见上：`ListTemplate.ListLevels` + `ApplyListTemplate` |
| 表格创建（二维批量填充） | 待测 | 12.1.0.29161 | **WPS supported** | 实测 2×3 建表与填充 |
| 表格增删行列 / 合并 / 拆分 | 待测 | 待测 | untested | 倒序删除；合并前校验矩形 |
| 表格底纹 / 单元格垂直对齐 | 待测 | 待测 | untested | — |
| 表格列宽/行高/边框/跨页行（V1） | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | 四项均写入并回读一致（见上表） |
| 页面设置（按 Section） | 待测 | 待测 | untested | 纸张/方向/边距(mm→pt) |
| 页眉页脚（primary/first/even） | 待测 | 待测 | untested | 独立 story |
| 页码 / 字段更新（PAGE/NUMPAGES/DATE/…） | 待测 | 待测 | untested | 头部/尾部字段分别更新 |
| PDF 导出（ExportAsFixedFormat） | 待测 | 待测 | untested | 导出后校验非空、不自动打开 |
| Track Changes（生成 Revisions） | 待测 | 12.1.0.29161 | **WPS supported** | 实测开启后替换产生 2 条修订 |
| Revisions 列表 / 接受 / 拒绝 | 待测 | 12.1.0.29161 | **WPS supported** | revision_ref 由 index+hash 校验；筛选批量实测见上 |
| Comments 增删列 | 待测 | 12.1.0.29161 | **WPS supported（间接）** | workflow 低置信降级为批注时写入成功 |
| UndoRecord（原生自定义撤销） | 待测 | 待测 | untested | WPS 不假定支持；有探测 |
| Application.Undo / Redo | 待测 | 待测 | untested | — |
| 媒体：图片/形状/文本框/WordArt（word_media） | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | 见上表；WordArt 缺失时返回 `unsupported_capability`，不降级 |
| 脚注（add/list） | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | 写入并回读 1 条 |
| 尾注（add/list） | 待测 | 待测 | untested | 与脚注同一代码路径，未真机实测 |
| 交叉引用 插入 | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | 书签引用写入成功 |
| 交叉引用 枚举（word_advanced.list_cross_references） | 待测 | 12.1.0.29161 | **WPS partial** | `Content.Fields.Count=0`，插入的 REF 域枚举不到 |
| 文档结构快照 / diff | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | 快照+diff 实测报出 2 类结构变更 |
| AI 工作流 prepare/validate/apply | 待测 | 12.1.0.29161 | **WPS supported / Word untested** | prepare/apply 实测；本地编排，依赖既有读写 |
| VBA status（Run/VBProject 探测） | 待测 | 12.1.0.29161 | **WPS supported** | `run=true project=true injection_ready=true` |
| VBA run_existing_macro | 待测 | 待测 | untested | ≤30 参；Empty 不等于失败 |
| VBA 临时注入（execute_body/execute_module） | 待测 | 12.1.0.29161 | **WPS supported（execute_body）/ Word untested** | 默认关闭；需 allow+confirm；`RsWordTmp_*` 创建后成功清理 |
| VBA 清理（cleanup_stale_temp_modules） | 待测 | 12.1.0.29161 | **WPS supported** | 只清理登记归属的 `RsWordTmp_*` |

## 只读探测（连接后自动）

`status.capabilities` 逐项给出探测结果（仅用 `GetIDsOfNames`，不执行宏、不修改文档或安全设置）：

- 应用级：`connect`、`application.documents`、`application.list_galleries`、`vba.run`、`history.undo_record`、
  以及本地实现的 `workflow.prepare_apply`、`advanced.structure_diff`。
- 活动文档存在时：`document.range/content/find/tables/sections/page_setup/export_pdf/track_revisions/revisions/comments/vb_project`。
- V1/V2 新增（**成员存在也一律标 `untested`**，真机实测后才由上表升级）：
  `media.inline_shapes`、`media.shapes`、`media.wordart`、`advanced.footnotes`、`advanced.endnotes`、
  `advanced.cross_reference`、`table.format_v1`、`vba.injection`、`list.custom_template`。
