# docxCleaner 运行状态记录

## 1. 文档定位

本文件用于记录 `docxCleaner` 项目的动态执行状态，包括：

- 历史已完成内容；
- 当前唯一 active 任务；
- 当前 active 任务的目标、范围、验收标准和阻塞项；
- 后续 next/planned 任务列表。

本文件可以频繁更新。每当 active 任务完成后，应执行以下更新规则：

1. 将当前 active 任务标记为已完成；
2. 保留历史记录，不覆盖已完成内容；
3. 将下一个任务标记为唯一 active；
4. 把新的 active 任务写清楚目标、范围、验收标准和缺失上下文；
5. 如果无法细化，必须明确标记缺少什么上下文。

重要规则：

- `runtime.md` 不写完整 Codex prompt；
- 完整 Codex prompt 只在聊天中输出，方便复制；
- `runtime.md` 只记录任务状态、范围、验收标准和必要上下文。

---

## 2. 当前项目状态摘要

当前仓库：

- GitHub 仓库：`https://github.com/smter6626/docxCleaner`
- 仓库名：`smter6626/docxCleaner`
- 默认分支：`main`
- 当前用途：维护 DOCX 元数据查看与整理工具

当前本地目录：

- `/Users/smter-mac/Documents/personalAPPS/docxCleaner`

当前已知本地文件：

- `DocxWhereisAI.py`：当前主线，只处理这个文件；
- `PPTwhereisAI.py`：暂时保留，但当前不处理；
- `stage6.0.docx`：本地测试文件，不提交到 GitHub。

当前文档策略：

- `static.md`：稳定目标与长期约束；
- `runtime.md`：动态进度、active 任务与后续步骤；
- 暂不写 `README.md`。

---

## 3. 已完成内容

### 2026-06-20：初始化双文档工作流

状态：已完成

完成内容：

1. 明确项目采用双文档工作流：
   - `static.md` 记录长期目标、范围和稳定约束；
   - `runtime.md` 记录历史进展、当前 active 任务和未来步骤。
2. 明确 `static.md` 只有在用户调整目标、范围、平台、技术栈或长期约束时才更新。
3. 明确 `runtime.md` 在每个 active 任务完成后更新，且历史记录不覆盖。
4. 明确当前暂不写 README。
5. 明确项目长期目标：把 DOCX 源数据查看器升级为 GUI 友好的 DOCX 元数据查看与编辑工具。
6. 明确开发环境优先级：
   - 主要开发环境为 macOS；
   - 长期希望 Windows 也能运行；
   - 如果跨平台一次性实现困难，则先保证 macOS。

---

### 2026-06-20：同步本地脚本到 GitHub 仓库

状态：已完成

完成内容：

1. 在本地目录 `/Users/smter-mac/Documents/personalAPPS/docxCleaner` 初始化 Git 仓库。
2. 将本地分支设置为 `main`。
3. 远程仓库已包含 `static.md` 与 `runtime.md`。
4. 本地提交 `DocxWhereisAI.py` 与 `PPTwhereisAI.py`。
5. HTTPS push 失败后改用 SSH remote。
6. 成功 push 到远程 `main`。
7. 本地 `git status` 显示 working tree clean。

相关提交：

- `c331eb3 Add static project specification`
- `b7c9bd9 Add runtime project state`
- `e2f0e98 Add initial scripts and project docs`

---

### 2026-06-20：审查 `DocxWhereisAI.py` 初始结构

状态：已完成

审查结论：

- 初始版本是命令行脚本；
- 路径通过 `FILE_PATH` 硬编码；
- 分析报告全部通过 `print()` 输出；
- 末尾通过 `input()` 询问是否修改时间戳；
- 已有 `modify_timestamps()`，但只支持非常基础的时间戳修改；
- 下一步应先做 GUI 报告展示，而不是直接做完整编辑器。

当前可复用分析函数：

- `check_core_properties`
- `check_app_properties`
- `check_rsid`
- `check_track_changes`
- `check_comments`
- `check_custom_properties`
- `check_zip_timestamps`
- `check_styles_consistency`
- `check_lang_distribution`
- `check_word_count`
- `check_embedded_media_exif`

---

### 2026-06-20：第一轮 GUI 改造

状态：已完成

完成内容：

1. 新增 `analyze_docx(path: Path) -> str`。
2. 使用 `io.StringIO` 与 `contextlib.redirect_stdout` 捕获原有 `print()` 报告输出。
3. 使用 `tempfile.TemporaryDirectory()` 替换固定临时目录。
4. 新增 `DocxMetadataApp` tkinter GUI 类。
5. GUI 已包含：
   - `选择 DOCX 文件` 按钮；
   - `分析` 按钮；
   - `修改元数据` 入口；
   - `一键清理` 入口；
   - 只读 `ScrolledText` 报告框；
   - 底部状态栏。
6. `main()` 现在只启动 GUI。
7. 主流程不再依赖 `FILE_PATH`。
8. 主流程不再使用 `input()`。
9. `修改元数据` 与 `一键清理` 暂时只提示“后续阶段实现”，不会真实修改文件。
10. 已通过用户本地 GUI 测试：能启动窗口、选择 DOCX、分析并查看报告。

相关提交：

- `6bcf98a Add tkinter report viewer for DOCX metadata`

---

### 2026-06-20：基础元数据修改窗口与 `_cleaned.docx` 输出

状态：已完成

完成内容：

1. 新增基础元数据修改窗口。
2. 新增底层函数 `write_basic_metadata_cleaned_docx(...)`。
3. 新增 `read_basic_metadata_defaults(...)` 和 XML/ZIP 辅助函数。
4. 移除旧 `modify_timestamps()`，不再保留 `_modified` 输出规则。
5. 可修改：
   - `dc:creator`；
   - `cp:lastModifiedBy`；
   - `dcterms:created`；
   - `dcterms:modified`；
   - `ep:TotalTime`。
6. 输出同目录 `_cleaned.docx`。
7. `一键清理` 仍是“后续阶段实现”。
8. 文件选择器已移除“所有文件”，只保留 `.docx`。
9. 用户已手动验证：基础数据修改可正常实现，未选择文件时会正确提示。

相关提交：

- `2612ab6 Add basic DOCX metadata editor`

---

## 4. 当前 Active 任务

### Active 任务 B4：RSID 异常值规范化

状态：active

### 4.1 任务目标

在现有 GUI 基础上，实现 RSID 异常值规范化功能。

本阶段不做“用户指定 RSID 数量”。第一版保持 RSID 数量不变，只处理异常或需要重写的 RSID 值，降低 settings.xml 与正文引用不一致的风险。

### 4.2 设计原则

RSID 是 Word 文档中的编辑会话标记痕迹，和编辑过程相关，但不能等同于真实编辑次数。

GUI 文案应使用以下表达：

- “RSID 数量可反映文档内部编辑会话痕迹，但不能等同于真实编辑次数。”
- “本功能用于规范化异常 RSID 值，并同步更新文档中的 RSID 引用。”

避免使用以下表达：

- “RSID 数量 = 文档编辑次数”；
- “伪装编辑历史”；
- “绕过检测”；
- “生成自然编辑痕迹”。

### 4.3 本阶段功能范围

需要实现：

1. 新增 RSID 处理入口，建议放在主窗口中，按钮文案可为：`处理 RSID`。
2. 未选择文件时点击 `处理 RSID`，弹出友好提示。
3. 已选择 `.docx` 后，打开 RSID 处理窗口。
4. RSID 处理窗口显示：
   - settings.xml 中声明的 RSID 数量；
   - document.xml 中正文使用到的独立 RSID 数量；
   - 异常 RSID 值列表。
5. 异常 RSID 初始定义：
   - `00000000`；
   - `FFFFFFFF`；
   - 非 8 位十六进制；
   - 空值；
   - settings.xml 与 document.xml 中存在明显不一致时，应在窗口中提示，不要求本轮完全修复所有不一致。
6. 提供按钮：`重新生成异常 RSID 值`。
7. 重新生成规则：
   - 保持 RSID 数量不变；
   - 生成 8 位大写十六进制字符串；
   - 不生成 `00000000` 或 `FFFFFFFF`；
   - 尽量避免与当前文档中已有 RSID 冲突；
   - 使用稳定映射：同一个旧 RSID 在所有位置映射到同一个新 RSID。
8. 同步更新：
   - `word/settings.xml` 中 `<w:rsid w:val="..."/>`；
   - `word/document.xml` 中 `w:rsidR`、`w:rsidRPr`、`w:rsidRDefault`、`w:rsidP`。
9. 输出 `_cleaned.docx`，不覆盖原文件。
10. 生成后可以重新分析，确认异常 RSID 值已被替换。

### 4.4 明确不做

本阶段不做：

- 不允许用户指定 RSID 数量；
- 不删除 RSID；
- 不增加 RSID 数量；
- 不逐属性完全随机；
- 不处理 header/footer/comments/footnotes/endnotes 中的 RSID，除非当前代码已易于扩展且不会增加风险；
- 不处理样式异常；
- 不处理一键清理；
- 不处理 PPT；
- 不写 README。

### 4.5 技术实现要求

建议新增或拆分以下函数：

- `collect_rsid_info(source_path: Path) -> dict`
- `generate_rsid_mapping(old_values: Iterable[str]) -> dict[str, str]`
- `write_rsid_cleaned_docx(source_path: Path, mapping: dict[str, str]) -> Path`

ZIP 重写要求：

- 只修改 `word/settings.xml` 与 `word/document.xml`；
- 其他 ZIP 条目内容不变；
- 尽量保留未修改条目的 `ZipInfo` 元信息；
- 输出文件仍为 `_cleaned.docx`；
- 如果 `_cleaned.docx` 已存在，可以覆盖 cleaned 文件，但不能覆盖原始文件。

### 4.6 验收标准

本 active 任务完成标准：

1. `python3 -m py_compile DocxWhereisAI.py` 通过；
2. `python3 DocxWhereisAI.py` 能启动 GUI；
3. 未选择文件时点击 `处理 RSID` 有友好提示；
4. 选择 `.docx` 后可以打开 RSID 处理窗口；
5. 窗口能显示 settings.xml 与 document.xml 中的 RSID 概况；
6. 异常 RSID 能被识别；
7. 点击 `重新生成异常 RSID 值` 后生成 `_cleaned.docx`；
8. 再次分析 `_cleaned.docx` 后，异常 RSID 值被替换；
9. RSID 数量保持不变；
10. 原文件不被覆盖；
11. `PPTwhereisAI.py` 未被修改；
12. 未提交 `stage6.0.docx`、`.venv/`、`__pycache__/` 或生成文件。

### 4.7 缺失上下文与阻塞项

当前无阻塞项。

需要用户本地手动验证：

- 真实测试文件中是否存在异常 RSID；
- 如果测试文件没有异常 RSID，是否需要临时生成测试 DOCX 或通过测试副本制造 `00000000` / `FFFFFFFF` 用于验证；
- Word 是否能正常打开生成的 `_cleaned.docx`。

---

## 5. 后续步骤

### Next 任务 B5：ZIP 内部时间戳分布整理

状态：planned

目标：

- 实现“处理时间戳”入口；
- 提供 ZIP 内部时间戳整理模式；
- 优先实现模式 C：按原始唯一时间戳分组映射到新时间戳；
- 保持 ZIP 内部时间戳结构一致、可解释、可验证。

计划方案：

1. 新增按钮：`处理时间戳`。
2. 新增时间戳处理窗口。
3. 提供至少三种模式：
   - 模式 A：保留原始 ZIP 时间戳；
   - 模式 B：统一为用户指定修改时间；
   - 模式 C：分布整理。
4. 模式 C 规则：
   - 不逐文件完全随机；
   - 按原始唯一时间戳分组；
   - 每个原始唯一时间戳映射到一个新时间戳；
   - 同一组文件仍共享同一个新时间戳；
   - 偏移范围由用户设置，例如 `0–180 秒` 或 `0–10 分钟`；
   - 默认基准时间使用用户选择的修改时间；
   - 不改变 ZIP entry 数量；
   - 尽量保留原始 ZIP 条目的 `compress_type`、`external_attr`、`extra`、`comment` 等信息。
5. GUI 文案避免使用“自然”“不可检测”“伪装”等表达。
6. 输出 `_cleaned.docx`，不覆盖原文件。
7. 修改完成后可再次分析验证 ZIP 时间戳分布变化。

### Next 任务 B6：样式 ID 与引用关系整理

状态：planned，需要用户确认细节后才能进入 active

目标：

- 处理样式异常，但不默认删除样式；
- 第一版只考虑可疑 `styleId` 重命名与引用同步；
- 将样式功能作为高风险功能单独推进。

进入 active 前必须和用户确认：

1. 如何定义“异常样式 ID”；
2. 是否只重命名可疑 `styleId`；
3. 是否同步更新 `styles.xml` 内部引用：`w:basedOn`、`w:next`、`w:link`、`w:styleLink`、`w:numStyleLink`；
4. 是否同步更新正文、页眉、页脚、批注、脚注、尾注等 part 中的样式引用；
5. 是否允许删除未使用样式。默认不删除。

### Next 任务 B7：一键整理

状态：planned

目标：

- 基于基础元数据修改、RSID 规范化、时间戳分布整理等底层函数组合实现一键整理；
- 默认值和修改范围必须明确；
- 不把不可逆修改隐藏在模糊按钮后。

### Next 任务 B8：macOS 稳定性检查

状态：planned

目标：

- 检查 macOS GUI 启动方式；
- 检查打开文件夹逻辑；
- 检查错误弹窗；
- 检查不同 DOCX 文件的兼容性。

### Next 任务 B9：Windows 兼容性检查

状态：planned

目标：

- 检查 Windows 路径处理；
- 检查 tkinter 控件显示；
- 检查打开文件夹逻辑；
- 明确 Windows 上的最低运行方式。

### Next 任务 B10：README 与发布说明

状态：planned

目标：

- 在核心功能稳定后再写 README；
- 说明用途边界、安装方式、运行方式和功能限制；
- 不提前写宣传式 README。

---

## 6. 当前禁止事项

在当前 active 任务完成前，不做以下事情：

- 不处理 `PPTwhereisAI.py`；
- 不写 README；
- 不做 RSID 数量编辑；
- 不做 ZIP 时间戳分布整理；
- 不做样式异常处理；
- 不做一键清理；
- 不添加安装包配置；
- 不添加批量处理；
- 不提交本地测试 DOCX；
- 不把项目描述成反检测或伪装工具。
