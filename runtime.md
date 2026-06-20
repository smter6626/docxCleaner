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

审查结论：

- B2 目标已满足，可以进入下一阶段；
- 当前只发现一个非阻塞细节：文件选择器仍包含“所有文件”选项，严格来说不是“仅显示 .docx”，可在下一轮顺手修正。

---

## 4. 当前 Active 任务

### Active 任务 B3：基础元数据修改窗口与 `_cleaned.docx` 输出

状态：active

### 4.1 任务目标

在现有 GUI 基础上，实现第一批可编辑元数据字段，并生成新的 `_cleaned.docx` 文件。

本阶段只处理基础单值字段，不处理 RSID，不做完整一键清理。

目标字段：

1. `docProps/core.xml` 中的 `dc:creator`；
2. `docProps/core.xml` 中的 `cp:lastModifiedBy`；
3. `docProps/core.xml` 中的 `dcterms:created`；
4. `docProps/core.xml` 中的 `dcterms:modified`；
5. `docProps/app.xml` 中的 `ep:TotalTime`。

输出规则：

- 输出文件为原文件 stem + `_cleaned.docx`；
- 输出文件保存到原文件相同目录；
- 不覆盖原文件；
- 修改完成后应能被本工具再次分析；
- 生成文件应能被 Microsoft Word 正常打开。

### 4.2 本阶段允许修改

允许修改：

- `DocxWhereisAI.py`

不允许修改：

- `PPTwhereisAI.py`
- `static.md`
- `runtime.md`，除非用户明确要求由 Codex 更新状态
- `README.md`
- 本地测试文件 `stage6.0.docx`
- `.venv/`
- `__pycache__/`
- 生成文件 `*_cleaned.docx` 或 `*_modified.docx`

### 4.3 本阶段功能范围

需要实现：

1. 点击主窗口 `修改元数据` 后打开 `Toplevel` 修改窗口；
2. 修改窗口提供以下控件：
   - 创建者 Entry；
   - 最后修改人 Entry；
   - 编辑时长 Spinbox，范围 0 到 9999；
   - 创建时间 Spinbox 组合：年、月、日、时、分、秒；
   - 修改时间 Spinbox 组合：年、月、日、时、分、秒；
   - 是否统一 ZIP 内部时间戳 Checkbutton，默认勾选；
   - 应用修改按钮；
   - 取消按钮。
3. 时间字段默认值使用当前 UTC 时间；
4. 保存时间格式为 `YYYY-MM-DDTHH:MM:SSZ`；
5. 点击应用后生成 `_cleaned.docx`；
6. 修改完成后弹窗显示输出路径；
7. 修改完成后不要自动覆盖 GUI 当前选中文件，除非用户后续明确要求。

### 4.4 实现约束

实现时应注意：

1. 保留现有分析功能；
2. 不破坏 `analyze_docx()`；
3. 不重新引入命令行 `input()`；
4. 不重新引入硬编码 `FILE_PATH` 主流程；
5. ZIP 重写时尽量保留未修改条目的原始 `ZipInfo` 元信息；
6. 只修改 `docProps/core.xml` 和 `docProps/app.xml`；
7. 如果某个 XML 文件缺失，应弹出友好错误或在 messagebox 中提示，不崩溃；
8. 本阶段不处理 RSID；
9. 本阶段不实现一键清理；
10. 文件选择器中“所有文件”选项可顺手移除，只保留 `.docx`。

### 4.5 验收标准

本 active 任务完成标准：

1. `python3 -m py_compile DocxWhereisAI.py` 通过；
2. `python3 DocxWhereisAI.py` 能启动 GUI；
3. 未选择文件时点击 `修改元数据` 有友好提示；
4. 选择 `.docx` 后点击 `修改元数据` 能打开修改窗口；
5. 修改窗口可填写创建者、最后修改人、编辑时长、创建时间和修改时间；
6. 点击应用后生成 `_cleaned.docx`；
7. 原文件不被覆盖；
8. 生成文件能再次被本工具分析；
9. 再次分析能看到创建者、最后修改人、编辑时长、创建时间、修改时间发生变化；
10. `PPTwhereisAI.py` 未被修改；
11. 未提交 `stage6.0.docx`、`.venv/`、`__pycache__/` 或生成文件。

### 4.6 缺失上下文与阻塞项

当前无阻塞项。

仍需用户本地手动验证：

- Word 是否能正常打开生成的 `_cleaned.docx`；
- GUI 修改窗口布局是否可接受；
- 修改后的字段是否符合用户预期。

---

## 5. 后续步骤

### Next 任务 B4：RSID 与 ZIP 时间戳整理

状态：planned

目标：

- 实现 RSID 列表编辑；
- 实现正文 RSID 属性替换；
- 实现 ZIP 内部时间戳统一；
- 将 ZIP 时间戳整理逻辑纳入更统一的清理流程。

### Next 任务 B5：一键整理

状态：planned

目标：

- 基于 B3 和 B4 的底层修改函数实现一键整理；
- 使用默认隐私整理值；
- 生成 `_cleaned.docx`；
- 避免重复代码。

### Next 任务 B6：macOS 稳定性检查

状态：planned

目标：

- 检查 macOS GUI 启动方式；
- 检查打开文件夹逻辑；
- 检查错误弹窗；
- 检查不同 DOCX 文件的兼容性。

### Next 任务 B7：Windows 兼容性检查

状态：planned

目标：

- 检查 Windows 路径处理；
- 检查 tkinter 控件显示；
- 检查打开文件夹逻辑；
- 明确 Windows 上的最低运行方式。

### Next 任务 B8：README 与发布说明

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
- 不做 RSID 修改；
- 不做一键清理；
- 不添加安装包配置；
- 不添加批量处理；
- 不提交本地测试 DOCX；
- 不把项目描述成反检测或伪装工具。
