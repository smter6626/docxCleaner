# docxCleaner 运行状态记录

## 1. 文档定位

本文件用于记录 `docxCleaner` 项目的动态执行状态，包括：

- 历史已完成内容；
- 当前唯一 active 任务；
- 当前任务的详细执行步骤；
- 缺失上下文或阻塞项；
- 后续计划。

本文件可以频繁更新。每当 active 任务完成后，应执行以下更新规则：

1. 将当前 active 任务移动或标记到“已完成”；
2. 保留历史记录，不覆盖已完成内容；
3. 将下一个任务标记为唯一 active；
4. 把新的 active 任务细化到可以直接照搬执行的程度；
5. 如果无法细化，必须明确标记缺少什么上下文。

---

## 2. 当前项目状态摘要

当前仓库：

- GitHub 仓库：`https://github.com/smter6626/docxCleaner`
- 仓库名：`smter6626/docxCleaner`
- 默认分支：`main`
- 当前用途：维护 DOCX 元数据查看与清理工具

当前本地目录：

- `/Users/smter-mac/Documents/personalAPPS/docxCleaner`

当前已知本地文件：

- `DocxWhereisAI.py`：当前主线，只处理这个文件；
- `PPTwhereisAI.py`：暂时保留，但当前不处理。

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

证据：

- 用户明确说明仓库地址为 `https://github.com/smter6626/docxCleaner`。
- 用户明确说明当前本地目录中有 `DocxWhereisAI.py` 与 `PPTwhereisAI.py`。
- 用户明确要求只写入 `static.md` 和 `runtime.md`，暂不写 README。

---

### 2026-06-20：同步本地脚本到 GitHub 仓库

状态：已完成

完成内容：

1. 在本地目录 `/Users/smter-mac/Documents/personalAPPS/docxCleaner` 初始化 Git 仓库。
2. 将本地分支设置为 `main`。
3. 先使用 HTTPS remote 拉取远程已有 `static.md` 与 `runtime.md`。
4. 本地提交 `DocxWhereisAI.py` 与 `PPTwhereisAI.py`。
5. HTTPS push 因 GitHub 不支持密码认证失败。
6. 将 remote 改为 SSH：`git@github.com:smter6626/docxCleaner.git`。
7. 使用 `ssh -T git@github.com` 验证 SSH 已能认证到 `smter6626`。
8. 成功 push 到远程 `main`。
9. 本地 `git status` 显示 working tree clean。

证据：

- 本地 `git log --oneline -5` 显示：
  - `e2f0e98 Add initial scripts and project docs`
  - `b7c9bd9 Add runtime project state`
  - `c331eb3 Add static project specification`
- 本地分支已 tracking `origin/main`。
- GitHub 仓库中已存在 `DocxWhereisAI.py`、`PPTwhereisAI.py`、`static.md`、`runtime.md`。

---

### 2026-06-20：审查 `DocxWhereisAI.py` 当前结构

状态：已完成

审查范围：

- 只审查 `DocxWhereisAI.py`；
- 暂不处理 `PPTwhereisAI.py`；
- 不修改代码功能。

当前脚本性质：

- 当前 `DocxWhereisAI.py` 是命令行脚本，不是 GUI 程序；
- 文件路径通过模块级常量 `FILE_PATH` 硬编码；
- 分析报告全部通过 `print()` 输出；
- 运行末尾通过 `input()` 询问是否修改时间戳；
- 已有一个非常基础的时间戳修改函数 `modify_timestamps()`；
- 当前输出文件后缀为 `_modified`，不符合未来目标 `_cleaned.docx`。

当前主要 import：

- `zipfile`
- `shutil`
- `re`
- `pathlib.Path`
- `datetime.datetime`
- `datetime.timezone`
- `collections.Counter`
- `xml.etree.ElementTree as ET`

当前命名空间常量：

- `NS_CP`
- `NS_DC`
- `NS_DCTERMS`
- `NS_EP`
- `NS_W`
- `NS_VT`
- `NS_CUSTOM`
- `NS`

当前分析函数清单：

1. `section(title)`
   - 打印分区标题；
   - 当前强依赖 `print()`。
2. `check_core_properties(z)`
   - 读取 `docProps/core.xml`；
   - 输出创建者、最后修改人、创建时间、修改时间、标题、主题、关键词、类别、版本；
   - 如果创建者与最后修改人不一致，会输出提示。
3. `check_app_properties(z)`
   - 读取 `docProps/app.xml`；
   - 输出应用程序、版本、公司、修订次数、总编辑时长、字数、字符数、段落数、页数、行数；
   - 会把 `TotalTime` 换算为小时。
4. `check_rsid(z)`
   - 读取 `word/settings.xml` 中 `<w:rsids>`；
   - 读取 `word/document.xml` 中 `rsidR`、`rsidRPr`、`rsidRDefault`、`rsidP`；
   - 输出 settings 中声明的 RSID 总数与正文中独立 RSID 数量。
5. `check_track_changes(z)`
   - 读取 `word/document.xml`；
   - 查找 `<w:ins>` 与 `<w:del>`；
   - 输出修订作者和修订时间。
6. `check_comments(z)`
   - 读取 `word/comments.xml`；
   - 输出批注作者、日期、缩写和内容预览。
7. `check_custom_properties(z)`
   - 读取 `docProps/custom.xml`；
   - 输出自定义属性名称、值、`fmtid` 和 `pid`。
8. `check_zip_timestamps(z)`
   - 遍历 ZIP 内部条目；
   - 输出关键 XML 文件的内部时间戳；
   - 统计唯一时间戳数量。
9. `check_styles_consistency(z)`
   - 读取 `word/styles.xml`；
   - 检测带数字后缀的可疑样式 ID。
10. `check_lang_distribution(z)`
    - 读取 `word/document.xml`；
    - 统计 `<w:lang>` 分布。
11. `check_word_count(z)`
    - 对比 `docProps/app.xml` 中记录的 Words/Characters 与 `word/document.xml` 实际文本统计。
12. `check_embedded_media_exif(z, tmp_dir)`
    - 可选导入 Pillow；
    - 检查 `word/media/` 中图片 EXIF；
    - Pillow 缺失时跳过。

当前修改函数：

- `modify_timestamps(original_path)`
  - 修改 `docProps/core.xml` 中创建时间与修改时间为当前 UTC 时间；
  - 将所有 ZIP 条目的内部时间戳设为当前时间；
  - 输出新文件名为原 stem + `_modified`；
  - 目前不修改创建者、最后修改人、编辑时长或 RSID；
  - 写入 ZIP 时使用新的 `ZipInfo`，没有完整保留原 ZIP 条目的压缩方式和外部属性。

当前入口：

- `main()` 使用硬编码 `FILE_PATH`；
- 创建固定临时目录 `/tmp/docx_forensics_tmp`；
- 依次调用所有检查函数；
- 清理临时目录；
- 最后通过 `input()` 询问是否修改时间戳；
- `if __name__ == "__main__": main()` 启动命令行流程。

主要风险点：

1. `print()` 分散在所有分析函数中，GUI 化前必须改为返回字符串，或用统一输出捕获器过渡。
2. `FILE_PATH` 硬编码，GUI 文件选择必须替代它。
3. `input()` 与 GUI 冲突，必须删除命令行交互。
4. `section()` 也直接 `print()`，需要重构为字符串生成函数。
5. 多个 XML 解析没有统一异常处理，损坏 DOCX 或 XML 缺失时可能崩溃。
6. 临时目录硬编码到 `/tmp/docx_forensics_tmp`，应改为 `tempfile.TemporaryDirectory()`。
7. 当前 `modify_timestamps()` 会重写 ZIP 条目，但没有完整保留原始 ZIP 元信息。
8. 当前修改输出后缀为 `_modified`，与项目目标 `_cleaned.docx` 不一致。
9. 当前脚本顶部描述带有“forensics/痕迹判断”倾向，后续 GUI 文案需要改为更中性的“元数据查看与整理”。

结论：

- 下一步不应直接做完整编辑器；
- 下一步应只做第一轮 GUI 改造：GUI 基础框架 + 文件选择 + 报告展示；
- 修改功能按钮可以先禁用或显示“后续阶段实现”；
- 当前所有分析逻辑应尽量保留，只调整输出方式；
- `modify_timestamps()` 暂时不要扩展，避免同时改 GUI 与 ZIP 写入逻辑导致错误定位困难。

---

## 4. 当前 Active 任务

### Active 任务 B2：第一轮 GUI 改造，只实现文件选择与报告展示

状态：active

#### 4.1 任务目标

将 `DocxWhereisAI.py` 从硬编码路径的命令行脚本，改造成最小可用的 tkinter GUI 程序。

本阶段只实现：

1. GUI 主窗口；
2. 文件选择；
3. 分析按钮；
4. 报告展示；
5. 状态栏；
6. 对缺失文件、损坏文件、XML 缺失的基础友好提示。

本阶段明确不实现：

- 不做元数据编辑窗口；
- 不做一键清理；
- 不做 RSID 修改；
- 不扩展 `modify_timestamps()`；
- 不处理 `PPTwhereisAI.py`；
- 不写 README。

#### 4.2 为什么这是当前 active

当前代码的核心问题不是“修改能力不足”，而是程序结构仍然是命令行脚本：

- 路径硬编码；
- 全部使用 `print()`；
- 末尾使用 `input()`；
- GUI 无法调用现有分析逻辑。

如果直接进入编辑器阶段，会同时改 GUI、XML 写入、ZIP 写入和交互流程，失败后难以定位。

因此必须先建立稳定的 GUI 报告层。

#### 4.3 给 Codex 的具体修改指令

把下面内容直接发给 Codex。要求 Codex 修改当前仓库中的 `DocxWhereisAI.py`。

```text
请只修改当前仓库中的 `DocxWhereisAI.py`，不要处理 `PPTwhereisAI.py`，不要写 README，不要创建新的主程序文件。

目标：完成第一轮 GUI 改造。只实现 tkinter GUI、文件选择和分析报告展示。暂时不要实现任何修改/清理功能。

当前代码状态：
- 这是命令行脚本；
- 文件路径通过 `FILE_PATH` 硬编码；
- 所有检查函数通过 `print()` 输出；
- `main()` 末尾通过 `input()` 询问是否修改时间戳；
- 已有 `modify_timestamps()`，但本阶段不要扩展它。

本阶段必须完成：

1. 添加 tkinter GUI。
   - 主窗口标题：`DOCX 元数据查看与清洗器`。
   - 使用标准库 `tkinter` 和 `tkinter.scrolledtext.ScrolledText`。
   - 主窗口包含：
     - `选择 DOCX 文件` 按钮；
     - `分析` 按钮；
     - `修改元数据` 按钮，但本阶段先禁用或点击后弹出“后续阶段实现”；
     - `一键清理` 按钮，但本阶段先禁用或点击后弹出“后续阶段实现”；
     - 只读 `ScrolledText` 报告框；
     - 底部状态栏 `Label`。
   - 初始状态栏文字：`请选择 DOCX 文件进行分析`。

2. 去除硬编码运行路径依赖。
   - 不再使用 `FILE_PATH` 作为主流程输入。
   - 用户通过文件选择对话框选择 `.docx` 文件。
   - 文件选择后状态栏显示当前文件路径。

3. 重构报告输出。
   - 不要让分析函数直接 `print()` 到终端。
   - 可以选择以下任一策略：
     A. 把 `section()` 和所有 `check_*` 函数改成返回字符串/字符串列表；
     B. 使用 `io.StringIO` + `contextlib.redirect_stdout` 作为过渡层捕获现有 `print()` 输出。
   - 本阶段优先保证改动小、能运行、报告完整。
   - 报告内容、分区标题、字段名和警告文本应尽量保持原样，不要求字符级完全一致。

4. 新增统一分析函数。
   - 添加类似 `analyze_docx(path: Path) -> str` 的函数。
   - 该函数打开 DOCX ZIP，依次执行现有检查：
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
   - 该函数返回完整报告字符串。

5. 临时目录处理。
   - EXIF 检查如需临时目录，使用 `tempfile.TemporaryDirectory()`。
   - 不再硬编码 `/tmp/docx_forensics_tmp`。
   - 临时目录必须自动清理。

6. 删除命令行交互。
   - `main()` 不再使用 `input()`。
   - `if __name__ == "__main__"` 中只启动 GUI。
   - 当前 `modify_timestamps()` 可以保留，但不在 GUI 主流程中调用。

7. 基础异常处理。
   - 未选择文件时点击“分析”，弹出 `messagebox.showwarning`。
   - 文件不存在时，报告区显示友好错误。
   - 文件不是合法 ZIP/DOCX 时，报告区显示友好错误。
   - 单个 XML 文件缺失时，不应导致整个程序崩溃，尽量沿用原来的“未找到 xxx.xml”提示。
   - 如果某个检查函数抛出异常，应在报告中显示该检查失败，但继续或安全结束，不让 GUI 崩溃。

8. GUI 状态管理。
   - 推荐使用 `DocxMetadataApp` 类管理：
     - root；
     - selected_path；
     - report_text；
     - status_label。
   - 避免新增大量全局变量。

9. 阶段标记。
   - 在代码中添加清晰注释：
     - `# ===== 阶段1完成：GUI 基础框架 =====`
     - `# ===== 阶段2完成：报告展示 =====`

本阶段验收方式：

1. 运行：
   python3 DocxWhereisAI.py
2. GUI 窗口能打开。
3. 未选择文件时点击“分析”，应弹出提示。
4. 选择 `.docx` 文件后，状态栏显示文件路径。
5. 点击“分析”后，报告显示在 GUI 的 ScrolledText 中。
6. 终端不再要求输入 `是否修改时间戳？(y/n)`。
7. `修改元数据` 与 `一键清理` 暂不执行真实修改。
8. Pillow 未安装时，EXIF 部分显示“未安装 Pillow，跳过 EXIF 检查。”，程序不崩溃。
```

#### 4.4 用户本地执行步骤

在 macOS 终端执行：

```bash
cd ~/Documents/personalAPPS/docxCleaner
git status
git pull
```

确认干净后，把 4.3 中的 Codex 指令发给 Codex。

Codex 修改完成后，执行：

```bash
python3 DocxWhereisAI.py
```

手动测试：

1. 窗口是否能打开；
2. 不选文件点“分析”是否提示；
3. 选择一个 `.docx` 文件后状态栏是否显示路径；
4. 点“分析”后报告是否显示在 GUI；
5. 终端是否不再出现 `是否修改时间戳？(y/n)`；
6. 关闭窗口后程序是否正常退出。

测试通过后执行：

```bash
git status
git diff -- DocxWhereisAI.py
```

如果 diff 只涉及 `DocxWhereisAI.py`，并且 GUI 测试通过，则提交：

```bash
git add DocxWhereisAI.py
git commit -m "Add tkinter report viewer for DOCX metadata"
git push
```

#### 4.5 本阶段验收标准

本 active 任务完成的标准：

1. `DocxWhereisAI.py` 能通过 `python3 DocxWhereisAI.py` 启动 GUI；
2. GUI 能选择 `.docx` 文件；
3. GUI 能显示完整分析报告；
4. 命令行不再要求用户输入是否修改时间戳；
5. `modify_timestamps()` 未被扩展为复杂编辑器；
6. `PPTwhereisAI.py` 未被修改；
7. 代码已提交并 push 到远程仓库。

#### 4.6 当前缺失上下文

当前缺失但不阻塞本任务的信息：

- 尚未提供一个固定测试用 `.docx` 文件；
- 暂未确认用户是否安装 Pillow；
- 暂未确认 Windows 上 tkinter 显示效果。

这些不影响当前 macOS 第一轮 GUI 改造。

---

## 5. 后续步骤

### Next 任务 B3：基础元数据修改

状态：planned

目标：

- 实现创建者、最后修改人、创建时间、修改时间、编辑时长的 GUI 修改；
- 生成 `_cleaned.docx`；
- 不覆盖原文件；
- 修改后能再次分析验证；
- 开始重构当前 `modify_timestamps()`，把它升级为通用修改函数。

### Next 任务 B4：RSID 与 ZIP 时间戳整理

状态：planned

目标：

- 实现 RSID 列表编辑；
- 实现正文 RSID 属性替换；
- 实现 ZIP 内部时间戳统一；
- 增加一键整理功能。

### Next 任务 B5：跨平台检查

状态：planned

目标：

- 检查 macOS 与 Windows 的路径处理；
- 检查打开文件夹逻辑；
- 检查 tkinter 控件显示；
- 明确 Windows 上的最低运行方式。

### Next 任务 B6：README 与发布说明

状态：planned

目标：

- 在核心功能稳定后再写 README；
- 说明用途边界、安装方式、运行方式和功能限制；
- 不提前写宣传式 README。

---

## 6. 当前禁止事项

在当前 active 任务完成前，不做以下事情：

- 不写 README；
- 不处理 `PPTwhereisAI.py`；
- 不做元数据编辑窗口；
- 不做一键清理；
- 不做 RSID 修改；
- 不添加安装包配置；
- 不添加批量处理；
- 不扩展到 PPTX 主线；
- 不把项目描述成反检测或伪装工具。
