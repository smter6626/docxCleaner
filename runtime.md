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

- `DocxWhereisAI.py`
- `PPTwhereisAI.py`

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
- 用户明确要求本次只写入 `static.md` 和 `runtime.md`，暂不写 README。

---

## 4. 当前 Active 任务

### Active 任务 A1：把本地现有脚本同步到 GitHub 仓库，并建立可执行的第一轮开发输入

状态：active

#### 4.1 任务目标

把当前本地已有代码文件同步到 GitHub 仓库，使后续 Codex 或其他开发工具可以基于真实代码进行修改，而不是基于抽象需求生成新脚本。

本任务完成后，仓库应至少包含：

- `DocxWhereisAI.py`
- `PPTwhereisAI.py`，如果决定暂时保留；
- `static.md`
- `runtime.md`

本任务不要求修改代码功能。

#### 4.2 为什么这是当前 active

当前 GitHub 仓库最初为空，但本地目录已经存在代码文件。如果不先同步现有代码，后续让 Codex 升级 GUI 时会出现两个问题：

1. Codex 只能根据需求重新生成代码，容易丢失原有分析函数；
2. `runtime.md` 无法基于真实代码状态制定可靠的阶段任务。

因此，当前必须先把真实代码放进仓库。

#### 4.3 具体执行步骤

在 macOS 终端中执行以下步骤。

##### 步骤 1：进入本地项目目录

```bash
cd ~/Documents/personalAPPS/docxCleaner
pwd
ls
```

预期结果：

```text
/Users/smter-mac/Documents/personalAPPS/docxCleaner
DocxWhereisAI.py    PPTwhereisAI.py
```

如果 `DocxWhereisAI.py` 不存在，停止执行，并标记为阻塞：缺少主程序文件。

##### 步骤 2：初始化 Git 仓库，如果尚未初始化

```bash
git status
```

如果出现：

```text
fatal: not a git repository
```

则执行：

```bash
git init
git branch -M main
```

如果已经是 Git 仓库，则不要重复初始化。

##### 步骤 3：绑定远程仓库

先查看当前 remote：

```bash
git remote -v
```

如果没有 `origin`，执行：

```bash
git remote add origin https://github.com/smter6626/docxCleaner.git
```

如果已有错误的 `origin`，执行：

```bash
git remote set-url origin https://github.com/smter6626/docxCleaner.git
```

再次确认：

```bash
git remote -v
```

预期包含：

```text
origin  https://github.com/smter6626/docxCleaner.git (fetch)
origin  https://github.com/smter6626/docxCleaner.git (push)
```

##### 步骤 4：拉取远程已有文档

因为远程仓库已经有 `static.md` 和 `runtime.md`，先执行：

```bash
git pull origin main --allow-unrelated-histories
```

如果提示合并信息，按默认保存即可。

如果出现冲突，停止执行，并先查看：

```bash
git status
```

冲突解决前不要继续提交。

##### 步骤 5：检查要提交的文件

```bash
git status
```

预期至少看到：

```text
Untracked files:
  DocxWhereisAI.py
  PPTwhereisAI.py
```

如果 `static.md` 或 `runtime.md` 出现为 modified，需要确认是否是本地误改；一般不要覆盖远程刚创建的版本。

##### 步骤 6：添加文件

```bash
git add DocxWhereisAI.py PPTwhereisAI.py static.md runtime.md
```

如果暂时不想把 `PPTwhereisAI.py` 放入仓库，可以改成：

```bash
git add DocxWhereisAI.py static.md runtime.md
```

当前建议：先加入 `PPTwhereisAI.py`，但在 `static.md` 中明确它不是当前主线。

##### 步骤 7：提交

```bash
git commit -m "Add initial docx cleaner scripts and project docs"
```

如果提示没有内容可提交，执行：

```bash
git status
```

确认文件是否已经被提交过。

##### 步骤 8：推送

```bash
git push -u origin main
```

如果需要 GitHub 登录认证，按终端提示完成认证。

##### 步骤 9：验证远程仓库

执行：

```bash
git status
git log --oneline -5
```

预期：

- 工作区干净；
- 最近提交中包含 `Add initial docx cleaner scripts and project docs`；
- GitHub 网页仓库中能看到 `DocxWhereisAI.py`、`PPTwhereisAI.py`、`static.md`、`runtime.md`。

#### 4.4 验收标准

本 active 任务完成的标准：

1. GitHub 仓库中存在 `DocxWhereisAI.py`；
2. GitHub 仓库中存在 `static.md`；
3. GitHub 仓库中存在 `runtime.md`；
4. 本地 `git status` 显示工作区干净；
5. 后续开发者可以直接基于仓库代码修改，而不是重新生成空白脚本。

#### 4.5 当前缺失上下文

当前缺失但不阻塞本任务的信息：

- `DocxWhereisAI.py` 的实际代码内容尚未审查；
- 当前脚本中的具体函数名、输出格式和修改逻辑尚未确认；
- 是否长期保留 `PPTwhereisAI.py` 尚未决定。

这些信息会在下一阶段代码审查时处理。

---

## 5. 后续步骤

### Next 任务 B1：审查 `DocxWhereisAI.py` 当前结构

状态：next

目标：

- 阅读主程序代码；
- 找出现有分析函数；
- 找出当前命令行入口；
- 找出时间戳修改函数；
- 判断哪些逻辑可直接复用，哪些需要重构。

预期输出：

- 当前函数清单；
- 分析函数与修改函数分组；
- GUI 改造风险点；
- 第一轮 Codex 修改指令。

### Next 任务 B2：第一轮 GUI 改造

状态：planned

目标：

- 只实现 GUI 基础框架；
- 只实现文件选择；
- 只实现 GUI 报告展示；
- 暂不实现修改功能。

验收标准：

- GUI 能启动；
- 能选择 `.docx`；
- 能在 ScrolledText 中显示原有完整分析报告；
- 原有命令行 `print` 输出逻辑已改为返回字符串或被统一捕获为字符串。

### Next 任务 B3：基础元数据修改

状态：planned

目标：

- 实现创建者、最后修改人、创建时间、修改时间、编辑时长的 GUI 修改；
- 生成 `_cleaned.docx`；
- 不覆盖原文件；
- 修改后能再次分析验证。

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

---

## 6. 当前禁止事项

在当前 active 任务完成前，不做以下事情：

- 不写 README；
- 不重写 `DocxWhereisAI.py`；
- 不添加安装包配置；
- 不添加批量处理；
- 不扩展到 PPTX 主线；
- 不把项目描述成反检测或伪装工具。
