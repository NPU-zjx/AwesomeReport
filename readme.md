# Word Experiment Report Skill

这是一个用于创建、编辑和整理 Word 实验报告的可复用工作包。

## 用户主要操作什么

新建报告时，用户主要使用并复制这个目录：

```text
workspace-template/
```

可以直接复制整个 `workspace-template/`，改名为本次报告项目目录，然后在其中放入材料和编辑报告。

```text
报告项目/
├─ input/                 # 放本次实验指南、任务书、题目和原始数据
├─ references/            # 放老师提供的实际 Word 模板和其他格式参考
├─ code/                  # 放 SQL、Python、脚本、备份和最终验证代码
├─ screenshots/           # 放提示词图、模型回答图、执行结果图和报错图
├─ working/               # 放正在编辑的 DOCX 和中间文件
├─ output/                # 放最终 DOCX 和 PDF
└─ qa/                    # 放渲染后的页面 PNG 和检查记录
```

如果没有老师提供的实际 Word 模板，可以参考：

`workspace-template/references/数学建模国赛标准论文Word模板.doc`

这个模板只用于参考字体、字号、行距、缩进、对齐、页边距和页面布局。

## 用户需要修改的内容

- 把本次实验的指导书和题目放入 `input/`。
- 把本次实验的模板或格式要求放入 `references/`。
- 把完整 SQL、程序、备份和验证代码放入 `code/`。
- 把每一问对应的结果图放入 `screenshots/`。
- 在 `working/` 中编辑报告。
- 将通过检查的 DOCX 和 PDF 放入 `output/`。

有实际 Word 模板时，优先使用实际模板；没有实际模板时，再参考数模模板。

## 用户通常不要修改的内容

以下内容属于 skill 本身的规则和工具，普通写报告时不需要修改：

- `SKILL.md`：AI 使用的核心规则
- `giude_for_human.md`：本工作包的详细使用说明
- `agents/`：Codex 界面显示信息
- `references/`：通用的编号、证据、排版检查和目录规范
- `scripts/`：创建报告工作目录的脚本

只有在要改造这个通用 skill 的行为时，才修改这些文件。

## 创建新项目

可以直接复制 `workspace-template/`，也可以运行：

```powershell
.\scripts\new_report_workspace.ps1 -ProjectRoot "D:\path\to\your-report" -CopyMathModelReference
```

如果有自己的实际模板，可以使用：

```powershell
.\scripts\new_report_workspace.ps1 -ProjectRoot "D:\path\to\your-report" -TemplatePath "D:\path\to\actual-template.docx"
```

## 最终提交

通常只提交具体报告项目中的：

- `output/` 中的最终 DOCX 和 PDF
- `code/` 中的完整代码
- `screenshots/` 中报告使用的结果图
- 为复现实验所需的 `input/` 和 `references/` 材料

不需要把整个 `word-experiment-report` skill 目录作为实验报告提交物。
