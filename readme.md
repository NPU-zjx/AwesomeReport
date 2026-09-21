# Word Experiment Report

这是一个用于创建和整理 Word 实验报告的可复用 skill 工作包。

## 目录

- `SKILL.md`：AI 创建和检查报告时使用的核心规则
- `guide_for_human.md`：给使用者看的文件放置和使用说明
- `workspace-template/`：新报告项目的目录模板
- `workspace-template/references/`：可选的数模版式参考
- `references/`：编号、证据、排版检查和项目结构说明
- `scripts/`：创建新报告工作目录的脚本
- `agents/`：Codex 界面显示信息

## 使用原则

有实际的 Word 报告模板时，优先使用实际模板；没有实际模板时，再参考数模模板的字体、字号、行距、缩进、对齐和页面布局。

实验指导书、题目、SQL 代码、截图和往年报告属于具体项目资料，应放在新建报告项目中，不直接写入通用 skill。

## 创建新项目

复制 `workspace-template/`，或运行：

```powershell
.\scripts\new_report_workspace.ps1 -ProjectRoot "D:\path\to\your-report" -CopyMathModelReference
```

最终的可编辑 DOCX 和 PDF 放在新项目的 `output/` 目录，代码和截图分别放在 `code/` 与 `screenshots/` 目录。
