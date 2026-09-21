---
name: word-experiment-report
description: Create, revise, and package reusable Word experiment reports from a guide, template, SQL/code, screenshots, and result evidence. Use when a report must be delivered as editable DOCX plus PDF with a complete reproducible project folder.
---

# Word Experiment Report

Use this skill for database, programming, or other practical-course reports that must remain editable and reproducible. The deliverable is a report project, not only a formatted document: preserve the source guide, template references, code, evidence images, working DOCX, QA renders, and final PDF in predictable locations.

## Start with the project workspace

Use `workspace-template/` as the default layout or run `scripts/new_report_workspace.ps1` to create a fresh copy. Keep these roles separate:

- `input/`: user-provided guide, assignment, raw data, or question list.
- `references/`: style templates, prior reports, and domain reference material.
- `code/`: executable SQL, scripts, backup files, and final verified code.
- `screenshots/`: question/prompt images, execution results, error and corrected-result images.
- `working/`: editable DOCX and intermediate source files.
- `output/`: final PDF and the final editable DOCX copied for delivery.
- `qa/`: rendered page PNGs, page-count checks, and other internal verification artifacts.

Read only the references needed for the current report. Start with `giude_for_human.md` for the input placement and optional-file rules. Then use `references/report-structure.md` for section hierarchy and numbering, `references/evidence-and-assets.md` for evidence completeness, `references/docx-qa.md` for rendering and visual review, and `references/project-layout.md` for packaging.

## Use source authority in this order

1. The experiment guide controls required questions, experiment scope, and submission requirements.
2. An actual user-supplied report/template DOCX is optional. When present, it controls the document's structure and style; when absent, use the math-model template only as a visual reference.
3. The math-model paper template in `workspace-template/references/` is optional and is used only for typography, font size, spacing, indentation, alignment, and general page balance when no actual report template is supplied. Do not copy its subject matter, headings, or claims into the report.
4. The user's verified SQL, execution logs, screenshots, and stated results control factual content.

If a guide and an attached example conflict, follow the guide and record the actual experiment scope. Do not invent results to fill a missing screenshot or code file.

## Build each report section from evidence

For every question or subquestion, include the question/prompt, the final SQL or code, a concise answer and execution judgment, and the corresponding result image. For model-assisted SQL work, preserve the prompt/output image and the Workbench or database execution image when both exist; keep error and corrected-result images when they explain the revision. Captions must identify the question and evidence type.

Keep top-level numbering unique. A common structure is:

1. 头歌平台实验部分
2. 线下实验（实验一、实验二等作为本节下的子节）
3. 实验中出现的问题以及解决方案
4. 实验总结

Do not create a second top-level “三” when adding another offline experiment. Use nested headings such as `2.1 线下实验一` and `2.2 线下实验二`, or the equivalent numbering already established by the retained template. Apply the same hierarchy to the Word navigation pane and visible headings.

For each practical subsection, explain the task, environment, steps, code, result, and judgment in that order. Keep code readable and complete enough to rerun. Link or name the corresponding source file instead of silently shortening a long SQL statement; if a shortened excerpt is shown, state where the complete code is stored.

## Author and package the DOCX

Copy the retained template into `working/`; never edit the original asset. Fill the report with the user's actual name, student number, class, date, experiment scope, SQL/code, screenshots, and results. Keep figure captions, question labels, and code formatting consistent. Preserve Chinese fonts, page margins, heading hierarchy, table widths, and page numbering from the retained template unless the user requests a change.

Before delivery, render the DOCX with the document skill's `render_docx.py`, inspect every rendered page PNG, fix clipping/overflow/broken tables/blank pages/missing glyphs, and render again after each layout-sensitive edit. If PDF is requested, emit it from the verified DOCX and place the final pair in `output/`. Do not treat text extraction alone as visual verification.

## Finish with a reproducible package

The final package should contain the verified PDF and editable DOCX plus the code, source guide/reference material, execution records, and evidence screenshots. Use `references/project-layout.md` as the packaging checklist. Keep raw materials and working/QA intermediates distinct from final output so a later student can reuse the project without overwriting the source template.


