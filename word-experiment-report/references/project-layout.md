# Reusable project layout

```text
report-project/
├─ input/                 # guide, assignment, raw question/data files
├─ references/            # copied templates, prior reports, style references
├─ code/                  # complete rerunnable SQL/scripts/backups
├─ screenshots/           # prompt, output, error, corrected-result images
├─ working/               # editable report and intermediate source
├─ output/                # final verified .docx and .pdf
└─ qa/                    # rendered PNGs and internal checks
```

The reusable math-model style template is stored in `workspace-template/references/` so it travels with a new report workspace. User-provided templates and project-specific guides belong in the new project's `references/` directory and should remain separate from the skill instructions.

A submission package should normally include the final PDF, editable DOCX, `code/`, `references/` or `input/` as needed for reproducibility, and the screenshots used to support the report. Keep `qa/` only when internal review records are requested.

