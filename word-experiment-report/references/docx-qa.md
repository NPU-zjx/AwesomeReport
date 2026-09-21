# DOCX and PDF quality gate

Use the bundled documents skill for Word authoring. For every meaningful create/edit batch:

1. Save the editable DOCX under `working/`.
2. Render it with the documents skill's `render_docx.py` into `qa/rendered/`; use `--emit_pdf` when the PDF is the requested output.
3. Inspect every page PNG at full resolution. Check heading numbering, Chinese glyphs, code wrapping, image size, captions, tables, page breaks, margins, headers/footers, and blank pages.
4. Correct the DOCX and render again whenever a defect appears.
5. Copy only the last verified DOCX and PDF to `output/`.

Text extraction, page count, or a successful LibreOffice exit code cannot replace visual inspection. QA renders remain internal unless the user requests them.
