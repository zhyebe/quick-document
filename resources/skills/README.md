# Embedded Office Skills

This directory is packaged with Quick Document and gives the desktop app a local skill library for higher-quality Office work.

- `documents/`: Word / DOCX creation, edit, review, redline, comments, OOXML patching, rendering, and visual QA guidance.
- `spreadsheets/`: Excel workbook planning, formulas, tables, charts, formatting, verification, and export guidance.
- `presentations/`: PowerPoint / Google Slides narrative design and editing. `.ppt` and `.pptx` both route here.

The app uses these files in two ways:

1. The AI planner receives a compact routing brief derived from these skills so generated document plans follow the right artifact discipline.
2. Future local executors can call the bundled scripts directly for Word rendering, comments, tracked changes, accessibility audits, spreadsheet verification, and presentation QA.

For packaging, `package.json` includes this folder as `extraResources`, so DMG and Windows EXE installers carry the skill library with the app.
