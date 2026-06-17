---
name: md-resume-to-pages
description: Convert Markdown resumes into polished Word `.docx` and Apple Pages `.pages` resumes using a repeatable md -> docx -> Pages workflow. Use when the user asks to turn a `.md` resume into a `.docx` or `.pages` resume, match an existing Pages resume layout, preserve a resume template, export through Apple Pages, or iterate on resume table/print layout such as a 3-row by 5-column identity header, basic information tables, project tables, and A4 visual QA.
---

# Markdown Resume To Pages

## Overview

Use this skill to convert a structured Markdown resume into a Word document and then export it to Apple Pages. Prefer the bundled scripts as starting points, then patch them for the current resume content, template, photo, and layout requirements.

## Workflow

1. Locate the source Markdown resume, target output name, optional portrait image, and optional reference `.pages` file.
2. Copy `scripts/build_pages_style_resume.py` into the workspace and patch constants, parser mappings, content fields, and layout details for the current task.
3. Build the Word file with the bundled workspace Python runtime, not system Python, when available.
4. Export the `.docx` to `.pages` with `scripts/export_to_pages09.scpt`:

```bash
osascript /Users/mqs/.codex/skills/md-resume-to-pages/scripts/export_to_pages09.scpt input.docx output.pages
```

5. Extract and inspect the Pages thumbnail:

```bash
unzip -p output.pages QuickLook/Thumbnail.jpg > /private/tmp/resume-thumb.jpg
```

6. Iterate until the generated `.pages` file visually matches the requested layout and has no overlap, clipping, unexpected grid lines, or awkward A4 print balance.

## Layout Rules

Read `references/layout.md` before making non-trivial layout edits. In particular, use a real 3-row by 5-column table for the top identity block when matching the Pages resume style from this workflow.

## Script Notes

- `scripts/build_pages_style_resume.py` is a working template from a completed Markdown-to-Pages resume task. Patch it in the workspace before use; do not edit the skill copy for task-specific personal data.
- `scripts/export_to_pages09.scpt` is parameterized and should be used for Pages exports.

## Verification

Always verify both generated files exist. For visual checks, inspect the Pages QuickLook thumbnail and, when possible, open the `.pages` file in Pages. If the user is tuning layout, prefer regenerating from the script rather than manually editing only the output artifact.
