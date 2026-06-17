# Resume Layout Reference

Use these rules when adapting `scripts/build_pages_style_resume.py` for a Markdown resume.

## Conversion Chain

1. Read the source `.md` resume and parse the profile table plus `##` sections.
2. Build a `.docx` with `python-docx`.
3. Export the `.docx` to `.pages` with Pages 09 export.
4. Extract `QuickLook/Thumbnail.jpg` from the `.pages` package and inspect it visually.
5. Iterate until the first-page layout has no overlap, clipping, unexpected table lines, or awkward spacing.

## Top Identity Table

The top identity block should be a real 3-row by 5-column table, not a loose visual simulation.

- Column A: portrait photo merged from A1:A3.
- B1: name.
- C1: gender.
- D1:E1: empty gray background, merged if needed.
- B2: current company label.
- C2:E2: current company value merged across columns.
- B3: current role label.
- C3: current role value.
- D3: work years label.
- E3: work years value.

Use a light gray fill for all cells. Remove visible borders with OOXML `w:val="nil"`. Keep the photo inside the merged A1:A3 cell with enough padding so Pages conversion does not overlap or clip it.

## Section Tables

- Basic information and job intention should be table-based and full-width enough for A4 output.
- Section titles can be one-cell borderless tables so Pages treats them consistently.
- Personal advantages and self-evaluation should be single-column borderless tables.
- Professional skills can use a two-column table with gray label cells and borderless cell edges.
- For each project, render company, project role, and tech stack as table rows rather than inline prose when the source data supports it.

## Visual QA

Use the generated Pages thumbnail for a quick first-page check:

```bash
unzip -p "output.pages" QuickLook/Thumbnail.jpg > /private/tmp/resume-thumb.jpg
```

Open the image and check:

- Top identity block is a true table-like gray band.
- No text overlaps the portrait or adjacent cells.
- No white internal grid lines remain unless explicitly requested.
- Right edge looks balanced for A4 print; avoid leaving a narrow unused strip.
- Section spacing is readable but compact enough for a resume.
