# Resume Layout Reference

Use these rules when adapting `scripts/build_pages_style_resume.py` for a Markdown resume.

## Conversion Chain

1. Read the source `.md` resume and parse the profile table plus `##` sections.
2. Build a `.docx` with `python-docx`.
3. Export the `.docx` to `.pages` with Pages 09 export.
4. Extract `QuickLook/Thumbnail.png` from the `.pages` package and inspect it visually.
5. Iterate until the first-page layout has no overlap, clipping, unexpected table lines, or awkward spacing.

## Top Identity Table

The top identity block should be a real 3-row by 5-column table, not a loose visual simulation.

- Column A: portrait photo merged from A1:A3.
- B1: name (large, bold).
- C1: gender.
- D1 and E1: empty gray background.
- B2: current company label ("目前公司：").
- C2: current company value. Keep D2 and E2 empty so the table remains a true 3×5 table after Pages import.
- B3: current role label ("目前职位：").
- C3: current role value.
- D3: work years label ("工作年限：").
- E3: work years value.

Use a light gray fill (`D9D9D9`) for all cells. Remove visible borders with OOXML `w:val="nil"`. Keep the photo inside the merged A1:A3 cell with enough padding so Pages conversion does not overlap or clip it.

For the reference resume copy, Pages reported these measured dimensions for the first table:

- Column widths: 81.35, 90.15, 130.75, 82.40, 116.95 pt.
- Row heights after Pages import: 28.00, 34.65, 45.05 pt.

When generating from Word, use approximately `1627, 1803, 2615, 1648, 2339` dxa for the five columns and `24, 26.65, 37.05` pt row heights so Pages imports to the measured row heights.

## Profile Table Parsing

The profile table at the top of the markdown (| 项目 | 信息 |) is parsed into a dictionary. The script maps Chinese field names through `PROFILE_FIELD_ALIASES`. For example:

| Markdown header | Mapped to |
|---|---|
| 姓名 / 名字 / Name | `姓名` |
| 手机 / 电话 / 手机号 | `手机` |
| 邮箱 / Email / 电子邮箱 | `邮箱` |
| 城市 / 所在地点 / 工作地点 | `城市` |
| 求职意向 / 期望职位 / 应聘职位 | `求职意向` |

See `PROFILE_FIELD_ALIASES` in the script for the full list.

## Section Tables

- **基本资料** and **职业意向** are 4-column table grids with a merged title row.
- **Section titles** are one-cell borderless tables so Pages treats them consistently.
- **个人优势** and **自我评价** are single-column borderless tables.
- **专业技能** uses a two-column table with gray label cells and borderless cell edges.
- **工作经历** uses gray bands for company/time rows, followed by a labeled text block and numbered bullet list.
- **项目经历** uses gray bands for project names, a key-value table for metadata (company, role, tech stack), then labeled blocks for description, duties, and results.
- **教育经历** supports two formats:
  - **Markdown table** (| 学校 | 学历 | 专业 | 时间 |): single-column rows with dates, school, degree.
  - **Dict format** (`### 学校名称` with key-value lines): uses gray bands like work experience.

## Font & Spacing Defaults

- Body text: 9.2 pt PingFang SC / Helvetica Neue
- Header identity area: 9 pt (name is 15 pt bold)
- Section titles: 10 pt bold
- Line spacing: 1.05 for table cells, 1.12 for paragraph blocks
- Page margins: top 1.9 cm, bottom 1.5 cm, left/right 1.55 cm
- A4 paper (210 × 297 mm)

## Blank Page Prevention

The script applies two fixes automatically:

1. **Docx fix**: Removes trailing empty paragraphs and zeroes the last paragraph's spacing before saving.
2. **Pages fix**: After Pages export, removes any trailing `<sf:section>` from `index.xml` inside the `.pages` package.

These prevent the common issue of an extra blank page at the end of the Pages document.

## Visual QA

Use the generated Pages thumbnail for a quick first-page check:

```bash
unzip -p "output.pages" QuickLook/Thumbnail.png > /private/tmp/resume-thumb.png
open /private/tmp/resume-thumb.png
```

Check:

- Top identity block is a true table-like gray band.
- No text overlaps the portrait or adjacent cells.
- No white internal grid lines remain unless explicitly requested.
- Right edge looks balanced for A4 print; avoid leaving a narrow unused strip.
- Section spacing is readable but compact enough for a resume.
- No extra blank page at the end.
