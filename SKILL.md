---
name: md-resume-to-pages
description: Convert Markdown resumes into polished Word `.docx` and Apple Pages `.pages` resumes using a repeatable md -> docx -> Pages workflow. Use when the user asks to turn a `.md` resume into a `.docx` or `.pages` resume, match an existing Pages resume layout, preserve a resume template, export through Apple Pages, or iterate on resume table/print layout such as a 3-row by 5-column identity header, basic information tables, project tables, and A4 visual QA.
---

# Markdown Resume To Pages

## Overview

Convert a structured Markdown resume into Word and Apple Pages. The skill provides a reusable build script that handles:

- **Markdown parsing** — profile tables, `##` sections, `###` sub-sections
- **A4 layout** — 3×5 top identity block with photo, basic info grid, career intention, work/project/education sections
- **Pages export** — via AppleScript automation, with blank page fix

## Quick Start

```bash
# 1. Copy the build script to your working directory
cp /path/to/skill/scripts/build_pages_style_resume.py ./

# 2. Edit the constants at the top of the script:
#    SOURCE = Path("your_resume.md")
#    OUTPUT = Path("your_resume.docx")
#    PHOTO = Path("your_photo.png")

# 3. Generate the Word file
python3 build_pages_style_resume.py

# 4. Export to Pages
osascript /path/to/skill/scripts/export_to_pages09.scpt \
  /path/to/your_resume.docx \
  /path/to/your_resume.pages
```

## Script Customization Guide

The template script (`build_pages_style_resume.py`) has all customization points at the top:

### Paths

```python
SOURCE = Path("简历.md")       # 源 Markdown 文件
OUTPUT = Path("简历.docx")     # 输出的 Word 文件
PHOTO = Path("photo.png")      # 证件照（不存在则跳过）
```

### Font Sizes

```python
BASE_SIZE = 9.2                # 正文字号
HEADER_FONT_SIZE = 9           # 顶部身份区字号
SECTION_TITLE_SIZE = 10        # 章节标题字号
```

### Field Name Mappings

The script auto-maps Chinese resume header names. If your resume uses field names not in the default list, add them to `PROFILE_FIELD_ALIASES`:

```python
PROFILE_FIELD_ALIASES = {
    "姓名": ["姓名", "名字", "Name"],
    "手机": ["手机", "电话", "手机号", "联系电话"],
    "城市": ["城市", "所在地点", "所在地", "工作地点"],
    "求职意向": ["求职意向", "期望职位", "应聘职位", "职位"],
    # ... add more as needed
}
```

## Markdown Resume Format

The script expects this structure:

```markdown
# 个人简历

| 项目 | 信息 |
|---|---|
| 姓名 | XXX    |     ← 解析为 profile 字段
| 手机 | 138... |
| ...  | ...   |

## 个人优势          ← sections 键名，必须匹配以下之一
(plain text lines)

## 专业技能
**后端开发：** Java、Spring Boot...   ← 解析为 标签：值

## 工作经历           ← 支持 ### 子标题
### 公司名称
**职位：** ...
**时间：** ...
(main text)
**主要工作内容：**
- 职责1
- 职责2

## 项目经历
### 项目名称
**所属公司：** ...
**技术栈：** ...
**项目描述：** ...
**主要职责：** ...
**项目成果：** ...

## 教育经历
| 学校 | 学历 | 专业 | 时间 |   ← 支持表格格式
|---|---|---|---|

## 自我评价
(plain text lines)
```

### Supported Section Names

| 中文 | 用途 |
|---|---|
| `个人优势` | 纯文本段落 |
| `专业技能` | 标签：值 列表 |
| `工作经历` | `###` 子标题，解析 职位/时间/主要工作内容 |
| `项目经历` | `###` 子标题，解析 项目描述/主要职责/项目成果 |
| `教育经历` | 表格格式 或 `###` 子标题 |
| `自我评价` | 纯文本段落 |

## Workflow

1. Copy `scripts/build_pages_style_resume.py` into your workspace
2. Edit constants (SOURCE, OUTPUT, PHOTO, BASE_SIZE)
3. Build the Word file: `python3 build_pages_style_resume.py`
4. Export to Pages:

```bash
osascript /path/to/skill/scripts/export_to_pages09.scpt \
  /path/to/input.docx \
  /path/to/output.pages
```

5. Fix blank page in the `.pages` file:

```bash
python3 build_pages_style_resume.py fix-pages /path/to/output.pages
```

6. Visual check:

```bash
unzip -p output.pages QuickLook/Thumbnail.png > /private/tmp/resume-thumb.png
open /private/tmp/resume-thumb.png
```

7. Iterate as needed

## Verification

- Check both `.docx` and `.pages` files exist
- Inspect the QuickLook thumbnail for layout issues (overlap, clipping, grid lines, spacing)
- The script automatically removes trailing empty paragraphs (docx) and empty sections (pages) to prevent extra blank pages in Pages export
