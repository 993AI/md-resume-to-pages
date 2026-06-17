import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


SOURCE = Path("resume.md")
OUTPUT = Path("resume.docx")
PHOTO = Path("portrait.png")

FONT_CN = "PingFang SC"
FONT_LATIN = "Helvetica Neue"
INK = RGBColor(0, 0, 0)
MUTED = RGBColor(90, 90, 90)
LIGHT_GRAY = "D9D9D9"
BORDER_GRAY = "CFCFCF"
CONTENT_WIDTH = 10000


def set_run(run, size=9.2, bold=False, color=INK):
    run.font.name = FONT_LATIN
    run._element.rPr.rFonts.set(qn("w:ascii"), FONT_LATIN)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def borders(cell, color=BORDER_GRAY, size="4"):
    tc_pr = cell._tc.get_or_add_tcPr()
    bdr = tc_pr.first_child_found_in("w:tcBorders")
    if bdr is None:
        bdr = OxmlElement("w:tcBorders")
        tc_pr.append(bdr)
    for edge in ("top", "left", "bottom", "right"):
        el = bdr.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            bdr.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def no_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    bdr = tc_pr.first_child_found_in("w:tcBorders")
    if bdr is None:
        bdr = OxmlElement("w:tcBorders")
        tc_pr.append(bdr)
    for edge in ("top", "left", "bottom", "right"):
        el = bdr.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            bdr.append(el)
        el.set(qn("w:val"), "nil")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")


def margins(cell, top=80, start=90, bottom=80, end=90):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for key, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        el = tc_mar.find(qn(f"w:{key}"))
        if el is None:
            el = OxmlElement(f"w:{key}")
            tc_mar.append(el)
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")


def widths(table, values):
    table.autofit = False
    grid = table._tbl.tblGrid
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        table._tbl.insert(0, grid)
    for child in list(grid):
        grid.remove(child)
    for value in values:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(value))
        grid.append(col)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Cm(values[i] / 567)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            margins(cell)


def set_row_heights(table, height_pt):
    for row in table.rows:
        row.height = Pt(height_pt)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY


def set_row_height(table, row_idx, height_pt):
    row = table.rows[row_idx]
    row.height = Pt(height_pt)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY


def clear_paragraph(paragraph):
    for run in paragraph.runs:
        run.text = ""


def cell_text(cell, text, size=9.2, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    clear_paragraph(p)
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, color=color)


def header_cell_text(cell, text, size=12, bold=False, color=INK):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    margins(cell, top=40, start=70, bottom=40, end=70)
    cell_text(cell, text, size=size, bold=bold, color=color)


def add_p(container, text="", size=9.2, bold=False, color=INK, before=0, after=2.2, indent=0):
    p = container.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.12
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if text:
        run = p.add_run(text)
        set_run(run, size=size, bold=bold, color=color)
    return p


def add_inline_label_p(container, label, text):
    p = add_p(container, after=2.2)
    r = p.add_run(label)
    set_run(r, size=9.2, bold=True)
    r = p.add_run(text)
    set_run(r, size=9.2)


def parse_resume(md):
    lines = md.splitlines()
    profile = {}
    sections = {}
    current = None
    current_h3 = None

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("|") and "项目" not in line and "---" not in line:
            cells = [c.strip() for c in line.strip("|").split("|")]
            if len(cells) == 2:
                profile[cells[0]] = cells[1]
        if line.startswith("## "):
            current = line[3:].strip()
            sections[current] = []
            current_h3 = None
        elif line.startswith("### ") and current:
            current_h3 = {"title": line[4:].strip(), "lines": []}
            sections[current].append(current_h3)
        elif current and line and not line.startswith("|") and not line.startswith("#"):
            target = current_h3["lines"] if current_h3 else sections[current]
            target.append(line)
        i += 1
    return profile, sections


def strip_md(text):
    return re.sub(r"\*\*([^*]+)\*\*", r"\1", text).strip()


def split_label(line):
    clean = strip_md(line)
    m = re.match(r"([^：:]{2,12})[：:]\s*(.*)", clean)
    if m:
        return m.group(1) + "：", m.group(2)
    return "", clean


def profile_value(profile, *keys, default=""):
    for key in keys:
        value = profile.get(key)
        if value:
            return value
    return default


def add_section_title(doc, title):
    add_p(doc, title, size=10.5, bold=True, before=12, after=6)


def add_section_title_row(doc, title):
    table = doc.add_table(rows=1, cols=1)
    widths(table, [CONTENT_WIDTH])
    cell = table.cell(0, 0)
    no_borders(cell)
    margins(cell, top=150, start=90, bottom=160, end=90)
    cell_text(cell, title, size=10.5, bold=True)


def add_info_grid(doc, profile):
    rows = [
        ("年　　龄：", profile_value(profile, "年龄"), "婚姻状况：", profile_value(profile, "婚姻状况")),
        ("手　　机：", profile_value(profile, "手机", "电话"), "邮　　箱：", profile_value(profile, "邮箱")),
        ("国　　籍：", profile_value(profile, "国籍"), "户　　籍：", profile_value(profile, "户籍")),
        ("目前状态：", profile_value(profile, "目前状态", "当前状态"), "所在地点：", profile_value(profile, "所在地点", "城市")),
        ("学　　历：", profile_value(profile, "学历"), "", ""),
    ]
    table = doc.add_table(rows=len(rows) + 1, cols=4)
    widths(table, [1250, 3300, 1550, 3900])

    title_cell = table.cell(0, 0).merge(table.cell(0, 3))
    no_borders(title_cell)
    margins(title_cell, top=90, start=90, bottom=160, end=90)
    cell_text(title_cell, "基本资料", size=10.5, bold=True)

    for row, values in zip(table.rows[1:], rows):
        for idx, value in enumerate(values):
            cell = row.cells[idx]
            no_borders(cell)
            is_label = idx in (0, 2)
            cell_text(cell, value, size=9.0, bold=is_label, color=MUTED if is_label else INK)


def add_intention(doc, profile):
    rows = [
        ("期望行业：", profile_value(profile, "期望行业", "行业")),
        ("期望职位：", profile_value(profile, "期望职位", "求职意向")),
        ("期望地点：", profile_value(profile, "期望地点", "城市"), "期望年薪：", profile_value(profile, "期望年薪", "薪资", default="面议")),
    ]
    table = doc.add_table(rows=4, cols=4)
    widths(table, [1250, 5450, 1550, 1750])
    title_cell = table.cell(0, 0).merge(table.cell(0, 3))
    no_borders(title_cell)
    margins(title_cell, top=90, start=90, bottom=160, end=90)
    cell_text(title_cell, "职业意向", size=10.5, bold=True)

    for r_idx, row in enumerate(table.rows[1:]):
        for c_idx, cell in enumerate(row.cells):
            no_borders(cell)
            value = rows[r_idx][c_idx] if c_idx < len(rows[r_idx]) else ""
            is_label = c_idx in (0, 2)
            cell_text(cell, value, size=9.0, bold=is_label, color=MUTED if is_label else INK)


def add_gray_band(doc, left, middle, right=None, second=None):
    table = doc.add_table(rows=1, cols=3)
    widths(table, [1900, 5000, 3100])
    for cell in table.rows[0].cells:
        shade(cell, LIGHT_GRAY)
        no_borders(cell)
    cell_text(table.cell(0, 0), left, size=9.0, bold=True)
    cell_text(table.cell(0, 1), middle + (f"\n{second}" if second else ""), size=9.0, bold=True)
    cell_text(table.cell(0, 2), right or "", size=9.0, bold=True)


def add_label_block(doc, label, lines, numbered=True):
    table = doc.add_table(rows=1, cols=2)
    widths(table, [1350, 8650])
    for cell in table.rows[0].cells:
        no_borders(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cell_text(table.cell(0, 0), label, size=8.8, bold=True)
    container = table.cell(0, 1)
    clear_paragraph(container.paragraphs[0])
    for idx, line in enumerate(lines, start=1):
        clean = strip_md(line[2:] if line.startswith("- ") else line)
        prefix = f"{idx}.　" if numbered else ""
        add_p(container, prefix + clean, size=8.8, after=1.4)


def add_key_value_table(doc, rows, label_width=1700, value_width=8300):
    if not rows:
        return
    table = doc.add_table(rows=len(rows), cols=2)
    widths(table, [label_width, value_width])
    for row, (label, value) in zip(table.rows, rows):
        label_cell, value_cell = row.cells
        for cell in row.cells:
            no_borders(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        shade(label_cell, LIGHT_GRAY)
        cell_text(label_cell, label, size=8.8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        clear_paragraph(value_cell.paragraphs[0])
        add_p(value_cell, value, size=8.8, after=1.2)


def add_single_column_table(doc, lines):
    if not lines:
        return
    table = doc.add_table(rows=len(lines), cols=1)
    widths(table, [CONTENT_WIDTH])
    for row, text in zip(table.rows, lines):
        cell = row.cells[0]
        no_borders(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        clear_paragraph(cell.paragraphs[0])
        add_p(cell, strip_md(text), size=8.8, after=1.2)


def add_labeled_line_block(doc, label, text):
    table = doc.add_table(rows=1, cols=2)
    widths(table, [1350, 8650])
    for cell in table.rows[0].cells:
        no_borders(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cell_text(table.cell(0, 0), label, size=8.8, bold=True)
    clear_paragraph(table.cell(0, 1).paragraphs[0])
    add_p(table.cell(0, 1), text, size=8.8, after=1.4)


def add_experience(doc, sections, profile):
    for item in sections.get("工作经历", []):
        lines = item["lines"]
        role = ""
        dates = ""
        intro = []
        bullets = []
        in_bullets = False
        for line in lines:
            label, value = split_label(line)
            if label == "职位：":
                role = value
            elif label == "时间：":
                dates = value
            elif "主要工作内容" in line:
                in_bullets = True
            elif line.startswith("- ") or in_bullets:
                if line.startswith("- "):
                    bullets.append(line)
            else:
                intro.append(strip_md(line))
        location = profile_value(profile, "工作地点", "所在地点", "城市")
        add_gray_band(doc, dates, item["title"], f"工作地点：{location}" if location else "", f"职位：{role}")
        add_labeled_line_block(doc, "工作概况：", "".join(intro))
        add_label_block(doc, "工作职责：", bullets[:7])


def add_projects(doc, sections):
    for item in sections.get("项目经历", []):
        lines = item["lines"]
        meta = []
        desc = []
        duties = []
        results = []
        target = meta
        for line in lines:
            clean = strip_md(line)
            if "项目描述" in clean:
                target = desc
                continue
            if "主要职责" in clean:
                target = duties
                continue
            if "项目成果" in clean:
                target = results
                continue
            target.append(line)
        add_gray_band(doc, "", item["title"], "")
        meta_rows = []
        for line in meta[:5]:
            label, value = split_label(line)
            if label:
                meta_rows.append((label.rstrip("："), value))
        add_key_value_table(doc, meta_rows)
        if desc:
            add_labeled_line_block(doc, "项目描述：", "".join(strip_md(x) for x in desc))
        if duties:
            add_label_block(doc, "主要职责：", duties[:7])
        if results:
            add_label_block(doc, "项目成果：", results[:4])


def add_education(doc, sections):
    for item in sections.get("教育经历", []):
        if isinstance(item, dict):
            school = item["title"]
            degree = ""
            major = ""
            dates = ""
            for line in item["lines"]:
                label, value = split_label(line)
                if label in ("学历：", "层次："):
                    degree = value
                elif label == "专业：":
                    major = value
                elif label in ("时间：", "在校时间："):
                    dates = value
            add_gray_band(doc, dates, school, degree)
            if major:
                add_inline_label_p(doc, "专业：", major)


def build():
    profile, sections = parse_resume(SOURCE.read_text(encoding="utf-8"))
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.9)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.55)
    section.right_margin = Cm(1.55)

    normal = doc.styles["Normal"]
    normal.font.name = FONT_LATIN
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
    normal.font.size = Pt(9.2)
    normal.paragraph_format.space_after = Pt(2)

    header = doc.add_table(rows=3, cols=5)
    widths(header, [1900, 1700, 3100, 1550, 1750])
    set_row_height(header, 0, 50)
    set_row_height(header, 1, 54)
    set_row_height(header, 2, 54)
    for row in header.rows:
        for cell in row.cells:
            shade(cell, LIGHT_GRAY)
            no_borders(cell)
    photo_cell = header.cell(0, 0).merge(header.cell(2, 0))
    photo_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    margins(photo_cell, top=35, start=55, bottom=35, end=35)
    clear_paragraph(photo_cell.paragraphs[0])
    photo_p = photo_cell.paragraphs[0]
    photo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    photo_p.paragraph_format.space_before = Pt(0)
    photo_p.paragraph_format.space_after = Pt(0)
    photo_p.paragraph_format.line_spacing = 1.0
    if PHOTO.exists():
        photo_run = photo_p.add_run()
        photo_run.add_picture(str(PHOTO), width=Cm(2.78), height=Cm(3.70))
    current_company = profile_value(profile, "目前公司", "当前公司")
    if not current_company and sections.get("工作经历") and isinstance(sections["工作经历"][0], dict):
        current_company = sections["工作经历"][0]["title"]
    current_title = profile_value(profile, "目前职位", "当前职位", "求职意向")
    current_title = current_title.split("/")[0].strip() if current_title else ""

    header_cell_text(header.cell(0, 1), profile_value(profile, "姓名", default="姓名"), size=19, bold=True)
    header_cell_text(header.cell(0, 2), profile_value(profile, "性别"), size=12)
    blank_top = header.cell(0, 3).merge(header.cell(0, 4))
    header_cell_text(blank_top, "", size=9.2)
    header_cell_text(header.cell(1, 1), "目前公司：", size=12, color=MUTED)
    company_cell = header.cell(1, 2).merge(header.cell(1, 4))
    header_cell_text(company_cell, current_company, size=12, bold=True)
    header_cell_text(header.cell(2, 1), "目前职位：", size=12, color=MUTED)
    header_cell_text(header.cell(2, 2), current_title, size=12, bold=True)
    header_cell_text(header.cell(2, 3), "工作年限：", size=12, bold=True)
    header_cell_text(header.cell(2, 4), profile_value(profile, "工作年限"), size=12, bold=True)

    add_info_grid(doc, profile)
    add_intention(doc, profile)
    add_section_title_row(doc, "个人优势")
    add_single_column_table(doc, sections.get("个人优势", []))
    add_section_title_row(doc, "专业技能")
    skill_rows = []
    for line in sections.get("专业技能", []):
        label, value = split_label(line)
        if label:
            skill_rows.append((label.rstrip("："), value))
    add_key_value_table(doc, skill_rows)
    add_section_title_row(doc, "工作经历")
    add_experience(doc, sections, profile)
    add_section_title_row(doc, "项目经历")
    add_projects(doc, sections)
    add_section_title_row(doc, "教育经历")
    add_education(doc, sections)
    add_section_title_row(doc, "自我评价")
    add_single_column_table(doc, sections.get("自我评价", []))

    doc.save(OUTPUT)


if __name__ == "__main__":
    build()
