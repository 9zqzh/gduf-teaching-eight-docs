"""Low-level python-docx helpers shared by the eight teaching documents.

All measurements are in centimetres unless noted. Chinese text uses 宋体,
Latin text and digits use Times New Roman, body size is 10.5pt by default,
matching the school templates.
"""

import os
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BODY_CN = "宋体"
BODY_EN = "Times New Roman"
TEXT_WIDTH_CM = 15.6
TWIPS_PER_CM = 567

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_MARKUP = re.compile(r"(\^\{[^{}]*\}|_\{[^{}]*\})")


def new_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    # Body pages of the school template use 2.8/2.6 cm margins (15.6 cm text width).
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)
    section.top_margin = Cm(2.8)
    section.bottom_margin = Cm(2.6)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.75)

    normal = doc.styles["Normal"]
    normal.font.name = BODY_EN
    normal.font.size = Pt(10.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_CN)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    return doc


def set_run_font(run, size=10.5, bold=False, cn=BODY_CN, en=BODY_EN, italic=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:cs"), en)


def _add_marked_runs(paragraph, text, size, bold, cn, en, italic=False):
    for token in _MARKUP.split(text):
        if not token:
            continue
        if token.startswith("^{") and token.endswith("}"):
            run = paragraph.add_run(token[2:-1])
            set_run_font(run, size=size, bold=bold, cn=cn, en=en, italic=italic)
            run.font.superscript = True
        elif token.startswith("_{") and token.endswith("}"):
            run = paragraph.add_run(token[2:-1])
            set_run_font(run, size=size, bold=bold, cn=cn, en=en, italic=italic)
            run.font.subscript = True
        else:
            run = paragraph.add_run(token)
            set_run_font(run, size=size, bold=bold, cn=cn, en=en, italic=italic)


def add_p(
    container,
    text="",
    size=10.5,
    bold=False,
    align="justify",
    cn=BODY_CN,
    en=BODY_EN,
    indent=0,
    line_spacing=1.5,
    space_before=0,
    space_after=0,
    keep_with_next=False,
    italic=False,
    style=None,
    left_cm=0,
):
    p = container.add_paragraph(style=style)
    fmt = p.paragraph_format
    if align:
        p.alignment = ALIGN[align]
    if left_cm:
        fmt.left_indent = Cm(left_cm)
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)
    if line_spacing:
        fmt.line_spacing = line_spacing
    else:
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if indent:
        fmt.first_line_indent = Pt(size * indent)
    fmt.keep_with_next = keep_with_next
    if text:
        _add_marked_runs(p, str(text), size, bold, cn, en, italic)
    return p


def add_title(doc, text, size=22, font=BODY_CN, bold=True, space_before=6, space_after=12):
    return add_p(
        doc,
        text,
        size=size,
        bold=bold,
        align="center",
        cn=font,
        line_spacing=1.5,
        space_before=space_before,
        space_after=space_after,
        keep_with_next=True,
    )


def add_section_heading(doc, text, space_before=10, space_after=4, font=BODY_CN):
    return add_p(
        doc,
        text,
        size=12,
        bold=False,
        align="left",
        cn=font,
        en=font,
        indent=2,
        line_spacing=1.5,
        space_before=space_before,
        space_after=space_after,
        keep_with_next=True,
    )


def add_blank(doc, count=1):
    for _ in range(count):
        add_p(doc, "", line_spacing=1.5)


def add_page_break(doc):
    from docx.enum.text import WD_BREAK

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def add_cover(doc, fields, maker="广东金融学院教务处  制", maker_size=14, images=None):
    """Cover page. ``images`` is a list of (path, width_cm) laid out in one line."""
    add_blank(doc, 3)
    if images:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        for path, width_cm in images:
            if os.path.exists(path):
                p.add_run().add_picture(path, width=Cm(width_cm))
        add_blank(doc, 3)
    for label, value in fields:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt = p.paragraph_format
        fmt.left_indent = Cm(2.36)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.5
        head, sep, tail = str(label).partition("：")
        label_run = p.add_run(head + sep if sep else head)
        set_run_font(label_run, size=16, bold=True)
        value_run = p.add_run(str(value))
        set_run_font(value_run, size=16, bold=True)
        value_run.font.underline = True
        add_p(doc, "", line_spacing=1.5)
    add_blank(doc, 3)
    add_p(doc, maker, size=maker_size, align="center", line_spacing=1.5)
    add_page_break(doc)


def add_page_number_footer(section, size=9):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    set_run_font(run, size=size)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _set_tbl_borders(table, sz=4, color="000000"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.insert_element_before(
            borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription"
        )
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if borders.find(qn(f"w:{edge}")) is not None:
            continue
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)


def _set_tbl_layout_fixed(table, widths_cm):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(sum(widths_cm) * TWIPS_PER_CM)))
    tbl_w.set(qn("w:type"), "dxa")

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.insert_element_before(layout, "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription")
    layout.set(qn("w:type"), "fixed")

    grid = tbl.tblGrid
    cols = grid.findall(qn("w:gridCol"))
    for col, width in zip(cols, widths_cm):
        col.set(qn("w:w"), str(int(width * TWIPS_PER_CM)))


def _set_row_cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def _mark_header_row(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))


def _clear_cell(cell):
    tc = cell._tc
    tc.clear_content()
    return cell.add_paragraph()


def _set_cell_content(cell, text, size=10.5, bold=False, align="center", cn=BODY_CN, en=BODY_EN):
    lines = str(text).split("\n") if text is not None else [""]
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first and cell.paragraphs else cell.add_paragraph()
        first = False
        p.alignment = ALIGN[align]
        fmt = p.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0
        fmt.first_line_indent = Pt(0)
        if line:
            _add_marked_runs(p, line, size, bold, cn, en)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_table(
    doc,
    rows,
    widths_cm,
    header_rows=0,
    size=10.5,
    body_align="left",
    header_align="center",
    repeat_header=False,
):
    """Build a bordered table.

    Each row is a list of cells. A cell is either a string, or a dict:
      {"text": str, "span": int, "vmerge": int, "bold": bool,
       "align": "left|center|right", "size": float}
    ``span`` merges horizontally, ``vmerge`` merges downward for N rows.
    Rows may be shorter than the column count when spans are used.
    """
    n_cols = len(widths_cm)
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = doc.styles["Table Grid"]
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_tbl_borders(table)
    _set_tbl_layout_fixed(table, widths_cm)

    for c, width in enumerate(widths_cm):
        table.columns[c].width = Cm(width)
        for r in range(len(rows)):
            table.cell(r, c).width = Cm(width)

    specs = {}
    for r, row in enumerate(rows):
        c = 0
        for cell in row:
            if cell is None:
                c += 1
                continue
            spec = cell if isinstance(cell, dict) else {"text": cell}
            specs[(r, c)] = spec
            c += int(spec.get("span", 1))

    for (r, c), spec in list(specs.items()):
        span = int(spec.get("span", 1))
        if span > 1:
            merged = table.cell(r, c).merge(table.cell(r, c + span - 1))
            merged.width = Cm(sum(widths_cm[c : c + span]))

    for (r, c), spec in list(specs.items()):
        vmerge = int(spec.get("vmerge", 1))
        if vmerge > 1:
            merged = table.cell(r, c).merge(table.cell(r + vmerge - 1, c))
            merged.width = Cm(widths_cm[c])

    for (r, c), spec in specs.items():
        cell = table.cell(r, c)
        _clear_cell(cell)
        is_header = r < header_rows
        _set_cell_content(
            cell,
            spec.get("text", ""),
            size=spec.get("size", size),
            bold=spec.get("bold", is_header),
            align=spec.get("align", header_align if is_header else body_align),
            cn=spec.get("cn", BODY_CN),
            en=spec.get("en", BODY_EN),
        )

    for r in range(len(rows)):
        _set_row_cant_split(table.rows[r])
        if repeat_header and r < header_rows:
            _mark_header_row(table.rows[r])

    return table


def add_field_line(doc, label, value, size=10.5, bold=False, indent=2):
    return add_p(doc, f"{label}{value}", size=size, bold=bold, align="left", indent=indent, line_spacing=1.5)


def add_numbered_items(doc, items, size=10.5, indent=2, start=1, numbered=True):
    for i, item in enumerate(items, start=start):
        prefix = f"[{i}] " if numbered else ""
        add_p(doc, prefix + str(item), size=size, align="justify", indent=indent, line_spacing=1.5)


def clean_text(text):
    if text is None:
        return ""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("###", "").strip()
    return text
