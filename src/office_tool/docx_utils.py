"""Small OOXML helpers for Word documents."""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .config import StyleSpec


ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def apply_style_to_paragraph(paragraph: Paragraph, style: StyleSpec, preserve_bold_italic: bool = True) -> None:
    paragraph.alignment = ALIGNMENT_MAP.get(style.alignment, WD_ALIGN_PARAGRAPH.LEFT)
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(style.space_before_pt)
    fmt.space_after = Pt(style.space_after_pt)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(style.line_spacing_pt or 28)
    set_first_line_chars(paragraph, style.first_line_chars)

    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        old_bold = run.bold
        old_italic = run.italic
        apply_font(run, style.font, style.size_pt, style.color)
        if style.bold is not None:
            run.bold = style.bold
        elif preserve_bold_italic:
            run.bold = old_bold
        if preserve_bold_italic:
            run.italic = old_italic


def apply_font(run: Run, font: str, size_pt: float, color: str | None = None) -> None:
    run.font.name = font
    run.font.size = Pt(size_pt)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for key in ("w:eastAsia", "w:ascii", "w:hAnsi", "w:cs"):
        r_fonts.set(qn(key), font)
    if color:
        run.font.color.rgb = RGBColor.from_string(color.strip().lstrip("#").upper())


def set_first_line_chars(paragraph: Paragraph, chars: int) -> None:
    fmt = paragraph.paragraph_format
    if chars <= 0:
        fmt.first_line_indent = None
        p_pr = paragraph._p.get_or_add_pPr()
        ind = p_pr.get_or_add_ind()
        for attr in ("w:firstLine", "w:firstLineChars"):
            if ind.get(qn(attr)) is not None:
                ind.attrib.pop(qn(attr), None)
        return

    fmt.first_line_indent = None
    p_pr = paragraph._p.get_or_add_pPr()
    ind = p_pr.get_or_add_ind()
    ind.set(qn("w:firstLineChars"), str(chars * 100))


def add_bottom_border(paragraph: Paragraph, color: str = "FF0000", size: str = "12") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = p_bdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        p_bdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), color.strip().lstrip("#").upper())


def add_page_number(paragraph: Paragraph, font: str, size_pt: float) -> None:
    paragraph.text = ""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    apply_font(run, font, size_pt)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)
