"""Small OOXML helpers for Word documents."""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

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
    if style.line_spacing_pt is None:
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
        fmt.line_spacing = 1.0
    else:
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Pt(style.line_spacing_pt)
    set_first_line_chars(paragraph, style.first_line_chars, style.size_pt)
    set_side_indent_chars(paragraph, style.left_indent_chars, style.right_indent_chars, style.size_pt)

    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        old_bold = run.bold
        old_italic = run.italic
        apply_font(run, style.font, style.size_pt, style.color, latin_font=style.latin_font)
        if style.bold is not None:
            run.bold = style.bold
        elif preserve_bold_italic:
            run.bold = old_bold
        if preserve_bold_italic:
            run.italic = old_italic


def apply_font(run: Run, font: str, size_pt: float, color: str | None = None, latin_font: str | None = None) -> None:
    run.font.name = font
    run.font.size = Pt(size_pt)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    latin = latin_font or font
    for key in ("w:eastAsia", "w:cs"):
        r_fonts.set(qn(key), font)
    for key in ("w:ascii", "w:hAnsi"):
        r_fonts.set(qn(key), latin)
    if color:
        run.font.color.rgb = RGBColor.from_string(color.strip().lstrip("#").upper())


def set_first_line_chars(paragraph: Paragraph, chars: int, size_pt: float = 16) -> None:
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
    ind.set(qn("w:firstLine"), str(int(chars * size_pt * 20)))
    ind.set(qn("w:firstLineChars"), str(chars * 100))


def set_side_indent_chars(paragraph: Paragraph, left_chars: float = 0, right_chars: float = 0, size_pt: float = 16) -> None:
    fmt = paragraph.paragraph_format
    fmt.left_indent = Pt(left_chars * size_pt) if left_chars > 0 else None
    fmt.right_indent = Pt(right_chars * size_pt) if right_chars > 0 else None
    p_pr = paragraph._p.get_or_add_pPr()
    ind = p_pr.get_or_add_ind()
    for attr, chars in (("w:leftChars", left_chars), ("w:rightChars", right_chars)):
        if chars > 0:
            ind.set(qn(attr), str(int(round(chars * 100))))
        elif ind.get(qn(attr)) is not None:
            ind.attrib.pop(qn(attr), None)


def add_floating_line(
    paragraph: Paragraph,
    *,
    width_cm: float,
    weight_pt: float,
    color: str,
    shape_id: str,
    compound: str | None = None,
    vertical_offset_pt: float = 0.0,
    vertical_relative: str = "line",
    clear_paragraph: bool = True,
) -> None:
    """Add a centered zero-height VML line in front of text."""
    if clear_paragraph:
        paragraph.text = ""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    width_pt = width_cm / 2.54 * 72

    run = OxmlElement("w:r")
    pict = OxmlElement("w:pict")
    vml_ns = "urn:schemas-microsoft-com:vml"
    office_ns = "urn:schemas-microsoft-com:office:office"
    word10_ns = "urn:schemas-microsoft-com:office:word"
    line = etree.Element(f"{{{vml_ns}}}line", nsmap={"v": vml_ns, "o": office_ns, "w10": word10_ns})
    line.set("id", shape_id)
    line.set("from", "0,0")
    line.set("to", f"{width_pt:.2f}pt,0")
    line.set("strokecolor", f"#{color.strip().lstrip('#').upper()}")
    line.set("strokeweight", f"{weight_pt:g}pt")
    line.set(
        "style",
        f"position:absolute;left:0;text-align:left;margin-left:0;margin-top:{vertical_offset_pt:g}pt;"
        f"width:{width_pt:.2f}pt;height:0;z-index:251659264;"
        "mso-position-horizontal:center;mso-position-horizontal-relative:margin;"
        f"mso-position-vertical:{'absolute' if vertical_relative == 'page' else 'top'};"
        f"mso-position-vertical-relative:{vertical_relative}",
    )
    line.set(f"{{{office_ns}}}allowincell", "f")
    line.set(f"{{{office_ns}}}allowoverlap", "f")
    if compound:
        stroke = etree.Element(f"{{{vml_ns}}}stroke")
        stroke.set("linestyle", compound)
        line.append(stroke)
    wrap = etree.Element(f"{{{word10_ns}}}wrap")
    wrap.set("type", "none")
    line.append(wrap)
    pict.append(line)
    run.append(pict)
    paragraph._p.append(run)


def clear_paragraph_frame(paragraph: Paragraph) -> None:
    p_pr = paragraph._p.pPr
    frame = p_pr.find(qn("w:framePr")) if p_pr is not None else None
    if frame is not None:
        p_pr.remove(frame)


def set_document_grid(
    section,
    chars_per_line: int,
    lines_per_page: int,
    char_space_pt: float = 15.8,
    line_pitch_pt: float = 28.95,
    normal_font_size_pt: float = 16.0,
) -> None:
    sect_pr = section._sectPr
    doc_grid = sect_pr.find(qn("w:docGrid"))
    if doc_grid is None:
        doc_grid = OxmlElement("w:docGrid")
        sect_pr.append(doc_grid)
    char_space = int(round((char_space_pt - normal_font_size_pt) * 4096))
    doc_grid.set(qn("w:type"), "lines")
    doc_grid.set(qn("w:charsPerLine"), str(chars_per_line))
    doc_grid.set(qn("w:linesPerPage"), str(lines_per_page))
    doc_grid.set(qn("w:charSpace"), str(char_space))
    doc_grid.set(qn("w:linePitch"), str(int(round(line_pitch_pt * 20))))


def get_document_grid(section) -> dict[str, str]:
    doc_grid = section._sectPr.find(qn("w:docGrid"))
    if doc_grid is None:
        return {}
    return {
        "type": doc_grid.get(qn("w:type"), ""),
        "charSpace": doc_grid.get(qn("w:charSpace"), ""),
        "linePitch": doc_grid.get(qn("w:linePitch"), ""),
        "charsPerLine": doc_grid.get(qn("w:charsPerLine"), ""),
        "linesPerPage": doc_grid.get(qn("w:linesPerPage"), ""),
    }


def set_even_and_odd_headers(doc: DocxDocument, enabled: bool = True) -> None:
    settings = doc.settings._element
    existing = settings.find(qn("w:evenAndOddHeaders"))
    if enabled and existing is None:
        settings.append(OxmlElement("w:evenAndOddHeaders"))
    elif not enabled and existing is not None:
        settings.remove(existing)


def add_page_number(
    paragraph: Paragraph,
    font: str,
    size_pt: float,
    alignment=WD_ALIGN_PARAGRAPH.CENTER,
    with_dashes: bool = True,
    edge_space_chars: int = 0,
) -> None:
    paragraph.text = ""
    paragraph.alignment = alignment
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    if alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        set_side_indent_chars(paragraph, right_chars=edge_space_chars, size_pt=size_pt)
    elif alignment == WD_ALIGN_PARAGRAPH.LEFT:
        set_side_indent_chars(paragraph, left_chars=edge_space_chars, size_pt=size_pt)
    else:
        set_side_indent_chars(paragraph, size_pt=size_pt)
    run = paragraph.add_run()
    apply_font(run, font, size_pt)
    if with_dashes:
        run.add_text("— ")

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
    if with_dashes:
        run.add_text(" —")


def add_body_text_box(
    paragraph: Paragraph,
    text: str,
    font: str,
    size_pt: float,
    color: str,
    line_spacing_pt: float = 40,
    width_cm: float = 15.5,
    height_cm: float = 1.4,
    character_spacing_pt: float = 0.0,
    top_cm: float = 3.0,
) -> None:
    """Anchor the letter-head text box in the first-page body."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    paragraph.text = ""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)

    color_hex = str(color).strip().lstrip("#").upper()
    width_pt = width_cm / 2.54 * 72
    height_pt = height_cm / 2.54 * 72
    top_pt = top_cm / 2.54 * 72

    r = OxmlElement("w:r")
    pict = OxmlElement("w:pict")
    vml_ns = "urn:schemas-microsoft-com:vml"
    office_ns = "urn:schemas-microsoft-com:office:office"
    word_ns = "urn:schemas-microsoft-com:office:word"

    shapetype = etree.Element(f"{{{vml_ns}}}shapetype")
    shapetype.set("id", "_x0000_t202")
    shapetype.set("coordsize", "21600,21600")
    shapetype.set(f"{{{office_ns}}}spt", "202")
    path = etree.Element(f"{{{vml_ns}}}path")
    path.set("gradientshapeok", "t")
    path.set(f"{{{office_ns}}}connecttype", "rect")
    shapetype.append(path)
    pict.append(shapetype)

    shape = etree.Element(f"{{{vml_ns}}}shape")
    shape.set("id", "OfficeToolLetterHeadTextBox")
    shape.set("type", "#_x0000_t202")
    shape.set(
        "style",
        f"position:absolute;left:0;text-align:left;margin-left:0;margin-top:{top_pt:.1f}pt;"
        f"width:{width_pt:.1f}pt;height:{height_pt:.1f}pt;z-index:251658240;"
        f"mso-position-horizontal:center;mso-position-horizontal-relative:page;"
        f"mso-position-vertical:absolute;mso-position-vertical-relative:page",
    )
    shape.set("stroked", "f")
    shape.set("fillcolor", "#" + color_hex)
    shape.set("filled", "f")

    textbox = etree.Element(f"{{{vml_ns}}}textbox")
    textbox.set("inset", "0,0,0,0")
    textbox.set("style", "mso-fit-shape-to-text:f")

    content = OxmlElement("w:txbxContent")
    text_p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(int(round(line_spacing_pt * 20))))
    spacing.set(qn("w:lineRule"), "exact")
    p_pr.extend([jc, spacing])
    text_p.append(p_pr)
    text_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for key in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(key), font)
    text_color = OxmlElement("w:color")
    text_color.set(qn("w:val"), color_hex)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), str(int(round(size_pt * 2))))
    size_cs = OxmlElement("w:szCs")
    size_cs.set(qn("w:val"), str(int(round(size_pt * 2))))
    r_pr.extend([fonts, text_color, size, size_cs])
    if character_spacing_pt:
        char_spacing = OxmlElement("w:spacing")
        char_spacing.set(qn("w:val"), str(int(round(character_spacing_pt * 20))))
        r_pr.append(char_spacing)
    text_run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    text_run.append(text_node)
    text_p.append(text_run)
    content.append(text_p)
    textbox.append(content)

    shape.append(textbox)
    pict.append(shape)
    r.append(pict)
    paragraph._p.append(r)
