"""Document-level features: header/footer, TOC, template loading, page setup,
default fonts, multi-column layout, tab stops, bookmarks, and footnotes."""
import itertools
import logging
import re

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.shared import Cm, Inches, Mm, Pt, Twips

from template_utils import find_docx_template
from .patterns import _PAGE_TOKEN_RE

logger = logging.getLogger(__name__)
def load_templates():
    """Resolve Word template path from custom/default template directories.
    Returns absolute path as string or None if not found.
    """
    path = find_docx_template()
    if path:
        logger.debug(f"Using Word template: {path}")
    else:
        logger.warning("No Word template found, will create a blank document")
    return path
# ---------------------------------------------------------------------------
# Header / footer / Word fields
# ---------------------------------------------------------------------------
def _add_field(paragraph, field_code):
    """Insert a Word field (PAGE, NUMPAGES, etc.) into a paragraph."""
    for fld_type, text in [('begin', None), (None, field_code), ('end', None)]:
        run = paragraph.add_run()
        if fld_type:
            fld = OxmlElement('w:fldChar')
            fld.set(qn('w:fldCharType'), fld_type)
            run._r.append(fld)
        else:
            elem = OxmlElement('w:instrText')
            elem.set(qn('xml:space'), 'preserve')
            elem.text = f' {text} '
            run._r.append(elem)
def set_header_footer(doc, text, kind='header'):
    """Set document header or footer text.
    Iterates over **all** document sections.  For each section the default
    header/footer is updated, and - when the section uses a different first-page
    header/footer - that variant is updated as well.
    Pre-existing paragraph formatting (alignment, style) from the template is
    preserved; only run content is replaced.
    Args:
        doc: The Word document.
        text: Content string.  Use ``{page}`` / ``{pages}`` for field tokens.
        kind: ``'header'`` or ``'footer'``.
    """
    _TOKEN_MAP = {'{page}': 'PAGE', '{pages}': 'NUMPAGES'}
    def _fill_paragraph(p, content):
        """Clear existing runs/fields and write *content* into paragraph *p*."""
        existing_alignment = p.alignment
        for child in list(p._p):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag in ('r', 'hyperlink', 'fldSimple'):
                p._p.remove(child)
        for part in _PAGE_TOKEN_RE.split(content):
            if part in _TOKEN_MAP:
                _add_field(p, _TOKEN_MAP[part])
            elif part:
                p.add_run(part)
        p.alignment = existing_alignment if existing_alignment is not None else WD_ALIGN_PARAGRAPH.CENTER
    def _update_part(section_part):
        """Update a single header or footer part."""
        section_part.is_linked_to_previous = False
        if section_part.paragraphs:
            _fill_paragraph(section_part.paragraphs[0], text)
        else:
            p = section_part.add_paragraph()
            _fill_paragraph(p, text)
    for section in doc.sections:
        _update_part(getattr(section, kind))
        if section.different_first_page_header_footer:
            first_kind = f'first_page_{kind}'
            first_part = getattr(section, first_kind, None)
            if first_part is not None:
                _update_part(first_part)
        even_kind = f'even_page_{kind}'
        even_part = getattr(section, even_kind, None)
        if even_part is not None and doc.settings.element.find(qn('w:evenAndOddHeaders')) is not None:
            _update_part(even_part)


# ---------------------------------------------------------------------------
# Table of Contents
# ---------------------------------------------------------------------------


def add_toc(doc):
    """Insert a Table of Contents field.
    The TOC is based on Heading styles 1-3 and will update when the document
    is opened in Word.
    """
    doc.add_heading('Table of Contents', level=1)
    p = doc.add_paragraph()
    run = p.add_run()
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'begin')
    run._r.append(fld)
    run = p.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    run._r.append(instr)
    run = p.add_run()
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'separate')
    run._r.append(fld)
    p.add_run('[Table of Contents - open in Word and press F9 to update]')
    run = p.add_run()
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'end')
    run._r.append(fld)
    doc.add_page_break()
    uf = OxmlElement('w:updateFields')
    uf.set(qn('w:val'), 'true')
    doc.settings.element.append(uf)


# ---------------------------------------------------------------------------
# Page size / orientation / margins
# ---------------------------------------------------------------------------

# Page dimensions in DXA (twentieths of a point / twips); 1440 = 1 inch.
_PAGE_SIZES_DXA = {
    'letter': (12240, 15840),
    'us-letter': (12240, 15840),
    'a4': (11906, 16838),
    'legal': (12240, 20160),
}

_LENGTH_RE = re.compile(r'^([\d.]+)\s*(in|cm|mm|pt)?$')


def _parse_length(token):
    """Parse a length token like ``1in``, ``2cm``, ``1.25in`` into a Length.

    Defaults to inches when no unit is given. Returns ``None`` on failure.
    """
    match = _LENGTH_RE.match((token or '').strip().lower())
    if not match:
        return None
    value = float(match.group(1))
    unit = match.group(2) or 'in'
    return {'in': Inches, 'cm': Cm, 'mm': Mm, 'pt': Pt}[unit](value)


def _apply_margin(section, margin_value):
    """Apply a ``margin`` directive value to *section*.

    Either a single length applied to all four sides (``1in``), or
    per-side assignments (``top=1in bottom=1in left=1.25in right=1.25in``).
    """
    margin_value = (margin_value or '').strip()
    if not margin_value:
        return
    if '=' not in margin_value:
        length = _parse_length(margin_value)
        if length is not None:
            section.top_margin = section.bottom_margin = length
            section.left_margin = section.right_margin = length
        return
    for token in margin_value.split():
        side, sep, raw = token.partition('=')
        if not sep:
            continue
        length = _parse_length(raw)
        if length is None:
            continue
        attr = f'{side.strip().lower()}_margin'
        if hasattr(section, attr):
            setattr(section, attr, length)
        else:
            logger.warning("Unknown margin side %r in directive value %r", side, margin_value)


def apply_page_setup(doc, directives=None, page_size=None, orientation=None, margin=None):
    """Apply page size, orientation, and margins to every section.

    *page_size*/*orientation*/*margin* are tool-parameter overrides; they win
    over the corresponding ``page``/``orientation``/``margin`` markdown
    directives when both are given. A no-op (leaves the template/python-docx
    default — normally A4 portrait with 1" margins) when none of these are
    present.
    """
    directives = directives or {}
    size_key = (page_size or directives.get('page') or '').strip().lower()
    orient_key = (orientation or directives.get('orientation') or '').strip().lower()
    margin_value = margin or directives.get('margin')
    dims = _PAGE_SIZES_DXA.get(size_key)
    landscape = orient_key == 'landscape'

    if size_key and dims is None:
        logger.warning("Unknown 'page' directive value %r; ignoring.", size_key)
    if not dims and not landscape and not margin_value:
        return

    for section in doc.sections:
        if dims is not None:
            width, height = dims
            if landscape:
                width, height = height, width
            section.page_width = Twips(width)
            section.page_height = Twips(height)
        elif landscape and section.page_width < section.page_height:
            section.page_width, section.page_height = section.page_height, section.page_width
        section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
        if margin_value:
            _apply_margin(section, margin_value)


# ---------------------------------------------------------------------------
# Default font / heading font override
# ---------------------------------------------------------------------------

def apply_default_font(doc, directives=None, default_font=None, template_in_use=False):
    """Apply a default body font and/or heading font override.

    *default_font* (tool parameter) wins over the ``font`` directive.
    ``font-size`` (points) and ``heading-font`` are directive-only.  When a
    Word template is in use, its own styles may still win over these settings
    in Word's own style-resolution order; a warning is logged so this is not
    silently confusing.
    """
    directives = directives or {}
    font_name = default_font or directives.get('font')
    font_size = directives.get('font-size')
    heading_font = directives.get('heading-font') or font_name
    if not font_name and not font_size and not heading_font:
        return
    if template_in_use and (font_name or font_size):
        logger.warning(
            "A Word template is in use together with a 'font'/'font-size' "
            "directive; the template's own styles may still take precedence "
            "for content that references them by name."
        )
    if font_name or font_size:
        try:
            normal = doc.styles['Normal']
        except KeyError:
            normal = None
        if normal is not None:
            if font_name:
                normal.font.name = font_name
            if font_size:
                try:
                    normal.font.size = Pt(float(font_size))
                except ValueError:
                    logger.warning("Invalid 'font-size' directive value: %r", font_size)
    if heading_font:
        for level in range(1, 7):
            try:
                style = doc.styles[f'Heading {level}']
            except KeyError:
                continue
            style.font.name = heading_font


# ---------------------------------------------------------------------------
# Multi-column layout
# ---------------------------------------------------------------------------

def apply_columns(doc, directives=None):
    """Apply a multi-column section layout from the ``columns`` directive.

    Recognised values:
      - ``2`` / ``3`` — that many equal-width columns.
      - ``2 sep`` — adds a vertical separator line between columns.
      - ``custom:5400,3240`` — explicit column widths in DXA (equalWidth=false).

    Applies to the last section only. Documents produced by this renderer are
    single-section, so column *breaks* mid-document (``SectionType.NEXT_COLUMN``
    in docx-js terms) are out of scope for v1 — the whole body flows through
    the configured columns.
    """
    value = ((directives or {}).get('columns') or '').strip()
    if not value:
        return
    section = doc.sections[-1]
    sectPr = section._sectPr
    existing = sectPr.find(qn('w:cols'))
    if existing is not None:
        sectPr.remove(existing)
    cols = OxmlElement('w:cols')
    if value.lower().startswith('custom:'):
        widths = [w.strip() for w in value.split(':', 1)[1].split(',') if w.strip()]
        if not widths:
            logger.warning("Invalid 'columns' directive value: %r", value)
            return
        cols.set(qn('w:equalWidth'), '0')
        cols.set(qn('w:num'), str(len(widths)))
        for raw_width in widths:
            try:
                width_dxa = int(float(raw_width))
            except ValueError:
                logger.warning("Invalid column width %r in 'columns' directive", raw_width)
                continue
            col = OxmlElement('w:col')
            col.set(qn('w:w'), str(width_dxa))
            cols.append(col)
    else:
        parts = value.split()
        try:
            num = int(parts[0])
        except (ValueError, IndexError):
            logger.warning("Invalid 'columns' directive value: %r", value)
            return
        cols.set(qn('w:num'), str(num))
        cols.set(qn('w:space'), '720')
        if 'sep' in [p.lower() for p in parts[1:]]:
            cols.set(qn('w:sep'), '1')
    sectPr.append(cols)


# ---------------------------------------------------------------------------
# Tab stops
# ---------------------------------------------------------------------------

_TAB_ALIGNMENTS = {'left', 'right', 'center', 'decimal'}
_TAB_LEADERS = {'dot', 'hyphen', 'underscore', 'none'}
_DEFAULT_CONTENT_WIDTH_TWIPS = 9360  # 6.5" usable width (US Letter, 1" margins)


def apply_tab_stop(paragraph, spec):
    """Configure an explicit right-margin tab stop from a ``tab`` directive.

    *spec* is e.g. ``right``, ``right dot``, ``center hyphen``. The stop is
    placed at the paragraph's section's content-width (right margin) — this
    covers the common "label ... value" / dot-leader TOC-style pattern the
    skill documents; arbitrary tab positions are out of scope for v1.
    """
    parts = (spec or '').strip().lower().split()
    if not parts:
        return
    align = parts[0] if parts[0] in _TAB_ALIGNMENTS else 'right'
    leader = next((p for p in parts[1:] if p in _TAB_LEADERS), None)
    try:
        section = paragraph.part.document.sections[-1]
        content_width_twips = int(
            (section.page_width - section.left_margin - section.right_margin) / 635
        )
    except Exception:
        content_width_twips = _DEFAULT_CONTENT_WIDTH_TWIPS
    pPr = paragraph._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), align)
    if leader and leader != 'none':
        tab.set(qn('w:leader'), leader)
    tab.set(qn('w:pos'), str(content_width_twips))
    tabs.append(tab)
    pPr.append(tabs)


# ---------------------------------------------------------------------------
# Bookmarks
# ---------------------------------------------------------------------------

_bookmark_id_counter = itertools.count(1)


def wrap_with_bookmark(doc, elements, name):
    """Wrap a run of already-inserted body *elements* with a named bookmark.

    *elements* must still be attached to the document body (this is only
    called from the non-``return_elements`` render path — see the
    ``<!-- bookmark: name -->`` handling in ``markdown_processor.py``; the
    template-placeholder path, which detaches elements for reinsertion
    elsewhere, does not support bookmarks in v1).
    """
    if not elements or not name:
        return
    bookmark_id = str(next(_bookmark_id_counter))
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), bookmark_id)
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), bookmark_id)
    try:
        elements[0].addprevious(start)
        elements[-1].addnext(end)
    except Exception:
        logger.warning("Failed to attach bookmark %r; element may be detached.", name)


# ---------------------------------------------------------------------------
# Footnotes
# ---------------------------------------------------------------------------

_FOOTNOTES_XML_TEMPLATE = (
    b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b'<w:footnote w:type="separator" w:id="-1">'
    b'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    b'<w:r><w:separator/></w:r></w:p></w:footnote>'
    b'<w:footnote w:type="continuationSeparator" w:id="0">'
    b'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    b'<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    b'</w:footnotes>'
)


def _get_or_create_footnotes_part(document_part):
    """Ensure the package has a ``word/footnotes.xml`` part; return its root element.

    python-docx (as of 1.2.0) has no built-in footnote API — unlike
    ``NumberingPart``/``CommentsPart`` there is no ``FootnotesPart`` class — so
    the part is created directly via the same low-level ``opc`` machinery
    python-docx itself uses for those parts. IDs ``0`` and ``-1`` are reserved
    by the OOXML spec for the separator/continuation-separator footnotes;
    user-visible footnotes start at id ``1``.
    """
    try:
        return document_part.part_related_by(RT.FOOTNOTES).element
    except KeyError:
        pass
    partname = PackURI('/word/footnotes.xml')
    element = parse_xml(_FOOTNOTES_XML_TEMPLATE)
    part = XmlPart(partname, CT.WML_FOOTNOTES, element, document_part.package)
    document_part.relate_to(part, RT.FOOTNOTES)
    return element


def add_footnote(document_part, text):
    """Append a new footnote containing *text* to the footnotes part.

    Returns the new footnote's numeric id (for the matching
    ``w:footnoteReference``).
    """
    footnotes_el = _get_or_create_footnotes_part(document_part)
    existing_ids = []
    for fn in footnotes_el.findall(qn('w:footnote')):
        raw_id = fn.get(qn('w:id'))
        try:
            existing_ids.append(int(raw_id))
        except (TypeError, ValueError):
            continue
    next_id = max([i for i in existing_ids if i > 0], default=0) + 1

    footnote = OxmlElement('w:footnote')
    footnote.set(qn('w:id'), str(next_id))
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'FootnoteText')
    pPr.append(pStyle)
    p.append(pPr)

    ref_run = OxmlElement('w:r')
    ref_rPr = OxmlElement('w:rPr')
    ref_style = OxmlElement('w:rStyle')
    ref_style.set(qn('w:val'), 'FootnoteReference')
    ref_rPr.append(ref_style)
    ref_run.append(ref_rPr)
    ref_run.append(OxmlElement('w:footnoteRef'))
    p.append(ref_run)

    text_run = OxmlElement('w:r')
    text_elem = OxmlElement('w:t')
    text_elem.set(qn('xml:space'), 'preserve')
    text_elem.text = ' ' + text
    text_run.append(text_elem)
    p.append(text_run)

    footnote.append(p)
    footnotes_el.append(footnote)
    return next_id


def add_footnote_reference_run(paragraph, footnote_id):
    """Append a superscript footnote-reference run to *paragraph*."""
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'FootnoteReference')
    rPr.append(rStyle)
    # Explicit superscript as a fallback for templates lacking a
    # 'FootnoteReference' character style — Word auto-numbers the
    # w:footnoteRef/w:footnoteReference regardless of style resolution.
    vert_align = OxmlElement('w:vertAlign')
    vert_align.set(qn('w:val'), 'superscript')
    rPr.append(vert_align)
    run.append(rPr)
    ref = OxmlElement('w:footnoteReference')
    ref.set(qn('w:id'), str(footnote_id))
    run.append(ref)
    paragraph._p.append(run)


def insert_footnote(paragraph, text):
    """Create a footnote with *text* and append its reference run to *paragraph*.

    Convenience wrapper combining :func:`add_footnote` and
    :func:`add_footnote_reference_run`, resolving the owning document part
    from *paragraph* itself (works from any content object with a ``.part``).
    """
    document_part = paragraph.part
    footnote_id = add_footnote(document_part, text)
    add_footnote_reference_run(paragraph, footnote_id)
