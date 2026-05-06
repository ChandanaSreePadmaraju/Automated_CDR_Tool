"""
post_processor.py
-----------------
Post-processing passes applied to the filled output document before saving.
All behaviour is driven by the "post_processing" block in prompts.json ΓÇö
nothing is hardcoded in this module.
"""

import json
import os
import re
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from src import detect_compliance_heading

_PROMPTS_FILE = os.path.join(os.path.dirname(__file__), "..", "prompts.json")
with open(_PROMPTS_FILE, "r", encoding="utf-8") as _f:
    _CFG: dict = json.load(_f).get("post_processing", {})

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _para_style_id(elem) -> str:
    pPr = elem.find(f"{{{_NS}}}pPr")
    if pPr is None:
        return ""
    pStyle = pPr.find(f"{{{_NS}}}pStyle")
    if pStyle is None:
        return ""
    return pStyle.get(f"{{{_NS}}}val", "")


def _para_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(f"{{{_NS}}}t")).strip()


def _heading_style_ids(doc: Document) -> set:
    return {s.style_id for s in doc.styles if s.name and s.name.startswith("Heading")}


# ---------------------------------------------------------------------------
# Pass 3 ΓÇö Remove paragraphs with specific styles (e.g. "Guidance")
# ---------------------------------------------------------------------------

def remove_styled_paragraphs(doc: Document) -> None:
    """Remove all body-level paragraphs whose style name is in the configured list."""
    targets: list[str] = _CFG.get("styles_to_remove", [])
    if not targets:
        return

    id_to_name = {s.style_id: (s.name or "") for s in doc.styles}
    body = doc.element.body
    to_remove = [
        e for e in list(body)
        if e.tag == f"{{{_NS}}}p" and id_to_name.get(_para_style_id(e), "") in targets
    ]
    for e in to_remove:
        body.remove(e)


# ---------------------------------------------------------------------------
# Pass 4 ΓÇö Remove template instruction bracket markers  (< / >)
# ---------------------------------------------------------------------------

def remove_template_instructions(doc: Document) -> None:
    """
    Remove template instruction blocks delimited by '<' ... '>' paragraphs.

    A block opens when a paragraph's text is exactly '<' or starts with '< '
    and closes when a paragraph's text is exactly '>' or starts with '> '.
    Every paragraph in the block (both markers inclusive) is removed.
    This handles multi-paragraph instruction notes such as:
        < note:
        The Product identification RX.Y ...
        >
    """
    if not _CFG.get("remove_template_instructions", False):
        return

    body = doc.element.body
    to_remove = []
    in_block = False

    for e in list(body):
        if e.tag != f"{{{_NS}}}p":
            continue
        text = _para_text(e)
        if not in_block:
            if text == "<" or text.startswith("< "):
                in_block = True
                to_remove.append(e)
                # Single-line block: e.g. '< DELETE THIS >'; close immediately
                if text.endswith(">") and len(text) > 1:
                    in_block = False
        else:
            to_remove.append(e)
            if text == ">" or text.startswith("> "):
                in_block = False

    for e in to_remove:
        body.remove(e)


# ---------------------------------------------------------------------------
# Pass 7 ΓÇö Remove empty table rows
# ---------------------------------------------------------------------------

def remove_empty_table_rows(doc: Document) -> None:
    """Remove rows (skipping the header row) where every cell is empty."""
    if not _CFG.get("remove_empty_table_rows", False):
        return

    ns = _NS
    for tbl in doc.element.body.iter(f"{{{ns}}}tbl"):
        rows = tbl.findall(f"{{{ns}}}tr")
        if len(rows) <= 1:
            continue
        for tr in rows[1:]:
            cells = tr.findall(f".//{{{ns}}}tc")
            if cells and all(
                not "".join(t.text or "" for t in tc.iter(f"{{{ns}}}t")).strip()
                for tc in cells
            ):
                tbl.remove(tr)


# ---------------------------------------------------------------------------
# Pass 10 ΓÇö Set all text colour to black
# ---------------------------------------------------------------------------

def set_all_text_black(doc: Document) -> None:
    """Force all text to render as black.

    Strategy:
    1. Remove any explicit <w:color> element (was just deleting the override).
    2. Insert <w:color w:val="000000"/> so the run's colour is always explicit
       black regardless of what a named style (e.g. GuidanceChar) would apply.
    3. Remove character-style references (w:rStyle) whose style name contains
       'Guidance', so those styles can no longer override colour/font.
    """
    if not _CFG.get("set_all_text_black", False):
        return

    # Build set of rStyle ids whose style name contains any of the
    # styles_to_remove names ΓÇö these are character-style variants
    # (e.g. "GuidanceChar") that must also be stripped so they cannot
    # override the forced black colour.  Driven by prompts.json so no
    # style names are hardcoded here.
    char_style_prefixes: list[str] = [
        s.lower() for s in _CFG.get("styles_to_remove", [])
    ]
    guidance_rStyle_ids: set[str] = set()
    for s in doc.styles:
        if (
            s.name
            and s.type.name == "CHARACTER"
            and any(prefix in s.name.lower() for prefix in char_style_prefixes)
        ):
            guidance_rStyle_ids.add(s.style_id)

    for rPr in doc.element.body.iter(f"{{{_NS}}}rPr"):
        # Skip runs inside tblHeader rows ΓÇö let them keep their natural style colour
        parent_tr = next(
            (a for a in rPr.iterancestors(f"{{{_NS}}}tr")), None
        )
        if parent_tr is not None:
            trPr = parent_tr.find(f"{{{_NS}}}trPr")
            if trPr is not None and trPr.find(f"{{{_NS}}}tblHeader") is not None:
                continue

        # 1. Remove GuidanceChar rStyle
        rStyle = rPr.find(f"{{{_NS}}}rStyle")
        if rStyle is not None and rStyle.get(f"{{{_NS}}}val", "") in guidance_rStyle_ids:
            rPr.remove(rStyle)

        # 2. Remove existing color element (might be non-black)
        color = rPr.find(f"{{{_NS}}}color")
        if color is not None:
            rPr.remove(color)

        # 3. Insert explicit black color as first child so it takes effect
        black = etree.Element(f"{{{_NS}}}color")
        black.set(f"{{{_NS}}}val", "000000")
        rPr.insert(0, black)

    # Also apply to headers and footers ΓÇö they are separate XML parts not in body
    for section in doc.sections:
        for part in (
            section.header, section.even_page_header, section.first_page_header,
            section.footer, section.even_page_footer, section.first_page_footer,
        ):
            try:
                for rPr in part._element.iter(f"{{{_NS}}}rPr"):
                    rStyle = rPr.find(f"{{{_NS}}}rStyle")
                    if rStyle is not None and rStyle.get(f"{{{_NS}}}val", "") in guidance_rStyle_ids:
                        rPr.remove(rStyle)
                    color = rPr.find(f"{{{_NS}}}color")
                    if color is not None:
                        rPr.remove(color)
                    black = etree.Element(f"{{{_NS}}}color")
                    black.set(f"{{{_NS}}}val", "000000")
                    rPr.insert(0, black)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Pass 13 ΓÇö Sort tables alphabetically under specified headings
# ---------------------------------------------------------------------------

def sort_tables_alphabetically(doc: Document) -> None:
    """
    For each heading listed in CFG["sort_table_alphabetically_under_headings"],
    find the first table that follows it and sort its rows alphabetically by
    the first column (header row is kept in place).
    """
    targets: list[str] = _CFG.get("sort_table_alphabetically_under_headings", [])
    if not targets:
        return

    h_ids = _heading_style_ids(doc)
    body = list(doc.element.body)
    targets_lower = [t.strip().lower() for t in targets]

    for i, elem in enumerate(body):
        if elem.tag != f"{{{_NS}}}p" or _para_style_id(elem) not in h_ids:
            continue
        if _para_text(elem).strip().lower() not in targets_lower:
            continue

        # Find first <w:tbl> after this heading (before next heading)
        for nxt in body[i + 1:]:
            if nxt.tag == f"{{{_NS}}}p" and _para_style_id(nxt) in h_ids:
                break
            if nxt.tag == f"{{{_NS}}}tbl":
                _sort_table(nxt)
                break


def _sort_table(tbl_elem) -> None:
    rows = tbl_elem.findall(f"{{{_NS}}}tr")
    if len(rows) <= 1:
        return

    header, *data_rows = rows

    def first_col_key(tr):
        tc = tr.find(f"{{{_NS}}}tc")
        if tc is None:
            return ""
        return "".join(t.text or "" for t in tc.iter(f"{{{_NS}}}t")).strip().lower()

    data_rows.sort(key=first_col_key)

    for r in rows:
        tbl_elem.remove(r)
    tbl_elem.append(header)
    for r in data_rows:
        tbl_elem.append(r)


# ---------------------------------------------------------------------------
# Pass 8 ΓÇö Reduce font size for NOTE paragraphs
# ---------------------------------------------------------------------------

def set_note_text_size(doc: Document) -> None:
    """Reduce the font size of NOTE paragraphs and all following paragraphs
    that belong to the same NOTE block (until the next NOTE prefix, a heading,
    or an empty paragraph that is not a list item).

    Applies to body paragraphs only (not table cells ΓÇö notes rarely appear
    inside tables).
    Prefixes and target size are read from prompts.json:
      note_text_prefixes   : list[str]  ΓÇô e.g. ["NOTE"]
      note_text_size_half_pt: int       ΓÇô half-points (16 = 8 pt)
    """
    prefixes  = [p.lower() for p in _CFG.get("note_text_prefixes", [])]
    size_val  = str(_CFG.get("note_text_size_half_pt", 18))
    if not prefixes:
        return

    W = f"{{{_NS}}}"
    h_ids = _heading_style_ids(doc)

    def _apply_size(p_elem) -> None:
        for r in p_elem.findall(f".//{W}r"):
            rPr = r.find(f"{W}rPr")
            if rPr is None:
                rPr = etree.Element(f"{W}rPr")
                r.insert(0, rPr)
            for tag in (f"{W}sz", f"{W}szCs"):
                for old in rPr.findall(tag):
                    rPr.remove(old)
            sz = etree.SubElement(rPr, f"{W}sz")
            sz.set(f"{W}val", size_val)
            szCs = etree.SubElement(rPr, f"{W}szCs")
            szCs.set(f"{W}val", size_val)

    # Process both body paragraphs and paragraphs inside table cells.
    # For body paragraphs: track NOTE blocks across consecutive paragraphs.
    # For table-cell paragraphs: apply size to the NOTE paragraph and all
    # following paragraphs in the same cell until an empty paragraph or heading.

    # --- Body-level paragraphs (block tracking) ---
    body_paras = [
        e for e in doc.element.body
        if e.tag == f"{W}p"
    ]

    in_note = False
    for p_elem in body_paras:
        p_text = "".join(t.text or "" for t in p_elem.iter(f"{W}t")).strip()
        style_id = _para_style_id(p_elem)

        if style_id in h_ids:
            in_note = False
            continue

        if any(p_text.lower().startswith(prefix) for prefix in prefixes):
            in_note = True

        if in_note:
            _apply_size(p_elem)
            if not p_text and style_id not in ("ListParagraph",):
                in_note = False

    # --- Table cell paragraphs ---
    for tbl in doc.element.body.iter(f"{W}tbl"):
        for tc in tbl.iter(f"{W}tc"):
            cell_paras = list(tc.findall(f"{W}p"))
            in_note = False
            for p_elem in cell_paras:
                p_text = "".join(t.text or "" for t in p_elem.iter(f"{W}t")).strip()
                style_id = _para_style_id(p_elem)

                if style_id in h_ids:
                    in_note = False
                    continue

                if any(p_text.lower().startswith(prefix) for prefix in prefixes):
                    in_note = True

                if in_note:
                    _apply_size(p_elem)
                    if not p_text and style_id not in ("ListParagraph",):
                        in_note = False


# ---------------------------------------------------------------------------
# Pass 9 ΓÇö Strip superscript formatting from list-marker runs
# ---------------------------------------------------------------------------

_MARKER_RE = re.compile(r'^\d+[.\s]*$')


def strip_superscript_list_markers(doc: Document) -> None:
    """Remove <w:vertAlign val="superscript"/> only from LEADING numeric runs.

    A run is considered a leading list marker if:
      - Its text matches the bare-number pattern (e.g. "1", "1.", "2."), AND
      - All runs that appear before it in the same paragraph have no visible
        text (i.e. the superscript is the very first visible content).

    This preserves legitimate trailing footnote-reference superscripts (e.g.
    a small "1" appended after "Instructions for Use Azurion R3.0") while
    removing numbered-list counters that Word copied as superscripts.

    Controlled by prompts.json: "strip_superscript_list_markers": true
    """
    if not _CFG.get("strip_superscript_list_markers", False):
        return

    W = f"{{{_NS}}}"
    for p_elem in doc.element.body.iter(f"{W}p"):
        runs = p_elem.findall(f".//{W}r")
        for idx, r in enumerate(runs):
            rPr = r.find(f"{W}rPr")
            if rPr is None:
                continue
            vert = rPr.find(f"{W}vertAlign")
            if vert is None or vert.get(f"{W}val") != "superscript":
                continue
            t = r.find(f"{W}t")
            text = (t.text or "").strip() if t is not None else ""
            if not _MARKER_RE.match(text):
                continue
            # Only strip if no visible text exists in any earlier run
            preceding_text = "".join(
                (rt.text or "")
                for prev_r in runs[:idx]
                for rt in prev_r.findall(f"{W}t")
            ).strip()
            if not preceding_text:
                rPr.remove(vert)


# ---------------------------------------------------------------------------
# Pass 14 ΓÇö Remove specified sections (heading + content)
# ---------------------------------------------------------------------------

def remove_sections(doc: Document) -> None:
    """
    Remove each heading listed in CFG["remove_sections"] together with all
    body elements that follow it (up to the next heading).
    """
    targets: list[str] = _CFG.get("remove_sections", [])
    if not targets:
        return

    h_ids = _heading_style_ids(doc)
    targets_lower = [t.strip().lower() for t in targets]
    ns = _NS
    body = doc.element.body

    for target in targets_lower:
        children = list(body)
        for i, elem in enumerate(children):
            if elem.tag == f"{{{ns}}}p" and _para_style_id(elem) in h_ids:
                if _para_text(elem).strip().lower() == target:
                    to_remove = [elem]
                    for nxt in children[i + 1:]:
                        if nxt.tag == f"{{{ns}}}p" and _para_style_id(nxt) in h_ids:
                            break
                        # Never delete layout paragraphs (inline sectPr) ΓÇö
                        # they carry the template's header/footer rId links.
                        if nxt.tag == f"{{{ns}}}p":
                            pPr = nxt.find(f"{{{ns}}}pPr")
                            if pPr is not None and pPr.find(f"{{{ns}}}sectPr") is not None:
                                continue
                        to_remove.append(nxt)
                    for e in to_remove:
                        body.remove(e)
                    break


def remove_headings_only(doc: Document) -> None:
    """
    Remove each heading listed in CFG["remove_headings_only"] together with
    any immediately following non-table paragraphs, but STOP (and keep) the
    first table or next heading encountered.  Use this instead of
    remove_sections when the heading's content (a table) must be preserved.
    """
    targets: list[str] = _CFG.get("remove_headings_only", [])
    if not targets:
        return

    h_ids = _heading_style_ids(doc)
    targets_lower = [t.strip().lower() for t in targets]
    ns = _NS
    body = doc.element.body

    for target in targets_lower:
        children = list(body)
        for i, elem in enumerate(children):
            if elem.tag == f"{{{ns}}}p" and _para_style_id(elem) in h_ids:
                if _para_text(elem).strip().lower() == target:
                    to_remove = [elem]
                    for nxt in children[i + 1:]:
                        # Stop at any table ΓÇö keep it
                        if nxt.tag == f"{{{ns}}}tbl":
                            break
                        # Stop at next heading ΓÇö keep it
                        if nxt.tag == f"{{{ns}}}p" and _para_style_id(nxt) in h_ids:
                            break
                        # Skip layout paragraphs
                        if nxt.tag == f"{{{ns}}}p":
                            pPr = nxt.find(f"{{{ns}}}pPr")
                            if pPr is not None and pPr.find(f"{{{ns}}}sectPr") is not None:
                                continue
                        to_remove.append(nxt)
                    for e in to_remove:
                        body.remove(e)
                    break


# Pass 11a ΓÇö Remove blank paragraphs immediately BEFORE any heading
def remove_blank_paragraph_before_headings(doc: Document) -> None:
    """Remove ALL consecutive empty paragraphs that appear immediately before a Heading.

    Keeps layout paragraphs (page-breaks / inline sectPr) and paragraphs
    with images intact.
    """
    if not _CFG.get("remove_blank_para_before_headings", False):
        return

    def _is_blank(elem) -> bool:
        if elem.tag != f"{{{_NS}}}p":
            return False
        pPr = elem.find(f"{{{_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{_NS}}}sectPr") is not None:
            return False  # layout paragraph ΓÇö keep
        for br in elem.iter(f"{{{_NS}}}br"):
            if br.get(f"{{{_NS}}}type") == "page":
                return False  # page-break ΓÇö keep
        text = ''.join(t.text or '' for t in elem.iter(f"{{{_NS}}}t")).strip()
        has_drawing = any(n.tag == qn('w:drawing') for n in elem.iter())
        return not text and not has_drawing

    h_ids = _heading_style_ids(doc)
    body_elem = doc.element.body

    changed = True
    while changed:
        changed = False
        children = list(body_elem)
        for i, elem in enumerate(children):
            if elem.tag != f"{{{_NS}}}p":
                continue
            if _para_style_id(elem) not in h_ids:
                continue
            # Remove all consecutive blank paragraphs directly before this heading
            j = i - 1
            while j >= 0 and _is_blank(children[j]):
                try:
                    body_elem.remove(children[j])
                    changed = True
                except Exception:
                    pass
                j -= 1


# Pass 11b
def remove_blank_paragraph_after_headings(doc: Document) -> None:
    """Remove ALL consecutive empty paragraphs that immediately follow a Heading.

    Keeps layout paragraphs (page-breaks / inline sectPr) and paragraphs
    with images intact.
    """
    if not _CFG.get("remove_blank_para_after_headings", False):
        return

    def _is_blank(elem) -> bool:
        if elem.tag != f"{{{_NS}}}p":
            return False
        pPr = elem.find(f"{{{_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{_NS}}}sectPr") is not None:
            return False  # layout paragraph ΓÇö keep
        for br in elem.iter(f"{{{_NS}}}br"):
            if br.get(f"{{{_NS}}}type") == "page":
                return False  # page-break ΓÇö keep
        text = ''.join(t.text or '' for t in elem.iter(f"{{{_NS}}}t")).strip()
        has_drawing = any(n.tag == qn('w:drawing') for n in elem.iter())
        return not text and not has_drawing

    h_ids = _heading_style_ids(doc)
    body_elem = doc.element.body

    # Iterate live over children; re-read list after each removal so indices
    # stay correct even though we're removing elements.
    changed = True
    while changed:
        changed = False
        children = list(body_elem)
        for i, elem in enumerate(children):
            if elem.tag != f"{{{_NS}}}p":
                continue
            if _para_style_id(elem) not in h_ids:
                continue
            # Remove all consecutive blank paragraphs directly after this heading
            j = i + 1
            while j < len(children) and _is_blank(children[j]):
                try:
                    body_elem.remove(children[j])
                    changed = True
                except Exception:
                    pass
                j += 1


# ---------------------------------------------------------------------------
# Pass 1 ΓÇö Fill placeholders throughout document
# ---------------------------------------------------------------------------

def fill_header_footer_placeholders(doc: Document, product_name: str | None) -> None:
    """
    Replace every placeholder listed in CFG["header_footer_placeholders"]
    with *product_name* throughout the entire document: body paragraphs,
    body tables (including cover page), and all headers/footers.

    Uses lxml directly so it works on both native python-docx elements and
    raw XML tables inserted at the zip level. Also handles placeholders that
    Word has split across multiple <w:r> runs.
    """
    if not product_name:
        return

    placeholders: list[str] = _CFG.get("header_footer_placeholders", [])
    if not placeholders:
        return

    def _replace_in_para_elem(p_elem) -> None:
        run_elems = p_elem.findall('.//' + qn('w:r'))
        if not run_elems:
            return
        for placeholder in placeholders:
            while True:
                texts = [''.join(t.text or '' for t in r.findall('.//' + qn('w:t')))
                         for r in run_elems]
                full = ''.join(texts)
                if placeholder not in full:
                    break
                ph_start = full.index(placeholder)
                ph_end   = ph_start + len(placeholder)
                offsets = [0]
                for t in texts[:-1]:
                    offsets.append(offsets[-1] + len(t))
                first_ri = next(i for i, off in enumerate(offsets)
                                if off + len(texts[i]) > ph_start)
                last_ri  = next(i for i in range(len(run_elems) - 1, -1, -1)
                                if offsets[i] < ph_end)
                pre  = texts[first_ri][:ph_start - offsets[first_ri]]
                post = texts[last_ri][ph_end - offsets[last_ri]:]
                new_text = pre + product_name + post
                # Write into first run's <w:t>
                wt_list = run_elems[first_ri].findall('.//' + qn('w:t'))
                if wt_list:
                    wt_list[0].text = new_text
                    for wt in wt_list[1:]:
                        wt.getparent().remove(wt)
                else:
                    wt = etree.SubElement(run_elems[first_ri], qn('w:t'))
                    wt.text = new_text
                # Clear text in remaining involved runs
                for ri in range(first_ri + 1, last_ri + 1):
                    for wt in run_elems[ri].findall('.//' + qn('w:t')):
                        wt.text = ''

    def _replace_in_xml_elem(container_elem) -> None:
        """Iterate every <w:p> inside *container_elem* and replace placeholders."""
        for p_elem in container_elem.iter(qn('w:p')):
            _replace_in_para_elem(p_elem)

    # Document body
    _replace_in_xml_elem(doc.element.body)

    # Headers and footers
    for section in doc.sections:
        for part in (
            section.header, section.even_page_header, section.first_page_header,
            section.footer, section.even_page_footer, section.first_page_footer,
        ):
            try:
                _replace_in_xml_elem(part._element)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Public API ΓÇö passes 2, 5, 6, 11, 12, 15
# ---------------------------------------------------------------------------

# Pass 2
def remove_preamble_before_first_heading(doc: Document) -> None:
    """
    Remove visible-text paragraphs and duplicate page-break paragraphs that
    appear BEFORE the first heading (the template instruction page).
    Tables (cover page) are always preserved.
    Only the LAST page-break paragraph before the first heading is kept so
    that the first content section still starts on a fresh page without an
    extra blank page between the cover and the first heading.
    """
    if not _CFG.get("remove_preamble_before_first_heading", False):
        return

    h_ids = _heading_style_ids(doc)
    body = doc.element.body

    def _has_page_break(p_elem) -> bool:
        for br in p_elem.iter(f"{{{_NS}}}br"):
            if br.get(f"{{{_NS}}}type") == "page":
                return True
        pPr = p_elem.find(f"{{{_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{_NS}}}sectPr") is not None:
            return True
        return False

    # Collect all paragraph elements before the first heading
    pre_heading_paras: list = []
    for elem in list(body):
        if elem.tag == f"{{{_NS}}}p" and _para_style_id(elem) in h_ids:
            break
        if elem.tag == f"{{{_NS}}}p":
            pre_heading_paras.append(elem)

    # Find the FIRST and LAST page-break paragraphs.
    # The template has: COVER ΓåÆ [break1] ΓåÆ instruction page ΓåÆ [break2] ΓåÆ TOC ΓåÆ [break3] ΓåÆ content.
    # After removing the instruction page, we want to keep break1 (coverΓåÆTOC) and
    # break3 (TOCΓåÆcontent). Any intermediate page-break paragraphs are removed.
    pb_indices = [i for i, e in enumerate(pre_heading_paras) if _has_page_break(e)]
    keep_pb = set()
    if pb_indices:
        keep_pb.add(pb_indices[0])   # first: cover ΓåÆ TOC
        keep_pb.add(pb_indices[-1])  # last:  TOC  ΓåÆ content

    to_remove = []
    for i, elem in enumerate(pre_heading_paras):
        text = ''.join(t.text or '' for t in elem.iter(f"{{{_NS}}}t")).strip()
        is_pb = _has_page_break(elem)
        if i in keep_pb:
            continue  # keep first and last page breaks
        if text or is_pb:
            to_remove.append(elem)

    for e in to_remove:
        body.remove(e)


# Pass 5
def strip_inline_angle_brackets(doc: Document) -> None:
    """
    Remove standalone '<' and '>' runs from every paragraph (including inside
    table cells), and strip inline '<...>' comment substrings.

    Handles three cases:
    1. Standalone bracket runs: a run whose only text is '<' or '>' ΓÇö removed.
    2. Single-run inline comment: '<...>' fully within one run's text ΓÇö stripped.
    3. Multi-run inline comment: '<' in one run and '>' in a later run of the
       same paragraph ΓÇö all runs from the opening '<' to the closing '>' are
       cleaned so the bracketed span is removed.
    """
    if not _CFG.get("strip_inline_angle_brackets", False):
        return

    _inline_re = re.compile(r'<[^<>]+>')

    def _strip_para(p_elem) -> None:
        runs = list(p_elem.findall(f"{{{_NS}}}r"))

        # Pass A ΓÇö standalone bracket runs and single-run inline comments
        for r in runs:
            t_elems = r.findall(f"{{{_NS}}}t")
            run_text = "".join(t.text or "" for t in t_elems).strip()
            if run_text in ("<", ">"):
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
            elif _inline_re.search(run_text):
                cleaned = _inline_re.sub("", run_text).strip()
                if t_elems:
                    t_elems[0].text = cleaned
                    for t in t_elems[1:]:
                        t.text = ""

        # Pass B ΓÇö multi-run inline comments (re-read after Pass A removals)
        runs = list(p_elem.findall(f"{{{_NS}}}r"))
        open_run_idx = None
        for idx, r in enumerate(runs):
            run_text = "".join(t.text or "" for t in r.findall(f"{{{_NS}}}t"))
            if open_run_idx is None:
                lt_pos = run_text.find("<")
                if lt_pos != -1:
                    # Check the rest of this run ΓÇö if '>' is also here, no span needed
                    after = run_text[lt_pos:]
                    if ">" not in after:
                        open_run_idx = idx
                        # Truncate this run at the '<'
                        for t in r.findall(f"{{{_NS}}}t"):
                            t.text = (t.text or "")[:lt_pos] if lt_pos == 0 else run_text[:lt_pos]
                            break
            else:
                gt_pos = run_text.find(">")
                if gt_pos != -1:
                    # Keep text after '>'
                    remainder = run_text[gt_pos + 1:]
                    for t in r.findall(f"{{{_NS}}}t"):
                        t.text = remainder
                        break
                    # Remove runs strictly between open_run_idx and idx
                    for mid in runs[open_run_idx + 1 : idx]:
                        parent = mid.getparent()
                        if parent is not None:
                            parent.remove(mid)
                    open_run_idx = None
                else:
                    # Entirely inside comment span ΓÇö blank it out
                    for t in r.findall(f"{{{_NS}}}t"):
                        t.text = ""

    for p_elem in doc.element.body.iter(f"{{{_NS}}}p"):
        _strip_para(p_elem)


# Pass 6
def remove_paragraphs_with_text(doc: Document) -> None:
    """
    Remove any body-level paragraph whose text contains one of the
    strings listed in CFG["remove_paragraphs_with_text"].
    """
    targets: list[str] = _CFG.get("remove_paragraphs_with_text", [])
    if not targets:
        return

    body = doc.element.body
    to_remove = []
    for elem in list(body):
        if elem.tag != f"{{{_NS}}}p":
            continue
        text = _para_text(elem)
        if any(t.lower() in text.lower() for t in targets):
            to_remove.append(elem)
    for e in to_remove:
        body.remove(e)


# Pass 12
def inject_definitions_fixed_rows(doc: Document) -> None:
    """
    Ensure specific rows (e.g. CDR, ISO, IGT-S) always exist in the
    Definitions & abbreviations table, then sorting is handled by the
    existing sort_tables_alphabetically pass.
    The rows are defined in CFG["definitions_fixed_rows"] as
    [["abbrev", "definition"], ...].
    CFG["definitions_heading"] names the heading to look under.
    """
    fixed: list[list[str]] = _CFG.get("definitions_fixed_rows", [])
    heading_text: str       = _CFG.get("definitions_heading", "")
    if not fixed or not heading_text:
        return

    h_ids = _heading_style_ids(doc)
    body  = list(doc.element.body)

    # Find the definitions table
    tbl_elem = None
    for i, elem in enumerate(body):
        if (elem.tag == f"{{{_NS}}}p"
                and _para_style_id(elem) in h_ids
                and _para_text(elem).strip().lower() == heading_text.lower()):
            for nxt in body[i + 1:]:
                if nxt.tag == f"{{{_NS}}}p" and _para_style_id(nxt) in h_ids:
                    break
                if nxt.tag == f"{{{_NS}}}tbl":
                    tbl_elem = nxt
                    break
            break

    if tbl_elem is None:
        return

    rows = list(tbl_elem.findall(f"{{{_NS}}}tr"))
    if not rows:
        return

    # Collect existing first-column values (lower-cased) to detect duplicates
    existing_keys: set[str] = set()
    for tr in rows[1:]:  # skip header
        tc = tr.find(f"{{{_NS}}}tc")
        if tc is not None:
            existing_keys.add(
                "".join(t.text or "" for t in tc.iter(f"{{{_NS}}}t")).strip().lower()
            )

    # Use the last data row as a style template for new rows
    template_row = rows[-1]

    for abbrev, definition in fixed:
        if abbrev.lower() in existing_keys:
            continue  # already present

        # Clone the template row and overwrite cell texts
        new_row = deepcopy(template_row)
        cells   = new_row.findall(f"{{{_NS}}}tc")
        texts   = [abbrev, definition]
        for ci, cell in enumerate(cells[:len(texts)]):
            # Capture run properties (font/size/bold) from the first run
            # before clearing, so the new run inherits the same formatting.
            existing_runs = cell.findall(f".//{{{_NS}}}r")
            saved_rPr = None
            for r in existing_runs:
                rPr = r.find(f"{{{_NS}}}rPr")
                if rPr is not None:
                    saved_rPr = deepcopy(rPr)
                    break

            # Clear existing runs
            for r in existing_runs:
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)

            # Re-add a run with the same rPr so font matches the table
            for p_elem in cell.findall(f"{{{_NS}}}p"):
                r_elem = etree.SubElement(p_elem, f"{{{_NS}}}r")
                if saved_rPr is not None:
                    r_elem.insert(0, saved_rPr)
                t_elem = etree.SubElement(r_elem, f"{{{_NS}}}t")
                t_elem.text = texts[ci]
                t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                break

        tbl_elem.append(new_row)

# ---------------------------------------------------------------------------
# Pass 16b ΓÇö Prepend header row to compliance checklist table
# ---------------------------------------------------------------------------

def prepend_compliance_table_header(doc: Document, template_path: str | None = None) -> None:
    """
    Prepend grey header rows (tblHeader) to the compliance table so they
    repeat on every page, styled as a compact grey label:
      - Grey text (808080), 9 pt, no bold
      - Cell-level thin grey borders (overrides table black borders)
      - tblHeader flag so Word repeats them at the top of every page

    Layout produced (on a fresh page):
      [page break]
      [main compliance table ΓÇö tblHeader rows at top repeat every page]
        Row 0 [tblHeader]: spanning cell ΓÇö standard name, centred, grey 9pt
        Row 1 [tblHeader]: CI. | RequirementΓÇôTest | ResultΓÇôRefΓÇôRemark | Verdict, grey 9pt
        Row 2+: data rows unchanged
    """
    cfg: dict = _CFG.get("compliance_table_header", {})

    heading_text: str = cfg.get("heading", "")
    columns: list     = cfg.get("columns", [])
    spanning_text: str = cfg.get("spanning_header", "")

    W = f"{{{_NS}}}"
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    # ------------------------------------------------------------------
    # Auto-detect heading_text, columns, spanning_header from template.
    # All three can be left blank/empty in prompts.json ΓÇö the template
    # docx is the single source of truth.
    # ------------------------------------------------------------------
    if not heading_text and template_path and os.path.isfile(template_path):
        heading_text = detect_compliance_heading(template_path)

    if template_path and os.path.isfile(template_path):
        try:
            _td      = Document(template_path)
            _td_body = list(_td.element.body)
            _td_h_ids = _heading_style_ids(_td)

            # Collect all tables inside the compliance section
            if heading_text:
                _in_sec     = False
                _sec_tables: list = []
                for _e in _td_body:
                    if _e.tag == f"{W}p" and _para_style_id(_e) in _td_h_ids:
                        if _para_text(_e).strip().lower() == heading_text.lower():
                            _in_sec = True
                        elif _in_sec:
                            break
                    elif _in_sec and _e.tag == f"{W}tbl":
                        _sec_tables.append(_e)

                # spanning_text ΓÇö from "Standard:" cell in any section table
                if not spanning_text:
                    for _tbl in _sec_tables:
                        for _row in _tbl.findall(f"{W}tr"):
                            _cells = _row.findall(f"{W}tc")
                            if len(_cells) >= 2:
                                _lbl = "".join(
                                    t.text or "" for t in _cells[0].iter(f"{W}t")
                                ).strip().lower()
                                if "standard" in _lbl:
                                    _raw = "".join(
                                        t.text or "" for t in _cells[-1].iter(f"{W}t")
                                    ).strip()
                                    # Strip edition suffix e.g. "(Ed. 1.0)"
                                    spanning_text = re.sub(
                                        r'\s*\(Ed\..*?\)\s*$', '', _raw
                                    ).strip()
                                    break
                        if spanning_text:
                            break

                # columns ΓÇö read only from explicit tblHeader rows in the biggest
                # template table. Never infer from data rows to avoid false positives.
                if not columns:
                    _best    = 0
                    _big_tbl = None
                    for _tbl in _sec_tables:
                        _n = len(_tbl.findall(f"{W}tr"))
                        if _n > _best:
                            _best    = _n
                            _big_tbl = _tbl
                    if _big_tbl is not None:
                        for _r in _big_tbl.findall(f"{W}tr"):
                            _r_trPr = _r.find(f"{W}trPr")
                            if _r_trPr is None or _r_trPr.find(f"{W}tblHeader") is None:
                                continue
                            _r_cells = _r.findall(f"{W}tc")
                            # skip spanning rows (any cell with gridSpan)
                            if any(_c.find(f"{W}tcPr/{W}gridSpan") is not None
                                   for _c in _r_cells):
                                continue
                            if len(_r_cells) >= 2:
                                _col_texts = [
                                    "".join(t.text or "" for t in _c.iter(f"{W}t")).strip()
                                    for _c in _r_cells
                                ]
                                if any(_col_texts):
                                    columns = _col_texts
                                    break
        except Exception:
            pass

    # bail if we still can't locate the target section
    if not heading_text:
        return

    # ------------------------------------------------------------------
    # Locate the main compliance table (largest table under heading)
    # ------------------------------------------------------------------
    h_ids = _heading_style_ids(doc)
    body_children = list(doc.element.body)

    out_tbl = None
    best = 0
    for i, elem in enumerate(body_children):
        if elem.tag == f"{W}p" and _para_style_id(elem) in h_ids:
            if _para_text(elem).strip().lower() == heading_text.lower():
                for nxt in body_children[i + 1:]:
                    if nxt.tag == f"{W}p" and _para_style_id(nxt) in h_ids:
                        break
                    if nxt.tag == f"{W}tbl":
                        n = len(nxt.findall(f"{W}tr"))
                        if n > best:
                            best = n
                            out_tbl = nxt
                break

    if out_tbl is None:
        return

    out_rows = out_tbl.findall(f"{W}tr")
    if not out_rows:
        return

    # ------------------------------------------------------------------
    # Idempotency ΓÇö skip if tblHeader rows already present in table
    # ------------------------------------------------------------------
    _check_val = spanning_text.lower() if spanning_text else (columns[0].lower() if columns else "")
    first_trPr = out_rows[0].find(f"{W}trPr")
    if first_trPr is not None and first_trPr.find(f"{W}tblHeader") is not None:
        # tblHeader rows already inserted
        return

    # ------------------------------------------------------------------
    # Page break: compliance table must start at top of a new page
    # ------------------------------------------------------------------
    _doc_body = out_tbl.getparent()
    _body_list = list(_doc_body)
    _tbl_idx   = _body_list.index(out_tbl)
    _prev_el   = _body_list[_tbl_idx - 1] if _tbl_idx > 0 else None
    _existing_br = _prev_el.find(f".//{W}br") if _prev_el is not None else None
    _has_pg_break = (
        _existing_br is not None
        and _existing_br.get(f"{W}type", "") == "page"
    )
    if not _has_pg_break:
        _pg_p = etree.Element(f"{W}p")
        _pg_r = etree.SubElement(_pg_p, f"{W}r")
        _pg_br = etree.SubElement(_pg_r, f"{W}br")
        _pg_br.set(f"{W}type", "page")
        _doc_body.insert(_tbl_idx, _pg_p)

    # ------------------------------------------------------------------
    # Read column widths from the first full-width data row.
    # A "full-width" row is the first row where no cell has gridSpan
    # and the cell count matches len(columns) (if known).
    # ------------------------------------------------------------------
    # Determine true grid column count from tblGrid
    grid_cols = len(out_tbl.findall(f"{W}tblGrid/{W}gridCol"))
    if not grid_cols:
        # fallback: scan rows for max cell count (accounting for gridspan)
        for r in out_rows:
            _cnt = sum(
                int((c.find(f"{W}tcPr/{W}gridSpan") or type("", (), {"get": lambda s, k, d="1": d})()).get(f"{W}val", "1"))
                for c in r.findall(f"{W}tc")
            )
            if _cnt > grid_cols:
                grid_cols = _cnt

    # Use columns count if provided, else fall back to grid
    n_cols = len(columns) if columns else grid_cols

    out_widths: list[tuple[str, str]] = []
    for r in out_rows:
        cells = r.findall(f"{W}tc")
        if len(cells) == n_cols:
            for c in cells:
                tcW = c.find(f"{W}tcPr/{W}tcW")
                out_widths.append((
                    tcW.get(f"{W}w", "0") if tcW is not None else "0",
                    tcW.get(f"{W}type", "dxa") if tcW is not None else "dxa",
                ))
            break
    while len(out_widths) < n_cols:
        out_widths.append(("0", "dxa"))

    try:
        total_w = sum(int(wv) for wv, _ in out_widths)
    except (ValueError, TypeError):
        total_w = 0
    w_type = out_widths[0][1] if out_widths else "dxa"

    # ------------------------------------------------------------------
    # Styling — sniff from the output table's first data row so headers
    # automatically match the template's own font, size, color, bold.
    # prompts.json style block can override individual values.
    # ------------------------------------------------------------------
    _style      = cfg.get("style", {})
    span_height = int(_style.get("span_row_height_twips", 480))
    col_height  = int(_style.get("col_row_height_twips",  300))

    # Sniff rPr from the first data cell in the output table
    _sniff_sz   = "22"   # 11 pt default
    _sniff_bold = True
    for _sr in out_rows:
        _sc = _sr.findall(f"{W}tc")
        if not _sc:
            continue
        _sp = _sc[0].find(f"{W}p")
        if _sp is None:
            continue
        _srPr = _sp.find(f"{W}pPr/{W}rPr")
        if _srPr is None:
            _srPr = _sp.find(f"{W}r/{W}rPr")
        if _srPr is not None:
            _sz_e = _srPr.find(f"{W}sz")
            if _sz_e is not None:
                _sniff_sz = _sz_e.get(f"{W}val", _sniff_sz)
            _sniff_bold = _srPr.find(f"{W}b") is not None
        break

    # Allow prompts.json overrides
    FONT_SZ     = str(int(_style.get("font_size_halfpt", int(_sniff_sz))))
    USE_BOLD    = _style.get("header_bold", _sniff_bold)

    def _make_hdr_cell(w_val, w_type_val, text, center=False, gridspan=None):
        """Build a tblHeader cell styled to match the table's own data rows."""
        tc = etree.Element(f"{W}tc")
        tcPr = etree.SubElement(tc, f"{W}tcPr")
        tcW_e = etree.SubElement(tcPr, f"{W}tcW")
        tcW_e.set(f"{W}w",    w_val)
        tcW_e.set(f"{W}type", w_type_val)
        if gridspan:
            gs = etree.SubElement(tcPr, f"{W}gridSpan")
            gs.set(f"{W}val", str(gridspan))

        p = etree.SubElement(tc, f"{W}p")
        pPr = etree.SubElement(p, f"{W}pPr")
        if center:
            jc = etree.SubElement(pPr, f"{W}jc")
            jc.set(f"{W}val", "center")
        pRpr = etree.SubElement(pPr, f"{W}rPr")
        if USE_BOLD:
            etree.SubElement(pRpr, f"{W}b")
            etree.SubElement(pRpr, f"{W}bCs")
        etree.SubElement(pRpr, f"{W}sz").set(f"{W}val", FONT_SZ)
        etree.SubElement(pRpr, f"{W}szCs").set(f"{W}val", FONT_SZ)

        r = etree.SubElement(p, f"{W}r")
        rPr = etree.SubElement(r, f"{W}rPr")
        if USE_BOLD:
            etree.SubElement(rPr, f"{W}b")
            etree.SubElement(rPr, f"{W}bCs")
        etree.SubElement(rPr, f"{W}sz").set(f"{W}val", FONT_SZ)
        etree.SubElement(rPr, f"{W}szCs").set(f"{W}val", FONT_SZ)
        t = etree.SubElement(r, f"{W}t")
        t.text = text
        t.set(f"{{{XML_NS}}}space", "preserve")
        return tc

    def _make_hdr_row(*cells_el, row_height_twips=None):
        """Wrap cells into a tblHeader row with optional fixed height."""
        row = etree.Element(f"{W}tr")
        trPr = etree.SubElement(row, f"{W}trPr")
        etree.SubElement(trPr, f"{W}cantSplit")
        etree.SubElement(trPr, f"{W}tblHeader")
        if row_height_twips:
            trH = etree.SubElement(trPr, f"{W}trHeight")
            trH.set(f"{W}val",  str(row_height_twips))
            trH.set(f"{W}hRule", "atLeast")
        for c in cells_el:
            row.append(c)
        return row

    # ------------------------------------------------------------------
    # Build and insert tblHeader rows into the MAIN compliance table
    # (so Word repeats them at the top of every page automatically)
    # ------------------------------------------------------------------
    insert_idx = list(out_tbl).index(out_rows[0])

    # Row 1: spanning cell ΓÇö standard name, centred (only if we have text)
    # Height 480 twips (~8.5 mm) ΓÇö visible, matches screenshot title row height
    if spanning_text:
        span_row = _make_hdr_row(
            _make_hdr_cell(str(total_w), w_type, spanning_text,
                           center=True, gridspan=n_cols),
            row_height_twips=span_height,
        )
        out_tbl.insert(insert_idx, span_row)
        insert_idx += 1

    # Row 2: column labels ΓÇö only if column names were provided/detected
    # Height 300 twips (~5.3 mm)
    if columns:
        col_cells = []
        for ci, col_text in enumerate(columns):
            w_val, w_type_val = out_widths[ci]
            is_last = (ci == len(columns) - 1)
            col_cells.append(_make_hdr_cell(w_val, w_type_val, col_text, center=is_last))
        col_row = _make_hdr_row(*col_cells, row_height_twips=col_height)
        out_tbl.insert(insert_idx, col_row)



# ---------------------------------------------------------------------------
# Pass 16 ΓÇö Normalize rFonts attributes to remove locale-specific overrides
# ---------------------------------------------------------------------------

def normalize_rfonts(doc: Document) -> None:
    """
    Remove specific w:rFonts attributes (e.g. eastAsia, cs) that can cause
    inconsistent font rendering when rows were typed with different locale
    settings in the source document.  Driven by CFG["normalize_rfonts_remove_attributes"].
    """
    attrs_to_remove: list[str] = _CFG.get("normalize_rfonts_remove_attributes", [])
    if not attrs_to_remove:
        return

    W = f"{{{_NS}}}"
    # Build fully-qualified attribute names
    fq_attrs = [f"{W}{a}" for a in attrs_to_remove]

    for rFonts in doc.element.body.iter(f"{W}rFonts"):
        for attr in fq_attrs:
            if attr in rFonts.attrib:
                del rFonts.attrib[attr]


# Pass 15
def mark_toc_dirty(doc: Document) -> None:
    """
    Mark every TOC field as dirty so Word auto-updates the Table of Contents
    when the document is first opened.
    """
    W = f"{{{_NS}}}"
    for p_elem in doc.element.body.iter(f"{W}p"):
        instr_texts = [(t.text or "").strip() for t in p_elem.iter(f"{W}instrText")]
        if not any(txt.upper().startswith("TOC") for txt in instr_texts):
            continue
        for fldChar in p_elem.iter(f"{W}fldChar"):
            if fldChar.get(f"{W}fldCharType") == "begin":
                fldChar.set(f"{W}dirty", "1")

# ---------------------------------------------------------------------------
# Pass 17 ΓÇö Fill ISO info table cells from source doc's matching table
#            Copies full paragraph XML (preserving fonts, colours, images)
# ---------------------------------------------------------------------------

_FILL_PLACEHOLDER_PREFIX = "rId_CDRCOPY_FILL_"
_fill_placeholder_counter = 0


def _remap_images_in_elem(elem, src_doc) -> list[dict]:
    """
    Deep-copy *elem* is assumed to have already been made.  Rewrite every
    a:blip r:embed attribute to a unique placeholder rId and collect
    { placeholder_rid, bytes, content_type } for later zip-level insertion.
    """
    global _fill_placeholder_counter
    R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    A_BLIP  = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}blip"
    # Use the drawingml namespace for blip
    BLIP_TAG = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"

    images: list[dict] = []
    orig_to_ph: dict[str, str] = {}

    for elem_node in elem.iter():
        rid = elem_node.get(R_EMBED)
        if rid is None:
            continue
        if rid not in orig_to_ph:
            _fill_placeholder_counter += 1
            ph = f"{_FILL_PLACEHOLDER_PREFIX}{_fill_placeholder_counter:06d}"
            orig_to_ph[rid] = ph
            try:
                img_part = src_doc.part.related_parts[rid]
                images.append({
                    "placeholder_rid": ph,
                    "bytes":           img_part.blob,
                    "content_type":    img_part.content_type,
                })
            except (KeyError, AttributeError):
                pass
        elem_node.set(R_EMBED, orig_to_ph[rid])

    return images


def fill_table_from_source(doc: Document, data_doc_path: str | None, template_path: str | None = None) -> list[dict]:
    """
    Fill cells in the preserved ISO info table by copying full paragraph XML
    (runs, formatting, images) from the matching rows in the source document.

    Returns a list of image_parts dicts (placeholder_rid / bytes / content_type)
    that must be passed to _copy_table_images for final zip-level wiring.

    Config key: "fill_table_from_source"
      target_heading : heading text under which the target table lives in output
      source_heading : heading text under which the source table lives in input
      row_label_map  : {target_label: source_label}
                       For 2-column target rows the value cell (col 1) is
                       replaced with the XML paragraphs from the source value cell.
                       For 1-column target rows the entire cell content is replaced
                       with the XML from the source cell (full cell clone).
    """
    from copy import deepcopy

    cfg: dict = _CFG.get("fill_table_from_source", {})
    if not cfg or not data_doc_path:
        return []

    target_heading: str = cfg.get("target_heading", "")
    source_heading: str = cfg.get("source_heading", "")
    row_label_map: dict = cfg.get("row_label_map", {})
    # Read fulfils normalisation prefix from config (no hardcoded strings)
    _FULFILS_PREFIX: str = cfg.get(
        "fulfils_row_prefix", "the product fulfils the requirements of"
    )
    # Auto-detect target_heading from template when blank
    if not target_heading and template_path and os.path.isfile(template_path):
        target_heading = detect_compliance_heading(template_path)
    # Prefixes for rows where only the checkbox state should be copied;
    # the template's own text is preserved unchanged.
    checkbox_only_prefixes: list[str] = [
        p.lower() for p in cfg.get("row_copy_checkbox_only", [])
    ]
    if not target_heading or not source_heading or not row_label_map:
        return []

    W = f"{{{_NS}}}"
    h_ids = _heading_style_ids(doc)
    all_image_parts: list[dict] = []

    def _cell_text(tc_elem) -> str:
        return "".join(t.text or "" for t in tc_elem.iter(f"{W}t")).strip()

    def _get_checkbox_state(tc_elem) -> str | None:
        """
        Return the effective checked state ('0' or '1') of a FORMCHECKBOX in a cell.

        Word uses two optional sub-elements inside <w:checkBox>:
          w:checked  ΓÇö current runtime state (takes priority when present)
          w:default  ΓÇö initial/default state (fallback)
        We read w:checked first; fall back to w:default.
        """
        for fldChar in tc_elem.iter(f"{W}fldChar"):
            if fldChar.get(f"{W}fldCharType") == "begin":
                ffData = fldChar.find(f"{W}ffData")
                if ffData is not None:
                    cb = ffData.find(f"{W}checkBox")
                    if cb is not None:
                        checked = cb.find(f"{W}checked")
                        if checked is not None:
                            return checked.get(f"{W}val", "0")
                        default = cb.find(f"{W}default")
                        if default is not None:
                            return default.get(f"{W}val", "0")
        return None

    def _set_checkbox_state(tc_elem, val: str) -> None:
        """
        Set the FORMCHECKBOX state in *tc_elem* to *val* ('0' or '1').

        Writes to w:checked if it already exists (keeps it consistent with
        w:default); otherwise writes to w:default.  Creates w:default if
        neither element is present.
        """
        for fldChar in tc_elem.iter(f"{W}fldChar"):
            if fldChar.get(f"{W}fldCharType") == "begin":
                ffData = fldChar.find(f"{W}ffData")
                if ffData is not None:
                    cb = ffData.find(f"{W}checkBox")
                    if cb is not None:
                        checked = cb.find(f"{W}checked")
                        if checked is not None:
                            checked.set(f"{W}val", val)
                        default = cb.find(f"{W}default")
                        if default is not None:
                            default.set(f"{W}val", val)
                        else:
                            d = etree.SubElement(cb, f"{W}default")
                            d.set(f"{W}val", val)
                        return

    def _find_table_after_heading_output(body_children, heading_text: str):
        for i, elem in enumerate(body_children):
            if elem.tag == f"{W}p" and _para_style_id(elem) in h_ids:
                if _para_text(elem).strip().lower() == heading_text.lower():
                    for nxt in body_children[i + 1:]:
                        if nxt.tag == f"{W}p" and _para_style_id(nxt) in h_ids:
                            break
                        if nxt.tag == f"{W}tbl":
                            return nxt
        return None

    def _find_source_table_by_content(body_children, key_labels: list):
        key_labels_lower = {k.lower() for k in key_labels}
        best_tbl, best_score = None, 0
        for elem in body_children:
            if elem.tag != f"{W}tbl":
                continue
            score = 0
            for tr in elem.findall(f"{W}tr"):
                cells = tr.findall(f"{W}tc")
                if not cells:
                    continue
                cell0 = _cell_text(cells[0]).lower()
                for kl in key_labels_lower:
                    if cell0.startswith(kl):
                        score += 1
                        break
            if score > best_score:
                best_score = score
                best_tbl = elem
        return best_tbl if best_score > 0 else None

    # Prefix used to normalise standard-specific "fulfils" rows so that inputs
    # with different standards (ISO 17664-2, DIN 6868-157, IEC 62304, ΓÇª) all
    # resolve to the same lookup key regardless of which standard is named.
    # Moved to config ΓÇö _FULFILS_PREFIX already set from cfg above.

    def _normalise_label(lbl: str) -> str:
        return _FULFILS_PREFIX if lbl.startswith(_FULFILS_PREFIX) else lbl

    # --- Load source doc ---
    from docx import Document as _Doc
    src_doc = _Doc(data_doc_path)
    src_children = list(src_doc.element.body)
    # Include the generic "fulfils" prefix so the scorer can find tables that
    # mention any standard (not just the ISO variant hard-coded in config).
    src_key_labels = list(row_label_map.values()) + [_FULFILS_PREFIX]
    src_tbl = _find_source_table_by_content(src_children, src_key_labels)
    if src_tbl is None:
        return []

    # Build map: lowercase source label ΓåÆ source tc element.
    # For "fulfils" rows also store the entry under the generic normalised key
    # so it can be found regardless of which standard the source doc names.
    src_label_to_tc: dict = {}
    for tr in src_tbl.findall(f"{W}tr"):
        cells = tr.findall(f"{W}tc")
        if not cells:
            continue
        cell_text = _cell_text(cells[0])
        if len(cells) == 2:
            key = cell_text.lower()
            src_label_to_tc[key] = cells[1]
            norm = _normalise_label(key)
            if norm != key:
                src_label_to_tc.setdefault(norm, cells[1])
        else:
            colon_pos = cell_text.find(":")
            label_key = (cell_text[:colon_pos + 1] if colon_pos != -1 else cell_text).lower()
            src_label_to_tc[label_key] = cells[0]
            norm = _normalise_label(label_key)
            if norm != label_key:
                src_label_to_tc.setdefault(norm, cells[0])

    row_label_map_lower = {k.lower(): v.lower() for k, v in row_label_map.items()}

    # --- Find target table in output ---
    out_children = list(doc.element.body)
    target_tbl = _find_table_after_heading_output(out_children, target_heading)
    if target_tbl is None:
        return []

    # --- For each mapped target row, replace cell content with source XML ---
    for tr in target_tbl.findall(f"{W}tr"):
        cells = tr.findall(f"{W}tc")
        if not cells:
            continue
        label_text = _cell_text(cells[0])
        label_lower = label_text.rstrip().lower()
        # For 1-column rows the template cell may have placeholder text after
        # the label (e.g. "Copy of marking plate:The artwork below...").
        # Match by the prefix up to (and including) the first ':' so we can
        # use clean short keys in row_label_map.
        if len(cells) == 1:
            colon_pos = label_lower.find(":")
            label_key = label_lower[:colon_pos + 1] if colon_pos != -1 else label_lower
        else:
            label_key = label_lower

        src_key = row_label_map_lower.get(label_key)
        # Fallback: normalise "fulfils" labels so any standard matches the
        # ISO-specific key written in row_label_map.
        if src_key is None:
            norm_key = _normalise_label(label_key)
            src_key = row_label_map_lower.get(norm_key)
        if src_key is None:
            continue

        src_tc = src_label_to_tc.get(src_key)
        # Fallback: look up by normalised key (handles DIN / IEC / ISO variants).
        if src_tc is None:
            src_tc = src_label_to_tc.get(_normalise_label(src_key))
        if src_tc is None:
            continue

        # ΓöÇΓöÇ Checkbox-only rows ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
        # For rows listed in row_copy_checkbox_only, preserve the template's
        # own text and only copy the checked/unchecked state from the source.
        target_cell = cells[1] if len(cells) == 2 else cells[0]
        is_checkbox_only = any(label_key.startswith(p) for p in checkbox_only_prefixes)
        if not is_checkbox_only:
            # Also check the raw un-truncated label text for 1-column rows
            is_checkbox_only = any(label_lower.startswith(p) for p in checkbox_only_prefixes)
        if is_checkbox_only:
            src_val = _get_checkbox_state(src_tc)
            if src_val is not None:
                _set_checkbox_state(target_cell, src_val)
            continue

        if len(cells) == 2:
            # 2-column row: replace the value cell (col 1) content
            # Copy all <w:p> elements from source value cell into target value cell
            val_tc = cells[1]
            # Remove existing paragraphs from target value cell (keep tcPr)
            for p in list(val_tc.findall(f"{W}p")):
                val_tc.remove(p)
            # Deep-copy all content children from source value cell
            # (paragraphs AND any nested tables that carry images)
            src_children = [c for c in src_tc if c.tag != f"{W}tcPr"]
            for src_child in src_children:
                new_child = deepcopy(src_child)
                imgs = _remap_images_in_elem(new_child, src_doc)
                all_image_parts.extend(imgs)
                val_tc.append(new_child)
            # Ensure there is at least one paragraph (Word requires it)
            if not val_tc.findall(f"{W}p"):
                val_tc.append(etree.SubElement(val_tc, f"{W}p"))
        else:
            # 1-column row: replace all content in the target cell with
            # a deep copy of the source cell's contents (paragraphs + nested tables)
            val_tc = cells[0]
            # Save tcPr if present
            tc_pr = val_tc.find(f"{W}tcPr")
            # Remove all children
            for child in list(val_tc):
                val_tc.remove(child)
            # Restore tcPr
            if tc_pr is not None:
                val_tc.append(deepcopy(tc_pr))
            # Copy all content children (not tcPr)
            src_children = [c for c in src_tc if c.tag != f"{W}tcPr"]
            for src_child in src_children:
                new_child = deepcopy(src_child)
                imgs = _remap_images_in_elem(new_child, src_doc)
                all_image_parts.extend(imgs)
                val_tc.append(new_child)
            # Ensure at least one paragraph
            if not val_tc.findall(f"{W}p"):
                val_tc.append(etree.SubElement(val_tc, f"{W}p"))

    return all_image_parts


def apply_all(doc: Document, product_name: str | None = None, data_doc_path: str | None = None, template_path: str | None = None) -> list[dict]:
    """Run every configured post-processing pass on *doc* in-place.
    Returns a list of image_parts dicts (may be empty) collected during
    fill_table_from_source that must be passed to _copy_table_images.
    """
    fill_header_footer_placeholders(doc, product_name)   # must run BEFORE bracket stripping
    remove_preamble_before_first_heading(doc)
    remove_styled_paragraphs(doc)
    remove_template_instructions(doc)
    strip_inline_angle_brackets(doc)
    remove_paragraphs_with_text(doc)
    remove_empty_table_rows(doc)
    set_note_text_size(doc)
    strip_superscript_list_markers(doc)
    remove_sections(doc)
    remove_headings_only(doc)
    set_all_text_black(doc)
    prepend_compliance_table_header(doc, template_path)   # after set_all_text_black so header keeps auto/style colour
    remove_blank_paragraph_before_headings(doc)
    remove_blank_paragraph_after_headings(doc)
    inject_definitions_fixed_rows(doc)
    sort_tables_alphabetically(doc)
    normalize_rfonts(doc)
    extra_image_parts = fill_table_from_source(doc, data_doc_path, template_path=template_path)
    mark_toc_dirty(doc)
    return extra_image_parts
