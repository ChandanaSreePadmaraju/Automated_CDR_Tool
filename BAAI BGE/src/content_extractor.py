"""
content_extractor.py
--------------------
Extract all content (paragraphs, tables) from the section in a .docx file
that starts at a named heading paragraph and ends just before the next heading
at the same or higher level.

Every body element is serialised as raw XML so all formatting, inline images,
run-level markup and nested structures are preserved exactly as-is.
"""

import copy
import itertools

from lxml import etree
from docx import Document
from docx.oxml.ns import qn

# Module-level counter so every placeholder rId is globally unique across all
# extract_section() calls in a single pipeline run.
_PLACEHOLDER_COUNTER = itertools.count(1)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _build_heading_level_map(doc: Document) -> dict:
    """
    Return  { id(para._p): level_int }  for every heading paragraph.
    Using object-identity on the lxml element avoids repeated string-parsing.
    """
    level_map: dict = {}
    for para in doc.paragraphs:
        style_name = para.style.name
        if style_name.startswith("Heading"):
            parts = style_name.split()
            level = int(parts[-1]) if parts[-1].isdigit() else 1
            level_map[id(para._p)] = level
    return level_map


# Prefix that is guaranteed never to appear in real Word rId values.
_PLACEHOLDER_PREFIX = "rId_CDRCOPY_"

# Large sentinel added to every <w:numId w:val="N"/> at extraction time so
# the values cannot collide with numbering definitions already in the template.
# _merge_numbering() in template_filler uses the same constant to locate and
# remap them back to real sequential IDs at ZIP-merge time.
_NUMID_OFFSET = 10_000


def _offset_numids(elem) -> None:
    """Add _NUMID_OFFSET to every <w:numId w:val="N"/> in *elem*."""
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    val_attr = f"{{{ns_w}}}val"
    for numId_elem in elem.findall(f".//{{{ns_w}}}numId"):
        val = numId_elem.get(val_attr)
        if val and val.isdigit():
            numId_elem.set(val_attr, str(int(val) + _NUMID_OFFSET))


def _extract_and_remap_images(elem, doc: Document) -> list[dict]:
    """
    For every <a:blip r:embed="rId…"> inside *elem*:
      1. Assign a globally-unique placeholder rId (e.g. rId_CDRCOPY_0003).
      2. Rewrite the blip's r:embed attribute in *elem* to the placeholder.
      3. Collect { "placeholder_rid", "bytes", "content_type" } for later
         zip-level insertion.

    Using placeholder rIds means the later global text-replace in
    _copy_table_images() only touches content extracted from the data doc
    and never accidentally overwrites template-original rId references
    (e.g. the cover page logo).
    """
    images: list[dict] = []
    original_to_placeholder: dict[str, str] = {}

    for blip in elem.iter(qn("a:blip")):
        orig_rid = blip.get(qn("r:embed"))
        if not orig_rid:
            continue

        if orig_rid not in original_to_placeholder:
            placeholder = f"{_PLACEHOLDER_PREFIX}{next(_PLACEHOLDER_COUNTER):06d}"
            original_to_placeholder[orig_rid] = placeholder
            try:
                img_part = doc.part.related_parts[orig_rid]
                images.append({
                    "placeholder_rid": placeholder,
                    "bytes":           img_part.blob,
                    "content_type":    img_part.content_type,
                })
            except (KeyError, AttributeError):
                # Image not resolvable — keep the placeholder so Word shows
                # a broken-image icon rather than silently referencing a
                # completely unrelated resource in the template.
                pass

        # Always rewrite the attribute to the placeholder
        blip.set(qn("r:embed"), original_to_placeholder[orig_rid])

    return images


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_section(
    doc_path: str,
    heading_para_index: int,
    heading_level: int,
) -> dict:
    """
    Extract the content of the section that *starts* at the paragraph whose
    index in ``doc.paragraphs`` is *heading_para_index*.

    The section ends just before the next paragraph whose heading level is
    ≤ *heading_level* (i.e. a sibling or ancestor heading).

    Every body element is returned as serialised raw XML so that all
    formatting, inline images, nested runs, and mixed content are preserved
    exactly — no structural assumptions about what a section contains.

    Returns
    -------
    {
        "items": [
            {
                "type":        "paragraph" | "table",
                "xml":         bytes,          – full serialised <w:p> or <w:tbl>
                "image_parts": list[dict],     – images referenced inside the element
            },
            …
        ]
    }
    """
    doc = Document(doc_path)
    paragraphs = doc.paragraphs

    if heading_para_index >= len(paragraphs):
        return {"items": []}

    heading_p_elem    = paragraphs[heading_para_index]._p
    heading_level_map = _build_heading_level_map(doc)

    body_children = list(doc.element.body)

    try:
        start_pos = body_children.index(heading_p_elem)
    except ValueError:
        return {"items": []}

    body_items: list[dict] = []

    for elem in body_children[start_pos + 1:]:

        if elem.tag == qn("w:p"):
            # Stop at the next sibling or ancestor heading
            if id(elem) in heading_level_map and heading_level_map[id(elem)] <= heading_level:
                break
            deep = copy.deepcopy(elem)
            _offset_numids(deep)
            image_parts = _extract_and_remap_images(deep, doc)
            body_items.append({"type": "paragraph", "xml": etree.tostring(deep), "image_parts": image_parts})

        elif elem.tag == qn("w:tbl"):
            deep = copy.deepcopy(elem)
            _offset_numids(deep)
            image_parts = _extract_and_remap_images(deep, doc)
            body_items.append({"type": "table", "xml": etree.tostring(deep), "image_parts": image_parts})

    # Strip trailing empty paragraphs (no text, no drawings) — avoids
    # inserting long runs of blank lines that appear at the end of sections
    # in the data doc (e.g. 18 empty ListParagraph lines after References table).
    while body_items:
        last = body_items[-1]
        if last["type"] == "paragraph":
            # Re-parse from xml bytes only when checking; reuse the deepcopy
            # already made above is not available here so a parse is needed —
            # but do it only once per loop iteration, not twice.
            p_elem      = etree.fromstring(last["xml"])
            has_text    = bool(''.join(t.text or '' for t in p_elem.iter(qn('w:t'))).strip())
            has_drawing = p_elem.find('.//' + qn('w:drawing')) is not None
            if not has_text and not has_drawing:
                body_items.pop()
                continue
        break

    return {"items": body_items}


def extract_pre_heading_content(doc_path: str) -> dict:
    """
    Extract every body element that appears **before the first heading**
    in *doc_path*.

    This captures document-level metadata text (e.g. "Philips Compliance
    Data Record  ISO …") that is written as plain paragraphs or tables at
    the very top of the file, before any section heading exists.

    Returns the same ``{"items": [...]}`` structure as :func:`extract_section`.
    """
    doc = Document(doc_path)
    heading_level_map = _build_heading_level_map(doc)
    body_children = list(doc.element.body)

    body_items: list[dict] = []

    for elem in body_children:
        # Stop as soon as we hit the first heading
        if elem.tag == qn("w:p") and id(elem) in heading_level_map:
            break

        if elem.tag == qn("w:p"):
            deep = copy.deepcopy(elem)
            _offset_numids(deep)
            image_parts = _extract_and_remap_images(deep, doc)
            body_items.append({"type": "paragraph", "xml": etree.tostring(deep), "image_parts": image_parts})

        elif elem.tag == qn("w:tbl"):
            deep = copy.deepcopy(elem)
            _offset_numids(deep)
            image_parts = _extract_and_remap_images(deep, doc)
            body_items.append({"type": "table", "xml": etree.tostring(deep), "image_parts": image_parts})

    # Strip leading and trailing empty paragraphs
    def _is_empty_para(item: dict) -> bool:
        if item["type"] != "paragraph":
            return False
        p_elem = etree.fromstring(item["xml"])
        has_text    = bool(''.join(t.text or '' for t in p_elem.iter(qn('w:t'))).strip())
        has_drawing = p_elem.find('.//' + qn('w:drawing')) is not None
        return not has_text and not has_drawing

    while body_items and _is_empty_para(body_items[0]):
        body_items.pop(0)
    while body_items and _is_empty_para(body_items[-1]):
        body_items.pop()

    return {"items": body_items}
