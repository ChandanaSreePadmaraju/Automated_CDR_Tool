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
            image_parts = _extract_and_remap_images(deep, doc)
            body_items.append({"type": "paragraph", "xml": etree.tostring(deep), "image_parts": image_parts})

        elif elem.tag == qn("w:tbl"):
            deep = copy.deepcopy(elem)
            image_parts = _extract_and_remap_images(deep, doc)
            body_items.append({"type": "table", "xml": etree.tostring(deep), "image_parts": image_parts})

    return {"items": body_items}
