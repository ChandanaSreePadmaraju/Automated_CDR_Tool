# BAAI BGE Updates

---

## April 20, 2026

### Cleanup

- Removed dev/debug scripts: `diag.py`, `quick_check.py`, `verify_output.py`
- Updated `.gitignore`: entire `output/` excluded (no pinned file), added `~$*` for Word lock files
- Output is now auto-versioned at runtime (`Filled_CDR.docx`, `_v2`, `_v3`, …)

### `product_name` moved to `prompts.json`

- `product_name` is now configured directly in `prompts.json` — no need to pass `--product-name` on the command line for every run
- `main.py` reads `product_name` from `prompts.json` at startup
- The `--product-name` CLI flag still works as a **one-off override** (e.g. when running for a different product without editing the config)
- Priority: CLI flag > `prompts.json`

### Removed `skip_template_headings`

- Removed the commented-out `skip_template_headings` key from `prompts.json` — it was never asked for
- Removed corresponding `_skip_headings` variable and filter from `template_filler.py`
- All matched headings now have their content filled — nothing is skipped

### Code cleanup — deferred imports moved to top level

All `import` statements that were inside functions have been moved to module-level:

| File | Removed from inside function | Now at top level |
|---|---|---|
| `main.py` | `import json as _json` inside `main()` | `import json` |
| `main.py` | `import re as _re` inside `main()` | `import re` |
| `template_filler.py` | `import json, os as _os` inside `fill_template()` | `import json`, `import os` |
| `template_filler.py` | `from lxml import etree as _etree` inside loop | uses already-imported `etree` |
| `template_filler.py` | `import io` removed | all usages use `from io import BytesIO` |
| `post_processor.py` | `from lxml import etree as _etree` (×3 functions) | `from lxml import etree` |
| `post_processor.py` | `from copy import deepcopy` inside function | `from copy import deepcopy` |

### Removed dead code — `add_page_break_before_headings`

- Removed `add_page_break_before_headings()` (~40 lines) from `post_processor.py`
- It was never called from `apply_all()` and the config key `page_break_before_headings` does not exist in `prompts.json`
- Removed the trailing comment about it in `apply_all()`

### Cleaned up verbose module docstring in `post_processor.py`

- Removed the function-list block from the module docstring — it duplicated what the code already says

---

## April 14, 2026

**Selected Model:** `bge-base-en-v1.5` → fast + accurate enough

---

## April 15, 2026

### Final Implementation — Full Pipeline

---

### Project Structure

```
BAAI BGE/
├── main.py                    # CLI entry point — runs the full pipeline
├── prompts.json               # All config (query prefix, heading mappings, post-processing flags)
├── requirements.txt           # Python dependencies
├── src/
│   ├── heading_extractor.py   # Step 1 — extract headings from any .docx
│   ├── heading_matcher.py     # Step 2 — BGE semantic heading matching
│   ├── content_extractor.py   # Step 3 — extract section content as raw XML
│   ├── template_filler.py     # Step 4 — insert content into template, save output
│   └── post_processor.py      # Step 5 — 11 post-processing cleanup passes
└── output/                    # Generated output (git-ignored)
```

---

### Bugs Fixed

| Bug | Root Cause | Fix |
|---|---|---|
| `"Compliance Checklist"` section not filled | Semantic score 0.509 < 0.55 threshold | `heading_mappings` override in `prompts.json` + explicit check in `heading_matcher.py` |
| Hardcoded document structure assumed | `<w:p>` extracted as plain text; images emitted as separate items | All elements serialised as raw XML bytes; single `_insert_xml_after` handles any element |
| Cover page `<ProductName RX.Y>` not replaced | Only headers/footers scanned; placeholder split across 4 `<w:r>` runs by Word | lxml-based multi-run collapse algorithm; added `doc.element.body` scan |
| Cover page logo replaced with data doc image | Global `str.replace('r:embed="rId12"')` overwrote template's `rId12` | Placeholder rIds (`rId_CDRCOPY_...`) injected at extraction time; only those replaced at zip level |

---

## Complete A-Z Pipeline — Technical Walkthrough

> Full technical trace from CLI invocation to final `.docx` — function signatures, data structures, XML namespaces, algorithm internals, and zip-level operations.

---

### Input / Output contract

| | Type | Example |
|---|---|---|
| **Template** | `.docx` — structured CDR skeleton: Heading-styled paragraphs, `<ProductName RX.Y>` placeholders, Guidance-styled instructions, empty sections | `Template_D001024021 CDR ISO 17664-2 (2021) ProductName RX.Y Rev C.docx` |
| **Data doc** | `.docx` — filled source document: real content under differently-named headings | `Input_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx` |
| **Output** | `.docx` — template structure + data doc content + all post-processing applied | `output/Filled_CDR.docx` (auto-versioned) |

Both documents are standard Office Open XML packages — zip archives containing `word/document.xml`, `word/_rels/document.xml.rels`, `word/media/`, etc.

---

### Call Graph

```
main()
  │
  ├─ parse_args()  →  argparse.Namespace
  │     flags: --template, --input, --output, --threshold (0.55), --product-name
  │
  ├─ json.load("prompts.json")  →  product_name  (CLI --product-name overrides)
  │
  ├─ re.sub(r'_v\d+$', '', base)  →  auto-versioning loop  →  output_path
  │
  ├─ [STEP 1]  extract_headings(template_path)  →  list[dict]
  │            extract_headings(data_path)       →  list[dict]
  │
  ├─ [STEP 2]  load_model()  →  SentenceTransformer("BAAI/bge-base-en-v1.5")
  │            match_headings(tmpl_headings, data_headings, model, threshold)
  │                →  list[{ template_heading, matched_heading, score }]
  │
  └─ [STEP 3+4+5]  fill_template(template_path, data_path, matches, output_path, product_name)
        │
        ├─ for each match (reverse paragraph_index order):
        │     extract_section(data_path, para_idx, level)  →  { "items": [...] }
        │     _clear_section(heading_elem)   or   partial-clear (keep_n_tables)
        │     _insert_xml_after(anchor, item["xml"])   ×N
        │
        ├─ post_processor.apply_all(template_doc, product_name)   [11 passes]
        │
        ├─ template_doc.save(BytesIO)   →   bytes
        └─ _copy_table_images(bytes, all_image_parts)   →   final bytes   →   file
```

---

### Step 1 — Heading Extraction (`heading_extractor.py`)

**Signature:**
```python
def extract_headings(doc_path: str) -> list[dict]:
```

**Algorithm:**
```python
doc = Document(doc_path)

for idx, para in enumerate(doc.paragraphs):
    style_name = para.style.name          # e.g. "Heading 1", "Heading 2"
    if not style_name.startswith("Heading"):
        continue
    text = para.text.strip()
    if not text:
        continue                          # skip blank heading paragraphs
    parts = style_name.split()
    level = int(parts[-1]) if parts[-1].isdigit() else 1
    headings.append({"text": text, "level": level, "paragraph_index": idx})
```

- `doc.paragraphs` only iterates body paragraphs — tables and headers/footers are not in this list
- `paragraph_index` is the **0-based offset into `doc.paragraphs`**, not the body-child index — used later in `extract_section()` to locate the start of a section

**Return type:** `list[dict]` — one entry per non-blank heading paragraph

```python
# Template
[
    {"text": "Introduction",         "level": 1, "paragraph_index": 3},
    {"text": "Compliance Checklist", "level": 2, "paragraph_index": 10},
    {"text": "Record history",       "level": 2, "paragraph_index": 45},
]

# Data doc
[
    {"text": "Test Record",               "level": 2, "paragraph_index": 8},
    {"text": "Document Revision History", "level": 2, "paragraph_index": 52},
]
```

---

### Step 2 — Semantic Heading Matching (`heading_matcher.py`)

**Signatures:**
```python
def load_model() -> SentenceTransformer:

def match_headings(
    template_headings: list[dict],
    data_headings:     list[dict],
    model:             SentenceTransformer,
    threshold:         float = 0.60,
) -> list[dict]:
```

**Module-level config** (read once at import time):
```python
_QUERY_PREFIX: str = _PROMPTS["query_prefix"]
# "Represent this sentence for searching relevant passages: "

_MODEL_NAME: str = _PROMPTS.get("model_name", "BAAI/bge-base-en-v1.5")
```

**Encoding:**
```python
# BGE asymmetric encoding: query side (template) gets the retrieval prefix
query_texts  = [_QUERY_PREFIX + h["text"] for h in template_headings]
corpus_texts = [h["text"] for h in data_headings]   # corpus side: plain text

query_embs  = model.encode(query_texts,  normalize_embeddings=True, show_progress_bar=False)
corpus_embs = model.encode(corpus_texts, normalize_embeddings=True, show_progress_bar=False)
# Each embedding: float32 array of shape (768,)
# normalize_embeddings=True  →  L2-normalised, so dot product == cosine similarity
```

**Similarity matrix:**
```python
sim_matrix: np.ndarray = cosine_similarity(query_embs, corpus_embs)
# Shape: (len(template_headings), len(data_headings))
# sim_matrix[i][j] = cosine similarity between template heading i and data heading j
```

**Matching logic — explicit overrides win over AI:**
```python
explicit: dict[str, str] = {
    k.lower(): v.lower()
    for k, v in _PROMPTS.get("heading_mappings", {}).items()
}
data_by_lower: dict[str, dict] = {h["text"].lower(): h for h in data_headings}

for i, tmpl_h in enumerate(template_headings):
    forced_target = explicit.get(tmpl_h["text"].lower())
    if forced_target and forced_target in data_by_lower:
        # score=1.0, bypasses AI entirely
        matches.append({...score: 1.0})
        continue

    best_idx   = int(np.argmax(sim_matrix[i]))
    best_score = float(sim_matrix[i][best_idx])
    matches.append({
        "matched_heading": data_headings[best_idx] if best_score >= threshold else None,
        "score": round(best_score, 4),
    })
```

**Return type:** `list[dict]` — same length as `template_headings`

```python
[
    {
        "template_heading": {"text": "Compliance Checklist", "paragraph_index": 10, "level": 2},
        "matched_heading":  {"text": "Test Record",          "paragraph_index": 8,  "level": 2},
        "score": 1.0,       # forced via heading_mappings; AI score was 0.509 (below threshold)
    },
    {
        "template_heading": {"text": "Record history",            "paragraph_index": 45, "level": 2},
        "matched_heading":  {"text": "Document Revision History", "paragraph_index": 52, "level": 2},
        "score": 1.0,       # forced via heading_mappings
    },
    {
        "template_heading": {"text": "Appendix", "paragraph_index": 80, "level": 1},
        "matched_heading":  None,    # best_score 0.31 < threshold 0.55
        "score": 0.31,
    },
]
```

**Why `heading_mappings` overrides exist:**  
`"Compliance Checklist"` vs `"Test Record"` — semantically unrelated strings, cosine score 0.509, below the 0.55 threshold. Lowering the global threshold causes false positives elsewhere; the explicit map targets only the known mismatches.

---

### Step 3 — Section Content Extraction (`content_extractor.py`)

**Signature:**
```python
def extract_section(
    doc_path:           str,
    heading_para_index: int,
    heading_level:      int,
) -> dict:
    # Returns: {"items": list[{"type", "xml", "image_parts"}]}
```

**Section boundary detection:**
```python
# Build a fast id(elem) → level lookup to avoid repeated style-name parsing
heading_level_map: dict = {}   # { id(para._p): int }

body_children = list(doc.element.body)
# doc.element.body children include <w:p>, <w:tbl>, <w:sectPr>, etc.

start_pos = body_children.index(heading_p_elem)   # O(n) scan by object identity

for elem in body_children[start_pos + 1:]:
    if elem.tag == qn("w:p"):
        # Stop if next heading at same or higher level (lower level number)
        if id(elem) in heading_level_map and heading_level_map[id(elem)] <= heading_level:
            break
        # --- extract paragraph ---

    elif elem.tag == qn("w:tbl"):
        # --- extract table ---
```

**Raw XML extraction — deep copy + serialise:**
```python
deep = copy.deepcopy(elem)          # lxml deep copy — isolates from live document
image_parts = _extract_and_remap_images(deep, doc)
body_items.append({
    "type":        "paragraph" | "table",
    "xml":         etree.tostring(deep),   # bytes: full <w:p> or <w:tbl> XML
    "image_parts": image_parts,
})
```

`etree.tostring()` preserves every namespace declaration, attribute, and nested child exactly — no information is lost through a Python object model.

**Image rId placeholder system:**

A `.docx` `<a:blip r:embed="rId7">` references a relationship defined in `word/_rels/document.xml.rels`. The same `rId7` in the template may point to a completely different image (e.g. the cover logo).

```python
_PLACEHOLDER_PREFIX  = "rId_CDRCOPY_"
_PLACEHOLDER_COUNTER = itertools.count(1)   # module-level; unique across all calls

def _extract_and_remap_images(elem, doc):
    for blip in elem.iter(qn("a:blip")):
        orig_rid = blip.get(qn("r:embed"))       # e.g. "rId7"
        placeholder = f"rId_CDRCOPY_{next(_PLACEHOLDER_COUNTER):06d}"
        blip.set(qn("r:embed"), placeholder)     # rewrite in-place on the deep copy
        img_part = doc.part.related_parts[orig_rid]
        images.append({
            "placeholder_rid": placeholder,      # e.g. "rId_CDRCOPY_000001"
            "bytes":           img_part.blob,    # raw image bytes
            "content_type":    img_part.content_type,  # e.g. "image/png"
        })
```

- `elem.iter(qn("a:blip"))` — `qn` expands to `{http://schemas.openxmlformats.org/drawingml/2006/main}blip`
- The counter is module-level so IDs are unique across all `extract_section()` calls in a single run

**Trailing blank paragraph strip:**
```python
while body_items:
    last = body_items[-1]
    if last["type"] == "paragraph":
        p_elem = etree.fromstring(last["xml"])
        has_text    = bool(''.join(t.text or '' for t in p_elem.iter(qn('w:t'))).strip())
        has_drawing = bool(list(p_elem.iter(qn('w:drawing'))))
        if not has_text and not has_drawing:
            body_items.pop()
            continue
    break
```

**Return value:**
```python
{
    "items": [
        {
            "type": "paragraph",
            "xml":  b'<w:p xmlns:w="..." ...><w:r><w:t>System passed all tests.</w:t></w:r></w:p>',
            "image_parts": []
        },
        {
            "type": "table",
            "xml":  b'<w:tbl ...>...<a:blip r:embed="rId_CDRCOPY_000001"/>...</w:tbl>',
            "image_parts": [
                {
                    "placeholder_rid": "rId_CDRCOPY_000001",
                    "bytes":           b'\x89PNG\r\n\x1a\n...',
                    "content_type":    "image/png"
                }
            ]
        }
    ]
}
```

---

### Step 4 — Template Filling (`template_filler.py`)

**Signature:**
```python
def fill_template(
    template_path: str,
    data_doc_path: str,
    matches:       list[dict],
    output_path:   str,
    product_name:  str | None = None,
) -> None:
```

#### 4a — Reverse-order sort

```python
valid_matches = [m for m in matches if m["matched_heading"] is not None]
valid_matches.sort(
    key=lambda m: m["template_heading"]["paragraph_index"],
    reverse=True,    # last heading in document first
)
```

Processing descending by `paragraph_index` means inserting content after heading at index 45 never shifts the index of heading at index 10.

#### 4b — Section clearing

**Full clear** (`_clear_section`):
```python
def _clear_section(heading_elem) -> None:
    body = heading_elem.getparent()
    body_children = list(body)
    start = body_children.index(heading_elem)
    to_remove = []
    for elem in body_children[start + 1:]:
        if _is_heading_elem(elem):   # any Heading style — stop
            break
        to_remove.append(elem)
    # Preserve trailing page-break / sectPr paragraphs for section spacing
    while to_remove and _is_layout_para(to_remove[-1]):
        to_remove.pop()
    for elem in to_remove:
        body.remove(elem)
```

`_is_layout_para` checks for `<w:br w:type="page"/>` or `<w:sectPr>` inside the paragraph — these define page boundaries and must survive clearing.

**Partial clear** (`sections_keep_first_n_tables` from `prompts.json`):
```python
# e.g. "Compliance Checklist": 1  →  keep first table, delete the rest
n = _keep_n[tmpl_h_lower]
tables_seen = 0
for elem in body_children[start + 1:]:
    if _is_heading_elem(elem): break
    if elem.tag == f"{{{ns}}}tbl":
        tables_seen += 1
        if tables_seen <= n:
            after_kept = elem   # advance anchor past kept table
            continue
    to_remove.append(elem)
```

#### 4c — XML insertion

```python
def _insert_xml_after(ref_elem, xml_bytes: bytes):
    new_elem = copy.deepcopy(etree.fromstring(xml_bytes))
    _strip_comments(new_elem)     # remove <w:commentRangeStart/End> and <w:commentReference>
    ref_elem.addnext(new_elem)    # lxml: insert immediately after ref_elem in parent
    return new_elem               # returned so caller can advance the anchor
```

The `anchor` variable advances with each insert:
```python
for item in content["items"]:
    anchor = _insert_xml_after(anchor, item["xml"])
    all_image_parts.extend(item.get("image_parts", []))
```

`sections_filter_heading_content`: for sections like `"Compliance Checklist"`, any extracted `<w:p>` whose `<w:pStyle w:val>` is a heading style ID is skipped (sub-headings like "Test Administration" are not inserted).

#### 4d — Image injection (zip level)

```python
def _copy_table_images(output_bytes: bytes, all_image_parts: list[dict]) -> bytes:
```

**Why zip-level?**  `python-docx` has no API to add new image relationships after a document is assembled. The only way is to directly manipulate the `.docx` zip.

**Algorithm:**
```python
# 1. Deduplicate: one entry per unique placeholder_rid
seen: dict[str, dict] = {}
for item in all_image_parts:
    seen.setdefault(item["placeholder_rid"], item)

# 2. Open output .docx (which is a zip) and read all entries into memory
with zipfile.ZipFile(BytesIO(output_bytes), "r") as zin:
    existing = {n: zin.read(n) for n in zin.namelist()}

rels_text = existing["word/_rels/document.xml.rels"].decode("utf-8")
doc_text  = existing["word/document.xml"].decode("utf-8")

# 3. Find highest existing rId number to avoid collisions
existing_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_text)]
next_id = max(existing_nums, default=0) + 1

# 4. For each unique image: assign final rId, write media file, add <Relationship>
for item in unique:
    new_rid    = f"rId{next_id}";  next_id += 1
    media_name = f"img_copied_{new_rid}{ext}"
    new_media[f"word/media/{media_name}"] = item["bytes"]
    new_rels.append(
        f'<Relationship Id="{new_rid}" Type="...relationships/image" '
        f'Target="media/{media_name}"/>'
    )
    rid_map[item["placeholder_rid"]] = new_rid

# 5. Text-replace ONLY rId_CDRCOPY_... strings — never touches template rIds
for placeholder, new_rid in rid_map.items():
    doc_text = doc_text.replace(f'r:embed="{placeholder}"', f'r:embed="{new_rid}"')

# 6. Inject <Relationship> entries before closing </Relationships> tag
insert_at = rels_text.rfind("</")
rels_text  = rels_text[:insert_at] + "\n  ".join(new_rels) + "\n" + rels_text[insert_at:]

# 7. Reassemble zip
with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
    for name, data in existing.items():
        if name not in {"word/_rels/document.xml.rels", "word/document.xml"}:
            zout.writestr(name, data)
    zout.writestr("word/_rels/document.xml.rels", rels_text.encode())
    zout.writestr("word/document.xml",            doc_text.encode())
    for name, data in new_media.items():
        zout.writestr(name, data)
```

---

### Step 5 — Post-Processing (`post_processor.py`)

**Module-level config** (read once at import time):
```python
_CFG: dict = json.load(open("prompts.json"))["post_processing"]
_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```

All passes operate on `doc.element.body` (raw lxml `_Element` tree) — not `python-docx`'s `.paragraphs` / `.tables` abstractions. This is necessary because content inserted via `_insert_xml_after` in Step 4 is raw lxml and does not appear in `doc.paragraphs` until the document is reloaded.

**Entry point:**
```python
def apply_all(doc: Document, product_name: str | None = None) -> None:
```

| # | Function | Key implementation detail |
|---|---|---|
| 1 | `fill_header_footer_placeholders` | Multi-run collapse: iterates every `<w:p>` in `doc.element.body` + all `doc.sections` header/footer parts. Concatenates run texts, finds the placeholder span by string index, collapses matching `<w:r>` nodes into one, deletes the rest. Handles Word splitting `<ProductName RX.Y>` across up to 4 separate `<w:r>` elements. |
| 2 | `remove_preamble_before_first_heading` | Finds first `<w:p>` with a Heading style in `body`. Removes all preceding body children except `<w:tbl>` (cover page) and paragraphs containing `<w:br w:type="page"/>` or `<w:sectPr>`. |
| 3 | `remove_styled_paragraphs` | Builds `{style_id: style_name}` from `doc.styles`. Removes `<w:p>` where `<w:pStyle w:val>` maps to a name in `_CFG["styles_to_remove"]` (e.g. `"Guidance"`). |
| 4 | `remove_template_instructions` | State-machine scan over body `<w:p>` elements. Opens a block when `_para_text(e)` == `"<"` or starts with `"< "`. Closes when text == `">"` or starts with `"> "`. Both markers and all paragraphs between them are removed. |
| 5 | `strip_inline_angle_brackets` | Removes any `<w:r>` whose sole `<w:t>` text is exactly `"<"` or `">"` from all body paragraphs. |
| 6 | `remove_paragraphs_with_text` | Removes `<w:p>` where `_para_text(e)` contains any string from `_CFG["remove_paragraphs_with_text"]`. `_para_text` joins all `<w:t>` text nodes. |
| 7 | `remove_empty_table_rows` | Iterates `doc.tables`. Skips the first row (header). Removes `<w:tr>` where every `<w:tc>` has empty text (joined `<w:t>` text == `""`). |
| 8 | `set_all_text_black` | Iterates all `<w:r>` in body. Removes any `<w:rStyle>` pointing to `"GuidanceChar"`. Upserts `<w:color w:val="000000"/>` inside `<w:rPr>`, creating `<w:rPr>` if absent. |
| 9 | `inject_definitions_fixed_rows` | Finds the definitions table by scanning `<w:p>` headings. For each `definitions_fixed_rows` entry: checks if first-column text already exists (case-insensitive). If missing: `deepcopy`s last row, replaces cell text nodes, appends to table. |
| 10 | `sort_tables_alphabetically` | For each heading in `sort_table_alphabetically_under_headings`: finds the first `<w:tbl>` after that heading. Extracts all `<w:tr>` except header. Sorts by `_para_text` of first `<w:tc>`. Re-appends rows in sorted order. |
| 11 | `remove_sections` | For each name in `_CFG["remove_sections"]`: finds the heading `<w:p>` by text match. Calls `_clear_section(heading_elem)` then removes the heading element itself. |

**Internal helpers used across passes:**
```python
def _para_style_id(elem) -> str:
    # elem.find(f"{{{_NS}}}pPr").find(f"{{{_NS}}}pStyle").get(f"{{{_NS}}}val")

def _para_text(elem) -> str:
    # "".join(t.text or "" for t in elem.iter(f"{{{_NS}}}t")).strip()

def _heading_style_ids(doc) -> set:
    # {s.style_id for s in doc.styles if s.name and s.name.startswith("Heading")}
```

---

### Data Flow — Object Types at Each Boundary

```
main()
  │
  │  argparse.Namespace
  │    .template    str  (abs path)
  │    .input       str  (abs path)
  │    .threshold   float
  │    .product_name  str | None
  │
  ├─ extract_headings(template_path)
  │    → list[{"text": str, "level": int, "paragraph_index": int}]   # N entries
  │
  ├─ extract_headings(data_path)
  │    → list[{"text": str, "level": int, "paragraph_index": int}]   # M entries
  │
  ├─ load_model()
  │    → SentenceTransformer   (weights in ~/.cache/huggingface/)
  │
  ├─ match_headings(tmpl_headings, data_headings, model, threshold)
  │    internal:
  │      query_embs  : np.ndarray  shape (N, 768)   float32  L2-normalised
  │      corpus_embs : np.ndarray  shape (M, 768)   float32  L2-normalised
  │      sim_matrix  : np.ndarray  shape (N, M)     float64  cosine similarity
  │    → list[{"template_heading": dict, "matched_heading": dict|None, "score": float}]
  │
  └─ fill_template(template_path, data_path, matches, output_path, product_name)
       │
       │  template       : docx.Document   (in-memory python-docx object)
       │  all_image_parts: list[{"placeholder_rid", "bytes", "content_type"}]
       │
       ├─ [for each match, reverse order]
       │    extract_section(data_path, para_idx, level)
       │    → {"items": [{"type": str, "xml": bytes, "image_parts": list}]}
       │         xml is etree.tostring() bytes of <w:p> or <w:tbl>
       │         image_parts: placeholder rIds written into the xml
       │
       │    _clear_section(heading._p)      OR      partial clear (keep_n_tables)
       │         operates on: heading._p.getparent()  (the <w:body> lxml element)
       │
       │    _insert_xml_after(anchor, xml_bytes)
       │         etree.fromstring(xml_bytes) → deepcopy → _strip_comments → addnext()
       │         anchor advances: anchor = new_elem
       │
       ├─ post_processor.apply_all(template, product_name)
       │         all 11 passes operate on template.element.body (lxml _Element)
       │
       ├─ template.save(BytesIO())
       │    → bytes   (valid .docx zip, but rId_CDRCOPY_... refs not yet resolved)
       │
       └─ _copy_table_images(bytes, all_image_parts)
            zipfile.ZipFile read → dict{name: bytes}
            find max rId in document.xml.rels → next_id
            for each unique placeholder:
                media_name = f"word/media/img_copied_rId{next_id}.{ext}"
                new_rels.append(<Relationship Id="rId{next_id}" .../>)
                rid_map[placeholder] = f"rId{next_id}"
            doc_text: str.replace(placeholder → new_rid)  [only rId_CDRCOPY_...]
            rels_text: insert new_rels before </Relationships>
            zipfile.ZipFile write → bytes
            → final bytes → open(output_path, "wb").write(...)
```

---

### Configuration Reference (`prompts.json`)

| Key | Type | What it controls |
|---|---|---|
| `model_name` | `str` | `SentenceTransformer` model identifier |
| `query_prefix` | `str` | Prepended to template heading texts before `model.encode()` |
| `product_name` | `str` | Passed as `product_name` arg to `fill_template()` and `apply_all()`; CLI `--product-name` overrides |
| `heading_mappings` | `dict[str, str]` | Forces `{ template_heading_text: data_heading_text }` matches at score 1.0 before AI scoring |
| `sections_keep_first_n_tables` | `dict[str, int]` | For named sections: keep first N `<w:tbl>` after heading, delete rest |
| `sections_filter_heading_content` | `list[str]` | For named sections: skip extracted `<w:p>` elements whose `<w:pStyle>` is a heading style |
| `post_processing.styles_to_remove` | `list[str]` | Style names → matched paragraphs removed from body |
| `post_processing.remove_template_instructions` | `bool` | Enable `< ... >` state-machine block removal |
| `post_processing.strip_inline_angle_brackets` | `bool` | Remove lone `<` / `>` `<w:r>` runs |
| `post_processing.remove_paragraphs_with_text` | `list[str]` | Substring match → paragraph removed |
| `post_processing.remove_preamble_before_first_heading` | `bool` | Delete body children before first heading `<w:p>` (except `<w:tbl>` and page-break paragraphs) |
| `post_processing.remove_empty_table_rows` | `bool` | Remove `<w:tr>` where all `<w:tc>` text is empty (skip row 0) |
| `post_processing.set_all_text_black` | `bool` | Upsert `<w:color w:val="000000"/>` in every `<w:rPr>`; remove `GuidanceChar` rStyle |
| `post_processing.sort_table_alphabetically_under_headings` | `list[str]` | Heading names → first `<w:tbl>` after each heading sorted by first-column text |
| `post_processing.definitions_fixed_rows` | `list[list[str, str]]` | `[[abbr, expansion], ...]` guaranteed rows in definitions table |
| `post_processing.definitions_heading` | `str` | Heading text used to locate the definitions `<w:tbl>` |
| `post_processing.remove_sections` | `list[str]` | Heading texts → heading `<w:p>` + all section content removed |
| `post_processing.header_footer_placeholders` | `list[str]` | Placeholder strings to find-and-replace with `product_name` |

---

### How to Run

```powershell
cd "BAAI BGE"

# Standard run — product_name read from prompts.json
python main.py `
  --template "Template_D001024021 CDR ISO 17664-2 (2021) ProductName RX.Y Rev C.docx" `
  --input    "Input_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx"

# One-off override
python main.py `
  --template "Template_D001024021 CDR ISO 17664-2 (2021) ProductName RX.Y Rev C.docx" `
  --input    "Input_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx" `
  --product-name "Azurion HW R4"
```

| Flag | Type | Default | Purpose |
|---|---|---|---|
| `--template` | `str` | required | Abs or relative path to template `.docx` |
| `--input` | `str` | required | Abs or relative path to data `.docx` |
| `--output` | `str` | `output/Filled_CDR.docx` | Output path; directory created with `os.makedirs(exist_ok=True)` |
| `--threshold` | `float` | `0.55` | Minimum cosine similarity to accept a BGE match |
| `--product-name` | `str` | `prompts.json["product_name"]` | Overrides config for this run only |

**Auto-versioning:**  `re.sub(r'_v\d+$', '', base)` strips any existing version suffix, then increments until a free filename is found — `Filled_CDR.docx` → `Filled_CDR_v2.docx` → …

---

### Key Technical Concepts

| Concept | Detail |
|---|---|
| **OOXML `.docx` structure** | ZIP archive: `word/document.xml` (body content), `word/_rels/document.xml.rels` (relationship map), `word/media/` (images), `word/styles.xml`, `word/header*.xml`, `word/footer*.xml` |
| **`w:` namespace** | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` — used by `<w:p>`, `<w:r>`, `<w:t>`, `<w:tbl>`, `<w:pStyle>`, `<w:rPr>`, `<w:color>`, `<w:br>`, etc. |
| **`a:` namespace** | `http://schemas.openxmlformats.org/drawingml/2006/main` — used by `<a:blip r:embed="...">` inside `<w:drawing>` |
| **`r:` namespace** | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` — `r:embed` attribute on `<a:blip>` holds the relationship ID |
| **`qn(tag)`** | `docx.oxml.ns.qn` expands Clark notation: `qn("w:p")` → `{http://...}p` |
| **BGE asymmetric encoding** | Query side (template headings) gets a retrieval prefix; corpus side (data headings) is plain text. Without the prefix on the query side, retrieval accuracy degrades. |
| **`normalize_embeddings=True`** | L2-normalises embeddings so `np.dot(a, b) == cosine_similarity(a, b)`. Required for `cosine_similarity()` from `sklearn` to give correct 0–1 scores. |
| **`lxml` vs `python-docx` API** | `python-docx` wraps lxml with a higher-level OO interface. Raw-inserted XML nodes (from `_insert_xml_after`) don't appear in `doc.paragraphs` / `doc.tables` until the document is saved and reloaded. All post-processor passes therefore use `doc.element.body` (lxml) directly. |
| **`rId_CDRCOPY_` prefix** | Guaranteed not to collide with Word-generated rIds (which are always `rId` + decimal digits only). Allows safe global `str.replace()` without regex. |
| **Zip-level image injection** | `python-docx` `.save()` finalises the document XML. After saving to `BytesIO`, the zip is reopened, images added, relationships updated, and the zip reassembled — all in memory. No temp files. |
