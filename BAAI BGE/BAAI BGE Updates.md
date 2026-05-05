# BAAI BGE Updates

---

## May 5, 2026 (Session 2)

### Multi-file Upload Support (`app.py`)

- **1 template + N input files** — `accept_multiple_files=True`; no file-count cap (removed `[:5]` limit)
- **Per-file expanders** — each input gets its own collapsible result panel with metrics + colour-coded match table
- **Individual download button** per output file (`op_<input_name>.docx`)
- **"Download All as ZIP"** button appears when ≥ 2 outputs succeed — bundles all via `zipfile.ZipFile`
- **Progress bar** shows `Processing 2/5: InputX.docx …` step-by-step across all inputs

### Auto Product-Name Detection (`app.py` + `main.py`)

- `product_name` is now **auto-detected** from each input document's Word core properties: tries `title` → `subject` → `description` in order
- If none found, falls back to `product_name` in `prompts.json` (key now optional / can be absent)
- **UI**: detected product name displayed as a caption inside each per-file expander
- **CLI**: detection runs on the `--input` path; `--product-name` flag still allows manual override
- Removed `auto_detect_product_name()` from `import io` dependency in `main.py` (uses `docx.Document` directly from a file path)

### Removed Hardcoded `product_name` from `prompts.json`

- `"product_name": "Azurion HW R3"` line deleted from `prompts.json`
- Config is now fully product-agnostic; the key is simply absent

### Threshold Slider in Sidebar (`app.py`)

- Live **match threshold slider** (0.0 – 1.0, step 0.05) in the left sidebar
- Default pre-filled from `prompts.json` `threshold` value at startup
- Adjustable before each Run — no restart needed
- Colour key shown in sidebar: 🟢 ≥ 0.85 · 🔵 ≥ 0.70 · 🔴 < 0.70

### `.gitignore` Updated

- Added `BAAI BGE/Input*.docx`, `BAAI BGE/Template*.docx`, `BAAI BGE/Review*.docx` exclusions
- Prevents binary sample/test data files from being tracked in git

### Code Cleanup (`main.py`)

- Removed unused `import re` and `import io`
- `auto_detect_product_name()` added as a standalone function before `main()`

---

## May 5, 2026

### Streamlit UI — Full Rework (`app.py`)

- **Wide layout** — `layout="wide"` so the full browser width is used
- **Colourful theme** — deep purple/blue gradient background, gradient title text, step badges, colour-coded score cells (green ≥ 0.85, blue ≥ 0.70, red < 0.70)
- **Session state** — results stored in `st.session_state`; cleared on every new Run click so stale results can never be downloaded
- **Error propagation** — `run_pipeline` returns a descriptive error string; shown to user on failure
- **Output filename** — always `op_<exact input filename>` — no product-name stripping, no `op1_`/`_v2` counter
- **Processed-file banner** — info box confirms which template + input were actually used after each run
- **`pandas` added** to `requirements.txt` (used for the colour-coded match table in the UI)
- **`applymap` → `map`** fix for pandas ≥ 2.1 compatibility
- **Output filename fix** — removed hardcoded `product_name` lookup from `derive_output_name`; filename is now fully automatic from the uploaded input file

### Auto-increment `op` index (`main.py`)

- When the output file already exists, the `op` prefix index is incremented (`op1_` → `op2_` → …) instead of appending `_v2`/`_v3`
- Old `_v2` suffix logic removed entirely

### `remove_blank_paragraph_before_headings` (Pass 11a, `post_processor.py`)

- New pass added before Pass 11b (`remove_blank_paragraph_after_headings`)
- Removes all consecutive blank paragraphs immediately **before** any heading
- Controlled by `"remove_blank_para_before_headings": true` in `prompts.json`
- Fixes unnecessary page gap between "Definitions & abbreviations" and "Compliance Checklist"

### Installed Package Versions (venv)

| Package | Version |
|---|---|
| python-docx | 1.2.0 |
| torch | 2.11.0 |
| sentence-transformers | 5.4.1 |
| scikit-learn | 1.8.0 |
| numpy | 2.4.4 |
| lxml | 6.1.0 |
| streamlit | 1.57.0 |
| pandas | 3.0.2 |

---

## May 4, 2026

### Streamlit Web UI (`app.py`)

- New `app.py` added — run with `streamlit run app.py`
- Two file-upload widgets: **Template (.docx)** and **Input / Test Record (.docx)**
- Full pipeline runs in-browser: heading extraction → BGE matching → template fill
- Live log panel shows step-by-step progress and the heading match table
- Download button produces `op_<input_stem>.docx` (product-name suffix stripped, no `op1_`/`op2_` counter — clean single output per run)
- Model loaded once and cached across runs via `@st.cache_resource`
- `streamlit>=1.35.0` added to `requirements.txt`

### `remove_blank_paragraph_before_headings` (Pass 11a)

- New post-processing pass added to `post_processor.py` before Pass 11b (`remove_blank_paragraph_after_headings`)
- Removes all consecutive blank paragraphs that appear immediately **before** any heading

- Eliminates the unwanted page gap that appeared between "Definitions & abbreviations" and "Compliance Checklist"
- Controlled by `"remove_blank_para_before_headings": true` in `prompts.json` (enabled by default)

### Output Filename: op-index increment replaces `_v2` suffix

- `main.py` auto-increment logic changed: when `op1_<name>.docx` already exists, the next run produces `op2_<name>.docx`, then `op3_`, etc.
- The old `_v2` / `_v3` versioning suffix is removed

---

## April 27, 2026

### Force-Black Text Extended to Headers / Footers

- `set_all_text_black()` (Pass 10) now loops over `doc.sections` and applies the same `<w:color w:val="000000"/>` injection to every header and footer part (`header`, `even_page_header`, `first_page_header`, `footer`, `even_page_footer`, `first_page_footer`)
- Fixes purple/pink coloured text that appeared in the document header (product name line) and footer (revision line)

### New Pass 11 — `remove_blank_paragraph_after_headings`

- New post-processing pass controlled by `"remove_blank_para_after_headings": true` in `prompts.json`
- Removes **all consecutive empty paragraphs** that appear immediately after any Heading-styled paragraph
- Uses a while-loop + re-scan approach so multiple blank lines after one heading are all removed in one call
- Preserves layout paragraphs (inline `<w:sectPr>`) and page-break paragraphs

### New Pass 15 — `mark_toc_dirty`

- New post-processing pass added unconditionally at the end of `apply_all()`
- Iterates all body paragraphs looking for field instructions (`<w:instrText>`) that start with `"TOC"`
- Sets `w:dirty="1"` on the `begin` `<w:fldChar>` so Word automatically refreshes the Table of Contents on first open

### `sections_filter_heading_content` Cleared

- Config key `"sections_filter_heading_content"` changed from `["Compliance Checklist"]` to `[]`
- Previous value caused sub-headings (**Test Administration**, **Test Measurements**, **Test Result**) inside the "Test Record" section to be silently dropped during extraction
- Now all heading-styled paragraphs inside extracted sections are preserved in the output

### Pass Numbering Cleanup

- All pass-number comments in `post_processor.py` now match the actual execution order in `apply_all()`:

| Pass | Function |
|---|---|
| 1 | `fill_header_footer_placeholders` |
| 2 | `remove_preamble_before_first_heading` |
| 3 | `remove_styled_paragraphs` |
| 4 | `remove_template_instructions` |
| 5 | `strip_inline_angle_brackets` |
| 6 | `remove_paragraphs_with_text` |
| 7 | `remove_empty_table_rows` |
| 8 | `set_note_text_size` |
| 9 | `strip_superscript_list_markers` |
| 10 | `set_all_text_black` |
| 11 | `remove_blank_paragraph_after_headings` |
| 12 | `inject_definitions_fixed_rows` |
| 13 | `sort_tables_alphabetically` |
| 14 | `remove_sections` |
| 15 | `mark_toc_dirty` |

- `mark_toc_dirty` moved to before `apply_all()` in the file (was incorrectly defined after the function that calls it)

---

## April 27, 2026

### Bullet / Numbered List Fix — `_merge_numbering` + `_NUMID_OFFSET`

- **Problem:** Extracted sections containing bullet or numbered lists rendered as plain paragraphs in the output — the numbering definitions (`abstractNum`/`num` entries in `word/numbering.xml`) were missing from the output document.
- `_NUMID_OFFSET = 10_000` sentinel added to `content_extractor.py` — every `<w:numId w:val="N"/>` in an extracted element is rewritten to `N + 10000` at extraction time via `_offset_numids(deep)`
- New zip-level pass `_merge_numbering()` in `template_filler.py`: reads the data doc's `word/numbering.xml`, copies needed `abstractNum`/`num` definitions into the output doc, rewrites sentinels to real sequential IDs
- Adds the `word/numbering.xml` relationship to the output doc if not already present

### Table Column Remapping — `sections_remap_table_columns`

- New config key `"sections_remap_table_columns"` in `prompts.json`
- When set for a section, `_build_remapped_table()` rebuilds the table using the **template's** header row, `tblPr`, column widths (`tcPr`), and borders — but fills data rows from the data doc using only the listed column indices
- Without this, "Record history" would paste the raw data-doc table (different column count and layout) instead of respecting the CDR template's table structure

### Pre-Heading Content Insertion

- `extract_pre_heading_content()` added to `content_extractor.py`
- Extracts all body elements that appear **before the first heading** in the data doc (e.g. document-title metadata paragraphs)
- Inserted immediately **before the first heading** in the template output (in original order)
- Leading and trailing empty paragraphs are stripped from the pre-heading block

### Inline `<w:sectPr>` Stripping — `_strip_inline_sectpr`

- New helper `_strip_inline_sectpr()` in `template_filler.py`
- Removes `<w:pPr><w:sectPr>` from every extracted element before insertion
- Prevents data-doc header/footer rIds from polluting the output document (would break the output's header/footer display)
- Called inside `_insert_xml_after()` so it applies automatically to every inserted element

### New Post-Processing Passes

| Pass | Function | What it does |
|---|---|---|
| 8 | `set_note_text_size` | Reduces font size of every paragraph starting with a NOTE prefix (e.g. `"NOTE"`, `"NOTE 1"`) to `note_text_size_half_pt` half-points (`18` = 9 pt). Sets `<w:sz>` and `<w:szCs>` on all runs. Applies to body paragraphs and table cells. |
| 9 | `strip_superscript_list_markers` | Removes `<w:vertAlign val="superscript"/>` **only from leading** numeric runs (e.g. `"1"`, `"1."`, `"2."`). Preserves trailing footnote-reference superscripts at the end of a paragraph. |

### Updated `remove_preamble_before_first_heading`

- Now keeps the **first** and **last** page-break paragraphs before the first heading (previously only the last was kept)
- Template layout: Cover → [break1] → instruction page → [break2] → TOC → [break3] → content. After removing the instruction page, both `break1` (Cover→TOC) and `break3` (TOC→content) must survive.

### `threshold` moved to `prompts.json`

- `"threshold": 0.6` is now a top-level key in `prompts.json`
- `main.py` reads it at startup; CLI `--threshold` still overrides for one-off runs
- Priority: CLI flag > `prompts.json`

### New `prompts.json` keys

| Key | Type | Purpose |
|---|---|---|
| `threshold` | `float` | Minimum cosine similarity; replaces the hardcoded default in `main.py` |
| `sections_remap_table_columns` | `dict[str, list[int]]` | Column-remap sections: keeps template table structure, maps data columns |
| `post_processing.note_text_prefixes` | `list[str]` | Paragraph prefixes that trigger font-size reduction (e.g. `"NOTE"`) |
| `post_processing.note_text_size_half_pt` | `int` | Target half-point size for NOTE paragraphs (18 = 9 pt) |
| `post_processing.strip_superscript_list_markers` | `bool` | Enable leading-superscript removal pass |

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
│   └── post_processor.py      # Step 5 — 13 post-processing cleanup passes
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
  │     flags: --template, --input, --output, --threshold (default: prompts.json["threshold"]), --product-name
  │
  ├─ json.load("prompts.json")  →  product_name, threshold  (CLI flags override)
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
        │     _clear_section(heading_elem)   or   partial-clear (keep_n_tables / remap_cols)
        │     _insert_xml_after(anchor, item["xml"])   ×N   [_strip_inline_sectpr applied]
        │
        ├─ extract_pre_heading_content(data_path)  →  pre items inserted before first heading
        │
        ├─ post_processor.apply_all(template_doc, product_name)   [13 passes]
        │
        ├─ template_doc.save(BytesIO)   →   bytes
        ├─ _copy_table_images(bytes, all_image_parts)   →   bytes  (images resolved)
        └─ _merge_numbering(bytes, data_path)   →   final bytes   →   file
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
    threshold:         float,
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

**Raw XML extraction — deep copy + serialise + offset numIds:**
```python
deep = copy.deepcopy(elem)          # lxml deep copy — isolates from live document
_offset_numids(deep)                # add _NUMID_OFFSET to every <w:numId> (list/bullet fix)
image_parts = _extract_and_remap_images(deep, doc)
body_items.append({
    "type":        "paragraph" | "table",
    "xml":         etree.tostring(deep),   # bytes: full <w:p> or <w:tbl> XML
    "image_parts": image_parts,
})
```

`_offset_numids()` adds `_NUMID_OFFSET = 10_000` to every `<w:numId w:val="N"/>` in the extracted element. This sentinel ensures the numId cannot collide with any existing definition in the template. `_merge_numbering()` in `template_filler.py` later reads these sentinels, locates the matching `abstractNum`/`num` entries in the data doc, copies them into the output doc with new sequential IDs, then rewrites the sentinels to the final IDs.

**Image rId placeholder system:**

A `.docx` `<a:blip r:embed="rId7">` references a relationship defined in `word/_rels/document.xml.rels`. The same `rId7` in the template may point to a completely different image (e.g. the cover logo). `_extract_and_remap_images()` rewrites every `r:embed` in the deep-copied element to a unique `rId_CDRCOPY_NNNNNN` placeholder before returning the image bytes. The counter is module-level — IDs are unique across all `extract_section()` calls in a single run.

**Trailing blank paragraph strip:** the last N empty (no text, no drawing) paragraphs are stripped before returning — avoids long runs of blank lines copied from the end of sections in the data doc.

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

**`extract_pre_heading_content(doc_path)`:**

New public function that extracts all `<w:p>` and `<w:tbl>` body elements appearing **before the first heading paragraph** in `doc_path`.

- Returns the same `{"items": [...]}` structure as `extract_section()`
- `_offset_numids` and `_extract_and_remap_images` are applied to every element
- Leading and trailing empty paragraphs are stripped
- Used in `fill_template()` to capture document-level metadata text (e.g. `"Philips Compliance Data Record  ISO 17664-2"`) that sits outside any section heading in the data doc

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

**Full clear** (`_clear_section`): removes all body elements after the heading element up to the next heading. Layout paragraphs (`<w:br w:type="page"/>` or inline `<w:sectPr>`) are always preserved — they define page boundaries and carry the template’s header/footer rId links.

**Partial clear** (`sections_keep_first_n_tables`): keeps the first N `<w:tbl>` elements after the heading and removes the rest — used when a section already has a template table that should be filled rather than replaced.

#### 4c — XML insertion

`_insert_xml_after(ref_elem, xml_bytes)` deserialises the bytes, strips comment markers and inline `<w:sectPr>` elements, then calls `ref_elem.addnext(new_elem)`. The `anchor` variable advances to the newly inserted element with each call, so subsequent items are appended in order.

`sections_filter_heading_content`: for sections like `"Compliance Checklist"`, any extracted `<w:p>` whose `<w:pStyle w:val>` is a heading style ID is skipped (sub-headings like “Test Administration” are not inserted).

#### 4d — Image injection (zip level)

`_copy_table_images(output_bytes, all_image_parts)` operates at zip level because `python-docx` has no API to add new image relationships after a document is assembled.

Algorithm: deduplicates image parts by `placeholder_rid` → reads the output zip → finds max existing rId number → writes each image to `word/media/img_copied_rId{n}.{ext}` → appends `<Relationship>` entries → text-replaces only `rId_CDRCOPY_...` strings in `word/document.xml` (never touches template rIds) → reassembles zip in memory.

---

### Step 5 — Post-Processing (`post_processor.py`)

**Module-level config** (read once at import time from `prompts.json["post_processing"]`). All passes operate on `doc.element.body` (raw lxml `_Element` tree) — not `python-docx`’s `.paragraphs` / `.tables` abstractions. This is necessary because content inserted via `_insert_xml_after` in Step 4 is raw lxml and does not appear in `doc.paragraphs` until the document is reloaded.

**Entry point:** `apply_all(doc, product_name)` — runs all 13 passes in order.

| # | Function | Key implementation detail |
|---|---|---|
| 1 | `fill_header_footer_placeholders` | Multi-run collapse: iterates every `<w:p>` in `doc.element.body` + all `doc.sections` header/footer parts. Concatenates run texts, finds the placeholder span by string index, collapses matching `<w:r>` nodes into one, deletes the rest. Handles Word splitting `<ProductName RX.Y>` across up to 4 separate `<w:r>` elements. |
| 2 | `remove_preamble_before_first_heading` | Finds first `<w:p>` with a Heading style in `body`. Removes all preceding body children except `<w:tbl>` (cover page). Keeps the **first** and **last** page-break paragraphs (preserves Cover→TOC and TOC→content breaks); removes all intermediate page-break and visible-text paragraphs. |
| 3 | `remove_styled_paragraphs` | Builds `{style_id: style_name}` from `doc.styles`. Removes `<w:p>` where `<w:pStyle w:val>` maps to a name in `_CFG["styles_to_remove"]` (e.g. `"Guidance"`). |
| 4 | `remove_template_instructions` | State-machine scan over body `<w:p>` elements. Opens a block when `_para_text(e)` == `"<"` or starts with `"< "`. Closes when text == `">"` or starts with `"> "`. Both markers and all paragraphs between them are removed. |
| 5 | `strip_inline_angle_brackets` | Removes any `<w:r>` whose sole `<w:t>` text is exactly `"<"` or `">"` from all body paragraphs. |
| 6 | `remove_paragraphs_with_text` | Removes `<w:p>` where `_para_text(e)` contains any string from `_CFG["remove_paragraphs_with_text"]`. `_para_text` joins all `<w:t>` text nodes. |
| 7 | `remove_empty_table_rows` | Iterates `doc.tables`. Skips the first row (header). Removes `<w:tr>` where every `<w:tc>` has empty text (joined `<w:t>` text == `""`). |
| 8 | `set_note_text_size` | Iterates every `<w:p>` in `doc.element.body`. If the joined text starts with a prefix from `_CFG["note_text_prefixes"]` (e.g. `"NOTE"`, `"NOTE 1"`): removes existing `<w:sz>`/`<w:szCs>` from all `<w:rPr>` children and inserts new ones with `val = note_text_size_half_pt` (18 = 9 pt). |
| 9 | `strip_superscript_list_markers` | Regex `r'^\d+[.\s]*$'` detects numeric list-marker runs. Removes `<w:vertAlign val="superscript"/>` only when the run is the first visible content in its paragraph (no preceding text). Trailing footnote-reference superscripts are preserved. |
| 10 | `set_all_text_black` | Iterates all `<w:rPr>` in body. Removes any `<w:rStyle>` pointing to a character style whose name contains a `styles_to_remove` prefix (e.g. `GuidanceChar`). Removes existing `<w:color>`, then inserts `<w:color w:val="000000"/>` as first child. |
| 11 | `inject_definitions_fixed_rows` | Finds the definitions table by scanning `<w:p>` headings. For each `definitions_fixed_rows` entry: checks if first-column text already exists (case-insensitive). If missing: `deepcopy`s last row, replaces cell text nodes (preserving run formatting), appends to table. |
| 12 | `sort_tables_alphabetically` | For each heading in `sort_table_alphabetically_under_headings`: finds the first `<w:tbl>` after that heading. Extracts all `<w:tr>` except header. Sorts by `_para_text` of first `<w:tc>`. Re-appends rows in sorted order. |
| 13 | `remove_sections` | For each name in `_CFG["remove_sections"]`: finds the heading `<w:p>` by text match. Removes heading + all following body elements up to the next heading (preserves inline `<w:sectPr>` paragraphs). |

---

### Configuration Reference (`prompts.json`)

| Key | Type | What it controls |
|---|---|---|
| `model_name` | `str` | `SentenceTransformer` model identifier |
| `query_prefix` | `str` | Prepended to template heading texts before `model.encode()` |
| `product_name` | `str` | Passed as `product_name` arg to `fill_template()` and `apply_all()`; CLI `--product-name` overrides |
| `threshold` | `float` | Minimum cosine similarity to accept a match; read at startup; CLI `--threshold` overrides |
| `heading_mappings` | `dict[str, str]` | Forces `{ template_heading_text: data_heading_text }` matches at score 1.0 before AI scoring |
| `sections_keep_first_n_tables` | `dict[str, int]` | For named sections: keep first N `<w:tbl>` after heading, delete rest |
| `sections_remap_table_columns` | `dict[str, list[int]]` | For named sections: rebuild table using template structure (header row, column widths) filled with the listed data-doc column indices |
| `sections_filter_heading_content` | `list[str]` | For named sections: skip extracted `<w:p>` elements whose `<w:pStyle>` is a heading style |
| `post_processing.styles_to_remove` | `list[str]` | Style names → matched paragraphs removed from body |
| `post_processing.remove_template_instructions` | `bool` | Enable `< ... >` state-machine block removal |
| `post_processing.strip_inline_angle_brackets` | `bool` | Remove lone `<` / `>` `<w:r>` runs |
| `post_processing.remove_paragraphs_with_text` | `list[str]` | Substring match → paragraph removed |
| `post_processing.remove_preamble_before_first_heading` | `bool` | Delete body children before first heading `<w:p>` (except `<w:tbl>` and page-break paragraphs) |
| `post_processing.remove_empty_table_rows` | `bool` | Remove `<w:tr>` where all `<w:tc>` text is empty (skip row 0) |
| `post_processing.note_text_prefixes` | `list[str]` | Paragraph prefixes that trigger font-size reduction (e.g. `"NOTE"`, `"NOTE 1"`) |
| `post_processing.note_text_size_half_pt` | `int` | Target half-point size for NOTE paragraphs (18 = 9 pt) |
| `post_processing.strip_superscript_list_markers` | `bool` | Remove `<w:vertAlign val="superscript"/>` from leading numeric runs only |
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
| `--threshold` | `float` | `prompts.json["threshold"]` | Minimum cosine similarity to accept a BGE match; reads config if not passed |
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
