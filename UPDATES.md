# Daily Updates

---

## April 8, 2026

- **Problem Statement:** Automate filling the template doc by existing relevant information from the data doc — including both **text content and images**.
- **How:** We will have headings in the template doc — by text similarity we will search similar headings in the data doc and paste the respective **text data and images** under the matched heading in the template.

---

## April 10, 2026

**Model Comparison Table:**

| Requirement      | GPT-4o-mini (Recommended) | LayoutLMv3               | Qwen2.5-VL                | BAAI BGE Embedding Models    |
|------------------|---------------------------|--------------------------|---------------------------|------------------------------|
| Heading Matching | Excellent (Semantic)      | Poor (Position-based)    | Good (Visual)             | Excellent (Semantic)         |
| Setup Speed      | Fast (Hours)              | Slow (Weeks of labeling) | Medium (Server setup)     | Fast (Hours)                 |
| Hardware Cost    | $0 (Cloud-based)          | High (Requires GPU)      | Very High (Requires A100) | $0 (CPU, no GPU needed)      |
| Usage Cost       | Low (Pay-per-use)         | $0 (Once hosted)         | $0 (Once hosted)          | $0 (Free, open-source)       |

> **GPT-4o-mini Variants:**
> - `gpt-4o-mini` → lightweight, fast, cheapest (recommended for this use case)
> - `gpt-4o` → higher accuracy, more expensive
> - `gpt-4-turbo` → older, slower, higher cost

> **LayoutLMv3 Variants:**
> - `layoutlmv3-base` → smaller, faster (less accurate)
> - `layoutlmv3-large` → best accuracy (needs more GPU VRAM)
> - `layoutlmv3-base-finetuned-funsd` → pre-fine-tuned on form understanding

> **Qwen2.5-VL Variants:**
> - `Qwen2.5-VL-3B` → lightest, runs on smaller GPUs
> - `Qwen2.5-VL-7B` → balanced accuracy/performance
> - `Qwen2.5-VL-72B` → best accuracy (requires A100/H100)

> **BGE Model Variants:**
> - `bge-base-en-v1.5` → balanced (faster, less RAM)
> - `bge-large-en-v1.5` → best accuracy (needs more RAM)

---

## April 14–27, 2026

**Model selected:** `bge-base-en-v1.5` (BAAI BGE) — CPU-only, free, fast enough for CDR heading counts.

### Full Pipeline Built (`BAAI BGE/`)

| Step | Module | What it does |
|---|---|---|
| 1 | `heading_extractor.py` | Extract all `Heading`-styled paragraphs from any `.docx` as `[{text, level, paragraph_index}]` |
| 2 | `heading_matcher.py` | Encode template + data headings with BGE; cosine-similarity matrix; explicit `heading_mappings` overrides at score 1.0 |
| 3 | `content_extractor.py` | Deep-copy + serialise every `<w:p>` / `<w:tbl>` in a section as raw XML bytes; remap image rIds to `rId_CDRCOPY_` placeholders; add `_NUMID_OFFSET` sentinel to all `<w:numId>` |
| 4 | `template_filler.py` | Insert extracted XML after matched template headings (reverse order); strip inline `<w:sectPr>`; column-remap tables; inject pre-heading content; post-process; zip-level image copy + numbering merge |
| 5 | `post_processor.py` | 15 cleanup passes: placeholder replacement, preamble removal, style removal, instruction block removal, angle-bracket stripping, paragraph text removal, empty row removal, NOTE font-size reduction, superscript marker removal, force-black text (body + headers/footers), blank-para-after-heading removal, definitions injection, table sort, section removal, TOC dirty-marking |

### Key Bugs Fixed

| Bug | Fix |
|---|---|
| `"Compliance Checklist"` not filled (score 0.509 < threshold) | `heading_mappings` explicit override in `prompts.json` |
| Cover page `<ProductName RX.Y>` not replaced | lxml multi-run collapse algorithm; scans `doc.element.body` in addition to headers/footers |
| Cover logo overwritten by data-doc image | `rId_CDRCOPY_` placeholder system — only placeholders are replaced at zip level |
| Bullet/numbered lists rendered as plain text | `_NUMID_OFFSET` sentinels + `_merge_numbering()` zip-level pass copies `abstractNum`/`num` from data doc |
| "Record history" table layout wrong | `sections_remap_table_columns` + `_build_remapped_table()` — keeps template column structure |
| Data-doc header/footer rIds in output | `_strip_inline_sectpr()` strips `<w:pPr><w:sectPr>` from all extracted elements before insertion |
| Numbered list items with superscript markers | `strip_superscript_list_markers` pass removes `<w:vertAlign val="superscript"/>` from leading numeric runs only |
| Purple/pink text in document headers/footers | `set_all_text_black` (Pass 10) extended to loop over all header/footer parts via `doc.sections` |
| TOC not updating on open | `mark_toc_dirty` (Pass 15) sets `w:dirty="1"` on TOC `begin` field chars |
| Extra blank lines after section headings | `remove_blank_paragraph_after_headings` (Pass 11) removes all consecutive empty paragraphs after any heading |
| Sub-headings dropped from Compliance Checklist | `sections_filter_heading_content` cleared to `[]`; all heading-styled elements now preserved during extraction |

### Configuration (`prompts.json`)

All pipeline behaviour is config-driven — no hardcoded document-specific values in code. Key config keys: `model_name`, `threshold`, `query_prefix`, `heading_mappings`, `sections_remap_table_columns`, `sections_keep_first_n_tables`, `sections_filter_heading_content`, and the full `post_processing` block (15 flags/lists).

> `product_name` is **no longer a config key** — it is auto-detected at runtime from each input document's Word core properties (title → subject → description). No hardcoded value needed in `prompts.json`.

### Streamlit UI (`app.py`)

- Run with `streamlit run "BAAI BGE/app.py"`
- Upload **1 CDR template** + **any number of test-record inputs**
- Pipeline runs per-input: heading extraction → BGE matching → template fill
- Per-file collapsible result panels with metrics + colour-coded match table (🟢 ≥ 0.85 · 🔵 ≥ 0.70 · 🔴 < 0.70)
- Individual download button per output; **ZIP download** when multiple outputs succeed
- **Threshold slider** in sidebar (default from `prompts.json`, live override before each run)
- **Product name** auto-detected from input doc's Word core properties; shown per-file in results



