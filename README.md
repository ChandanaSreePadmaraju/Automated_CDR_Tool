# Automated CDR Tool

Automatically fills **Compliance Data Record (CDR) template documents** by extracting relevant content — text, tables, images, and lists — from a source test-record document and inserting it under the correct headings, using AI-powered semantic matching.

---

## What It Does

In medical device compliance workflows, engineers manually copy content from a filled *Test Record* document into a structured *CDR template*. This tool automates that process end-to-end:

1. **Extracts headings** from both the CDR template and the source test-record (`.docx`)
2. **Semantically matches** headings using the [`BAAI/bge-base-en-v1.5`](https://huggingface.co/BAAI/bge-base-en-v1.5) embedding model — so headings with different wording but the same meaning are correctly paired (e.g. *"Compliance Checklist"* ↔ *"Test Record"*)
3. **Extracts section content** (paragraphs, tables, images, bullet/numbered lists) from the test record
4. **Fills the template** with that content, preserving layout, images, and formatting
5. **Post-processes** the output: cleans up instructions, fixes fonts, forces black text, refreshes the Table of Contents, and more

---

## Project Structure

```
BAAI BGE/
├── app.py              # Streamlit web UI — recommended for most users
├── main.py             # CLI entry point
├── prompts.json        # All pipeline configuration (no hardcoded values in code)
├── requirements.txt    # Python dependencies
├── Dockerfile          # Docker image definition
├── .dockerignore
├── src/
│   ├── heading_extractor.py   # Step 1 — extract headings from .docx
│   ├── heading_matcher.py     # Step 2 — semantic heading matching via BGE
│   ├── content_extractor.py   # Step 3 — extract section content as raw XML
│   ├── template_filler.py     # Step 4 — insert content into template
│   └── post_processor.py      # Step 5 — 15 post-processing cleanup passes
└── output/                    # Generated output files (git-ignored)
docker-compose.yml
```

---

## Requirements

- Python 3.11+
- No GPU required — runs entirely on CPU

---

## Installation

```bash
# 1. Clone the repository
git clone <repository link>.git
cd Automated_CDR_Tool/Qwen

# 2. Create and activate a virtual environment
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate

# 3. Install dependencies
pip install -r "BAAI BGE/requirements.txt"
```

The first run will automatically download the `BAAI/bge-base-en-v1.5` model (~440 MB) from Hugging Face and cache it locally.

---

## Usage — Web UI (Recommended)

```bash
streamlit run "BAAI BGE/app.py"
```

Then open `http://localhost:8501` in your browser.

### Steps

1. **Upload CDR Template** — the structured `.docx` skeleton with Heading-styled sections and `<ProductName RX.Y>` placeholders
2. **Upload Test Record(s)** — one or more filled source `.docx` files
3. *(Optional)* Adjust the **match threshold** slider in the sidebar (default: 0.60)
4. Click **Run** — the pipeline processes each input file sequentially
5. **Download** individual output files or use **Download All as ZIP** for multiple outputs

Each result panel shows:
- Auto-detected product name
- Colour-coded heading match table (🟢 ≥ 0.85 · 🔵 ≥ 0.70 · 🔴 < 0.70)
- Match metrics (matched / unmatched / skipped)

---

## Usage — Command Line

```bash
cd "BAAI BGE"

# Single template + single input
python main.py --template Template.docx --input "Test Record.docx"

# Multiple templates → produces op1_..., op2_..., op3_... outputs
python main.py --templates t1.docx t2.docx t3.docx --input "Test Record.docx"

# Override product name manually
python main.py --template Template.docx --input "Test Record.docx" --product-name "Azurion HW R3"

# Override match threshold
python main.py --template Template.docx --input "Test Record.docx" --threshold 0.65
```

Output files are saved to the `output/` folder, named `op1_<input filename>.docx`, `op2_...`, etc.

---

## Usage — Docker

```bash
# Build and start
docker-compose up --build

# Access the UI
open http://localhost:8501
```

---

## How Product Name Is Detected

The tool automatically extracts the product name from the input filename using the pattern `(YYYY) <Product Name>`:

```
Input_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx
                                                 ↑ detected → "Azurion HW R3"
```

This replaces the `<ProductName RX.Y>` placeholder throughout the output document. You can override this with `--product-name` on the CLI or by editing the detected value in the UI.

---

## Output File Naming

Output files are named based on the input filename with the product name portion stripped:

```
Input1_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx
→ op1_Input1_D001352871 Test Record ISO 17664-2 (2021).docx
```

If the same output already exists, the `op` index increments automatically (`op1_`, `op2_`, ...).

---

## Notes

- **Upload size limit (UI):** Streamlit's default is 200 MB total across all uploads. Typical `.docx` files are 1–5 MB (40–200 files per batch). To raise the limit: `streamlit run "BAAI BGE/app.py" --server.maxUploadSize 500`
- **First run is slower** due to model download and caching. Subsequent runs are fast.
- Sample input/template `.docx` files are excluded from the repository (`.gitignore`) — provide your own documents.
