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


