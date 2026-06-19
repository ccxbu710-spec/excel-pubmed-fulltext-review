---
name: excel-pubmed-fulltext-review
description: Download and organize academic full texts from an Excel workbook containing PubMed links, PMIDs, DOIs, titles, or review citation order; use scansci-pdf, PubMed metadata, open-access sources, and institutional access workflows to fill missing PDFs, then generate per-paper key interpretation Markdown files plus a final Excel summary. Use when the user asks to batch download papers from Excel/PubMed, apply through an institutional library or WebVPN/泉方/本地PubMed for failed full texts, name PDFs by sequence and title, or summarize/classify downloaded literature into a workbook.
---

# Excel PubMed Fulltext Review

## Workflow

Use this skill for ordered literature-review batches driven by an Excel workbook.

1. Inspect the workbook and identify the sheet, citation-order column, article-title column, PMID/PubMed-link column, DOI column if present, and any existing review-extraction columns.
2. Run `scripts/run_pipeline.py` to create output folders, fetch PubMed metadata, download legal/open PDFs with PMC, Unpaywall, and `scansci-pdf`, extract PDF text, create per-paper key-reading Markdown files, write a JSON report, and build a summary workbook.
3. For failed PDFs, retry in this order unless the user specifies otherwise: PMC/PMCID and Unpaywall, `scansci-pdf` smart/fastest or legal-only download, `scansci-pdf` Sci-Hub/Tor or institutional login when appropriate, then browser-based institutional access.
4. For institutional failures, generate and use the institutional checklist. If the user has an authenticated browser session, use the institution/library site exactly as requested. For 泉方/本地PubMed, follow `references/quanfang_workflow.md`; its current preferred route is to try `申请全文` before `全文链接` when `全文链接` leads to publisher or institution verification.
5. Move institutionally downloaded PDFs into `下载全文/` with the script's exact target filename from the report/checklist.
6. Re-run `scripts/run_pipeline.py` on the same workbook/output directory. Existing valid PDFs are reused, new files are interpreted, and the final summary workbook is refreshed.
7. Verify counts before final delivery: workbook row count, PDF count, Markdown interpretation count, summary row count, and report success count must match unless the user accepts unresolved failures.

## Quick Start

Prefer the bundled script when the workbook has conventional Chinese columns:

```bash
python3 /path/to/excel-pubmed-fulltext-review/scripts/run_pipeline.py \
  --input "/absolute/path/literature.xlsx" \
  --sheet "综述引用（按引文顺序）" \
  --output-dir "/absolute/path/output"
```

If column names differ, pass explicit mappings:

```bash
python3 /path/to/excel-pubmed-fulltext-review/scripts/run_pipeline.py \
  --input "/absolute/path/literature.xlsx" \
  --sheet "References" \
  --seq-col "序号" \
  --title-col "文章题目" \
  --pmid-col "PMID" \
  --pubmed-url-col "PubMed链接" \
  --doi-col "DOI" \
  --output-dir "/absolute/path/output"
```

The script writes:

- `下载全文/`: full-text PDFs named as `序号_文章题目.pdf`
- `全文文本提取/`: extracted text files
- `文献重点解读/`: one Markdown interpretation per paper
- `文献全文下载报告.json`: status, target paths, metadata, and errors
- `全文文献重点解读汇总.xlsx`: ordered summary by sequence, title, and interpretation
- `机构全文补全清单.xlsx`: only when some PDFs remain missing

## Institutional Fill

Use institutional access only for PDFs that remain failed after legal/open-source attempts. Keep the user's requested access route authoritative. Do not switch to a different document-delivery platform if the user specifies one.

When using a browser-based institutional route:

1. Open the authenticated institution/library site in Chrome if the task depends on existing login state.
2. Search by PMID first; use DOI or exact title only if PMID search fails.
3. Apply/request full text for all failed rows.
4. Wait the platform's required processing time.
5. Open the personal center/requested-literature list, download each successful full text, and move it to the target `PDF路径` or `目标PDF文件名` from `文献全文下载报告.json` or `机构全文补全清单.xlsx`.
6. Re-run the pipeline and verify all statuses.

Read `references/quanfang_workflow.md` when the user mentions 泉方, 本地PubMed, tsgyun, `pm.yuntsg.com`, or “我申请的文献”.

## scansci-pdf Retry Pattern

Use `scansci-pdf` before institutional browser work unless the user explicitly asks to skip it.

- Prefer `scansci_pdf_smart_download` or CLI `python -m scansci_pdf get DOI --output ...` for DOI rows.
- Use `strategy="legal_only"` when the user wants only legal/open routes; otherwise smart/fastest may try PMC, Unpaywall, OpenAlex, publisher direct links, LibGen/Sci-Hub, and Tor depending on configuration.
- For paywalled publisher failures, `scansci_pdf_login(identifier=DOI)` can open a browser login flow; the user must complete SSO/CAPTCHA steps manually, then retry download with the same DOI.
- For ScienceDirect/Elsevier-heavy batches, note that an Elsevier API key can avoid many browser challenges, but do not configure or store keys unless the user provides them.
- When a scansci result succeeds, validate the file starts with `%PDF-`, copy it to the report's `目标PDF路径`, and rerun the pipeline.

## Interpretation Rules

The script creates a structured first-pass interpretation from workbook extraction columns, PubMed abstract, DOI/PMCID, PDF path, and extracted-text length. If the user asks for deeper scientific reading, enrich each Markdown file using the extracted full text, preserving the same filename and keeping the summary workbook synchronized afterward.

For medical literature, keep interpretations evidence-focused:

- State purpose, design/type, methods, cohort/object, main findings, conclusion/significance, novelty, and relevance to the review topic.
- Do not invent unavailable sample sizes, endpoints, or mechanistic claims.
- Flag rows with weak text extraction or mismatched titles for manual review.

## Verification

Before final response, run or reproduce these checks:

```bash
find "/output/下载全文" -maxdepth 1 -type f -name '*.pdf' | wc -l
find "/output/文献重点解读" -maxdepth 1 -type f -name '*.md' | wc -l
python3 - <<'PY'
import json, pandas as pd
from pathlib import Path
out = Path("/output")
report = json.loads((out / "文献全文下载报告.json").read_text(encoding="utf-8"))
print(len(report), sum(r.get("下载状态") == "success" for r in report))
df = pd.read_excel(out / "全文文献重点解读汇总.xlsx")
print(len(df), df["下载状态"].value_counts(dropna=False).to_dict())
PY
```

Report final paths, success count, any unresolved failures, and whether the summary workbook was refreshed.
