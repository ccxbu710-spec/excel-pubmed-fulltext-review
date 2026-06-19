# 泉方（本地 PubMed）机构全文补全流程

Use this reference when the user requests 泉方、tsgyun、本地 PubMed、`pm.yuntsg.com`、个人中心、我申请的文献, or a similar library document-delivery workflow.

## Browser Route

1. Use Chrome when the user has already logged in and the workflow depends on cookies/session state.
2. Open the institution portal or direct resource page. Common URLs seen in this workflow:
   - Institution portal: `https://user.tsgyun.com/user/login?insid=34`
   - 泉方本地 PubMed: `https://pm.yuntsg.com/`
   - Personal center/request shelf: `https://user.tsgyun.com/center/bookshelf?activeLi=1`
3. Confirm the page shows the user as logged in before applying.
4. In 本地 PubMed, search each failed item by PMID. If PMID fails, search by DOI, then exact title.
5. Select the matching result. If both `全文链接` and `申请全文` are visible, prefer
   `申请全文` first when `全文链接` appears to require publisher or institutional
   verification; `申请全文` can directly open a PDF download/viewer page.
6. If `申请全文` opens a PDF/download viewer, use the visible download/save action
   and monitor `~/Downloads`. If it only submits a request, record success/failure
   per PMID. A typical success message says the literature has been opened in a new
   page and added to `个人中心-我申请的文献`.
7. If `全文链接` was clicked and lands on an institution verification or publisher
   login page, go back to the result page and try `申请全文` before continuing with
   institution selection.
8. Wait the user/platform-required interval when the item is queued, often 5 minutes.
9. Open `个人中心` -> `我申请的文献`.
10. For each item with `申请成功`, click `打开全文`, then click `保存PDF`.
11. Move the downloaded PDF from the browser downloads folder into the pipeline's
    `下载全文/` target path. Use the exact target filename from
    `文献全文下载报告.json` or `机构全文补全清单.xlsx`.
12. Re-run `scripts/run_pipeline.py` to reuse the newly downloaded PDFs and refresh
    interpretations/report/summary.

## Publisher Redirect Notes

- ScienceDirect: if an `Are you a robot?` Cloudflare page appears, pause for the
  user to solve it manually. After the user confirms completion, wait 30-60 seconds,
  reopen the canonical article URL, click `View PDF`, then monitor `~/Downloads`.
  Do not automate the CAPTCHA click and do not store signed temporary PDF URLs.
- Wiley: a `/doi/pdf/...` page may expose an `Open` or `pdfdirect` link. Opening
  `pdfdirect` can trigger Chrome's PDF download even when browser automation reports
  `net::ERR_ABORTED`.
- Lancet/Cell/Elsevier partner journals: full text pages often expose
  `/action/showPdf?pii=...`; open that URL in Chrome and monitor downloads.

## Practical Notes

- Keep the Excel citation order as the authoritative order.
- Prefer PMID as the unique search key because title punctuation and early-online metadata often differ.
- If duplicate PDFs appear in Downloads, move only the newest matching file into the project folder unless the user asks to clean personal downloads.
- Do not expose, inspect, or export cookies/passwords. Use only the visible authenticated browser session.
- Validate every copied PDF with `%PDF-`, page count when possible, and size > 10 KB.
- If a row still has no `申请成功` after the wait, leave it in the checklist and explain the status.
