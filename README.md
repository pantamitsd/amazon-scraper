# amazon-scraper

Fully automated Amazon product scraper. GitHub Actions reads `queries.xlsx`
from this repo, scrapes each URL with headless Chrome, and commits the
results back to `output.xlsx` — no Streamlit, no manual upload/download.

Same pattern as `maps-scraper`, minus the batching/state files (this scraper
re-runs the full `queries.xlsx` list every time; see "Notes" if you want
incremental/batched behavior like maps-scraper).

## Repository structure

```
amazon-scraper/
│
├── .github/
│   └── workflows/
│       └── scrape.yml
│
├── amazon_scraper.py
├── queries.xlsx
├── output.xlsx
├── requirements.txt
├── .gitignore
└── README.md
```

## Setup steps (exact)

1. **Create the repo** (or reuse an existing one) and push these files to the
   `main` branch.
2. **Add `queries.xlsx`** at the repo root with a column named `url` (or
   `Link`) — one Amazon product URL or 10-character ASIN per row. A starter
   file is included.
3. **Repo Settings → Actions → General → Workflow permissions** → select
   **"Read and write permissions"**. This is required or the final push step
   in the workflow will fail with a 403, exactly like the maps-scraper repo
   needed `permissions: contents: write`.
4. **Repo Settings → Actions → General → Actions permissions** → make sure
   Actions are allowed to run (Allow all actions and reusable workflows, or
   at minimum allow `actions/*` and `browser-actions/setup-chrome`).
5. Commit an empty/placeholder `output.xlsx` so the first run has something
   to diff against (already included).
6. Go to the **Actions** tab → **Amazon Scraper** workflow → **Run workflow**
   to trigger it manually the first time, or just wait for the schedule.

No API keys or secrets are needed for this scraper — it only reads public
Amazon product pages, so there's nothing to add under **Settings → Secrets
and variables → Actions**.

## GitHub Actions configuration

`.github/workflows/scrape.yml` runs on:
- `workflow_dispatch` — manual trigger from the Actions tab
- `schedule: cron: "0 */6 * * *"` — every 6 hours (edit the cron expression
  to change frequency; GitHub's scheduler can run a few minutes late,
  especially at busy times)

It does, in order:
1. Checkout the repo
2. Set up Python 3.11
3. Install current stable Google Chrome (`browser-actions/setup-chrome`)
4. `pip install -r requirements.txt`
5. `python amazon_scraper.py` (reads `queries.xlsx`, writes `output.xlsx`)
6. `git add output.xlsx` + commit + push, only if the file actually changed

## Permissions / secrets required

| Requirement | Where | Why |
|---|---|---|
| `permissions: contents: write` | already set in `scrape.yml` | lets the job push commits |
| Workflow permissions = Read and write | Repo Settings → Actions → General | without this the push step gets `remote: Permission denied` / 403, same failure mode you hit on maps-scraper |
| No secrets needed | — | scraper only hits public Amazon pages |

## What changed vs. the Streamlit version (and why)

| Change | Reason |
|---|---|
| Removed `import streamlit`, `main()` UI, `st.file_uploader`, `st.progress`, `st.download_button` | Streamlit is a local web-server UI framework — it has nothing to attach to on a headless GitHub Actions runner and isn't needed for a scheduled job. |
| `scrape_amazon_file(uploaded_file)` → `scrape_amazon_file(input_path)`, reads `queries.xlsx` directly via `pd.read_excel(input_path)` | No user is present to upload a file; the workflow's checkout step already puts `queries.xlsx` on disk. |
| `st.progress(...)` / `st.error(...)` → `print(...)` | Progress needs to show up in the Actions log, not a browser UI. |
| Added `out.to_excel(OUTPUT_FILE, index=False)` + `main()` entrypoint | Replaces the download button — the workflow commits this file instead of a user clicking download. |
| `uc.Chrome(options=opts, version_main=149)` → `uc.Chrome(options=opts)` (auto-detect) | The original pinned `version_main=149` to match one specific local Chrome install. GitHub's runner installs whatever the current stable Chrome is (and it changes over time), so a hardcoded version throws `SessionNotCreatedException: ChromeDriver only supports Chrome version 149, current browser version is X`. Auto-detection makes this self-healing. |
| Added optional `opts.binary_location` from `CHROME_PATH` env var | `browser-actions/setup-chrome` exposes the exact Chrome binary path it installed; pointing Chrome at it directly avoids any ambiguity if more than one Chrome-like binary exists on the runner's `PATH`. No-op locally. |
| Nothing else touched | All extraction logic (`extract_price`, `extract_rating_block`, `extract_stock_message`), retry logic (`safe_get`), delays, headless flags, and the anti-detection Chrome options are byte-for-byte identical to the original. |

## Common errors and how to fix them

**`SessionNotCreatedException: This version of ChromeDriver only supports
Chrome version X, current browser version is Y`**
Caused by a hardcoded `version_main`. Already fixed here by leaving it
unset — if it recurs, delete any `version_main=` argument from
`build_driver()`.

**Push step fails with `remote: Permission denied` / `403`**
Repo Settings → Actions → General → Workflow permissions isn't set to
"Read and write". Fix it there (see Setup step 3).

**`git commit` step says "nothing to commit" every run even though prices
should have changed**
Normal — the workflow only commits when `output.xlsx`'s bytes actually
differ. If Amazon returned the same values, there's nothing to push.

**Workflow times out or hangs**
Amazon serving a CAPTCHA/bot-check page instead of the product page is the
usual cause on shared-IP CI runners. Options: reduce `queries.xlsx` size per
run, increase delays in `DELAY`, or route requests through a
residential/rotating proxy (not included here — ask if you want this
wired in).

**`ModuleNotFoundError` for `undetected_chromedriver` / `bs4` / etc.**
`requirements.txt` wasn't installed — check the "Install Python
dependencies" step succeeded in the Actions log.

**Chrome fails to start with `DevToolsActivePort file doesn't exist`**
Almost always missing `--no-sandbox` / `--disable-dev-shm-usage` on a CI
runner. Both are already present in `build_driver()`; if you added new
Chrome options, make sure you didn't remove them.

## Notes / possible next steps

- This script re-scrapes every row in `queries.xlsx` on every run. If your
  list gets large, mirror maps-scraper's pattern: track scraped rows in a
  `completed_queries.txt`/state file and only process the remainder each
  run, so a single Actions run doesn't have to redo everything.
- `output.xlsx` is overwritten each run rather than appended. If you want a
  history of price changes over time, write to a timestamped file
  (`output_YYYY-MM-DD.xlsx`) instead, or append rows with a `scraped_at`
  column.
