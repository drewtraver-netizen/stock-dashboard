# stock-dashboard

Simple free dashboard hosted on GitHub Pages, with local Excel sync.

## Source
- Excel file: `/Users/georgetraver/Library/Mobile Documents/com~apple~CloudDocs/Drew/Financial/Investing/Drew Port 2.xlsm`
- Sheet: `Stock Model`

## Files
- `index.html`, `styles.css`, `app.js` → dashboard UI
- `data/data.json` → published data file
- `scripts/sync_excel.py` → reads Excel and updates JSON when content changes

## 1) Install dependency (once)
```bash
python3 -m pip install openpyxl
```

## 2) Run sync manually
From this folder:
```bash
python3 scripts/sync_excel.py \
  --excel "/Users/georgetraver/Library/Mobile Documents/com~apple~CloudDocs/Drew/Financial/Investing/Drew Port 2.xlsm" \
  --sheet "Stock Model" \
  --repo .
```

## 3) Auto-commit + push on changes
```bash
python3 scripts/sync_excel.py \
  --excel "/Users/georgetraver/Library/Mobile Documents/com~apple~CloudDocs/Drew/Financial/Investing/Drew Port 2.xlsm" \
  --sheet "Stock Model" \
  --repo . \
  --commit
```

## 4) Publish on GitHub Pages
1. Create a GitHub repo and push this folder.
2. In GitHub repo settings → Pages:
   - Source: `Deploy from a branch`
   - Branch: `main`
   - Folder: `/ (root)`
3. Your site will be at: `https://<username>.github.io/<repo>/`

## 5) Cron schedule (next step)
Use OpenClaw cron to run the sync command every 10–15 minutes.
