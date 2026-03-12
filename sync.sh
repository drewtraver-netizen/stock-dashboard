#!/bin/bash
# sync.sh - Manually sync Excel spreadsheet to stock dashboard and push to GitHub

set -e

EXCEL="/Users/georgetraver/Library/Mobile Documents/com~apple~CloudDocs/Drew/Financial/Investing/Drew Port 2.xlsm"
SHEET="DailyChartData"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "📊 Syncing stock dashboard from spreadsheet..."
cd "$REPO_DIR"

python3 scripts/sync_excel.py \
  --excel "$EXCEL" \
  --sheet "$SHEET" \
  --repo . \
  --commit

echo "✅ Done!"
