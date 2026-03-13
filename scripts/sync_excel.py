#!/usr/bin/env python3
import argparse
import datetime as dt
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path


def stable_hash(obj) -> str:
    data = json.dumps(obj, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


def load_workbook_xlwings(excel_path: Path):
    """Connect to the already-open workbook in Excel for live calculated values.
    Falls back to opening a new visible Excel instance if not already open."""
    import xlwings as xw

    filename = excel_path.name

    # Try to connect to an already-running Excel instance with the file open
    for app in xw.apps:
        for wb in app.books:
            if wb.name == filename:
                print(f"Connected to already-open workbook: {wb.name}")
                return None, wb  # None = don't quit the app when done

    # Not open — launch a new visible instance so formulas recalculate properly
    print("Workbook not open in Excel, opening now...")
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(str(excel_path))
    app.calculate()
    return app, wb


def load_sheet(wb, sheet_name: str):
    """Read headers and rows from a sheet using an open xlwings workbook."""
    import xlwings as xw

    if sheet_name not in [s.name for s in wb.sheets]:
        available = ", ".join(s.name for s in wb.sheets)
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {available}")

    ws = wb.sheets[sheet_name]

    # Find last row and column with actual data using End key navigation
    last_row = ws.range("A1").end("down").row
    last_col = ws.range("A1").end("right").column

    if last_row > 100000:
        # Fallback: sheet is likely empty beyond first row
        last_row = 1

    data_range = ws.range((1, 1), (last_row, last_col))
    all_values = data_range.value

    if not all_values:
        return [], []

    # Normalize to list of lists
    if not isinstance(all_values[0], list):
        all_values = [[v] for v in all_values]

    raw_headers = all_values[0]
    headers = []
    for i, h in enumerate(raw_headers):
        name = str(h).strip() if h is not None else ""
        headers.append(name or f"Column_{i+1}")

    records = []
    for row in all_values[1:]:
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        obj = {}
        for idx, header in enumerate(headers):
            val = row[idx] if idx < len(row) else None
            if isinstance(val, (dt.datetime, dt.date, dt.time)):
                val = val.isoformat()
            obj[header] = val
        records.append(obj)

    return headers, records


def read_portfolios(wb):
    """Read model score, S&P quote, YTD return, portfolio weights from Portfolios sheet."""
    ws = wb.sheets["Portfolios"]

    model_score = ws.range("J17").value
    sp_quote    = ws.range("V5").value
    ytd_return  = ws.range("P21").value

    portfolio_weights = []
    for r in range(28, 44):
        symbol = ws.cells(r, 13).value   # M
        weight = ws.cells(r, 17).value   # Q
        if symbol and weight and float(weight) > 0:
            portfolio_weights.append({"symbol": str(symbol).strip(), "weight": float(weight)})

    return model_score, sp_quote, ytd_return, portfolio_weights


def read_daily_returns(wb):
    """Read daily returns from Portfolios sheet: S&P T16, QQQ T17, IWM T18, DIA T19, GLD T23, my return P18."""
    ws = wb.sheets["Portfolios"]
    index_daily = [
        {"symbol": "S&P",  "return": ws.range("T16").value},
        {"symbol": "QQQ",  "return": ws.range("T17").value},
        {"symbol": "IWM",  "return": ws.range("T18").value},
        {"symbol": "DIA",  "return": ws.range("T19").value},
        {"symbol": "GLD",  "return": ws.range("T23").value},
    ]
    index_daily = [x for x in index_daily if x["return"] is not None]
    for x in index_daily:
        x["return"] = float(x["return"])
    my_return = ws.range("P18").value
    return {
        "indexDaily": index_daily,
        "myReturn": float(my_return) if my_return is not None else None,
    }


def read_index_returns(wb):
    """Read index symbols and returns from Models (Annual) sheet rows 13-14."""
    ws = wb.sheets["Models (Annual)"]
    index_returns = []
    for col in range(1, 50):
        symbol = ws.cells(13, col).value
        ret    = ws.cells(14, col).value
        if symbol and ret is not None:
            index_returns.append({"symbol": str(symbol).strip(), "return": float(ret)})
    return index_returns


def git_commit(repo_root: Path, message: str):
    subprocess.run(["git", "add", "data/data.json", "data/data.js", ".state/sync_state.json"],
                   cwd=repo_root, check=True)
    diff = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=repo_root, check=False)
    if diff.returncode == 0:
        print("No staged changes to commit.")
        return False
    subprocess.run(["git", "commit", "-m", message], cwd=repo_root, check=True)
    subprocess.run(["git", "push"], cwd=repo_root, check=True)
    return True


def main():
    parser = argparse.ArgumentParser(description="Sync Excel sheet to dashboard JSON via xlwings")
    parser.add_argument("--excel",  required=True, help="Path to .xlsx/.xlsm file")
    parser.add_argument("--sheet",  required=True, help="Sheet name for chart data")
    parser.add_argument("--repo",   default=".",   help="Dashboard repo root")
    parser.add_argument("--commit", action="store_true", help="Commit and push when changed")
    args = parser.parse_args()

    repo_root  = Path(args.repo).resolve()
    data_path  = repo_root / "data" / "data.json"
    js_path    = repo_root / "data" / "data.js"
    state_path = repo_root / ".state" / "sync_state.json"
    data_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.parent.mkdir(parents=True, exist_ok=True)

    excel_path = Path(os.path.expanduser(args.excel)).resolve()
    if not excel_path.exists():
        print(f"Excel file not found: {excel_path}", file=sys.stderr)
        sys.exit(1)

    print("Opening workbook in Excel (live values)...")
    app, wb = load_workbook_xlwings(excel_path)
    if app is not None:
        app.calculate()  # Force full recalculation if we opened a new instance

    try:
        headers, rows = load_sheet(wb, args.sheet)
        model_score, sp_quote, ytd_return, portfolio_weights = read_portfolios(wb)
        index_returns = read_index_returns(wb)
        daily_returns = read_daily_returns(wb)
    finally:
        if app is not None:
            # We opened this instance ourselves — close it
            wb.close()
            app.quit()

    payload_core = {
        "sheet":            args.sheet,
        "source":           str(excel_path),
        "headers":          headers,
        "rows":             rows,
        "modelScore":       model_score,
        "spQuote":          sp_quote,
        "ytdReturn":        ytd_return,
        "portfolioWeights": portfolio_weights,
        "indexReturns":     index_returns,
        "dailyReturns":     daily_returns,
    }
    content_hash = stable_hash(payload_core)

    prev_hash = None
    if state_path.exists():
        try:
            prev_hash = json.loads(state_path.read_text(encoding="utf-8")).get("contentHash")
        except Exception:
            prev_hash = None

    if content_hash == prev_hash:
        print("No data change detected.")
        return

    payload = {
        **payload_core,
        "generatedAt": dt.datetime.now(dt.timezone.utc).isoformat(),
        "rowCount":    len(rows),
    }

    data_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    js_path.write_text(
        "window.DASHBOARD_DATA = " + json.dumps(payload, indent=2, ensure_ascii=False) + ";\n",
        encoding="utf-8",
    )
    state_path.write_text(
        json.dumps({"contentHash": content_hash, "updatedAt": payload["generatedAt"]}, indent=2),
        encoding="utf-8",
    )
    print(f"Updated data files with {len(rows)} rows. YTD: {ytd_return}")

    if args.commit:
        try:
            committed = git_commit(repo_root, f"Update dashboard data ({args.sheet})")
            if committed:
                print("Committed and pushed to GitHub.")
        except subprocess.CalledProcessError as e:
            print(f"Git command failed: {e}", file=sys.stderr)
            sys.exit(e.returncode)


if __name__ == "__main__":
    main()
