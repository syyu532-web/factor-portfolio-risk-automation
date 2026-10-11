#!/usr/bin/env python3
"""Build a quarterly U.S. listed-company SIC -> Fama-French 48 snapshot from SEC EDGAR bulk data."""

from __future__ import annotations

import argparse
import csv
import io
import json
import re
import sys
import time
import urllib.request
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path

SUBMISSIONS_URL = "https://www.sec.gov/Archives/edgar/daily-index/bulkdata/submissions.zip"
TICKERS_URL = "https://www.sec.gov/files/company_tickers_exchange.json"
DEFAULT_EXCHANGES = {"Nasdaq", "NYSE", "NYSE American", "NYSE Arca", "CboeBZX"}
HEADER_RE = re.compile(r"^\s*(\d{1,2})\s+(\S+)\s+(.+?)\s*$")
RANGE_RE = re.compile(r"^\s*(\d{4})-(\d{4})\b")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Create a listed U.S. company SIC-to-FF48 CSV from SEC bulk data.")
    p.add_argument("--email", required=True, help="Your contact email; required in the SEC User-Agent.")
    p.add_argument("--ff48-file", default="Siccodes48.txt", help="Path to Siccodes48.txt.")
    p.add_argument("--output-dir", default="output", help="Directory for downloads and CSV outputs.")
    p.add_argument("--snapshot-date", help="Date label in YYYY-MM-DD format; default is today's UTC date.")
    p.add_argument("--all-exchanges", action="store_true", help="Keep all exchange values in the SEC ticker file, not only major U.S. exchanges.")
    p.add_argument("--no-download", action="store_true", help="Reuse already-downloaded SEC files in output-dir/raw.")
    return p.parse_args()


def download(url: str, destination: Path, user_agent: str) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(url, headers={"User-Agent": user_agent})
    tmp = destination.with_suffix(destination.suffix + ".part")
    print(f"Downloading {url}")
    try:
        with urllib.request.urlopen(request, timeout=180) as response, open(tmp, "wb") as f:
            while True:
                block = response.read(1024 * 1024)
                if not block:
                    break
                f.write(block)
        tmp.replace(destination)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def parse_ff48(path: Path) -> list[dict]:
    if not path.exists():
        raise FileNotFoundError(f"FF48 definition file not found: {path}")
    ranges = []
    current = None
    for raw_line in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        header = HEADER_RE.match(raw_line)
        if header:
            current = {
                "ff48_code": int(header.group(1)),
                "ff48_short": header.group(2),
                "ff48_name": header.group(3).strip(),
            }
            continue
        code_range = RANGE_RE.match(raw_line)
        if code_range and current:
            ranges.append({
                **current,
                "sic_start": int(code_range.group(1)),
                "sic_end": int(code_range.group(2)),
            })
    if not ranges:
        raise ValueError("No SIC ranges were parsed from Siccodes48.txt.")
    return ranges


def ff48_for_sic(sic, ranges: list[dict]) -> dict | None:
    try:
        value = int(str(sic).strip())
    except (TypeError, ValueError):
        return None
    for item in ranges:
        if item["sic_start"] <= value <= item["sic_end"]:
            return item
    return None


def read_ticker_exchange_map(path: Path, allowed_exchanges: set[str] | None) -> dict[str, list[dict]]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    fields = raw["fields"]
    positions = {field: i for i, field in enumerate(fields)}
    required = {"cik", "ticker", "exchange"}
    if not required.issubset(positions):
        raise ValueError(f"Unexpected SEC ticker file fields: {fields}")
    output: dict[str, list[dict]] = {}
    for record in raw["data"]:
        cik = str(record[positions["cik"]]).zfill(10)
        ticker = str(record[positions["ticker"]]).strip()
        exchange = str(record[positions["exchange"]]).strip()
        name = str(record[positions["name"]]).strip() if "name" in positions else ""
        if not ticker or (allowed_exchanges is not None and exchange not in allowed_exchanges):
            continue
        output.setdefault(cik, []).append({"ticker": ticker, "exchange": exchange, "ticker_name": name})
    return output


def quarter_label(d: date) -> str:
    return f"{d.year}Q{(d.month - 1) // 3 + 1}"


def build_snapshot(submissions_zip: Path, listed: dict[str, list[dict]], ranges: list[dict], snapshot_date: date) -> tuple[list[dict], list[dict], list[dict]]:
    rows, exceptions, missing = [], [], []
    seen_ciks = set()
    with zipfile.ZipFile(submissions_zip) as zf:
        members = [name for name in zf.namelist() if name.lower().endswith(".json")]
        total = len(members)
        for i, member in enumerate(members, start=1):
            if i % 5000 == 0 or i == total:
                print(f"Reading SEC company files: {i:,}/{total:,}")
            try:
                company = json.load(io.TextIOWrapper(zf.open(member), encoding="utf-8"))
                cik = str(company.get("cik", "")).zfill(10)
                ticker_records = listed.get(cik)
                if not ticker_records:
                    continue
                seen_ciks.add(cik)
                sic = company.get("sic")
                match = ff48_for_sic(sic, ranges)
                base = {
                    "snapshot_date": snapshot_date.isoformat(),
                    "cik": cik,
                    "company_name": company.get("name") or ticker_records[0]["ticker_name"],
                    "sic": sic if sic not in (None, "") else "",
                    "sic_description": company.get("sicDescription") or "",
                }
                if match is None:
                    exceptions.append({**base, "reason": "missing_sic" if not base["sic"] else "sic_not_in_ff48_definition"})
                for ticker_record in ticker_records:
                    rows.append({
                        **base,
                        "ticker": ticker_record["ticker"],
                        "exchange": ticker_record["exchange"],
                        "ff48_code": match["ff48_code"] if match else "",
                        "ff48_short": match["ff48_short"] if match else "",
                        "ff48_name": match["ff48_name"] if match else "",
                    })
            except (json.JSONDecodeError, UnicodeDecodeError, KeyError, OSError) as exc:
                exceptions.append({"snapshot_date": snapshot_date.isoformat(), "cik": "", "company_name": member, "sic": "", "sic_description": "", "reason": f"read_error: {exc}"})
    for cik, ticker_records in listed.items():
        if cik not in seen_ciks:
            for record in ticker_records:
                missing.append({"snapshot_date": snapshot_date.isoformat(), "cik": cik, **record, "reason": "CIK_not_found_in_submissions_zip"})
    rows.sort(key=lambda x: (x["exchange"], x["ticker"]))
    return rows, exceptions, missing


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    out = Path(args.output_dir)
    raw = out / "raw"
    ff48_path = Path(args.ff48_file)
    if args.snapshot_date:
        snapshot = datetime.strptime(args.snapshot_date, "%Y-%m-%d").date()
    else:
        snapshot = datetime.now(timezone.utc).date()
    user_agent = f"SIC FF48 academic research {args.email}"
    submissions_zip = raw / "submissions.zip"
    tickers_json = raw / "company_tickers_exchange.json"
    if not args.no_download:
        download(SUBMISSIONS_URL, submissions_zip, user_agent)
        time.sleep(0.2)
        download(TICKERS_URL, tickers_json, user_agent)
    elif not submissions_zip.exists() or not tickers_json.exists():
        raise FileNotFoundError("--no-download was set, but a required raw SEC file is missing in output-dir/raw.")

    ranges = parse_ff48(ff48_path)
    allowed = None if args.all_exchanges else DEFAULT_EXCHANGES
    listed = read_ticker_exchange_map(tickers_json, allowed)
    print(f"Eligible ticker records: {sum(len(v) for v in listed.values()):,}; CIKs: {len(listed):,}")
    rows, exceptions, missing = build_snapshot(submissions_zip, listed, ranges, snapshot)
    label = quarter_label(snapshot)
    main_file = out / f"us_listed_sic_ff48_{label}.csv"
    exception_file = out / f"mapping_exceptions_{label}.csv"
    missing_file = out / f"missing_submissions_{label}.csv"
    metadata_file = out / f"run_metadata_{label}.json"
    write_csv(main_file, rows, ["snapshot_date", "cik", "ticker", "exchange", "company_name", "sic", "sic_description", "ff48_code", "ff48_short", "ff48_name"])
    write_csv(exception_file, exceptions, ["snapshot_date", "cik", "company_name", "sic", "sic_description", "reason"])
    write_csv(missing_file, missing, ["snapshot_date", "cik", "ticker", "exchange", "ticker_name", "reason"])
    metadata = {
        "snapshot_date": snapshot.isoformat(),
        "created_at_utc": datetime.now(timezone.utc).isoformat(),
        "submissions_url": SUBMISSIONS_URL,
        "ticker_exchange_url": TICKERS_URL,
        "exchange_filter": "all SEC exchange values" if args.all_exchanges else sorted(DEFAULT_EXCHANGES),
        "ff48_definition_file": str(ff48_path),
        "ff48_sic_ranges_parsed": len(ranges),
        "output_rows": len(rows),
        "unmapped_company_rows": len(exceptions),
        "listed_ciks_missing_from_submissions_zip": len(missing),
    }
    metadata_file.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\nCompleted")
    print(f"Main file:       {main_file}")
    print(f"Exceptions:      {exception_file}")
    print(f"Missing CIKs:    {missing_file}")
    print(f"Run metadata:    {metadata_file}")
    print(f"Rows written:    {len(rows):,}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("Cancelled.")
    except Exception as exc:
        sys.exit(f"ERROR: {exc}")
