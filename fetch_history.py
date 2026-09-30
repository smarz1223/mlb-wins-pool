#!/usr/bin/env python3
"""
MLB Wins Pool - League History fetcher
Runs nightly via GitHub Actions right after fetch_data.py.

Reads the published Google Sheet (xlsx export), uses ONLY the "Data" tab
(one row per owner per season), recomputes everything from W/L/run diff,
and writes history.json for the History page.

Rules baked in (from the league rules + project decisions):
  - Standings order = Win %, then run differential as the tiebreaker
  - Finish (rank) uses that same tiebreak, so tied owners get different finishes
  - Games Back is recomputed from W-L (the sheet's 2018 GB column is wrong)
  - Titles = rank 1 in a season

Standard library only (no pip installs). Safe to fail: on any error it prints
the reason, exits non-zero, and leaves the existing history.json untouched.

Usage:
    python fetch_history.py                 # pulls the published sheet
    python fetch_history.py some_file.xlsx  # reads a local xlsx (testing)
"""

import datetime
import io
import json
import os
import re
import sys
import time
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

HISTORY_XLSX_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vT_dreBx17dUOxpr3wHpy_WxAdiY4xoeuNzO_4wanQX17faaUxnm9TEjhqVK1ea4vWWZoshFOPtS3hk"
    "/pub?output=xlsx"
)
OUTPUT_FILE = "history.json"

DATA_SHEET = "Data"          # the tab this script reads
SUMMARY_SHEET = "Summary"    # only used for a soft cross-check of the Titles list

# Header names in the Data tab (matched case-insensitively, punctuation ignored)
HEADER_ALIASES = {
    "year":   ["year"],
    "owner":  ["owners", "owner"],
    "rd":     ["rundiff", "rundifferential"],
    "wins":   ["wins", "w"],
    "losses": ["losses", "l"],
}

# A season with fewer games than this share of the largest season gets a footnote
SHORT_SEASON_SHARE = 0.75

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"


# ─────────────────────────────────────────────
# XLSX READER (standard library only)
# ─────────────────────────────────────────────

def download(url, attempts=3, timeout=45):
    last = None
    for i in range(attempts):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read()
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  download attempt {i + 1} failed: {e}")
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"Could not download the sheet: {last}")


def col_index(ref):
    """'B12' -> 1 (0-based column)."""
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def row_number(ref):
    return int(re.search(r"\d+", ref).group(0))


def read_shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall(f"{NS_MAIN}si"):
        out.append("".join(t.text or "" for t in si.iter(f"{NS_MAIN}t")))
    return out


def sheet_paths(zf):
    """Map sheet name -> path inside the zip."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for rel in rels.findall(f"{NS_PKG}Relationship"):
        t = rel.get("Target")
        t = t.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        targets[rel.get("Id")] = t
    out = {}
    for s in wb.find(f"{NS_MAIN}sheets").findall(f"{NS_MAIN}sheet"):
        out[s.get("name")] = targets[s.get(f"{NS_REL}id")]
    return out


def read_sheet(zf, path, shared):
    """Return {(row, col): value} with row 1-based, col 0-based."""
    cells = {}
    root = ET.fromstring(zf.read(path))
    for c in root.iter(f"{NS_MAIN}c"):
        ref = c.get("r")
        t = c.get("t")
        v = c.find(f"{NS_MAIN}v")
        val = None
        if t == "inlineStr":
            val = "".join(x.text or "" for x in c.iter(f"{NS_MAIN}t"))
        elif v is not None and v.text is not None:
            if t == "s":
                val = shared[int(v.text)]
            elif t in ("str", "e"):
                val = v.text if t == "str" else None
            elif t == "b":
                val = bool(int(v.text))
            else:
                try:
                    val = float(v.text)
                except ValueError:
                    val = None
        if val is not None and val != "":
            cells[(row_number(ref), col_index(ref))] = val
    return cells


def norm(s):
    return re.sub(r"[^a-z]", "", str(s).lower())


# ─────────────────────────────────────────────
# DATA TAB -> ROWS
# ─────────────────────────────────────────────

def read_data_rows(cells):
    # Find columns from the header row (row 1)
    header = {col: norm(v) for (r, col), v in cells.items() if r == 1 and isinstance(v, str)}
    colmap = {}
    for key, aliases in HEADER_ALIASES.items():
        for col, name in header.items():
            if name in aliases:
                colmap[key] = col
                break
    missing = [k for k in HEADER_ALIASES if k not in colmap]
    if missing:
        raise RuntimeError(f"Data tab is missing expected columns: {missing}. Found headers: {sorted(header.values())}")

    rows, seen = [], set()
    max_row = max(r for (r, _c) in cells)
    for r in range(2, max_row + 1):
        year = cells.get((r, colmap["year"]))
        owner = cells.get((r, colmap["owner"]))
        wins = cells.get((r, colmap["wins"]))
        losses = cells.get((r, colmap["losses"]))
        rd = cells.get((r, colmap["rd"]), 0.0)
        if not isinstance(year, float) or not isinstance(owner, str) or not owner.strip():
            continue
        if not isinstance(wins, float) or not isinstance(losses, float):
            continue
        year = int(year)
        owner = owner.strip()
        if wins + losses <= 0:
            continue
        if (year, owner) in seen:
            raise RuntimeError(f"Duplicate row for {owner} in {year}")
        seen.add((year, owner))
        rows.append({"year": year, "owner": owner, "w": wins, "l": losses, "rd": rd if isinstance(rd, float) else 0.0})
    if not rows:
        raise RuntimeError("No season rows found in the Data tab")
    return rows


# ─────────────────────────────────────────────
# COMPUTE
# ─────────────────────────────────────────────

def num(x):
    """Whole floats -> int (keeps JSON tidy); halves stay 282.5."""
    if isinstance(x, float) and x == int(x):
        return int(x)
    return x


def close(a, b):
    return abs(a - b) < 1e-9


def build_seasons(rows):
    by_year = defaultdict(list)
    for r in rows:
        r = dict(r)
        r["g"] = r["w"] + r["l"]
        r["pct"] = r["w"] / r["g"]
        by_year[r["year"]].append(r)

    seasons = {}
    warnings = []
    for year in sorted(by_year):
        rs = sorted(by_year[year], key=lambda r: (-r["pct"], -r["rd"], r["owner"]))
        lead = rs[0]
        for i, r in enumerate(rs):
            r["rank"] = i + 1
            r["gb"] = ((lead["w"] - lead["l"]) - (r["w"] - r["l"])) / 2
            r["tb"] = False
        # Flag owners who were tied on Win % with a neighbour (decided by run diff)
        for i, r in enumerate(rs):
            for j in (i - 1, i + 1):
                if 0 <= j < len(rs) and close(r["pct"], rs[j]["pct"]):
                    r["tb"] = True
                    if close(r["rd"], rs[j]["rd"]) and j > i:
                        warnings.append(f"{year}: {r['owner']} and {rs[j]['owner']} tied on Win % AND run differential")
        seasons[year] = rs
    return seasons, warnings


def build_history(rows):
    seasons, warnings = build_seasons(rows)
    years = sorted(seasons)
    latest = years[-1]

    games_by_year = {y: max(r["g"] for r in seasons[y]) for y in years}
    biggest = max(games_by_year.values())
    notes = [
        f"{y} was a shortened season ({num(games_by_year[y])} games per owner)."
        for y in years if games_by_year[y] < SHORT_SEASON_SHARE * biggest
    ]

    # Season tables
    season_out = {}
    for y in years:
        season_out[str(y)] = {
            "games": num(games_by_year[y]),
            "standings": [
                {
                    "rank": r["rank"], "owner": r["owner"],
                    "w": num(r["w"]), "l": num(r["l"]),
                    "pct": round(r["pct"], 6), "rd": num(r["rd"]),
                    "gb": num(r["gb"]), "tb": r["tb"],
                }
                for r in seasons[y]
            ],
        }

    # Champions
    champions = []
    for y in years:
        c, ru = seasons[y][0], seasons[y][1]
        champions.append({
            "year": y, "owner": c["owner"],
            "w": num(c["w"]), "l": num(c["l"]), "pct": round(c["pct"], 6), "rd": num(c["rd"]),
            "runner_up": ru["owner"], "runner_up_pct": round(ru["pct"], 6),
        })

    # All-time per owner
    per_owner = defaultdict(list)
    for y in years:
        for r in seasons[y]:
            per_owner[r["owner"]].append((y, r))

    all_time = []
    for owner, lst in per_owner.items():
        gp = sum(r["g"] for _y, r in lst)
        w = sum(r["w"] for _y, r in lst)
        l = sum(r["l"] for _y, r in lst)
        rd = sum(r["rd"] for _y, r in lst)
        finishes = [r["rank"] for _y, r in lst]
        all_time.append({
            "owner": owner,
            "seasons": len(lst),
            "gp": num(gp), "w": num(w), "l": num(l),
            "pct": round(w / gp, 6), "rd": num(rd),
            "titles": sum(1 for f in finishes if f == 1),
            "avg_finish": round(sum(finishes) / len(finishes), 4),
            "best_finish": min(finishes),
            "sub500": sum(1 for _y, r in lst if r["w"] * 2 < r["g"]),
            "active": any(y == latest for y, _r in lst),
        })
    all_time.sort(key=lambda o: (-o["pct"], -o["rd"], o["owner"]))

    # Grids (same owner order as the all-time table)
    grid = []
    for o in all_time:
        by_year = {}
        for y, r in per_owner[o["owner"]]:
            by_year[str(y)] = {"pct": round(r["pct"], 6), "rd": num(r["rd"]), "finish": r["rank"], "tb": r["tb"]}
        grid.append({
            "owner": o["owner"], "active": o["active"],
            "all_time_pct": o["pct"], "all_time_rd": o["rd"], "avg_finish": o["avg_finish"],
            "by_year": by_year,
        })

    return {
        "years": years,
        "latest_year": latest,
        "notes": notes,
        "champions": champions,
        "all_time": all_time,
        "grid": grid,
        "seasons": season_out,
    }, warnings


# ─────────────────────────────────────────────
# SOFT CROSS-CHECK vs the sheet's hand-entered Titles list (warn only)
# ─────────────────────────────────────────────

def check_titles(zf, shared, paths, champions):
    msgs = []
    try:
        if SUMMARY_SHEET not in paths:
            return msgs
        cells = read_sheet(zf, paths[SUMMARY_SHEET], shared)
        hdr = None
        for (r, c), v in cells.items():
            if c == 1 and isinstance(v, str) and v.strip().lower() == "year":
                if str(cells.get((r, 2), "")).strip().lower() == "owner":
                    hdr = r
                    break
        if hdr is None:
            return msgs
        listed = {}
        r = hdr + 1
        while isinstance(cells.get((r, 1)), float):
            listed[int(cells[(r, 1)])] = str(cells.get((r, 2), "")).strip()
            r += 1
        derived = {c["year"]: c["owner"] for c in champions}
        for y in sorted(set(listed) | set(derived)):
            if listed.get(y) != derived.get(y):
                msgs.append(f"Titles list on the Summary tab says {y} = {listed.get(y)!r}, but the Data tab works out to {derived.get(y)!r}")
    except Exception as e:  # noqa: BLE001
        msgs.append(f"(skipped Titles cross-check: {e})")
    return msgs


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def strip_updated(d):
    d = dict(d)
    d.pop("updated", None)
    return d


def main():
    local = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("HISTORY_XLSX_PATH")
    if local:
        print(f"Reading local file: {local}")
        with open(local, "rb") as f:
            raw = f.read()
    else:
        print("Downloading published history sheet...")
        raw = download(HISTORY_XLSX_URL)
    print(f"  {len(raw):,} bytes")

    zf = zipfile.ZipFile(io.BytesIO(raw))
    shared = read_shared_strings(zf)
    paths = sheet_paths(zf)
    if DATA_SHEET not in paths:
        raise RuntimeError(f"No '{DATA_SHEET}' tab found. Tabs: {list(paths)}")

    rows = read_data_rows(read_sheet(zf, paths[DATA_SHEET], shared))
    print(f"  {len(rows)} owner-season rows")

    history, warnings = build_history(rows)
    warnings += check_titles(zf, shared, paths, history["champions"])

    history["updated"] = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    history["source"] = "MLB Wins All-Time Google Sheet (Data tab)"

    for w in warnings:
        print(f"  WARNING: {w}")

    # Only rewrite when the content actually changed (keeps nightly commits quiet)
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE) as f:
                old = json.load(f)
            if strip_updated(old) == strip_updated(history):
                print("No changes since last run; history.json left as is.")
                return
        except (ValueError, OSError):
            pass

    with open(OUTPUT_FILE, "w") as f:
        json.dump(history, f, indent=1)
    print(f"Wrote {OUTPUT_FILE}: {len(history['years'])} seasons ({history['years'][0]}-{history['latest_year']}), "
          f"{len(history['all_time'])} owners")
    top = history["champions"][-1]
    print(f"Latest champion: {top['owner']} ({top['year']})")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print(f"ERROR: {e}")
        sys.exit(1)
