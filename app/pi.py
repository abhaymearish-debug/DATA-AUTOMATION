"""Purchase Instruction: KSBC's monthly buying instruction, shop by shop.

WHAT THE RAWS ARE
-----------------
Once a month the ERP issues one "Month Wise Purchase Instruction" per shop -
about 295 files, each an HTML table saved with an .xls extension. Every file
carries the shop, its warehouse, the report month, the shop's capacity, and
one row per KSD brand with four figures:

    L3MS  previous three months' sale, in BOTTLES
    RL    in cases
    RQ    in cases
    MQ    in cases, and always RL + RQ

MQ is the number that matters: it is what KSBC will let that shop buy.

THE PARSER IS THE BUILD SCRIPT'S
--------------------------------
parse_file below is build_pi_analysis.py's parser, regex for regex. The app
and the workbook builder must never disagree about what a raw says, and the
cheapest way to guarantee that is to read it the same way rather than a
better way.

THE TRAP THIS REPORT EXISTS TO CATCH
------------------------------------
A shop whose PI had not been generated when the export was taken comes back
as a file with no table at all - and printed as zeros it is indistinguishable
from a shop KSBC decided to give nothing. In September 2026 that was 33 shops
carrying 926 cases of August MQ, which was most of a 1,076-case "drop" that
had not actually happened. So a blank PI is never shown as a zero here: it is
counted, named and totalled separately.
"""

from __future__ import annotations

import glob
import html
import os
import re
import shutil
import threading
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

from . import config

# Brand codes as build_pi_analysis.py maps them. A code outside this set means
# KSBC has listed a new SKU, which is a thing to be told about, not to average
# away - so it is kept under its own name rather than dropped.
BRAND_FULL: dict[str, str] = {
    "1139703": "BCB NO.1 CLASSIC BRANDY",
    "1139707": "BLENDER'S CHOICE NO.1 BRANDY",
    "1139708": "MORNING WALKERS XO BRANDY",
    "1139715": "CHAIRMAN'S CHOICE XO BRANDY",
    "1339703": "OLD PEARL NO.1 MATURED XXX RUM",
    "1339704": "ROYAL OLD FORT NO.1 XXX RUM",
    "1339710": "K.S 99 LIFE TIME MATURED XXX RUM",
    "1339718": "MAGIC BLEND RESERVED XXX RUM",
}
BRAND_SHORT: dict[str, str] = {
    "1139703": "BCB", "1139707": "Blender's", "1139708": "Morning Walker",
    "1139715": "Chairman's", "1339703": "Old Pearl", "1339704": "Royal Old Fort",
    "1339710": "K.S 99", "1339718": "Magic Blend",
}
# The order the bond sheets print in.
BRAND_ORDER = ["1339703", "1139707", "1139703", "1339710",
               "1339704", "1139708", "1339718", "1139715"]

MEASURES = [("l3ms", "L3MS", "btl"), ("rl", "RL", "cs"),
            ("rq", "RQ", "cs"), ("mq", "MQ", "cs")]

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]
_MON3 = {m[:3].upper(): i for i, m in enumerate(MONTH_NAMES, 1)}


# ---------------------------------------------------------------------------
# Parsing (build_pi_analysis.py's, verbatim in behaviour)
# ---------------------------------------------------------------------------


def _cells_of(t: str) -> list[str]:
    return [html.unescape(re.sub(r"<[^>]+>", "", c)).strip()
            for c in re.findall(r"<t[dh][^>]*>(.*?)(?=<t[dh][^>]*>|$)", t, re.S)]


def parse_file(path) -> dict:
    path = Path(path)
    txt = path.read_text(encoding="utf-8", errors="replace")
    shop = re.search(r"Shop\s*:\s*<b>([^<]+)</b>", txt)
    wh = re.search(r"Warehouse\s*:\s*<b>([^<]+)</b>", txt)
    mon = re.search(r"Report Month\s*:\s*<b>([^<]+)</b>", txt)
    cap = re.search(r"Shop Capacity\(Cases\)\s*:\s*([\d,]+)\s*,\s*Shop Area\s*:\s*([\d,]+)", txt)

    rows: list[tuple] = []
    cs = _cells_of(txt)
    i = 0
    while i < len(cs) - 6:
        c = cs[i:i + 7]
        if re.fullmatch(r"\d{1,3}", c[0]) and re.fullmatch(r"\d{6,8}", c[1]) \
           and c[2] and not c[2].isdigit():
            nums, ok = [], True
            for x in c[3:7]:
                x = (x or "0").replace(",", "") or "0"
                if not re.fullmatch(r"\d+", x):
                    ok = False
                    break
                nums.append(int(x))
            if ok:
                rows.append((c[1], c[2], *nums))
                i += 7
                continue
        i += 1

    m = re.match(r"(\d+)[\s-]+(.*)", shop.group(1).strip()) if shop else None
    return {
        "file": path.name,
        "shop_code": m.group(1) if m else None,
        "shop_name": m.group(2).strip() if m else None,
        "warehouse": wh.group(1).strip() if wh else None,
        "month": mon.group(1).strip() if mon else None,
        "capacity": int(cap.group(1).replace(",", "")) if cap else None,
        "area": int(cap.group(2).replace(",", "")) if cap else None,
        "rows": rows,
    }


_FILE_DATE_RE = re.compile(r"Report[ _](\d{1,2})-([A-Za-z]{3})-(\d{4})", re.IGNORECASE)


def month_key(parsed: dict, filename: str = "") -> str:
    """'2026-09' for a PI file.

    The raw names its month but not its year, and the year is in the export
    filename. A December export of a January instruction is the one case where
    those disagree, so it is handled rather than assumed away.
    """
    name = parsed.get("month") or ""
    if not name:
        return ""
    try:
        mnum = MONTH_NAMES.index(name.strip().capitalize()) + 1
    except ValueError:
        return ""
    m = _FILE_DATE_RE.search(filename or parsed.get("file") or "")
    if not m:
        return f"{date.today().year:04d}-{mnum:02d}"
    year, fmon = int(m.group(3)), _MON3.get(m.group(2).upper(), mnum)
    if fmon == 12 and mnum == 1:
        year += 1
    return f"{year:04d}-{mnum:02d}"


def month_label(key: str) -> str:
    if not key or len(key) != 7:
        return key or ""
    return f"{MONTH_NAMES[int(key[5:7]) - 1]} {key[:4]}"


# ---------------------------------------------------------------------------
# Storage
# ---------------------------------------------------------------------------


def root() -> Path:
    return config.CLAUDE_ROOT / "PURCHASE INSTRUCTION" / "_months"


def month_dir(key: str) -> Path:
    return root() / key


def _blank_of(d: dict) -> bool:
    return not d["rows"]


def _tidy_shop(name: str) -> str:
    """Master writes '10001-PONNANI'; a filed shop reads '10001 - PONNANI'."""
    return re.sub(r"^\s*(\d+)\s*-\s*", r"\1 - ", name or "").strip()


def _shop_of(d: dict) -> str:
    code, name = d.get("shop_code") or "", d.get("shop_name") or ""
    return f"{code} - {name}".strip(" -") or (d.get("file") or "")


def store(paths: list[Path], expect: str = "") -> dict:
    """File a batch of raws under the month they are for, and say what landed.

    A month is replaced whole rather than merged: a second export of the same
    month is a correction, and half of one export beside half of another is
    the one state nobody could reason about.

    `expect` is the month picked in the dialog. The files still name their own
    month - ~295 of them agreeing is the better answer - so the pick is used
    as a check, not an override: files for another month are set aside and
    named, rather than filed under a month they do not belong to.

    The return is a receipt: every file is accounted for as filed, blank,
    a duplicate, for another month, or unreadable.
    """
    seen: list[tuple[Path, dict, str]] = []
    problems: list[dict] = []

    for p in paths:
        parsed = parse_file(p)
        key = month_key(parsed, p.name)
        if not parsed["shop_code"] or not key:
            why = ("no shop could be read out of it" if not parsed["shop_code"]
                   else "it names no report month")
            problems.append({"name": p.name, "state": "failed", "why": why})
            continue
        seen.append((p, parsed, key))

    if not seen:
        return {"error": f"None of those {len(paths)} files look like a KSBC purchase "
                         "instruction export - no shop or report month could be read "
                         "out of any of them."}

    counts = Counter(k for _p, _d, k in seen)
    if expect:
        key = expect
        if key not in counts:
            found = ", ".join(month_label(k) for k in sorted(counts))
            return {"error": f"You picked {month_label(expect)}, but these files are "
                             f"for {found}. Change the month, or pick the right files."}
    elif len(counts) > 1:
        months = ", ".join(month_label(k) for k in sorted(counts))
        return {"error": f"Those files cover more than one month ({months}). "
                         "Upload one month's instruction at a time."}
    else:
        key = next(iter(counts))

    keep: list[tuple[Path, dict]] = []
    for p, d, k in seen:
        if k == key:
            keep.append((p, d))
        else:
            problems.append({"name": p.name, "state": "other_month",
                             "shop": _shop_of(d), "why": f"is for {month_label(k)}"})

    # Two files for one shop would be counted twice by load(), so only one is
    # filed. An instruction beats an empty sheet; failing that, the later file
    # wins, since a re-export is a correction.
    by_shop: dict[str, tuple[Path, dict]] = {}
    for p, d in keep:
        code = d["shop_code"]
        old = by_shop.get(code)
        if old is None:
            by_shop[code] = (p, d)
            continue
        drop = old if (_blank_of(old[1]) and not _blank_of(d)) else (p, d)
        win = (p, d) if drop is old else old
        by_shop[code] = win
        problems.append({"name": drop[0].name, "state": "duplicate",
                         "shop": _shop_of(d), "why": f"same shop as {win[0].name}"})

    folder = month_dir(key)
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)
    for p, _d in by_shop.values():
        shutil.copy2(p, folder / p.name)

    blanks = [d for _p, d in by_shop.values() if _blank_of(d)]
    for d in sorted(blanks, key=_shop_of):
        problems.append({"name": d["file"], "state": "blank", "shop": _shop_of(d),
                         "why": "no instruction lines - nothing to buy"})

    _CACHE.pop(key, None)
    return _receipt(key, by_shop, len(paths), problems)


_CACHE: dict = {}
_LOCK = threading.Lock()


def months() -> list[dict]:
    """Stored PI months, newest first."""
    folder = root()
    if not folder.is_dir():
        return []
    out = []
    for path in sorted(folder.iterdir(), reverse=True):
        if not path.is_dir() or not re.fullmatch(r"\d{4}-\d{2}", path.name):
            continue
        out.append({"key": path.name, "label": month_label(path.name),
                    "files": len(list(path.glob("*.xls")))})
    return out


def load(key: str) -> list[dict]:
    """Every shop's parsed PI for a month. Cached on the folder's mtime."""
    folder = month_dir(key)
    if not folder.is_dir():
        return []
    stamp = folder.stat().st_mtime_ns
    with _LOCK:
        hit = _CACHE.get(key)
    if hit and hit[0] == stamp:
        return hit[1]

    parsed = [parse_file(p) for p in sorted(folder.glob("*.xls"))]
    parsed = [d for d in parsed if d["shop_code"]]
    with _LOCK:
        _CACHE[key] = (stamp, parsed)
    return parsed


# ---------------------------------------------------------------------------
# The variance report
# ---------------------------------------------------------------------------


def _blank_cells() -> dict:
    return {code: {"l3ms": 0, "rl": 0, "rq": 0, "mq": 0} for code in BRAND_ORDER}


def _add(into: dict, code: str, l3ms: int, rl: int, rq: int, mq: int) -> None:
    cell = into.setdefault(code, {"l3ms": 0, "rl": 0, "rq": 0, "mq": 0})
    cell["l3ms"] += l3ms
    cell["rl"] += rl
    cell["rq"] += rq
    cell["mq"] += mq


def _totals(cells: dict) -> dict:
    out = {"l3ms": 0, "rl": 0, "rq": 0, "mq": 0}
    for cell in cells.values():
        for k in out:
            out[k] += cell[k]
    return out


def pi_master() -> dict:
    """Master data keyed the way a PI names a shop.

    KSBC has two codes for the same shop: the six-digit one its sales exports
    use (104001) and the short one on the instruction (4001). Master carries
    the long code in Shop Code and the short one at the front of Shop Name, so
    that is what this joins on - exactly as build_pi_analysis.py does.
    """
    from . import reports_api

    out: dict[str, dict] = {}
    for _long, info in reports_api.load_master().items():
        m = re.match(r"\s*(\d+)", info.get("name") or "")
        if m:
            out[m.group(1)] = info
    return out


def variance(month: str, view: str = "bond", cluster: int | None = None,
             prior: str = "") -> dict:
    """One month's instruction, grouped, with last month beside it."""
    from . import reports_api

    current = load(month)
    if not current:
        return {"error": f"No purchase instruction is stored for {month_label(month)}."}

    master = pi_master()
    of_cluster = reports_api.cluster_of_bond()
    clusters = reports_api.bond_clusters()

    # A shop that is not in master AND has an empty instruction is a stray
    # export, not a missing PI - there is no shop behind it to chase. One that
    # is in master stays even when closed, because a closed shop with an
    # instruction is worth seeing.
    current = [d for d in current if d["shop_code"] in master or d["rows"]]

    def group_of(d: dict) -> str:
        if view == "warehouse":
            return reports_api.canonical_warehouse(d.get("warehouse") or "") or "(no warehouse)"
        return master.get(d["shop_code"], {}).get("bond", "") or "(unmapped)"

    prior_rows = load(prior) if prior else []
    prior_mq = defaultdict(int)
    prior_seen: dict[str, bool] = {}
    for d in prior_rows:
        prior_seen[d["shop_code"]] = bool(d["rows"])
        for _code, _brand, _l3, _rl, _rq, mq in d["rows"]:
            prior_mq[d["shop_code"]] += mq

    groups: dict[str, dict] = {}
    shops: dict[str, list] = defaultdict(list)
    blanks: list[dict] = []
    unknown_codes: set = set()

    for d in current:
        key = group_of(d)
        if view == "bond" and cluster in (1, 2, 3) and of_cluster.get(key) != cluster:
            continue
        if view == "warehouse" and cluster in (1, 2, 3) \
           and reports_api.CLUSTER_OF_WAREHOUSE.get(key) != cluster:
            continue

        cells = _blank_cells()
        for code, brand, l3ms, rl, rq, mq in d["rows"]:
            if code not in BRAND_FULL:
                unknown_codes.add(f"{code} {brand}")
            _add(cells, code, l3ms, rl, rq, mq)

        blank = not d["rows"]
        was = prior_mq.get(d["shop_code"], 0)
        shop = {
            "code": d["shop_code"],
            # KSBC writes the shop as '4024-THRIKKUNNAPUZHA' and every sheet in
            # the office reads it that way, so the code travels with the name.
            "name": f'{d["shop_code"]}-{d["shop_name"]}' if d["shop_name"] else d["shop_code"],
            "bond": master.get(d["shop_code"], {}).get("bond", ""),
            "capacity": d["capacity"],
            "blank": blank,
            "cells": cells,
            "total": _totals(cells),
            "prior_mq": was if prior else None,
        }
        shops[key].append(shop)

        # A blank PI is a missing instruction, not a zero one. It is named here
        # so the page can say which shops are missing rather than printing a
        # row of noughts that reads like a decision KSBC took.
        if blank:
            blanks.append({"group": key, "code": shop["code"], "name": shop["name"],
                           "bond": shop["bond"], "prior_mq": was})

        g = groups.setdefault(key, {"cells": _blank_cells(), "shops": 0,
                                    "blank": 0, "prior_mq": 0})
        for code, cell in cells.items():
            _add(g["cells"], code, cell["l3ms"], cell["rl"], cell["rq"], cell["mq"])
        g["shops"] += 1
        g["blank"] += 1 if blank else 0
        g["prior_mq"] += was

    for key in shops:
        shops[key].sort(key=lambda s: (-s["total"]["mq"], s["name"]))

    def row(label, kind, keys, cl=None):
        cells = _blank_cells()
        n = nblank = was = 0
        for k in keys:
            g = groups.get(k)
            if not g:
                continue
            for code, cell in g["cells"].items():
                _add(cells, code, cell["l3ms"], cell["rl"], cell["rq"], cell["mq"])
            n += g["shops"]
            nblank += g["blank"]
            was += g["prior_mq"]
        total = _totals(cells)
        return {"kind": kind, "label": label, "key": keys[0] if kind == "group" else "",
                "cluster": cl, "cells": cells, "total": total,
                "shops": n, "blank": nblank,
                "prior_mq": was if prior else None,
                "delta": (total["mq"] - was) if prior else None}

    rows: list[dict] = []
    shown: list[str] = []
    if view == "bond":
        for cid in sorted(clusters):
            if cluster in (1, 2, 3) and cid != cluster:
                continue
            members = [b for b in sorted(clusters[cid]) if b in groups]
            if not members:
                continue
            for bond in members:
                rows.append(row(bond, "group", [bond], cid))
            rows.append(row(f"CLUSTER {cid}", "cluster", members, cid))
            shown += members
        loose = sorted(set(groups) - set(shown))
    else:
        loose = sorted(groups)

    for key in loose:
        rows.append(row(key, "group", [key], None))
    shown += loose
    if cluster in (None, 0) or view == "warehouse":
        rows.append(row("GRAND TOTAL", "grand", shown, None))

    return {
        "month": month, "month_label": month_label(month),
        "prior": prior, "prior_label": month_label(prior) if prior else "",
        "view": view, "cluster": cluster or 0,
        "brands": [{"code": c, "short": BRAND_SHORT[c], "full": BRAND_FULL[c]}
                   for c in BRAND_ORDER],
        "measures": [{"key": k, "label": l, "unit": u} for k, l, u in MEASURES],
        "rows": rows,
        "shops": {k: v for k, v in shops.items()},
        "blanks": sorted(blanks, key=lambda b: -b["prior_mq"]),
        "unknown_codes": sorted(unknown_codes),
        "months": months(),
        "source_block": source_for(month, prior),
    }


def source_for(month: str, prior: str = "") -> dict:
    """The sheets behind one month's instruction, and the month beside it.

    KSBC exports purchase instruction one shop at a time, so a month is a few
    hundred files that all say the same thing about themselves. The month, the
    count filed and how many came back blank is the answer worth reading; the
    Receipt button beside this opens the file-by-file version.
    """
    from . import reports_api as _r

    def leg(key: str, label: str, tone: str = "") -> dict:
        folder = month_dir(key) if key else None
        if not folder or not folder.is_dir():
            return {}
        files = [f for f in folder.glob("*.xls*") if not f.name.startswith("~$")]
        if not files:
            return {}
        filed = load(key)
        blank = sum(1 for d in filed if _blank_of(d))
        note = f"{len(files)} shop sheet{'' if len(files) == 1 else 's'} filed"
        if blank:
            note += f", {blank} returned blank"
        return _r.src_leg(label, [{"title": month_label(key),
                                   "kind": "purchase instruction",
                                   "stream": _r.STREAM["purchase"],
                                   "name": note}], tone=tone)

    legs = [l for l in (leg(month, "This month"),
                        leg(prior, "Compared against", "green")) if l]
    return _r.src_block(legs)


def _receipt(key: str, by_shop: dict, files: int, problems: list) -> dict:
    """What landed, shop by shop - the same answer at upload and months later.

    Every shop in bond mapping, and what became of it. Four states, and the
    fourth is the one that used to be invisible: KSBC keeps exporting for
    shops that have closed, and those files land here with nothing in master
    to hang them on. They are filed - a closed shop with an instruction is
    worth seeing - but they are named as strays rather than counted as though
    the mapping knew about them.
    """
    blanks = [d for _p, d in by_shop.values() if _blank_of(d)]
    try:
        master = pi_master()
    except Exception:
        master = {}

    roster: list[dict] = []
    for code, (_p, d) in by_shop.items():
        info = master.get(code)
        roster.append({
            "code": code, "name": _shop_of(d),
            "bond": str((info or {}).get("bond") or ""),
            "warehouse": str(d.get("warehouse") or ""),
            # Two separate facts: what its sheet said, and whether the mapping
            # knows the shop at all. A closed shop KSBC still exports for can
            # be either, and counting it as one hid the other.
            "state": "blank" if _blank_of(d) else "filed",
            "stray": info is None,
        })
    for code, info in master.items():
        if code in by_shop:
            continue
        roster.append({
            "code": code, "name": _tidy_shop(str(info.get("name") or code)),
            "bond": str(info.get("bond") or ""),
            "warehouse": str(info.get("warehouse") or ""),
            "state": "absent", "stray": False,
        })
    roster.sort(key=lambda r: (r["bond"] or "~", r["name"]))

    def count(state: str) -> int:
        return sum(1 for r in roster if r["state"] == state)

    return {
        "month": key, "label": month_label(key),
        "files": files,
        "shops": len(by_shop),
        "mapped": len(master),
        "ok": len(by_shop) - len(blanks),
        "blank": len(blanks),
        "stray": sum(1 for r in roster if r.get("stray")),
        "absent": count("absent"),
        "duplicate": sum(1 for x in problems if x["state"] == "duplicate"),
        "other_month": sum(1 for x in problems if x["state"] == "other_month"),
        "failed": sum(1 for x in problems if x["state"] == "failed"),
        "warehouses": len({(d.get("warehouse") or "").strip()
                           for _p, d in by_shop.values() if d.get("warehouse")}),
        "roster": roster,
        "problems": problems[:400],
    }


def receipt_for(month: str) -> dict:
    """The receipt for a month already filed, read back off the files.

    What the upload itself saw - a file that would not open, one for another
    month, two sheets for the same shop - is only in the receipt that upload
    wrote. This reads what is on disk now, so it always answers for filed,
    blank, absent and stray, and says it is a reading rather than a record.
    """
    folder = month_dir(month)
    if not folder.is_dir():
        return {"error": f"Nothing is filed for {month_label(month)}."}

    by_shop: dict[str, tuple] = {}
    unreadable = 0
    for path in sorted(folder.glob("*.xls*")):
        try:
            parsed = parse_file(path)
        except Exception:
            unreadable += 1
            continue
        code = parsed.get("shop_code")
        if not code:
            unreadable += 1
            continue
        by_shop[code] = (path, parsed)

    got = _receipt(month, by_shop, len(by_shop) + unreadable, [])
    got["rebuilt"] = True
    got["failed"] = unreadable
    return got
