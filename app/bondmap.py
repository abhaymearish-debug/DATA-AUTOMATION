"""Bond mapping: which shop belongs to which bond, and which bond to which cluster.

Two different kinds of fact live here, so they are stored differently.

  * Shop -> bond is KSBC's own reference data. It already lives in the 'Bond'
    column of MASTER DATA CONFIRMED.xlsx, which the 81 build scripts read for
    themselves. Editing anywhere else would give the app one answer and the
    workbooks another, so this writes back to that column — and takes a
    timestamped backup of the file first, every time.

  * Bond -> cluster is not in any workbook; it was a constant in the code. It
    becomes a small JSON the app owns, so the ASM split can be changed without
    a deploy. The built-in split is the fallback.
"""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl

from . import config

SHEET = "16-4-25"
BOND_COL = "Bond"

# The split as it has always been. An override file replaces it wholesale.
DEFAULT_CLUSTERS: dict[str, list[str]] = {
    "1": ["ALAPPUZHA", "ATTINGAL", "KOLLAM", "KOTTARAKARA", "NEDUMANGAD", "PATHANAMTHITTA"],
    "2": ["ALUVA", "KOTTAYAM", "THODUPUZHA", "THRISSUR", "TRIPUNITHURA"],
    "3": ["KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA"],
}


def clusters_file() -> Path:
    return config.WORKSPACE_ROOT / "_seed" / "bond_clusters.json"


def load_clusters() -> dict[str, list[str]]:
    path = clusters_file()
    if path.is_file():
        try:
            data = json.loads(path.read_text())
            out = {str(k): [str(b).strip().upper() for b in v if str(b).strip()]
                   for k, v in data.items() if str(k) in ("1", "2", "3")}
            if out:
                return {k: out.get(k, []) for k in ("1", "2", "3")}
        except (OSError, ValueError):
            pass
    return {k: list(v) for k, v in DEFAULT_CLUSTERS.items()}


def save_clusters(mapping: dict[str, list[str]]) -> None:
    """Replace the cluster split, having checked it is one.

    This is written from a request body and decides how every report groups
    its rows, so a malformed or empty body used to be able to blank the split
    with nothing to restore it from. Three checks and a backup: the bonds have
    to be bonds this install knows, no bond may sit in two clusters, and the
    result may not be empty. The previous file is kept beside the new one.
    """
    if not isinstance(mapping, dict):
        raise ValueError("The cluster mapping must be an object of 1/2/3 to bond lists.")

    clean = {k: sorted({str(b).strip().upper() for b in (mapping.get(k) or [])
                        if str(b).strip()})
             for k in ("1", "2", "3")}

    if not any(clean.values()):
        raise ValueError("That would leave every cluster empty.")

    seen: dict[str, str] = {}
    for k, bonds in clean.items():
        for b in bonds:
            if b in seen:
                raise ValueError(f"{b} is in cluster {seen[b]} and cluster {k}.")
            seen[b] = k

    known = {b.strip().upper() for bonds in DEFAULT_CLUSTERS.values() for b in bonds}
    known |= {b.strip().upper() for bonds in load_clusters().values() for b in bonds}
    strangers = sorted(set(seen) - known)
    if strangers:
        raise ValueError("Not a bond on this install: " + ", ".join(strangers))

    path = clusters_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.is_file():
        try:
            path.with_suffix(path.suffix + ".bak").write_text(path.read_text())
        except OSError:
            pass
    path.write_text(json.dumps(clean, indent=2))


def _sheet(wb):
    return wb[SHEET] if SHEET in wb.sheetnames else wb[wb.sheetnames[-1]]


def _header_row(ws) -> tuple[int, dict[str, int]]:
    """Find the header and its column numbers (1-based), by name not position."""
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=12, values_only=True), 1):
        names = [str(v).strip() if v is not None else "" for v in row]
        if "Shop Code" in names:
            return r, {n: i + 1 for i, n in enumerate(names) if n}
    raise RuntimeError("MASTER DATA has no 'Shop Code' header row.")


def load_shops() -> dict:
    """Every shop in master data, with the bond it is currently mapped to."""
    path = config.MASTER_DATA
    if not path.is_file():
        return {"error": f"MASTER DATA CONFIRMED.xlsx is not in the workspace."}

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = _sheet(wb)
    head_row, ix = _header_row(ws)

    def cell(row, name):
        i = ix.get(name)
        return "" if not i or i > len(row) or row[i - 1] is None else str(row[i - 1]).strip()

    shops = []
    for row in ws.iter_rows(min_row=head_row + 1, values_only=True):
        code = cell(row, "Shop Code")
        if not code:
            continue
        shops.append({
            "code": code,
            "name": cell(row, "Shop Name"),
            "warehouse": cell(row, "Warehouse Name"),
            "cat": cell(row, "CAT").upper(),
            "staff": cell(row, "Field Staff"),
            "bond": cell(row, "Bond").upper(),
            "status": cell(row, "Status"),
        })
    wb.close()

    clusters = load_clusters()
    bonds = sorted({s["bond"] for s in shops if s["bond"]}
                   | {b for v in clusters.values() for b in v})
    return {"shops": shops, "bonds": bonds, "clusters": clusters, "source": path.name}


def save_bonds(changes: dict[str, str]) -> dict:
    """Write shop -> bond back into master data. Backs the file up first."""
    path = config.MASTER_DATA
    if not path.is_file():
        return {"error": "MASTER DATA CONFIRMED.xlsx is not in the workspace."}
    if not changes:
        return {"changed": 0, "backup": ""}

    lock = path.parent / f"~${path.name}"
    if lock.exists():
        return {"error": f"'{path.name}' is open in Excel. Close it and save again."}

    backups = path.parent / "_master_backups"
    backups.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backups / f"{path.stem}.bak_{stamp}{path.suffix}"
    shutil.copy2(path, backup)

    # Not read_only: this one writes. openpyxl is pinned at the version the
    # build scripts use, so a round-trip here is the same round-trip they do.
    wb = openpyxl.load_workbook(path)
    ws = _sheet(wb)
    head_row, ix = _header_row(ws)
    code_col, bond_col = ix.get("Shop Code"), ix.get(BOND_COL)
    if not code_col or not bond_col:
        wb.close()
        return {"error": "MASTER DATA is missing a 'Shop Code' or 'Bond' column."}

    changed, unknown = 0, []
    wanted = {str(k).strip(): str(v).strip().upper() for k, v in changes.items()}
    seen = set()
    for r in range(head_row + 1, ws.max_row + 1):
        code = ws.cell(row=r, column=code_col).value
        if code is None:
            continue
        code = str(code).strip()
        if code not in wanted:
            continue
        seen.add(code)
        cell = ws.cell(row=r, column=bond_col)
        if str(cell.value or "").strip().upper() != wanted[code]:
            cell.value = wanted[code]
            changed += 1
    unknown = sorted(set(wanted) - seen)

    wb.save(path)
    wb.close()
    return {"changed": changed, "unknown": unknown, "backup": backup.name}
