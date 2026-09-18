"""Monthly sales targets, and the brand families they are set against.

Two things live here.

THE FAMILIES
------------
A target is set against a brand, but KSBC's raws carry SKU-level item names
and several spellings of the same brand. 'BCB NO.1 CLASSIC BRANDY' and
'B.C.B NO.1 CLASSIC BRANDY' are one brand to the office and two strings to a
computer, and the raws still carry a dozen names that have not sold a case
this year. So the report works in *families*: the eight the targets are set
in, plus an 'Other' bucket that only appears when something outside the eight
actually sells. Nothing is dropped silently - a brand that comes back from
the dead shows up in Other rather than vanishing from the total.

THE TARGETS THEMSELVES
----------------------
Bond x family x month, in cases, typed into the app. This is not KSBC data
and it is in no workbook; it is a management decision, so the app owns it as
a small JSON per month and nothing else needs to change when it is revised.
Cluster and network totals are never stored - they are the bonds added up,
which is the only way they can never disagree with the rows above them.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from . import config

# key, label as the target sheet writes it, and the column header on screen.
FAMILIES: list[tuple[str, str]] = [
    ("BCB",        "BCB No.1"),
    ("BLENDERS",   "Blender's Choice"),
    ("CHAIRMANS",  "Chairman's Choice"),
    ("KS99",       "K.S 99"),
    ("MAGIC",      "Magic Blend"),
    ("MORNING",    "Morning Walker"),
    ("OLDPEARL",   "Old Pearl"),
    ("OLDFORT",    "Royal Old Fort"),
]
OTHER = "OTHER"
OTHER_LABEL = "Other"

FAMILY_KEYS = [k for k, _ in FAMILIES]
FAMILY_LABEL = dict(FAMILIES) | {OTHER: OTHER_LABEL}

# Matched against the item name with punctuation and spaces taken out, so
# "B.C.B NO.1", "BCB NO.1" and "BCB  No1" are one thing. Order matters only
# in that the first hit wins; the fragments below do not overlap.
_RULES: list[tuple[str, str]] = [
    ("OLDPEARL",  "OLDPEARL"),
    ("OLDFORT",   "OLDFORT"),       # ROYAL OLD FORT and the legacy K S OLD FORT
    ("BLENDERS",  "BLENDER"),
    ("CHAIRMANS", "CHAIRMAN"),
    ("MORNING",   "MORNINGWALKER"),
    ("MAGIC",     "MAGICBLEND"),
    ("KS99",      "KS99"),
    ("BCB",       "BCB"),
]

_TIDY = re.compile(r"[^A-Z0-9]+")


def family_of(brand: str) -> str:
    """Which family an item name belongs to, or OTHER."""
    flat = _TIDY.sub("", (brand or "").upper())
    for key, fragment in _RULES:
        if fragment in flat:
            return key
    return OTHER


def columns(used_other: bool = False) -> list[tuple[str, str]]:
    """The columns a grid should show: the eight, and Other only if it sold."""
    return FAMILIES + ([(OTHER, OTHER_LABEL)] if used_other else [])


# ---------------------------------------------------------------------------
# Storage
# ---------------------------------------------------------------------------

_MONTH = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")


def _folder() -> Path:
    return config.WORKSPACE_ROOT / "_targets"


def _file(month: str) -> Path:
    if not _MONTH.match(month or ""):
        raise ValueError(f"{month!r} is not a YYYY-MM month.")
    return _folder() / f"{month}.json"


def load(month: str) -> dict:
    """bond -> family -> cases. Missing months are empty, never an error."""
    try:
        path = _file(month)
    except ValueError:
        return {}
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text())
    except (OSError, ValueError):
        return {}
    out: dict[str, dict[str, float]] = {}
    for bond, cells in (data.get("bonds") or {}).items():
        row = {}
        for key, value in (cells or {}).items():
            if key not in FAMILY_KEYS:
                continue
            try:
                num = float(value)
            except (TypeError, ValueError):
                continue
            if num:
                row[key] = num
        if row:
            out[str(bond).strip().upper()] = row
    return out


def save(month: str, bonds: dict, by: str = "") -> dict:
    """Replace a month's targets. A bond set to all zeros drops out."""
    path = _file(month)
    clean: dict[str, dict[str, float]] = {}
    for bond, cells in (bonds or {}).items():
        name = str(bond).strip().upper()
        if not name:
            continue
        row = {}
        for key in FAMILY_KEYS:
            try:
                num = float((cells or {}).get(key) or 0)
            except (TypeError, ValueError):
                num = 0.0
            if num < 0:
                raise ValueError(f"{name} {FAMILY_LABEL[key]}: a target cannot be negative.")
            if num:
                row[key] = round(num, 3)
        if row:
            clean[name] = row

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "month": month,
        "updated": datetime.now().isoformat(timespec="seconds"),
        "by": by,
        "bonds": clean,
    }, indent=2))
    return {"month": month, "bonds": len(clean),
            "total": round(sum(sum(r.values()) for r in clean.values()), 3)}


def months() -> list[str]:
    """Months that have targets, newest first."""
    folder = _folder()
    if not folder.is_dir():
        return []
    found = [p.stem for p in folder.glob("*.json") if _MONTH.match(p.stem)]
    return sorted(found, reverse=True)


def meta(month: str) -> dict:
    """When a month's targets were last written, and by whom."""
    try:
        path = _file(month)
    except ValueError:
        return {}
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text())
    except (OSError, ValueError):
        return {}
    return {"updated": data.get("updated", ""), "by": data.get("by", "")}
