#!/usr/bin/env python3
"""
Assemble `aging_live_template.html` from its parts.

WHY THIS EXISTS.  Per CLAUDE.md the KSD live dashboards share ONE design
language, and the way that is enforced is that each new artifact copies the
liquidation `<style>` block VERBATIM and appends only a module-specific
addendum.  Hand-porting the styling produces drift; this script makes the
copy mechanical and repeatable.

  liquidation_live_template.html  <style> … </style>   (the canonical base)
+ _aging_addendum.css                                  (aging-only rules)
+ _aging_body.html                                     (markup)
+ _aging_app.js                                        (behaviour)
= aging_live_template.html                             (what the refresh
                                                        script fills)

Re-run this whenever the liquidation styling changes, so the aging artifact
picks the change up instead of slowly diverging.

    python3 .claude/scripts/build_aging_template.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
LIQ = HERE / "liquidation_live_template.html"
OUT = HERE / "aging_live_template.html"

# The knocked-out KSD shield is byte-identical in every dashboard template, so
# take it from whichever is present.  It used to come from the commitment
# template alone -- which was moved into `_reverted/` mid-build on 8 Aug 2026
# and broke the build.  Liquidation is first because the style block already
# comes from there, so the builder needs one fewer file to exist.
LOGO_SOURCES = [
    LIQ,
    HERE / "warehouse_live_template.html",
    HERE / "commitment_live_template.html",
    HERE / "_reverted" / "commitment_live_template.html",
    OUT,
]

PARTS = {
    "css":  HERE / "_aging_addendum.css",
    "body": HERE / "_aging_body.html",
    "js":   HERE / "_aging_app.js",
}

# Chart.js -- the exact tag the artifact sandbox allows. Must match byte for
# byte (integrity + crossorigin) or the sandbox blocks the load.
CHART_JS = ('<script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/'
            'chart.umd.js" integrity="sha384-iU8HYtnGQ8Cy4zl7gbNMOhsDTTKX02BT'
            'XptVP/vqAWIaTfM7isw76iyZCsjL2eVi" crossorigin="anonymous">'
            '</script>')


def extract_style(path: Path) -> str:
    """Pull the contents of the FIRST top-level <style> block, sanitised.

    KNOWN DEFECT IN THE SOURCE (found 8 Aug 2026 while building this):
    liquidation_live_template.html's style block ends with FOUR orphaned
    declaration blocks -- their selector lines were lost in some earlier edit,
    so the file tails off with bare `background:…; }` rules under a
    "daily-trend chips" comment.  A browser discards malformed top-level rules,
    which is why nobody noticed: those chip styles are simply dead in the
    liquidation artifact today.

    We must NOT copy them here.  An unbalanced block sits immediately before
    the aging addendum, and a CSS parser recovering from an error can swallow
    the rule that follows -- which would be the first rule of OUR stylesheet.
    So the trailing orphan run is detected, dropped, and reported.
    """
    txt = path.read_text(encoding="utf-8")
    m = re.search(r"<style>\n(.*?)\n\s*</style>", txt, re.S)
    if not m:
        raise ValueError(f"no <style> block found in {path.name}")

    lines = m.group(1).split("\n")
    bal, last_ok = 0, -1
    for i, ln in enumerate(lines):
        bal += ln.count("{") - ln.count("}")
        if bal < 0:
            break            # first orphaned close -- everything from the last
        if bal == 0:         # balanced point onward is unusable
            last_ok = i

    if bal == 0 and last_ok == len(lines) - 1:
        return "\n".join(lines)

    kept, dropped = lines[:last_ok + 1], lines[last_ok + 1:]
    css = "\n".join(kept)
    if css.count("{") != css.count("}"):
        raise ValueError(
            f"{path.name}: style block cannot be balanced even after trimming "
            f"({css.count('{')} open vs {css.count('}')} close). Inspect it "
            "by hand -- do not ship a guess.")
    print(f"  ! {path.name}: dropped {len(dropped)} trailing lines of orphaned "
          f"CSS (declarations with no selector). These rules are already dead "
          f"in {path.name} itself and are not inherited here.")
    for d in dropped:
        if d.strip():
            print(f"      · {d.strip()[:88]}")
    return css


def extract_logo() -> tuple[str, str]:
    """The knocked-out KSD shield, already a base64 data URI.

    Probes every dashboard template rather than depending on one file; the
    image is byte-identical in all of them.
    """
    for p in LOGO_SOURCES:
        if not p.exists():
            continue
        m = re.search(r'src="(data:image/png;base64,[A-Za-z0-9+/=]{500,})"',
                      p.read_text(encoding="utf-8"))
        if m:
            return m.group(1), p.name
    raise ValueError("no base64 KSD logo found in any dashboard template: "
                     + ", ".join(p.name for p in LOGO_SOURCES))


def main() -> int:
    for p in [LIQ, *PARTS.values()]:
        if not p.exists():
            print(f"ERROR: missing part {p}", file=sys.stderr)
            return 1

    base = extract_style(LIQ)
    add = PARTS["css"].read_text(encoding="utf-8")
    logo, logo_src = extract_logo()
    body = PARTS["body"].read_text(encoding="utf-8").replace("__LOGO__", logo)
    js = PARTS["js"].read_text(encoding="utf-8")

    if "__PAYLOAD__" not in js:
        print("ERROR: _aging_app.js has no __PAYLOAD__ placeholder",
              file=sys.stderr)
        return 1

    html = (
        "<!doctype html>\n<html lang=\"en\">\n<head>\n"
        "<meta charset=\"utf-8\" />\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n"
        "<title>K.S. Distillery &mdash; Aging Stock &middot; Live</title>\n"
        f"{CHART_JS}\n"
        "<style>\n" + base + "\n" + add + "\n</style>\n"
        "</head>\n" + body + "\n<script>\n" + js + "\n</script>\n"
        "</body>\n</html>\n"
    )
    OUT.write_text(html, encoding="utf-8")

    css_all = base + add
    print(f"built {OUT.name}")
    print(f"  base css   : {base.count(chr(10))+1} lines (from {LIQ.name})")
    print(f"  addendum   : {add.count(chr(10))+1} lines")
    print(f"  logo       : from {logo_src}")
    print(f"  braces     : {css_all.count('{')} open / {css_all.count('}')} close")
    print(f"  total      : {len(html):,} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
