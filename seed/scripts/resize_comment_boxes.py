#!/usr/bin/env python3
"""
resize_comment_boxes.py — size every DETAIL-sheet receipt note box to fit its
text.  PURE ZIP/VML SURGERY: it rewrites only the comment-shape width/height in
the vmlDrawing parts and leaves every other byte of the workbook untouched
(charts, conditional formatting, cell values, the region-insight panels and the
comment TEXT are all byte-identical).

WHY THIS EXISTS (root cause, verified openpyxl 3.1.5):
  openpyxl's loader does NOT read a legacy comment's VML shape size — on every
  load it resets each box to its hard-coded default 144x79 px, and re-writes
  that default on save.  So the careful per-note height that
  ksbc_receipt_notes.py sets (width=240, height up to 320) is silently reverted
  the moment ANY later openpyxl pass reloads the workbook — in the KSBC pipeline
  that is the final ksbc_region_insights.py save.  Result: all boxes ship at
  144x79, the ~50-char title wraps to 3 lines, and every note with 2+ receipt
  dates is clipped after its first date (a 5 cs / 4-day receipt reads as
  "1 cs"; a 92 cs shop-total reads as "3 cs").  The note NUMBERS were always
  correct — only the box was too small.

THE FIX MUST RUN AFTER THE LAST openpyxl SAVE.  Because this pass edits the VML
directly and never re-loads via openpyxl, the sizes it writes are the sizes
Excel displays.  It is therefore invoked as the very last action of
ksbc_region_insights.py (the pipeline's final write step), and can also be run
standalone on any workbook.

Mapping: within a worksheet, xl/comments/comment{K}.xml (cell ref -> text) and
xl/drawings/commentsDrawing{K}.vml (shape -> Row/Column) share the index K, so
each shape is matched to its comment by cell reference (openpyxl-independent).

Usage:  python3 resize_comment_boxes.py "<workbook.xlsx>"          (in place)
        python3 resize_comment_boxes.py "<src.xlsx>" "<dst.xlsx>"  (copy)
Returns (n_boxes, h_min, h_max).  No-op (0 boxes) if the workbook has no
comments — never an error, never a build gate.
"""
from __future__ import annotations
import os, re, sys, html, zipfile, tempfile
from openpyxl.utils import get_column_letter

WIDTH = 260          # px — wide enough that a ~55-char title wraps to <=2 lines
CPL   = 42           # approx chars/line at WIDTH, 9pt Tahoma
LINE_PX = 15         # px per displayed line
PAD_PX = 16          # top+bottom padding
H_MIN, H_MAX = 60, 420

_SHAPE_RE = re.compile(
    r'(<(\w+):shape\b[^>]*type="#_x0000_t202"[^>]*>)(.*?)(</\2:shape>)', re.S)
_ROW_RE = re.compile(r'<\w+:Row>(\d+)</\w+:Row>')
_COL_RE = re.compile(r'<\w+:Column>(\d+)</\w+:Column>')
_STYLE_RE = re.compile(r'(width:)\d+(?:\.\d+)?px(;height:)\d+(?:\.\d+)?px')
_COMMENT_RE = re.compile(r'<comment ref="([A-Z]+\d+)"[^>]*>(.*?)</comment>', re.S)
_T_RE = re.compile(r'<t[^>]*>(.*?)</t>', re.S)


def _plain_text(block: str) -> str:
    return html.unescape("".join(_T_RE.findall(block)))


def _needed_h(text: str) -> int:
    disp = 0
    for ln in text.split("\n"):
        disp += max(1, -(-len(ln) // CPL))      # ceil division
    return max(H_MIN, min(H_MAX, PAD_PX + LINE_PX * disp))


def _parse_comments(xml: str) -> dict:
    return {m.group(1): _plain_text(m.group(2)) for m in _COMMENT_RE.finditer(xml)}


def _resize_vml(vml: str, text_by_ref: dict, stats: list) -> str:
    def repl(m):
        opentag, _pfx, body, closetag = m.groups()
        rm, cm = _ROW_RE.search(body), _COL_RE.search(body)
        if not (rm and cm):
            return m.group(0)
        ref = f"{get_column_letter(int(cm.group(1)) + 1)}{int(rm.group(1)) + 1}"
        txt = text_by_ref.get(ref)
        if txt is None:
            return m.group(0)
        h = _needed_h(txt)
        newopen = _STYLE_RE.sub(
            lambda s: f"{s.group(1)}{WIDTH}px{s.group(2)}{h}px", opentag)
        stats.append(h)
        return newopen + body + closetag
    return _SHAPE_RE.sub(repl, vml)


def resize_boxes(src: str, dst: str | None = None):
    """Resize every comment box in `src` to fit its text; write to `dst`
    (default: in place, atomically). Returns (n_boxes, h_min, h_max)."""
    dst = dst or src
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        infos = zin.infolist()
        data = {n: zin.read(n) for n in names}
    comments = {}
    for n in names:
        m = re.match(r'xl/comments/comment(\d+)\.xml$', n)
        if m:
            comments[m.group(1)] = _parse_comments(data[n].decode('utf-8'))
    heights = []
    for n in names:
        m = re.match(r'xl/drawings/commentsDrawing(\d+)\.vml$', n)
        if m and m.group(1) in comments:
            data[n] = _resize_vml(
                data[n].decode('utf-8'), comments[m.group(1)], heights).encode('utf-8')
    # atomic write: temp file in the destination dir, then os.replace
    d = os.path.dirname(os.path.abspath(dst)) or '.'
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=d)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                zout.writestr(info, data[info.filename])
        os.replace(tmp, dst)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    if not heights:
        return 0, 0, 0
    return len(heights), min(heights), max(heights)


def main(argv):
    if len(argv) not in (2, 3):
        print("usage: python3 resize_comment_boxes.py <workbook.xlsx> [<out.xlsx>]")
        return 1
    src = argv[1]
    dst = argv[2] if len(argv) == 3 else None
    n, lo, hi = resize_boxes(src, dst)
    if n:
        print(f"COMMENT BOXES SIZED · {n} boxes · height {lo}..{hi}px "
              f"(was fixed 79px) · width {WIDTH}px (was 144px)")
    else:
        print("COMMENT BOXES · none found (no-op)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
