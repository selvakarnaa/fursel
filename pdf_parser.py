"""
pdf_parser.py
-------------
Column-aware parser for Rolls-Royce / ATA-style engine manuals.

KEY DESIGN: Single-pass extraction using page.get_text("dict") which returns
text blocks (type=0) and image blocks (type=1) INTERLEAVED in Y-position order.
This means images are captured at their EXACT position within the subtask content
flow — same ordering as they appear in the PDF page.

Image naming:  {subtask_id_safe}__{page:03d}__img{seq}.{ext}
               e.g.  72_41_31_110_066_001__p039__img0.png
"""

import re
import uuid
import argparse
from pathlib import Path
from dataclasses import dataclass, field

import fitz  # PyMuPDF

from database import Database

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

MIN_IMAGE_WIDTH  = 100
MIN_IMAGE_HEIGHT = 100
ROW_TOL          = 6.0

ID_CODE = r"\d{2}-\d{2}-\d{2}(?:-\d{2,3}){1,3}"

SUBTASK_HEADER = re.compile(rf"^\s*SUBTASK\s+({ID_CODE})", re.I)
SUBTASK_ANY    = re.compile(rf"SUBTASK\s+({ID_CODE})", re.I)
TASK_ANY       = re.compile(rf"(?:OP\s+|SPM\s+|EM\s+)?TASK\s+({ID_CODE})", re.I)
FIGURE_REF     = re.compile(rf"Fig(?:ure)?\s+({ID_CODE})", re.I)
# Caption pattern: "Figure 72-00-00-995-830, Sheet 1: Title text"
FIGURE_CAPTION = re.compile(rf"^Figure\s+({ID_CODE})(?:[,\s].*)?$", re.I)
DATA_CARD      = re.compile(r"([A-Z]{2,}-DC-[A-Z0-9]+)", re.I)
RTV_PAT        = re.compile(r"\bRTV\d+\b", re.I)

CALLOUT        = re.compile(r"^\s*(CAUTION|WARNING|NOTE|NOTICE)\b\s*:?\s*", re.I)
CALLOUT_LABEL  = re.compile(r"^\s*(CAUTION|WARNING|NOTE|NOTICE)\s*:?\s*$", re.I)

STEP_PAT       = re.compile(r"^\s*(?:[A-Z]\.|\(\d+\)|\([a-z]\)|\d{1,2}\.)\ ")
NUM_TITLE_PAT  = re.compile(r"^\s*(\d{1,2})\.\s+(.*)")

ACCOUNTABILITY = re.compile(
    r"^(PART ACCOUNTABILITY|OMAT ACCOUNTABILITY|TOOLING ACCOUNTABILITY)", re.I)

ID_ONLY_PAT    = re.compile(rf"^\s*({ID_CODE})\s*$")
PAGE_FOOTER    = re.compile(r"^\s*Page\s+\d+\s+of\s+\d+", re.I)
NORETAIN       = re.compile(r"DO\s+NOT\s+RETAIN", re.I)
FILTERING      = re.compile(r"^\s*Filtering\s+is\s+on", re.I)

JUNK_PATTERNS = [
    r"Filtering is on", r"DO\s+NOT\s+RETAIN", r"System\s+PRINTED", r"\bPRINTED\s*:",
    r"INITIATED\s+USER", r"Check and Rectify Manual", r"Engine Manual",
    r"Standard Practices Manual", r"Overhaul Processes Manual",
    r"Export Rating", r"Model\s*:\s*Trent", r"Effectivity\s*:", r"RevDate\s*:",
    r"Page\s+\d+\s+of\s+\d+", r"ROLLS[- ]ROYCE", r"PROPRIETARY",
    r"NOT FOR MANUFACTURE", r"INSPECTION\s*/\s*CHECK", r"\bREPAIR\s+\d{3}\b",
    r"^\s*(ASSEMBLY|INSTALLATION|STORAGE)\s*$",
]
JUNK_RE = re.compile("|".join(JUNK_PATTERNS), re.I)

_MONTHS = {
    "JAN": "Jan", "FEB": "Feb", "MAR": "Mar", "APR": "Apr", "MAY": "May",
    "JUN": "Jun", "JUL": "Jul", "AUG": "Aug", "SEP": "Sep", "OCT": "Oct",
    "NOV": "Nov", "DEC": "Dec", "01": "Jan", "02": "Feb", "03": "Mar",
    "04": "Apr", "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
}
REVDATE_RE   = re.compile(r"RevDate\s*[:\s]\s*(\d{1,2})[ /\-]([A-Za-z]{3,}|\d{2})[ /\-](\d{4})", re.I)
SPACEDATE_RE = re.compile(r"\b(\d{1,2})\s+([A-Z]{3})\s+(\d{4})\b")

_LIGATURES = {
    "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl",
    "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "ft", "\ufb06": "st",
}


def is_junk(text: str) -> bool:
    return bool(JUNK_RE.search(text))


DISPOSITION_RE = re.compile(
    r"\b(Accept|Reject|Repair,|CIR\s+TASK|SPM\s+TASK|OP\s+TASK|EM\s+TASK)\b", re.I)
MEASURE_RE = re.compile(r"\d[\d,\.]*\s*mm|in\.\)|Refer\s+to\s+Fig", re.I)


def looks_like_content(text: str) -> bool:
    if STEP_PAT.match(text): return True
    if DISPOSITION_RE.search(text): return True
    if MEASURE_RE.search(text): return True
    return False


def parse_revdate(text: str) -> str:
    m = REVDATE_RE.search(text)
    if m:
        d, mon, y = m.groups()
        key = mon.upper()[:3] if mon.isalpha() else mon
        return f"{d.zfill(2)} {_MONTHS.get(key, mon[:3].capitalize())} {y}"
    m = SPACEDATE_RE.search(text)
    if m:
        d, mon, y = m.groups()
        return f"{d.zfill(2)} {_MONTHS.get(mon.upper(), mon.capitalize())} {y}"
    return ""


def _clean_title(text: str) -> str:
    text = text.lstrip("| ").strip()
    m = NUM_TITLE_PAT.match(text)
    return m.group(2).strip() if m else text


# ─────────────────────────────────────────────
# Data class
# ─────────────────────────────────────────────

@dataclass
class SubtaskRecord:
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    source_pdf: str = ""
    task_id: str = ""
    subtask_id: str = ""
    title: str = ""
    revision_date: str = ""
    # content_items holds the ORDERED sequence of text and images as they
    # appear in the PDF.  Each item is a dict:
    #   text  item: {"type": "text"|"step"|"callout"|"row", "text": str}
    #   image item: {"type": "image", "path": str, "label": str}
    content_items: list = field(default_factory=list)
    # Flat convenience lists (derived from content_items during finalise)
    procedure_steps: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    cautions: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    figure_refs: list = field(default_factory=list)
    figure_images: list = field(default_factory=list)   # [{label, path}] — kept for DB compat
    data_cards: list = field(default_factory=list)
    cross_refs: list = field(default_factory=list)
    rtv_refs: list = field(default_factory=list)
    accountability: dict = field(default_factory=dict)
    raw_text: str = ""


# ─────────────────────────────────────────────
# Text cleaning
# ─────────────────────────────────────────────

def _clean(text: str) -> str:
    text = re.sub(r"([A-Za-z])-\n([A-Za-z])", r"\1\2", text)
    for k, v in _LIGATURES.items():
        text = text.replace(k, v)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def _collapse(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


# ─────────────────────────────────────────────
# Block extraction (text + images, position-ordered)
# ─────────────────────────────────────────────

def _raw_blocks(page) -> list[dict]:
    """
    Return ALL blocks (text AND image) sorted by Y position (top→bottom).
    Text blocks  → {"type": "text",  "text": str,   "bbox": tuple}
    Image blocks → {"type": "image", "bytes": bytes, "ext": str,
                    "width": int, "height": int, "bbox": tuple}
    """
    out = []
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        bbox = tuple(b.get("bbox", (0, 0, 0, 0)))

        if b.get("type") == 1:
            # ── Image block ──────────────────────────────────────────
            w = b.get("width",  0)
            h = b.get("height", 0)
            if w >= MIN_IMAGE_WIDTH and h >= MIN_IMAGE_HEIGHT:
                out.append({
                    "type":   "image",
                    "bytes":  b.get("image", b""),
                    "ext":    b.get("ext", "png"),
                    "width":  w,
                    "height": h,
                    "bbox":   bbox,
                })
            continue

        if b.get("type") != 0:
            continue

        # ── Text block ───────────────────────────────────────────────
        lines = b.get("lines", [])
        if not lines:
            continue
        line_ys = [ln["bbox"][1] for ln in lines]
        side_by_side = (len(lines) >= 2 and max(line_ys) - min(line_ys) <= ROW_TOL)

        if side_by_side:
            for ln in lines:
                spans = [sp.get("text", "") for sp in ln.get("spans", [])
                         if sp.get("text", "").strip()]
                if spans:
                    text = _clean(" ".join(spans))
                    if text:
                        out.append({"type": "text", "text": text, "bbox": tuple(ln["bbox"])})
        else:
            lines_text = []
            for ln in lines:
                spans = [sp.get("text", "") for sp in ln.get("spans", [])
                         if sp.get("text", "").strip()]
                if spans:
                    lines_text.append(" ".join(spans))
            if lines_text:
                text = _clean("\n".join(lines_text))
                if text:
                    out.append({"type": "text", "text": text, "bbox": bbox})

    # Sort everything by Y position so text and images interleave naturally
    out.sort(key=lambda b: b["bbox"][1])
    return out


def _group_text_rows(blocks: list[dict]) -> list[list[dict]]:
    """Group text blocks that share the same Y into column rows."""
    if not blocks:
        return []
    rows, anchor = [], None
    for b in blocks:
        y0 = b["bbox"][1]
        if rows and anchor is not None and abs(y0 - anchor) <= ROW_TOL:
            rows[-1].append(b)
        else:
            rows.append([b])
            anchor = y0
    for row in rows:
        row.sort(key=lambda b: b["bbox"][0])
    return rows


def _is_footer(text: str) -> bool:
    return bool(PAGE_FOOTER.match(text) or NORETAIN.search(text) or FILTERING.match(text))


# ─────────────────────────────────────────────
# Single-pass page extraction
# ─────────────────────────────────────────────

def extract_page_items(
    page,
    page_idx: int,
    images_dir: Path,
    current_sid: str,
) -> tuple[list[tuple], str]:
    """
    Returns (items, page_task_code).

    items is an ORDERED list of tuples preserving the exact PDF sequence:
      ("subtask",  text)
      ("callout",  text)
      ("step",     text)
      ("text",     text)
      ("row",      text)
      ("image",    image_path_str)   ← NEW — at the exact Y position

    Images are written to disk here so they can be slotted inline.
    """
    height = page.rect.height
    raw    = _raw_blocks(page)

    page_task_code = ""
    img_seq        = 0
    items          = []

    # We need to process text blocks in column-row groups but preserve
    # their position relative to image blocks.
    # Strategy: walk blocks in Y order; when we hit an image, emit it
    # immediately; accumulate consecutive text blocks and flush as row groups.

    pending_text: list[dict] = []

    def flush_text():
        nonlocal pending_text
        if not pending_text:
            return
        for row in _group_text_rows(pending_text):
            _emit_text_row(row)
        pending_text = []

    def _emit_text_row(row):
        nonlocal page_task_code
        in_footer = row[0]["bbox"][1] > height * 0.90
        in_header = row[0]["bbox"][3] < height * 0.08

        if len(row) == 1:
            txt = _collapse(row[0]["text"])
            if not txt:
                return
            if in_footer and not SUBTASK_HEADER.match(txt):
                mm = re.search(ID_CODE, txt)
                if mm and not page_task_code:
                    page_task_code = mm.group(0)
            if (in_footer or in_header) and not SUBTASK_HEADER.match(txt):
                if _is_footer(txt) or ID_ONLY_PAT.match(txt) or len(txt) < 4:
                    return
            if SUBTASK_HEADER.match(txt):
                items.append(("subtask", txt))
            elif CALLOUT.match(txt) and not CALLOUT_LABEL.match(txt):
                items.append(("callout", txt))
            elif STEP_PAT.match(txt):
                items.append(("step", txt))
            else:
                items.append(("text", txt))

        elif len(row) == 2:
            left  = _collapse(row[0]["text"])
            right = _collapse(row[1]["text"])
            in_footer = row[0]["bbox"][1] > height * 0.90
            in_header = row[0]["bbox"][3] < height * 0.08
            if in_footer and not SUBTASK_HEADER.match(left):
                mm = re.search(ID_CODE, left)
                if mm and not page_task_code:
                    page_task_code = mm.group(0)
            if (in_footer or in_header) and not SUBTASK_HEADER.match(left):
                if _is_footer(left) or ID_ONLY_PAT.match(left):
                    return
            if SUBTASK_HEADER.match(left):
                items.append(("subtask", left))
                if right:
                    items.append(("text", right))
            elif CALLOUT_LABEL.match(left) or CALLOUT.match(left):
                label = CALLOUT.match(left).group(1).upper() if CALLOUT.match(left) else "NOTE"
                items.append(("callout", f"{label}: {right}".strip()))
            elif STEP_PAT.match(left):
                merged = left if not right else f"{left}\n        {right}"
                items.append(("step", merged))
            else:
                items.append(("row", f"{left}  |  {right}" if right else left))
        else:
            cells = [_collapse(b["text"]) for b in row if _collapse(b["text"])]
            if cells:
                items.append(("row", " | ".join(cells)))

    for blk in raw:
        if blk["type"] == "image":
            # Flush any accumulated text above this image first
            flush_text()
            # Save image to disk at this exact position
            nonlocal_sid = current_sid  # subtask active when this page started
            if nonlocal_sid:
                fname = f"{_safe_sid(nonlocal_sid)}__p{page_idx:03d}__img{img_seq}.{blk['ext']}"
            else:
                fname = f"unassigned__{uuid.uuid4()}.{blk['ext']}"
            dest = images_dir / fname
            dest.write_bytes(blk["bytes"])
            items.append(("image", str(dest)))
            img_seq += 1
        else:
            pending_text.append(blk)

    # Flush any remaining text after last image
    flush_text()

    return items, page_task_code


# ─────────────────────────────────────────────
# Assemble subtask records
# ─────────────────────────────────────────────

def build_records(
    doc_items: list[tuple],
    source_pdf: str,
) -> list["SubtaskRecord"]:
    """
    doc_items: ordered (kind, text_or_path, page_task_code) for whole document.
    Kinds:  subtask | callout | step | text | row | image
    """
    records: list[SubtaskRecord] = []

    def finalize(rec: SubtaskRecord, raw_lines: list[str]):
        rec.raw_text    = "\n".join(raw_lines)
        rec.figure_refs = sorted(set(FIGURE_REF.findall(rec.raw_text)))
        rec.data_cards  = sorted(set(DATA_CARD.findall(rec.raw_text)))
        rec.rtv_refs    = sorted(set(RTV_PAT.findall(rec.raw_text)))
        subs  = [m for m in SUBTASK_ANY.findall(rec.raw_text) if m != rec.subtask_id]
        tasks = TASK_ANY.findall(rec.raw_text)
        rec.cross_refs  = sorted(set(subs + tasks))
        # Build figure_images from content_items for DB storage
        rec.figure_images = [
            {"label": Path(ci["path"]).name, "path": ci["path"]}
            for ci in rec.content_items if ci["type"] == "image"
        ]
        records.append(rec)

    cur: SubtaskRecord | None = None
    title_set = False
    skipping  = False
    raw_lines: list[str] = []

    for kind, text, page_task in doc_items:

        # ── Image item — always emit inline, never skip ──────────────
        if kind == "image":
            if cur is not None:
                cur.content_items.append({"type": "image", "path": text, "label": Path(text).name})
            continue

        # ── Subtask boundary ─────────────────────────────────────────
        if kind == "subtask":
            if cur is not None:
                finalize(cur, raw_lines)
            sid = SUBTASK_ANY.search(text).group(1)
            cur = SubtaskRecord(subtask_id=sid, task_id=page_task or "", source_pdf=source_pdf)
            title_set = False
            skipping  = False
            raw_lines = [text]
            continue

        if cur is None:
            continue

        rv = parse_revdate(text)
        if rv and not cur.revision_date:
            cur.revision_date = rv
        if not cur.task_id:
            idm = ID_ONLY_PAT.match(text)
            if idm:
                cur.task_id = idm.group(1)

        # ── Callouts ─────────────────────────────────────────────────
        if kind == "callout":
            skipping = False
            raw_lines.append(text)
            low = text.lower()
            cur.content_items.append({"type": "callout", "text": text})
            if low.startswith("caution"):
                cur.cautions.append(text)
            elif low.startswith("warning") or low.startswith("notice"):
                cur.warnings.append(text)
            else:
                cur.notes.append(text)
            continue

        # ── Steps ────────────────────────────────────────────────────
        if kind == "step":
            skipping = False
            if not title_set:
                cur.title = _clean_title(text)
                title_set = True
                raw_lines.append(text)
                continue
            raw_lines.append(text)
            cur.procedure_steps.append(text)
            cur.content_items.append({"type": "step", "text": text})
            continue

        # ── Accountability headers ────────────────────────────────────
        if ACCOUNTABILITY.match(text):
            skipping = False
            cur.accountability[text.upper().split()[0]] = []
            continue

        # ── Junk filtering ────────────────────────────────────────────
        content = looks_like_content(text)
        if not content and (is_junk(text) or ID_ONLY_PAT.match(text)):
            skipping = True
            continue
        if not content and skipping:
            continue

        skipping = False

        if not title_set:
            m = NUM_TITLE_PAT.match(text)
            if m:
                cur.title = m.group(2).strip()
                title_set = True
                raw_lines.append(text)
                continue

        raw_lines.append(text)
        cur.procedure_steps.append(text)
        cur.content_items.append({"type": "step", "text": text})

    if cur is not None:
        finalize(cur, raw_lines)

    return records


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _safe_sid(subtask_id: str) -> str:
    return subtask_id.replace("-", "_")


# ─────────────────────────────────────────────
# Figure index — scan whole PDF for figure pages
# ─────────────────────────────────────────────

def build_figure_index(
    doc,
    source_pdf: str,
    images_dir: Path,
    start: int = 0,
    end: int = None,
) -> list[dict]:
    """
    Scan every page for a figure caption line like:
        "Figure 72-00-00-995-830, Sheet 1: Visually examine IP bleed valve"
    When found, capture the largest image on that page and register it
    in the figure index.

    Returns a list of dicts: {figure_id, image_path, caption, page_num, source_pdf}
    """
    last = doc.page_count if end is None else min(end, doc.page_count)
    index: list[dict] = []
    seen: set[str] = set()

    for i in range(start, last):
        page = doc[i]
        # Collect all text on the page to find caption lines
        page_text = page.get_text("text")
        caption_match = None
        figure_id = None
        for line in page_text.splitlines():
            line = line.strip()
            m = FIGURE_CAPTION.match(line)
            if m:
                figure_id = m.group(1)
                caption_match = line
                break

        if not figure_id:
            continue  # No figure caption on this page

        # Avoid duplicate registrations (Sheet 2 of same figure)
        if figure_id in seen:
            continue
        seen.add(figure_id)

        # Find the largest image on this page
        best_img = None
        best_area = 0
        d = page.get_text("dict")
        for b in d.get("blocks", []):
            if b.get("type") == 1:
                w = b.get("width", 0)
                h = b.get("height", 0)
                area = w * h
                if area > best_area and w >= MIN_IMAGE_WIDTH and h >= MIN_IMAGE_HEIGHT:
                    best_area = area
                    best_img = b

        if best_img is None:
            # No inline image found — render the whole page as an image
            pix = page.get_pixmap(dpi=150)
            safe_fid = figure_id.replace("-", "_")
            fname = f"figidx__{safe_fid}__p{i:03d}.png"
            dest = images_dir / fname
            pix.save(str(dest))
            index.append({
                "figure_id": figure_id,
                "image_path": str(dest),
                "caption": caption_match or "",
                "page_num": i,
                "source_pdf": source_pdf,
            })
        else:
            ext = best_img.get("ext", "png")
            safe_fid = figure_id.replace("-", "_")
            fname = f"figidx__{safe_fid}__p{i:03d}.{ext}"
            dest = images_dir / fname
            dest.write_bytes(best_img["image"])
            index.append({
                "figure_id": figure_id,
                "image_path": str(dest),
                "caption": caption_match or "",
                "page_num": i,
                "source_pdf": source_pdf,
            })

    return index


# ─────────────────────────────────────────────
# Top-level extraction
# ─────────────────────────────────────────────

def extract_pdf(pdf_path: str, db_path: str = "output/saesl.db",
                start: int = 0, end=None) -> list[SubtaskRecord]:
    source_path = Path(pdf_path)
    pdf_name    = source_path.stem.replace(" ", "_")
    output_dir  = Path("output") / pdf_name
    images_dir  = output_dir / "images"
    output_dir.mkdir(parents=True, exist_ok=True)
    images_dir.mkdir(exist_ok=True)

    print(f"\n{'='*60}\nProcessing: {pdf_path}\n{'='*60}")

    doc   = fitz.open(pdf_path)
    last  = doc.page_count if end is None else min(end, doc.page_count)

    # Single pass: text + images extracted together in position order
    doc_items: list[tuple] = []
    last_task = ""
    last_sid  = ""

    img_count = 0

    for i in range(start, last):
        page  = doc[i]

        # Pre-scan: does this page open a new subtask header?
        # If so, update last_sid BEFORE image extraction so images get the right name.
        pre_text = page.get_text("text")
        for line in pre_text.splitlines():
            m = SUBTASK_HEADER.match(line)
            if m:
                last_sid = m.group(1)
                break

        items, page_task = extract_page_items(
            page,
            page_idx    = i,
            images_dir  = images_dir,
            current_sid = last_sid,
        )
        if page_task:
            last_task = page_task

        for item in items:
            kind = item[0]
            val  = item[1]
            if kind == "subtask":
                m2 = SUBTASK_ANY.search(val)
                if m2:
                    last_sid = m2.group(1)
            if kind == "image":
                img_count += 1
            doc_items.append((kind, val, page_task or last_task))

    doc.close()

    print(f"  Pages processed : {last - start}")
    print(f"  Images extracted: {img_count}")

    records = build_records(doc_items, source_path.name)
    print(f"  Subtasks parsed : {len(records)}")

    # ── Build figure index (whole-PDF scan for figure caption pages) ──
    doc2 = fitz.open(pdf_path)
    figure_index = build_figure_index(doc2, source_path.name, images_dir, start, end if end else None)
    doc2.close()
    print(f"  Figures indexed : {len(figure_index)}")

    db = Database(db_path)
    for rec in records:
        db.save_subtask(rec)
    db.save_figure_index(figure_index)
    db.close()
    print(f"  Saved {len(records)} records to: {db_path}\n{'='*60}\n")

    return records


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Parse Rolls-Royce / ATA engine manual PDF")
    ap.add_argument("--pdf",   required=True)
    ap.add_argument("--db",    default="output/saesl.db")
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--end",   type=int, default=None)
    args = ap.parse_args()

    records = extract_pdf(args.pdf, args.db, args.start, args.end)
    print(f"\nSummary: {len(records)} subtasks")
    for r in records[:40]:
        img_count = sum(1 for ci in r.content_items if ci["type"] == "image")
        print(f"  • {r.subtask_id} | '{r.title[:45]}' | steps={len(r.procedure_steps)} imgs={img_count}")
    if len(records) > 40:
        print(f"  ... and {len(records) - 40} more")
