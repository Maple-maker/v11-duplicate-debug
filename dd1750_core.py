"""DD1750 core: parse BOM PDFs and render DD Form 1750 overlays.

FIXED VERSION:
- 18 items per page (not 40)
- Correct NSN assignment (Item 1 gets its own NSN)
- Clean descriptions (remove material IDs like C_75Q65)
"""

from __future__ import annotations

import io
import math
import os
import re
from dataclasses import dataclass
from typing import List, Tuple

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics


# --- Constants derived from the supplied blank template (letter: 612x792)
PAGE_W, PAGE_H = 612.0, 792.0

# Column x-bounds (points)
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

# Table y-bounds (points)
Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5

# FIXED: Standard DD1750 has 18 lines, not 40
ROWS_PER_PAGE = 18

# Compute row height from table bounds
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE

# Text padding inside cells
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


_B_PREFIX_RE = re.compile(r"^B\s+(.+?)\s+(\d+)\s*$")
_NSN_RE = re.compile(r"^(\d{9})$")


def _clean_desc(desc: str) -> str:
    """Clean description to only include essential nomenclature (red box text)."""
    desc = desc.strip()
    desc = re.sub(r"\s+", " ", desc)
    
    # Remove material IDs and part numbers that appear at start
    # These typically look like: C_75Q65, 1354640W, etc.
    desc = re.sub(r"^[A-Z0-9_\-]{6,}\s*[\~\-\s]*\s*", "", desc)
    
    # Remove common BOM header fragments
    desc = desc.replace("COMPONENT LISTING / HAND RECEIPT", "").strip()
    desc = re.sub(r"\bWTY\b.*\bAuth\b\s*Qty\b", "", desc, flags=re.IGNORECASE).strip()
    
    # Remove any remaining coded columns (usually 2-3 character codes)
    desc = re.sub(r"\b[A-Z0-9]{1,3}\s+[A-Z0-9]{1,3}\s+[A-Z0-9]{1,3}\b", "", desc)
    
    # Clean up multiple spaces
    desc = re.sub(r"\s{2,}", " ", desc).strip()
    
    # Truncate if too long for DD1750
    if len(desc) > 120:
        desc = desc[:117] + "..."
    
    return desc


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract line items from a text-based BOM PDF with FIXED NSN assignment."""
    
    items: List[Tuple[str, str, int]] = []
    
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[start_page:]
        for p in pages:
            text = p.extract_text() or ""
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            i = 0
            while i < len(lines):
                ln = lines[i]
                m = _B_PREFIX_RE.match(ln)
                if m:
                    desc_raw, qty_s = m.group(1), m.group(2)
                    qty = int(qty_s)
                    
                    # Clean description FIRST
                    desc = _clean_desc(desc_raw)
                    
                    nsn = ""
                    # FIXED: Look ahead for NSN, but stop at next item
                    for j in range(i + 1, min(i + 5, len(lines))):
                        # Stop if we hit another item
                        if _B_PREFIX_RE.match(lines[j]):
                            break
                        # Check for NSN
                        if _NSN_RE.match(lines[j]):
                            nsn = lines[j]
                            break
                    
                    items.append((desc, nsn, qty))
                i += 1
    
    # Convert to BomItem objects
    out: List[BomItem] = []
    for idx, (desc, nsn, qty) in enumerate(items, start=1):
        out.append(BomItem(line_no=idx, description=desc, nsn=nsn, qty=qty))
    return out


def _wrap_to_width(text: str, font: str, size: float, max_w: float, max_lines: int) -> List[str]:
    """Greedy word-wrap using actual font metrics."""
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return [""]
    
    words = text.split(" ")
    lines: List[str] = []
    cur = ""
    
    def fits(s: str) -> bool:
        return pdfmetrics.stringWidth(s, font, size) <= max_w
    
    for w in words:
        if not cur:
            trial = w
        else:
            trial = cur + " " + w
        
        if fits(trial):
            cur = trial
            continue
        
        if not cur:
            chunk = w
            while chunk:
                lo, hi = 1, len(chunk)
                best = 1
                while lo <= hi:
                    mid = (lo + hi) // 2
                    cand = chunk[:mid]
                    if fits(cand):
                        best = mid
                        lo = mid + 1
                    else:
                        hi = mid - 1
                lines.append(chunk[:best])
                chunk = chunk[best:]
                if len(lines) >= max_lines:
                    return lines[:max_lines]
            cur = ""
        else:
            lines.append(cur)
            cur = w if fits(w) else w
            if len(lines) >= max_lines:
                return lines[:max_lines]
    
    if cur:
        lines.append(cur)
    
    if len(lines) > max_lines:
        lines = lines[:max_lines]
    
    return lines


def _draw_center(c: canvas.Canvas, txt: str, x_l: float, x_r: float, y: float, font: str, size: float):
    c.setFont(font, size)
    x = (x_l + x_r) / 2.0
    c.drawCentredString(x, y, txt)


def _build_overlay_page(items: List[BomItem], page_num: int, total_pages: int) -> bytes:
    """Return a PDF bytes for a single overlay page with 18 rows max."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    
    FONT_MAIN = "Helvetica"
    FONT_SMALL = "Helvetica"
    
    # Adjust baseline for 18 rows
    first_row_top = Y_TABLE_TOP_LINE - 5.0
    
    max_content_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X
    
    for row_idx in range(ROWS_PER_PAGE):  # Now 18 rows
        item_idx = row_idx
        y_row_top = first_row_top - row_idx * ROW_H
        y_desc = y_row_top - 7.0
        y_nsn = y_row_top - 12.2
        
        if item_idx >= len(items):
            continue
        
        it = items[item_idx]
        
        # Box number
        _draw_center(c, str(it.line_no), X_BOX_L, X_BOX_R, y_desc, FONT_MAIN, 8)
        
        # Contents: Keep it simple - description on top, NSN below if space
        desc_lines = _wrap_to_width(it.description, FONT_MAIN, 7.0, max_content_w, max_lines=1)
        
        c.setFont(FONT_MAIN, 7.0)
        c.drawString(X_CONTENT_L + PAD_X, y_desc, desc_lines[0])
        
        # NSN on separate line if present
        if it.nsn:
            c.setFont(FONT_SMALL, 6.0)
            nsn_text = f"NSN: {it.nsn}"
            # Ensure NSN fits
            if pdfmetrics.stringWidth(nsn_text, FONT_SMALL, 6.0) <= max_content_w:
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, nsn_text)
        
        # Quantities
        _draw_center(c, "EA", X_UOI_L, X_UOI_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, str(it.qty), X_INIT_L, X_INIT_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, "0", X_SPARES_L, X_SPARES_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, str(it.qty), X_TOTAL_L, X_TOTAL_R, y_desc, FONT_MAIN, 8)
    
    c.showPage()
    c.save()
    return buf.getvalue()


def generate_dd1750_from_pdf(
    bom_pdf_path: str,
    template_pdf_path: str,
    out_pdf_path: str,
    start_page: int = 0,
) -> Tuple[str, int]:
    """Generate DD1750 PDF with all fixes applied."""
    
    items = extract_items_from_pdf(bom_pdf_path, start_page=start_page)
    item_count = len(items)
    
    if item_count == 0:
        # Create a single-page copy of the template
        reader = PdfReader(template_pdf_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(out_pdf_path, "wb") as f:
            writer.write(f)
        return out_pdf_path, 0
    
    total_pages = math.ceil(item_count / ROWS_PER_PAGE)
    
    writer = PdfWriter()
    
    for p in range(total_pages):
        chunk = items[p * ROWS_PER_PAGE : (p + 1) * ROWS_PER_PAGE]
        overlay_pdf = _build_overlay_page(chunk, page_num=p + 1, total_pages=total_pages)
        overlay_reader = PdfReader(io.BytesIO(overlay_pdf))
        
        fresh_template = PdfReader(template_pdf_path).pages[0]
        fresh_template.merge_page(overlay_reader.pages[0])
        writer.add_page(fresh_template)
    
    with open(out_pdf_path, "wb") as f:
        writer.write(f)
    
    return out_pdf_path, item_count
