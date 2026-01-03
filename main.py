"""DD1750 core: parse BOM PDFs and render DD Form 1750 overlays.

FIXED VERSION:
- 18 items per page (not 40)
- Correct NSN assignment (NSN from Material column goes with item ABOVE)
- Only extract Description column text
- Ignore WTY, ARC, CIIC, UI, SCMC columns
"""

from __future__ import annotations

import io
import math
import os
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional

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


def _clean_description(text: str) -> str:
    """Clean description to only include essential nomenclature."""
    if not text:
        return ""
    
    # Remove any WTY, ARC, CIIC, UI, SCMC codes that might have leaked in
    text = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC)\b.*?\d*', '', text, flags=re.IGNORECASE)
    
    # Remove material IDs like C_75Q65 ~ 1354640W
    text = re.sub(r'C_[A-Z0-9]+\s*~\s*[A-Z0-9]+', '', text)
    
    # Remove COEI codes
    text = re.sub(r'COEI-\d+', '', text)
    
    # Remove any remaining numbers that are likely part numbers
    text = re.sub(r'\b\d{6,}\b', '', text)
    
    # Clean up whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Remove trailing commas, colons, or dashes
    text = re.sub(r'[,\-:]+\s*$', '', text)
    
    # Truncate if too long for DD1750
    if len(text) > 120:
        text = text[:117] + "..."
    
    return text


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract line items from BOM PDF with CORRECT NSN assignment."""
    
    items: List[BomItem] = []
    
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[start_page:]
        
        for page_num, page in enumerate(pages):
            # Extract the entire page text to understand structure
            text = page.extract_text() or ""
            lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
            
            # Parse line by line to match your BOM format
            i = 0
            current_item = None
            pending_nsn = None
            
            while i < len(lines):
                line = lines[i]
                
                # Look for item patterns
                # Pattern 1: Item with description (may have "B" in LV column)
                if re.match(r'^[AB]\s+', line) or any(x in line.upper() for x in ['ASSEMBLY', 'CABLE', 'CONNECTOR', 'SWITCH']):
                    # If we have a pending item, save it with its NSN
                    if current_item and pending_nsn:
                        current_item.nsn = pending_nsn
                        items.append(current_item)
                        pending_nsn = None
                    elif current_item:
                        items.append(current_item)
                    
                    # Extract description (skip first 2-3 columns which are Material, LV)
                    parts = line.split()
                    if len(parts) >= 3:
                        # Join remaining parts as description
                        desc_start = 2  # Skip Material and LV columns
                        description = ' '.join(parts[desc_start:])
                        
                        # Clean up - remove any column codes that got included
                        description = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\b', '', description, flags=re.IGNORECASE)
                        description = _clean_description(description)
                        
                        # Extract quantity if present at end
                        qty = 1
                        if description and description.split()[-1].isdigit():
                            qty = int(description.split()[-1])
                            description = ' '.join(description.split()[:-1])
                        
                        current_item = BomItem(
                            line_no=len(items) + 1,
                            description=description,
                            nsn="",  # Will be filled from next line
                            qty=qty
                        )
                
                # Pattern 2: NSN line (9-digit number)
                elif re.match(r'^\d{9}$', line):
                    pending_nsn = line
                
                # Pattern 3: Material line with NSN (e.g., "011800996 C_75Q65 ~ 1354640W")
                elif re.search(r'\b\d{9}\b', line):
                    nsn_match = re.search(r'\b(\d{9})\b', line)
                    if nsn_match:
                        pending_nsn = nsn_match.group(1)
                
                i += 1
            
            # Don't forget the last item
            if current_item:
                if pending_nsn:
                    current_item.nsn = pending_nsn
                items.append(current_item)
    
    # Alternative approach: Try table extraction if line parsing fails
    if len(items) == 0:
        items = _extract_via_tables(pdf_path, start_page)
    
    return items


def _extract_via_tables(pdf_path: str, start_page: int) -> List[BomItem]:
    """Alternative extraction using table structure."""
    items: List[BomItem] = []
    
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[start_page:]
        
        for page in pages:
            tables = page.extract_tables()
            
            if not tables:
                continue
            
            for table in tables:
                if len(table) < 2:
                    continue
                
                # Process rows
                for row_idx, row in enumerate(table):
                    if row_idx == 0:  # Skip header
                        continue
                    
                    # Look for description in likely columns
                    description = ""
                    nsn = ""
                    qty = 1
                    
                    for cell in row:
                        if not cell:
                            continue
                        
                        cell_text = str(cell).strip()
                        
                        # Check for description (contains words, not just codes)
                        if (len(cell_text) > 10 and 
                            re.search(r'[a-zA-Z]', cell_text) and
                            not re.match(r'^\d{9}$', cell_text) and
                            not re.match(r'^[AB]$', cell_text) and
                            not re.match(r'^(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)$', cell_text, re.IGNORECASE)):
                            description = _clean_description(cell_text)
                        
                        # Check for NSN (9-digit number)
                        elif re.match(r'^\d{9}$', cell_text):
                            nsn = cell_text
                        
                        # Check for quantity
                        elif cell_text.isdigit() and 1 <= int(cell_text) <= 100:
                            qty = int(cell_text)
                    
                    if description:
                        items.append(BomItem(
                            line_no=len(items) + 1,
                            description=description,
                            nsn=nsn,
                            qty=qty
                        ))
    
    return items


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
    
    # DEBUG: Print what was extracted
    print(f"DEBUG: Extracted {item_count} items")
    for i, item in enumerate(items[:10], 1):
        print(f"  Item {i}: '{item.description[:50]}...' | NSN: {item.nsn} | Qty: {item.qty}")
    
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
