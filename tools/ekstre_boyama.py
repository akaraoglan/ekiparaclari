import fitz
import os
import re
import uuid

BEYANNAME_RE = re.compile(r'\b\d{8}(IM|EX|AN)\d+\b', re.IGNORECASE)
DATE_RE = re.compile(r'\d{2}\.\d{2}\.\d{4}')

YELLOW     = (1.0, 1.0, 0.0)
LIGHT_BLUE = (0.53, 0.81, 0.98)
LIGHT_RED  = (1.0, 0.60, 0.60)
OPACITY    = 0.38
TRANSACTION_NAME = "Gümrük Vergi Tahsilatı"


def _extract_transaction_date(text: str) -> str:
    m = DATE_RE.search(text)
    return m.group(0) if m else None


def _out_filename(original: str, extracted_date: str = None) -> str:
    if extracted_date:
        return f"{extracted_date} VAKIFBANK_boyanmis.pdf"

    m = re.match(r'(\d{2}\.\d{2}\.\d{4})', os.path.basename(original))
    date = m.group(1) if m else ""
    return f"{date} VAKIFBANK_boyanmis.pdf" if date else "VAKIFBANK_boyanmis.pdf"


def _desc_bottom(page, hit_y1: float, page_w: float) -> float:
    """Return the y-bottom of the description lines below a hit."""
    clip = fitz.Rect(0, hit_y1 - 2, page_w, hit_y1 + 80)
    blocks = page.get_text("blocks", clip=clip, sort=True)
    bottom = hit_y1 + 26  # safe fallback (~2 description lines)
    for blk in blocks:
        if blk[6] != 0:
            continue
        txt = blk[4].strip()
        # stop if a new transaction row starts (date pattern)
        if re.match(r'\d{2}\.\d{2}\.\d{4}\s', txt):
            break
        # stop at page footer markers
        if txt.startswith("***") or txt.startswith("www."):
            break
        bottom = max(bottom, blk[3])
    return bottom


def paint_vakifbank_pdf(input_path: str, original_filename: str, output_dir: str) -> dict:
    """
    Paint Gümrük Vergi Tahsilatı rows:
      - IM beyanname → yellow
      - EX beyanname → light blue
      - Missing or unrecognized beyanname type → light red
    Returns dict with output_path, out_filename, extracted_date, im_count,
    ex_count and other_count.
    """
    doc = fitz.open(input_path)
    im_count = ex_count = other_count = 0
    extracted_date = None

    for page in doc:
        pw = page.rect.width
        hits = page.search_for(TRANSACTION_NAME)
        if not hits:
            continue

        shape = page.new_shape()

        for index, hit in enumerate(hits):
            next_hit = hits[index + 1] if index + 1 < len(hits) else None
            y_bottom = _desc_bottom(page, hit.y1, pw)
            if next_hit:
                y_bottom = min(y_bottom, next_hit.y0 - 2)

            # Limit classification to the current transaction. Otherwise a
            # following row's declaration code could classify this row.
            clip = fitz.Rect(0, hit.y1 - 2, pw, y_bottom + 2)
            desc_text = page.get_text("text", clip=clip)

            if not extracted_date:
                extracted_date = _extract_transaction_date(desc_text)

            m = BEYANNAME_RE.search(desc_text)
            code = m.group(1).upper() if m else None
            if code in ("IM", "AN"):
                fill = YELLOW
                im_count += 1
            elif code == "EX":
                fill = LIGHT_BLUE
                ex_count += 1
            else:
                fill = LIGHT_RED
                other_count += 1

            rect = fitz.Rect(8, hit.y0 - 1, pw - 8, y_bottom + 2)
            shape.draw_rect(rect)
            shape.finish(color=None, fill=fill, fill_opacity=OPACITY)

        shape.commit(overlay=False)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"vakifbank_boyama_{uuid.uuid4().hex[:8]}.pdf")
    doc.save(output_path)
    doc.close()

    out_filename = _out_filename(original_filename, extracted_date)

    return {
        "output_path": output_path,
        "out_filename": out_filename,
        "extracted_date": extracted_date,
        "im_count": im_count,
        "ex_count": ex_count,
        "other_count": other_count,
    }
