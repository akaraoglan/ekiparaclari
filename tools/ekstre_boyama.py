import fitz
import os
import re
import tempfile

BEYANNAME_RE = re.compile(r'\b\d{8}(IM|EX|AN)\d+\b', re.IGNORECASE)

YELLOW     = (1.0, 1.0, 0.0)
LIGHT_BLUE = (0.53, 0.81, 0.98)
OPACITY    = 0.38
TRANSACTION_NAME = "Gümrük Vergi Tahsilatı"


def _out_filename(original: str) -> str:
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


def paint_vakifbank_pdf(input_path: str, original_filename: str) -> dict:
    """
    Paint Gümrük Vergi Tahsilatı rows:
      - IM beyanname → yellow
      - EX beyanname → light blue
    Saves to SAVE_DIR if available.
    Returns dict with tmp_path, saved_path, out_filename, im_count, ex_count.
    """
    out_filename = _out_filename(original_filename)
    doc = fitz.open(input_path)
    im_count = ex_count = 0

    for page in doc:
        pw = page.rect.width
        hits = page.search_for(TRANSACTION_NAME)
        if not hits:
            continue

        shape = page.new_shape()

        for hit in hits:
            # Extract description text clipped below the hit
            clip = fitz.Rect(0, hit.y1 - 2, pw, hit.y1 + 80)
            desc_text = page.get_text("text", clip=clip)

            m = BEYANNAME_RE.search(desc_text)
            if not m:
                continue

            code = m.group(1).upper()
            if code in ("IM", "AN"):
                fill = YELLOW
                im_count += 1
            else:
                fill = LIGHT_BLUE
                ex_count += 1

            y_bottom = _desc_bottom(page, hit.y1, pw)
            rect = fitz.Rect(8, hit.y0 - 1, pw - 8, y_bottom + 2)
            shape.draw_rect(rect)
            shape.finish(color=None, fill=fill, fill_opacity=OPACITY)

        shape.commit(overlay=False)

    tmp_path = tempfile.mktemp(suffix=".pdf")
    doc.save(tmp_path)
    doc.close()

    return {
        "tmp_path": tmp_path,
        "out_filename": out_filename,
        "im_count": im_count,
        "ex_count": ex_count,
    }
