import os
import uuid
from pypdf import PdfReader, PdfWriter
from tools.common import parse_page_ranges

def rotate_pdf(input_path: str, output_dir: str, rotation: int = 90, pages: str = "") -> str:
    if rotation not in [90, 180, 270]:
        raise ValueError("Döndürme açısı sadece 90, 180 veya 270 olabilir.")

    reader = PdfReader(input_path)
    writer = PdfWriter()
    total_pages = len(reader.pages)

    targets = set(parse_page_ranges(pages, total_pages)) if pages else set(range(total_pages))

    for idx, page in enumerate(reader.pages):
        if idx in targets:
            rotated = page.rotate(rotation)
            writer.add_page(rotated)
        else:
            writer.add_page(page)

    output_path = os.path.join(output_dir, f"pdf_dondur_{uuid.uuid4().hex[:8]}.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path
