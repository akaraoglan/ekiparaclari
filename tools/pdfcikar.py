import os
import uuid
from pypdf import PdfReader, PdfWriter
from tools.common import parse_page_ranges

def remove_pdf_pages(input_path: str, output_dir: str, pages_to_remove: str) -> str:
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    remove_indices = set(parse_page_ranges(pages_to_remove, total_pages))

    if len(remove_indices) >= total_pages:
        raise ValueError("Tüm sayfalar çıkarılamaz. En az bir sayfa kalmalıdır.")

    writer = PdfWriter()
    for idx in range(total_pages):
        if idx not in remove_indices:
            writer.add_page(reader.pages[idx])

    output_path = os.path.join(output_dir, f"pdf_cikar_{uuid.uuid4().hex[:8]}.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path
