import os
import uuid
from pypdf import PdfReader, PdfWriter
from tools.common import parse_page_ranges, zip_files

def split_pdf_by_range(input_path: str, output_dir: str, page_ranges: str) -> str:
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    selected_pages = parse_page_ranges(page_ranges, total_pages)

    writer = PdfWriter()
    for idx in selected_pages:
        writer.add_page(reader.pages[idx])

    output_path = os.path.join(output_dir, f"pdf_ayir_{uuid.uuid4().hex[:8]}.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

def split_pdf_to_single_pages(input_path: str, output_dir: str) -> str:
    reader = PdfReader(input_path)
    created = []
    for i, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        page_path = os.path.join(output_dir, f"sayfa_{i}_{uuid.uuid4().hex[:4]}.pdf")
        with open(page_path, "wb") as f:
            writer.write(f)
        created.append(page_path)

    zip_path = os.path.join(output_dir, f"pdf_tek_sayfalar_{uuid.uuid4().hex[:8]}.zip")
    return zip_files(created, zip_path)
