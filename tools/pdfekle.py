import os
import uuid
from pypdf import PdfReader, PdfWriter

def insert_pdf(main_path: str, insert_path: str, after_page: int, output_dir: str) -> str:
    main_reader = PdfReader(main_path)
    insert_reader = PdfReader(insert_path)

    total_pages = len(main_reader.pages)
    if after_page < 0 or after_page > total_pages:
        raise ValueError(f"Sayfa numarası 0 ile {total_pages} arasında olmalıdır. (0 = başa ekle)")

    writer = PdfWriter()

    for i in range(after_page):
        writer.add_page(main_reader.pages[i])

    for page in insert_reader.pages:
        writer.add_page(page)

    for i in range(after_page, total_pages):
        writer.add_page(main_reader.pages[i])

    output_path = os.path.join(output_dir, f"pdf_ekle_{uuid.uuid4().hex[:8]}.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path
