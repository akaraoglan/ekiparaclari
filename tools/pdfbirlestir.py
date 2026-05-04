import os
import uuid
from pypdf import PdfReader, PdfWriter

def merge_pdfs(file_paths, output_dir: str) -> str:
    writer = PdfWriter()
    for path in file_paths:
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)

    output_path = os.path.join(output_dir, f"pdf_birlestir_{uuid.uuid4().hex[:8]}.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path
