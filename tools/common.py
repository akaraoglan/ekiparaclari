import os
import re
import time
import zipfile
import unicodedata

def ensure_dirs(*dirs):
    for d in dirs:
        os.makedirs(d, exist_ok=True)

def secure_tr_filename(filename: str) -> str:
    filename = filename.strip().replace("\\", "_").replace("/", "_")
    stem, dot, ext = filename.rpartition(".")
    if not dot:
        stem = filename
        ext = ""
    normalized = unicodedata.normalize("NFKD", stem)
    ascii_name = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_name = re.sub(r"[^A-Za-z0-9._-]+", "_", ascii_name).strip("._-")
    if not ascii_name:
        ascii_name = "dosya"
    if ext:
        clean_ext = re.sub(r"[^A-Za-z0-9]+", "", ext.lower())
        return f"{ascii_name}.{clean_ext}" if clean_ext else ascii_name
    return ascii_name

def parse_page_ranges(page_text: str, total_pages: int):
    result = []
    parts = [p.strip() for p in page_text.split(",") if p.strip()]
    if not parts:
        raise ValueError("Geçerli sayfa aralığı girilmedi.")
    for part in parts:
        if "-" in part:
            halves = part.split("-", 1)
            try:
                start = int(halves[0])
                end = int(halves[1])
            except ValueError:
                raise ValueError(f"Geçersiz aralık biçimi: '{part}'. Örnek: 1-3")
            if start < 1 or end > total_pages or start > end:
                raise ValueError(
                    f"Geçersiz aralık: '{part}'. Sayfa numaraları 1 ile {total_pages} arasında olmalı ve başlangıç ≤ bitiş."
                )
            result.extend(range(start - 1, end))
        else:
            try:
                page = int(part)
            except ValueError:
                raise ValueError(f"Geçersiz sayfa numarası: '{part}'. Sayısal bir değer girin.")
            if page < 1 or page > total_pages:
                raise ValueError(f"Geçersiz sayfa: '{part}'. Dosyada {total_pages} sayfa var.")
            result.append(page - 1)
    dedup = []
    for p in result:
        if p not in dedup:
            dedup.append(p)
    return dedup

def zip_files(file_paths, output_zip):
    with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in file_paths:
            zf.write(path, arcname=os.path.basename(path))
    return output_zip

def cleanup_old_files(folder: str, max_age_hours: int = 12):
    now = time.time()
    max_age_seconds = max_age_hours * 3600
    if not os.path.isdir(folder):
        return
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        try:
            if os.path.isfile(path) and (now - os.path.getmtime(path) > max_age_seconds):
                os.remove(path)
        except OSError:
            pass
