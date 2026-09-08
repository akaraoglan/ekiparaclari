import os
import re
import threading
import unicodedata
import uuid
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from statistics import median

import fitz
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


_OCR_ENGINE = None
_OCR_LOCK = threading.Lock()
_MONEY_RE = re.compile(r"(?<!\d)\d{1,3}(?:\.\d{3})*,\d{2}(?!\d)|(?<!\d)\d+,\d{2}(?!\d)")


@dataclass(frozen=True)
class TextBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2


def _plain(text: str) -> str:
    text = text.casefold().replace("ı", "i")
    text = "".join(
        char for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def _parse_money(text: str) -> Decimal | None:
    match = _MONEY_RE.search(text.replace(" ", ""))
    if not match:
        return None
    try:
        return Decimal(match.group(0).replace(".", "").replace(",", "."))
    except InvalidOperation:
        return None


def format_tr_money(value: Decimal) -> str:
    rendered = f"{value:,.2f}"
    return rendered.replace(",", "_").replace(".", ",").replace("_", ".")


def _native_boxes(page) -> list[TextBox]:
    return [
        TextBox(str(word[4]), float(word[0]), float(word[1]), float(word[2]), float(word[3]))
        for word in page.get_text("words")
        if str(word[4]).strip()
    ]


def _run_rapidocr(image):
    try:
        from rapidocr import RapidOCR
    except ImportError as exc:
        raise RuntimeError(
            "OCR paketleri eksik. Sunucuda 'pip install -r requirements.txt' komutunu çalıştırın."
        ) from exc

    global _OCR_ENGINE
    try:
        with _OCR_LOCK:
            if _OCR_ENGINE is None:
                _OCR_ENGINE = RapidOCR()
            result = _OCR_ENGINE(image)
    except Exception as exc:
        detail = str(exc).splitlines()[0][:160]
        if "libgl" in detail.casefold():
            raise RuntimeError(
                "OCR başlatılamadı. Linux sunucuda libGL1 paketini kurup servisi yeniden başlatın."
            ) from exc
        raise RuntimeError(f"OCR çalıştırılamadı ({type(exc).__name__}: {detail}).") from exc
    return result


def _ocr_region_boxes(page, clip: fitz.Rect, scale: float) -> list[TextBox]:
    try:
        import numpy as np
    except ImportError as exc:
        raise RuntimeError(
            "OCR paketleri eksik. Sunucuda 'pip install -r requirements.txt' komutunu çalıştırın."
        ) from exc

    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale),
        clip=clip,
        colorspace=fitz.csRGB,
        alpha=False,
    )
    image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height, pixmap.width, pixmap.n
    )
    result = _run_rapidocr(image)
    offset_x = clip.x0 * scale
    offset_y = clip.y0 * scale

    boxes = []
    if result.boxes is not None:
        for points, text in zip(result.boxes, result.txts):
            xs = [float(point[0]) + offset_x for point in points]
            ys = [float(point[1]) + offset_y for point in points]
            boxes.append(TextBox(str(text), min(xs), min(ys), max(xs), max(ys)))
    return boxes


def _ocr_boxes(page, scale: float = 2.0) -> tuple[list[TextBox], float, float]:
    width = float(page.rect.width * scale)
    height = float(page.rect.height * scale)
    top = fitz.Rect(0, 0, page.rect.width, page.rect.height * 0.42)
    extra_payments = fitz.Rect(
        0,
        page.rect.height * 0.38,
        page.rect.width * 0.52,
        page.rect.height * 0.90,
    )
    boxes = _ocr_region_boxes(page, top, scale)
    boxes.extend(_ocr_region_boxes(page, extra_payments, scale))
    return boxes, width, height


def _group_lines(boxes: list[TextBox], page_height: float) -> list[list[TextBox]]:
    lines: list[list[TextBox]] = []
    box_heights = [box.y1 - box.y0 for box in boxes]
    typical_height = median(box_heights) if box_heights else 8.0
    tolerance = max(3.0, typical_height * 0.48)
    for box in sorted(boxes, key=lambda item: (item.center_y, item.x0)):
        best_line = None
        best_distance = None
        for line in lines:
            line_y = sum(item.center_y for item in line) / len(line)
            distance = abs(box.center_y - line_y)
            if distance <= tolerance and (best_distance is None or distance < best_distance):
                best_line = line
                best_distance = distance
        if best_line is None:
            lines.append([box])
        else:
            best_line.append(box)
    for line in lines:
        line.sort(key=lambda item: item.x0)
    return sorted(lines, key=lambda line: sum(item.center_y for item in line) / len(line))


def _line_text(line: list[TextBox]) -> str:
    return " ".join(item.text for item in line)


def _find_worker_x(lines: list[list[TextBox]], width: float) -> float:
    for line in lines:
        for item in line:
            normalized = _plain(item.text).replace("1", "i")
            if normalized in {"isci", "isci i"} or "isci" in normalized:
                if item.center_y < max(box.center_y for row in lines for box in row) * 0.55:
                    return item.center_x
    return width * 0.80


def _amount_on_named_line(
    lines: list[list[TextBox]], label: str, worker_x: float, width: float
) -> Decimal | None:
    label_plain = _plain(label)
    for line in lines:
        text_plain = _plain(_line_text(line))
        if label_plain not in text_plain:
            continue
        candidates = [
            (abs(item.center_x - worker_x), amount)
            for item in line
            if (amount := _parse_money(item.text)) is not None
            and width * 0.69 <= item.center_x <= width * 0.88
        ]
        if candidates:
            return min(candidates, key=lambda pair: pair[0])[1]
    return None


def _fallback_worker_amounts(
    lines: list[list[TextBox]], worker_x: float, width: float
) -> tuple[Decimal | None, Decimal | None]:
    header_y = None
    for line in lines:
        if any("isci" in _plain(item.text).replace("1", "i") for item in line):
            header_y = sum(item.center_y for item in line) / len(line)
            break
    if header_y is None:
        return None, None

    amounts: list[tuple[float, Decimal]] = []
    for line in lines:
        line_y = sum(item.center_y for item in line) / len(line)
        if line_y <= header_y:
            continue
        candidates = [
            (abs(item.center_x - worker_x), _parse_money(item.text))
            for item in line
            if width * 0.69 <= item.center_x <= width * 0.88
            and _parse_money(item.text) is not None
        ]
        if candidates:
            amounts.append((line_y, min(candidates, key=lambda pair: pair[0])[1]))
        if len(amounts) >= 5:
            break
    if len(amounts) >= 5:
        return amounts[3][1], amounts[4][1]
    return None, None


def _extract_gross_amount(
    lines: list[list[TextBox]], width: float, height: float
) -> tuple[Decimal, Decimal]:
    gross_header_x = width * 0.40
    gross_header_y = None
    for line in lines:
        line_text = _plain(_line_text(line))
        if "brut tutar" in line_text:
            gross_header_y = sum(item.center_y for item in line) / len(line)
            matching = [item for item in line if "brut" in _plain(item.text)]
            if matching:
                gross_header_x = matching[0].center_x
            break

    if gross_header_y is None:
        raise ValueError("Çalışmalar bölümündeki Brüt Tutar başlığı okunamadı.")

    base_gross = None
    for line in lines:
        line_y = sum(item.center_y for item in line) / len(line)
        if line_y <= gross_header_y or line_y >= height * 0.36:
            continue
        if "toplam" not in _plain(_line_text(line)):
            continue
        candidates = [
            (abs(item.center_x - gross_header_x), amount)
            for item in line
            if (amount := _parse_money(item.text)) is not None
            and width * 0.34 <= item.center_x <= width * 0.50
        ]
        if candidates:
            base_gross = min(candidates, key=lambda pair: pair[0])[1]
            break

    if base_gross is None:
        raise ValueError("Çalışmalar bölümündeki toplam brüt tutar okunamadı.")

    yk_fee = Decimal("0")
    for line in lines:
        normalized = _plain(_line_text(line))
        if not re.search(r"\byk\b.*\bucreti\b", normalized):
            continue
        candidates = [
            (item.center_x, amount)
            for item in line
            if (amount := _parse_money(item.text)) is not None
            and width * 0.28 <= item.center_x <= width * 0.50
        ]
        if not candidates:
            raise ValueError("YK Ücreti satırındaki tutar okunamadı.")
        yk_fee = max(candidates, key=lambda pair: pair[0])[1]
        break

    return base_gross + yk_fee, yk_fee


def extract_page_values(page) -> tuple[Decimal, Decimal, Decimal, Decimal]:
    native = _native_boxes(page)
    native_text = " ".join(item.text for item in native)
    if "gelir" in _plain(native_text) and "damga" in _plain(native_text):
        boxes = native
        width = float(page.rect.width)
        height = float(page.rect.height)
    else:
        boxes, width, height = _ocr_boxes(page)

    lines = _group_lines(boxes, height)
    gross, yk_fee = _extract_gross_amount(lines, width, height)
    worker_x = _find_worker_x(lines, width)
    income = _amount_on_named_line(lines, "gelir vergisi", worker_x, width)
    stamp = _amount_on_named_line(lines, "damga vergisi", worker_x, width)

    if income is None or stamp is None:
        fallback_income, fallback_stamp = _fallback_worker_amounts(lines, worker_x, width)
        income = income if income is not None else fallback_income
        stamp = stamp if stamp is not None else fallback_stamp

    missing = []
    if income is None:
        missing.append("gelir vergisi")
    if stamp is None:
        missing.append("damga vergisi")
    if missing:
        raise ValueError(f"İşçi sütunundaki {' ve '.join(missing)} okunamadı.")
    return gross, income, stamp, yk_fee


def analyze_tax_pdfs(pdf_files: list[tuple[str, str]]) -> dict:
    rows = []
    warnings = []
    income_total = Decimal("0")
    stamp_total = Decimal("0")
    gross_total = Decimal("0")

    for path, original_name in pdf_files:
        try:
            document = fitz.open(path)
        except Exception as exc:
            warnings.append(f"{original_name}: PDF açılamadı ({exc}).")
            continue

        try:
            for page_number, page in enumerate(document, start=1):
                try:
                    gross, income, stamp, yk_fee = extract_page_values(page)
                except RuntimeError:
                    raise
                except Exception as exc:
                    warnings.append(f"{original_name} - Sayfa {page_number}: {exc}")
                    continue
                income_total += income
                stamp_total += stamp
                gross_total += gross
                rows.append({
                    "filename": original_name,
                    "page": page_number,
                    "gross": gross,
                    "gross_text": format_tr_money(gross),
                    "yk_fee": yk_fee,
                    "income": income,
                    "income_text": format_tr_money(income),
                    "stamp": stamp,
                    "stamp_text": format_tr_money(stamp),
                })
        finally:
            document.close()

    return {
        "rows": rows,
        "warnings": warnings,
        "gross_total": gross_total,
        "gross_total_text": format_tr_money(gross_total),
        "income_total": income_total,
        "income_total_text": format_tr_money(income_total),
        "stamp_total": stamp_total,
        "stamp_total_text": format_tr_money(stamp_total),
    }


def create_tax_excel(result: dict, output_dir: str) -> str:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Vergi Toplamları"
    sheet.sheet_view.showGridLines = False

    dark_blue = "1F4E78"
    medium_blue = "D9EAF7"
    light_blue = "EAF3F8"
    white = "FFFFFF"
    border_color = "B8C7D1"
    thin_border = Border(
        left=Side(style="thin", color=border_color),
        right=Side(style="thin", color=border_color),
        top=Side(style="thin", color=border_color),
        bottom=Side(style="thin", color=border_color),
    )

    sheet.merge_cells("A1:E1")
    title = sheet["A1"]
    title.value = "Brüt Ücret, İşçi Gelir ve Damga Vergisi Toplamı"
    title.fill = PatternFill("solid", fgColor=dark_blue)
    title.font = Font(name="Arial", size=15, bold=True, color=white)
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 28

    header_row = 3
    headers = [
        "Dosya",
        "Sayfa",
        "Brüt Ücret (YK Dahil)",
        "Gelir Vergisi (İşçi)",
        "Damga Vergisi (İşçi)",
    ]
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(header_row, column, header)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.font = Font(name="Arial", bold=True, color=white)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    first_data_row = header_row + 1
    for row_number, row in enumerate(result["rows"], start=first_data_row):
        values = [
            row["filename"],
            row["page"],
            float(row["gross"]),
            float(row["income"]),
            float(row["stamp"]),
        ]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row_number, column, value)
            cell.font = Font(name="Arial", size=10)
            cell.border = thin_border
            cell.alignment = Alignment(
                horizontal="right" if column >= 2 else "left",
                vertical="center",
            )
            if row_number % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=light_blue)
        sheet.cell(row_number, 3).number_format = "#,##0.00"
        sheet.cell(row_number, 4).number_format = "#,##0.00"
        sheet.cell(row_number, 5).number_format = "#,##0.00"

    last_data_row = first_data_row + len(result["rows"]) - 1
    total_row = last_data_row + 1
    sheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    total_label = sheet.cell(total_row, 1, "GENEL TOPLAM")
    total_label.alignment = Alignment(horizontal="right", vertical="center")
    sheet.cell(total_row, 3, f"=SUM(C{first_data_row}:C{last_data_row})")
    sheet.cell(total_row, 4, f"=SUM(D{first_data_row}:D{last_data_row})")
    sheet.cell(total_row, 5, f"=SUM(E{first_data_row}:E{last_data_row})")
    for column in range(1, 6):
        cell = sheet.cell(total_row, column)
        cell.fill = PatternFill("solid", fgColor=medium_blue)
        cell.font = Font(name="Arial", size=11, bold=True, color=dark_blue)
        cell.border = thin_border
    sheet.cell(total_row, 3).number_format = "#,##0.00"
    sheet.cell(total_row, 4).number_format = "#,##0.00"
    sheet.cell(total_row, 5).number_format = "#,##0.00"

    sheet.auto_filter.ref = f"A{header_row}:E{last_data_row}"
    sheet.freeze_panes = f"A{first_data_row}"
    widths = [42, 10, 24, 24, 24]
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    sheet.print_title_rows = f"1:{header_row}"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    if result["warnings"]:
        warning_sheet = workbook.create_sheet("Uyarılar")
        warning_sheet.sheet_view.showGridLines = False
        warning_sheet["A1"] = "Kontrol Gereken Sayfalar"
        warning_sheet["A1"].fill = PatternFill("solid", fgColor="9C6500")
        warning_sheet["A1"].font = Font(name="Arial", size=13, bold=True, color=white)
        warning_sheet["A1"].alignment = Alignment(vertical="center")
        warning_sheet.row_dimensions[1].height = 25
        for row_number, warning in enumerate(result["warnings"], start=3):
            cell = warning_sheet.cell(row_number, 1, warning)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        warning_sheet.column_dimensions["A"].width = 110
        warning_sheet.freeze_panes = "A3"

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(
        output_dir,
        f"isci_vergi_toplami_{uuid.uuid4().hex[:8]}.xlsx",
    )
    workbook.save(output_path)
    return output_path
