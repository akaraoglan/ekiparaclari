import os
import math
import re
import threading
import unicodedata
import uuid
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from statistics import median

import fitz
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


_OCR_ENGINE = None
_OCR_LOCK = threading.Lock()
_MONEY_RE = re.compile(r"(?<!\d)[0-9O]+(?:[.,'/][0-9O]+)+(?!\d)", re.IGNORECASE)
_STANDARD_TR_MONEY_RE = re.compile(
    r"[0-9O]{1,3}(?:\.[0-9O]{3})*,[0-9O]{2}|[0-9O]+,[0-9O]{2}",
    re.IGNORECASE,
)
_EXCEL_COLUMNS = [
    ("gross", "Brüt Ücret (YK Dahil)"),
    ("income", "Gelir Vergisi (İşçi)"),
    ("stamp", "Damga Vergisi (İşçi)"),
    ("overtime_total", "Fazla Mesailer Toplamı"),
    ("bonus", "İkramiye"),
    ("other_extra_payments", "Diğer Ek Ödeme ve Yardımlar"),
    ("worker_sgk", "İşçi SGK"),
    ("worker_unemployment", "İşçi İşsizlik"),
    ("employer_sgk", "İşveren SGK"),
    ("employer_unemployment", "İşveren İşsizlik"),
    ("income_tax_incentive", "Gelir Vergisi Teşviki"),
    ("stamp_tax_incentive", "Damga Vergisi Teşviki"),
    ("advance", "Avans"),
    ("other_advances", "Avans 1-2-3"),
    ("account_369_deductions", "369 Kesintiler"),
    ("net_paid", "Net Ödenen"),
    ("incentives_total", "Teşvikler Toplamı"),
    ("employer_cost", "İşveren Maliyeti (Teşvikli)"),
]


class OCRUnavailableError(RuntimeError):
    """OCR dependencies or models could not be initialized."""


@dataclass(frozen=True)
class TextBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    confidence: float = 1.0

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


def _text_matches(text: str, label: str) -> bool:
    normalized = _plain(text)
    wanted = _plain(label)
    if wanted in normalized:
        return True
    if wanted.replace(" ", "") in normalized.replace(" ", ""):
        return True

    words = normalized.split()
    wanted_words = wanted.split()
    if not words or not wanted_words:
        return False
    window_size = len(wanted_words)
    for start in range(max(1, len(words) - window_size + 1)):
        candidate = " ".join(words[start:start + window_size])
        if SequenceMatcher(None, wanted, candidate).ratio() >= 0.90:
            return True
    return False


def _is_total_label(text: str) -> bool:
    # A scan cropped at the left edge may lose the first two letters.
    # These are structural total labels, never extra-payment categories.
    return _text_matches(text, "toplam") or _plain(text) == "plam"


def _parse_money(text: str) -> Decimal | None:
    compact = text.replace(" ", "")
    for match in _MONEY_RE.finditer(compact):
        candidate = match.group(0).upper().replace("O", "0")
        if not set(candidate) - {"0", ".", ",", "'"}:
            return Decimal("0")
        separator_index = max(
            candidate.rfind("."),
            candidate.rfind(","),
            candidate.rfind("'"),
            candidate.rfind("/"),
        )
        decimals = candidate[separator_index + 1:]
        if len(decimals) != 2:
            continue
        integer_part = re.sub(r"[.,'/]", "", candidate[:separator_index])
        try:
            return Decimal(f"{integer_part}.{decimals}")
        except InvalidOperation:
            continue
    return None


def _parse_upside_down_money(text: str) -> Decimal | None:
    compact = re.sub(r"[^0-9O.,'/]", "", text.upper().replace("O", "0"))
    if not compact:
        return None
    source_variants = [compact[1:], compact] if compact.startswith("1") else [compact]
    for source in source_variants:
        rotated = source[::-1].translate(str.maketrans({"6": "9", "9": "6"}))
        parsed = _parse_money(rotated)
        if parsed is not None:
            return parsed
        digits = re.sub(r"\D", "", rotated)
        if len(digits) >= 3:
            try:
                return Decimal(f"{digits[:-2]}.{digits[-2:]}")
            except InvalidOperation:
                continue
    return None


def format_tr_money(value: Decimal) -> str:
    rendered = f"{value:,.2f}"
    return rendered.replace(",", "_").replace(".", ",").replace("_", ".")


def _native_boxes(page) -> list[TextBox]:
    boxes = []
    for word in page.get_text("words"):
        if not str(word[4]).strip():
            continue
        # Native PDF coordinates ignore page rotation; rendered OCR does not.
        rect = fitz.Rect(word[:4]) * page.rotation_matrix
        boxes.append(TextBox(str(word[4]), rect.x0, rect.y0, rect.x1, rect.y1))
    return boxes


def _run_rapidocr(image, recognition_only: bool = False):
    try:
        from rapidocr import RapidOCR
    except ImportError as exc:
        raise OCRUnavailableError(
            "OCR paketleri eksik. Sunucuda 'pip install -r requirements.txt' komutunu çalıştırın."
        ) from exc

    global _OCR_ENGINE
    with _OCR_LOCK:
        if _OCR_ENGINE is None:
            try:
                _OCR_ENGINE = RapidOCR()
            except Exception as exc:
                detail = str(exc).splitlines()[0][:160]
                if "libgl" in detail.casefold():
                    raise OCRUnavailableError(
                        "OCR başlatılamadı. Linux sunucuda libGL1 paketini kurup servisi yeniden başlatın."
                    ) from exc
                raise OCRUnavailableError(
                    f"OCR başlatılamadı ({type(exc).__name__}: {detail})."
                ) from exc
        try:
            return _OCR_ENGINE(
                image, use_det=not recognition_only, use_cls=not recognition_only,
            )
        except Exception as exc:
            detail = str(exc).splitlines()[0][:160]
            # A page-specific inference failure should allow other reading
            # strategies and subsequent pages, unlike missing dependencies.
            raise ValueError(
                f"OCR görüntüyü okuyamadı ({type(exc).__name__}: {detail})."
            ) from exc


def _ocr_region_boxes(
    page, clip: fitz.Rect, scale: float, remove_colored_ink: bool = False,
    recognition_only: bool = False,
) -> list[TextBox]:
    try:
        import numpy as np
    except ImportError as exc:
        raise OCRUnavailableError(
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
    if remove_colored_ink:
        # Keep dark printed strokes and suppress colored signatures/stamps.
        # This filters existing pixels; it never reconstructs missing digits.
        image = np.repeat(image.max(axis=2, keepdims=True), 3, axis=2)
        # A border keeps the detector from clipping a tightly cropped digit.
        image = np.pad(image, ((10, 10), (10, 10), (0, 0)), constant_values=255)
    result = _run_rapidocr(image, recognition_only=recognition_only)
    if recognition_only:
        return [
            TextBox(text, clip.x0 * scale, clip.y0 * scale,
                    clip.x1 * scale, clip.y1 * scale, float(score))
            for text, score in zip(
                result.txts if result.txts is not None else (),
                result.scores if result.scores is not None else (),
            )
        ]
    border = 10 if remove_colored_ink else 0
    offset_x = clip.x0 * scale - border
    offset_y = clip.y0 * scale - border

    boxes = []
    if result.boxes is not None:
        result_scores = getattr(result, "scores", None)
        scores = (
            result_scores
            if result_scores is not None
            else [0.0] * len(result.txts)
        )
        for points, text, score in zip(result.boxes, result.txts, scores):
            xs = [float(point[0]) + offset_x for point in points]
            ys = [float(point[1]) + offset_y for point in points]
            boxes.append(
                TextBox(
                    str(text),
                    min(xs),
                    min(ys),
                    max(xs),
                    max(ys),
                    float(score),
                )
            )
    return boxes


def _ocr_boxes(page, scale: float = 2.0) -> tuple[list[TextBox], float, float]:
    width = float(page.rect.width * scale)
    height = float(page.rect.height * scale)
    boxes = _ocr_region_boxes(page, page.rect, scale)
    return boxes, width, height


def _ocr_banded_boxes(page, scale: float = 3.0) -> tuple[list[TextBox], float, float]:
    """Read overlapping bands so full-page downscaling cannot hide small text."""
    width, height = float(page.rect.width), float(page.rect.height)
    boxes = []
    band_count = 4
    padding = 24.0
    for index in range(band_count):
        start = height * index / band_count
        end = height * (index + 1) / band_count
        clip = fitz.Rect(0, max(0, start - padding), width, min(height, end + padding))
        region_boxes = _ocr_region_boxes(page, clip, scale)
        # Overlap protects rows at crop edges. Ownership by center prevents
        # counting the same payment twice when both bands read it.
        boxes.extend(
            box for box in region_boxes
            if start * scale <= box.center_y < end * scale
        )
    return boxes, width * scale, height * scale


def _deskew_page_image(page):
    """Correct a measured table slope on a copy, without clipping scan edges."""
    try:
        import cv2
        import numpy as np
    except ImportError as exc:
        raise OCRUnavailableError("OCR görüntü işleme paketleri eksik.") from exc
    scale = 2.0
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False)
    image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, 3)
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    lines = cv2.HoughLinesP(
        cv2.Canny(gray, 50, 150), 1, np.pi / 1800,
        threshold=100, minLineLength=pixmap.width * 0.30, maxLineGap=30,
    )
    if lines is None:
        return None
    angles = []
    for x0, y0, x1, y1 in lines.reshape(-1, 4):
        angle = math.degrees(math.atan2(float(y1 - y0), float(x1 - x0)))
        if abs(angle) <= 5 and pixmap.height * 0.02 < (y0 + y1) / 2 < pixmap.height * 0.98:
            angles.append(angle)
    if len(angles) < 4:
        return None
    angle = median(angles)
    agreeing = [value for value in angles if abs(value - angle) <= 0.5]
    if len(agreeing) < max(4, len(angles) * 0.65) or abs(angle) < 0.25:
        return None
    angle = median(agreeing)
    matrix = cv2.getRotationMatrix2D((pixmap.width / 2, pixmap.height / 2), angle, 1)
    width = math.ceil(abs(matrix[0, 0]) * pixmap.width + abs(matrix[0, 1]) * pixmap.height)
    height = math.ceil(abs(matrix[0, 1]) * pixmap.width + abs(matrix[0, 0]) * pixmap.height)
    matrix[0, 2] += (width - pixmap.width) / 2
    matrix[1, 2] += (height - pixmap.height) / 2
    corrected = cv2.warpAffine(
        image, matrix, (width, height), borderMode=cv2.BORDER_CONSTANT,
        borderValue=(255, 255, 255),
    )
    encoded, png = cv2.imencode(".png", cv2.cvtColor(corrected, cv2.COLOR_RGB2BGR))
    if not encoded:
        raise ValueError("Eğimi düzeltilmiş görüntü oluşturulamadı.")
    return png.tobytes(), width / scale, height / scale, angle


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


def _line_y(line: list[TextBox]) -> float:
    return sum(item.center_y for item in line) / len(line)


def _regional_text(
    line: list[TextBox], width: float, x_min: float, x_max: float
) -> str:
    return " ".join(
        item.text for item in line
        if width * x_min <= item.center_x <= width * x_max
    )


def _regional_amounts(
    line: list[TextBox], width: float, x_min: float, x_max: float
) -> list[tuple[float, Decimal]]:
    return [
        (item.center_x, amount)
        for item in line
        if width * x_min <= item.center_x <= width * x_max
        and (amount := _parse_money(item.text)) is not None
    ]


def _money_text_needs_reread(text: str, amount: Decimal) -> str | None:
    if amount == 0:
        return None
    compact = text.replace(" ", "").upper().replace("O", "0")
    if re.fullmatch(r"0\.\d{3},\d{2}", compact):
        return "missing_leading_digit"
    if _STANDARD_TR_MONEY_RE.fullmatch(compact):
        return None
    return "noisy"


def _reread_money_with_consensus(
    page,
    item: TextBox,
    amount: Decimal,
    base_scale: float,
) -> Decimal:
    reread_kind = _money_text_needs_reread(item.text, amount)
    if reread_kind is None:
        return amount

    if reread_kind == "missing_leading_digit":
        variants = ((1.0, 2.8), (1.0, 3.2))
    else:
        variants = ((5.0, 3.0), (9.0, 5.0))

    original_y0 = item.y0 / base_scale
    original_y1 = item.y1 / base_scale
    original_center_y = item.center_y / base_scale
    readings: list[Decimal] = []
    for padding, scale in variants:
        clip = fitz.Rect(
            page.rect.width * 0.68,
            max(0, original_y0 - padding),
            page.rect.width,
            min(page.rect.height, original_y1 + padding),
        )
        reread_boxes = _ocr_region_boxes(page, clip, scale)
        candidates = []
        for reread_box in reread_boxes:
            reread_amount = _parse_money(reread_box.text)
            if reread_amount is None:
                continue
            distance = abs(reread_box.center_y / scale - original_center_y)
            candidates.append((distance, reread_amount))
        if candidates:
            readings.append(min(candidates, key=lambda pair: pair[0])[1])

    if len(readings) == len(variants) and len(set(readings)) == 1:
        return readings[0]
    raise ValueError(
        f"Şüpheli tutar ({item.text}) iki görüntü okumasında doğrulanamadı."
    )


def _find_section_y(
    lines: list[list[TextBox]],
    label: str,
    width: float,
    x_min: float,
    x_max: float,
    after_y: float = 0,
) -> float | None:
    for line in lines:
        if _line_y(line) <= after_y:
            continue
        if _text_matches(_regional_text(line, width, x_min, x_max), label):
            return _line_y(line)
    return None


def _named_amount(
    lines: list[list[TextBox]],
    label: str,
    width: float,
    label_x: tuple[float, float],
    amount_x: tuple[float, float],
    y_min: float = 0,
    y_max: float = float("inf"),
    allow_next_line: bool = False,
) -> Decimal | None:
    for line in lines:
        line_y = _line_y(line)
        if not y_min < line_y < y_max:
            continue
        line_label = _regional_text(line, width, *label_x)
        if not _text_matches(line_label, label):
            continue
        amounts = _regional_amounts(line, width, *amount_x)
        if amounts:
            return max(amounts, key=lambda pair: pair[0])[1]
        if allow_next_line:
            box_heights = [
                item.y1 - item.y0
                for candidate_line in lines
                for item in candidate_line
            ]
            max_distance = max(
                8.0,
                (median(box_heights) if box_heights else 8.0) * 2.8,
            )
            nearby_lines = sorted(
                (
                    candidate_line
                    for candidate_line in lines
                    if 0 < _line_y(candidate_line) - line_y <= max_distance
                    and _line_y(candidate_line) < y_max
                ),
                key=_line_y,
            )
            for nearby_line in nearby_lines:
                nearby_amounts = _regional_amounts(
                    nearby_line, width, *amount_x
                )
                if nearby_amounts:
                    return max(nearby_amounts, key=lambda pair: pair[0])[1]
    return None


def _reread_named_amount(
    page, lines, label, width, label_x, amount_x, y_min, y_max, base_scale,
) -> Decimal | None:
    """Recover an unreadable printed amount only when two crops agree."""
    if page is None:
        return None
    label_rows = [
        line for line in lines
        if y_min < _line_y(line) < y_max
        and _text_matches(_regional_text(line, width, *label_x), label)
    ]
    if len(label_rows) != 1:
        return None
    line = label_rows[0]
    label_boxes = [item for item in line if width * label_x[0] <= item.center_x <= width * label_x[1]]
    center_y = median(item.center_y for item in label_boxes) / base_scale
    typical_height = median(item.y1 - item.y0 for item in label_boxes) / base_scale
    padding = max(4.0, typical_height * 0.75)
    neighboring_distances = [
        abs(_line_y(row) / base_scale - center_y)
        for row in lines if row is not line
        and _regional_text(row, width, *label_x).strip()
    ]
    if neighboring_distances:
        padding = min(padding, max(4.0, min(neighboring_distances) / 2 - 1.0))
    numeric_boxes = [
        item for item in line
        if width * amount_x[0] <= item.center_x <= width * amount_x[1]
        and any(char.isdigit() for char in item.text)
    ]
    if len(numeric_boxes) > 1:
        return None
    x0, x1 = page.rect.width * amount_x[0], page.rect.width * amount_x[1]
    if numeric_boxes:
        x0 = max(x0, numeric_boxes[0].x0 / base_scale - 2.0)
        x1 = min(x1, numeric_boxes[0].x1 / base_scale + 2.0)
    clip = fitz.Rect(x0, max(0, center_y - padding), x1, min(page.rect.height, center_y + padding))
    readings = []
    for scale in (3.2, 4.0):
        boxes = _ocr_region_boxes(
            page, clip, scale, remove_colored_ink=True, recognition_only=True,
        )
        amounts = {
            value for box in boxes
            if box.confidence >= 0.85
            and _STANDARD_TR_MONEY_RE.fullmatch(box.text.replace(" ", ""))
            and (value := _parse_money(box.text)) is not None
        }
        if len(amounts) != 1:
            return None
        readings.append(amounts.pop())
    return readings[0] if readings[0] == readings[1] else None


def _section_total(
    lines: list[list[TextBox]],
    width: float,
    y_min: float,
    y_max: float,
    label_x: tuple[float, float],
    amount_x: tuple[float, float],
    target_x: float | None = None,
) -> Decimal | None:
    for line in lines:
        line_y = _line_y(line)
        if not y_min < line_y < y_max:
            continue
        if not _is_total_label(_regional_text(line, width, *label_x)):
            continue
        amounts = _regional_amounts(line, width, *amount_x)
        if not amounts:
            continue
        if target_x is None:
            return max(amounts, key=lambda pair: pair[0])[1]
        return min(amounts, key=lambda pair: abs(pair[0] - target_x))[1]
    return None


def _canonical_extra_payment_label(
    raw_label: str,
    confidence: float = 1.0,
) -> tuple[str, str, bool, str | None, float]:
    """Keep the observed heading; never infer an extra payment category."""
    display = re.sub(r"\s+", " ", raw_label).strip()
    # Only spacing and letter case are equivalent. Accents, punctuation and
    # qualifiers may distinguish actual payment types and must remain intact.
    identity = unicodedata.normalize("NFC", display).casefold().replace("i\u0307", "i")
    key = f"ocr_{identity}" if identity else "ocr_belirsiz"
    needs_review = "�" in display or confidence < 0.85 or len(_plain(display)) < 3
    reason = (
        "Başlık eksik veya düşük güvenle okundu; ödeme türü tahmin edilmedi."
        if needs_review else None
    )
    return key, display or "Başlık okunamadı", needs_review, reason, confidence


def _find_extra_payment_total(
    lines: list[list[TextBox]],
    width: float,
    y_min: float,
    y_max: float,
    target_x: float,
) -> tuple[float, Decimal] | None:
    for line in lines:
        line_y = _line_y(line)
        if not y_min < line_y < y_max:
            continue
        label = _regional_text(line, width, 0.00, 0.20)
        if not _is_total_label(label):
            continue
        amounts = _regional_amounts(line, width, 0.28, 0.50)
        if amounts:
            amount = min(amounts, key=lambda pair: abs(pair[0] - target_x))[1]
            return line_y, amount
    return None


def _reread_extra_payment_label(
    page,
    line: list[TextBox],
    base_scale: float,
) -> tuple[str, float] | None:
    if not line:
        return None
    original_y0 = min(item.y0 for item in line) / base_scale
    original_y1 = max(item.y1 for item in line) / base_scale
    original_center_y = _line_y(line) / base_scale
    scale = 3.2
    clip = fitz.Rect(
        0,
        max(0, original_y0 - 3.0),
        page.rect.width * 0.31,
        min(page.rect.height, original_y1 + 3.0),
    )
    reread_boxes = _ocr_region_boxes(page, clip, scale)
    reread_lines = _group_lines(reread_boxes, page.rect.height * scale)
    if not reread_lines:
        return None
    reread_line = min(
        reread_lines,
        key=lambda row: abs(_line_y(row) / scale - original_center_y),
    )
    text = _line_text(reread_line).strip()
    if not text:
        return None
    confidence = median(item.confidence for item in reread_line)
    return text, confidence


def _extract_extra_payment_items(
    lines: list[list[TextBox]],
    width: float,
    extra_y: float,
    total_y: float,
    page=None,
    ocr_scale: float = 1.0,
) -> tuple[dict[str, Decimal], dict[str, dict]]:
    payments: dict[str, Decimal] = {}
    metadata: dict[str, dict] = {}

    for line in lines:
        line_y = _line_y(line)
        if not extra_y < line_y < total_y:
            continue
        amounts = _regional_amounts(line, width, 0.28, 0.50)
        if not amounts:
            continue
        raw_label = _regional_text(line, width, 0.00, 0.30).strip()
        if _is_total_label(raw_label):
            continue
        label_items = [
            item for item in line
            if width * 0.00 <= item.center_x <= width * 0.30
        ]
        confidence = (
            median(item.confidence for item in label_items)
            if label_items
            else 0.0
        )
        if not raw_label:
            confidence = 0.0
            raw_label = f"Başlık okunamadı (satır {len(payments) + 1})"

        key, header, estimated, reason, score = _canonical_extra_payment_label(
            raw_label, confidence
        )
        if estimated and page is not None:
            reread = _reread_extra_payment_label(page, line, ocr_scale)
            if reread is not None:
                reread_text, reread_confidence = reread
                reread_result = _canonical_extra_payment_label(
                    reread_text, reread_confidence
                )
                reread_key, _, reread_estimated, _, reread_score = reread_result
                if (
                    (reread_key == key and (
                        not reread_estimated or reread_score > score + 0.05
                    ))
                    or (estimated and not reread_estimated)
                ):
                    key, header, estimated, reason, score = reread_result
                    raw_label = reread_text

        amount = max(amounts, key=lambda pair: pair[0])[1]
        payments[key] = payments.get(key, Decimal("0")) + amount
        current = metadata.get(key)
        if current is None:
            metadata[key] = {
                "header": header,
                "estimated": estimated,
                "raw_label": raw_label,
                "reason": reason,
            }
        else:
            current["estimated"] = current["estimated"] or estimated
            if estimated:
                current["raw_label"] = raw_label
                current["reason"] = reason or current.get("reason")

    return payments, metadata


def _reread_extra_payment_section(
    page,
    extra_y: float,
    total_y: float,
    base_scale: float,
) -> tuple[dict[str, Decimal], dict[str, dict]] | None:
    scale = 3.2
    clip = fitz.Rect(
        0,
        max(0, extra_y / base_scale - 5.0),
        page.rect.width * 0.52,
        min(page.rect.height, total_y / base_scale + 5.0),
    )
    boxes = _ocr_region_boxes(page, clip, scale, remove_colored_ink=True)
    width = float(page.rect.width * scale)
    height = float(page.rect.height * scale)
    lines = _group_lines(boxes, height)
    reread_extra_y = _find_section_y(
        lines, "ek odemeler", width, 0.00, 0.48
    )
    if reread_extra_y is None:
        return None
    reread_total = _find_extra_payment_total(
        lines,
        width,
        reread_extra_y,
        height * 0.95,
        width * 0.40,
    )
    if reread_total is None:
        return None
    reread_total_y, _ = reread_total
    return _extract_extra_payment_items(
        lines,
        width,
        reread_extra_y,
        reread_total_y,
        page=None,
        ocr_scale=scale,
    )


def _find_worker_x(lines: list[list[TextBox]], width: float) -> float:
    for line in lines:
        for item in line:
            normalized = _plain(item.text).replace("1", "i")
            if normalized in {"isci", "isci i"} or "isci" in normalized:
                if item.center_y < max(box.center_y for row in lines for box in row) * 0.55:
                    return item.center_x
    return width * 0.80


def _find_employer_x(lines: list[list[TextBox]], width: float) -> float:
    page_bottom = max(box.center_y for row in lines for box in row)
    for line in lines:
        for item in line:
            if (
                item.center_y < page_bottom * 0.25
                and "isveren" in _plain(item.text).replace("1", "i")
            ):
                return item.center_x
    return width * 0.92


def _deduction_row_amount(
    lines: list[list[TextBox]],
    label: str,
    target_x: float,
    width: float,
    height: float,
) -> Decimal | None:
    wanted = _plain(label)
    for line in lines:
        if _line_y(line) >= height * 0.30:
            break
        label_items = [
            item for item in line
            if width * 0.43 <= item.center_x <= width * 0.61
        ]
        label_text = _plain(" ".join(item.text for item in label_items))
        if " " in wanted:
            label_matches = label_text.endswith(wanted)
        else:
            label_matches = any(_plain(item.text) == wanted for item in label_items)
        if not label_matches:
            continue
        amounts = _regional_amounts(line, width, 0.68, 0.99)
        if amounts:
            return min(amounts, key=lambda pair: abs(pair[0] - target_x))[1]
    return None


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
        if _line_y(line) < height * 0.25 and "brut tutar" in line_text:
            gross_header_y = sum(item.center_y for item in line) / len(line)
            matching = [item for item in line if "brut" in _plain(item.text)]
            if matching:
                gross_header_x = matching[0].center_x
            break

    if gross_header_y is None:
        gross_header_y = _find_section_y(
            lines, "calismalar", width, 0.00, 0.52
        )
    if gross_header_y is None:
        raise ValueError("Çalışmalar bölümünün başlığı okunamadı.")

    base_gross = None
    for line in lines:
        line_y = sum(item.center_y for item in line) / len(line)
        if line_y <= gross_header_y or line_y >= height * 0.36:
            continue
        if not _is_total_label(_regional_text(line, width, 0.00, 0.20)):
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


def _extract_additional_values(
    lines: list[list[TextBox]],
    width: float,
    height: float,
    yk_fee: Decimal = Decimal("0"),
    page=None,
    ocr_scale: float = 1.0,
) -> dict:
    gross_header_x = width * 0.40
    for line in lines:
        if (
            _line_y(line) < height * 0.25
            and "brut tutar" in _plain(_regional_text(line, width, 0.28, 0.48))
        ):
            matching = [item for item in line if "brut" in _plain(item.text)]
            if matching:
                gross_header_x = matching[0].center_x
            break

    overtime_y = _find_section_y(lines, "fazla mesailer", width, 0.00, 0.48)
    extra_y = _find_section_y(lines, "ek odemeler", width, 0.00, 0.48)
    gross_payments_y = _find_section_y(lines, "brut odemeler", width, 0.00, 0.48)
    if overtime_y is None or extra_y is None:
        raise ValueError("Fazla Mesailer veya Ek Ödemeler bölüm sınırları okunamadı.")

    extra_search_end = gross_payments_y or height * 0.95
    extra_total_info = _find_extra_payment_total(
        lines,
        width,
        extra_y,
        extra_search_end,
        gross_header_x,
    )
    if extra_total_info is None:
        raise ValueError("Ek Ödemeler toplamı okunamadı.")
    extra_total_y, extra_total = extra_total_info

    overtime_total = _section_total(
        lines, width, overtime_y, extra_y, (0.00, 0.18), (0.28, 0.50), gross_header_x
    )
    bonus = _named_amount(
        lines,
        "ikramiye",
        width,
        (0.00, 0.28),
        (0.28, 0.50),
        extra_y,
        extra_total_y,
    )
    if overtime_total is None or extra_total is None:
        raise ValueError("Fazla Mesailer veya Ek Ödemeler toplamı okunamadı.")
    bonus = bonus if bonus is not None else Decimal("0")

    extra_payments, extra_payment_meta = _extract_extra_payment_items(
        lines,
        width,
        extra_y,
        extra_total_y,
        page=page,
        ocr_scale=ocr_scale,
    )
    item_total = sum(extra_payments.values(), Decimal("0"))
    extra_payment_warning = None
    if (
        abs(item_total - extra_total) > Decimal("0.01")
        and page is not None
    ):
        reread = _reread_extra_payment_section(
            page,
            extra_y,
            extra_total_y,
            ocr_scale,
        )
        if reread is not None:
            reread_payments, reread_metadata = reread
            reread_total = sum(reread_payments.values(), Decimal("0"))
            if abs(reread_total - extra_total) <= Decimal("0.01"):
                extra_payments = reread_payments
                extra_payment_meta = reread_metadata
                item_total = reread_total
    if abs(item_total - extra_total) > Decimal("0.01"):
        extra_payment_warning = (
            "Ek ödeme kalemleri toplamı "
            f"{format_tr_money(item_total)} TL, bordrodaki Ek Ödemeler toplamı "
            f"{format_tr_money(extra_total)} TL. Yeni detay sütunları kontrol edilmelidir."
        )

    worker_x = _find_worker_x(lines, width)
    employer_x = _find_employer_x(lines, width)
    worker_sgk = _deduction_row_amount(lines, "sgk", worker_x, width, height)
    worker_sgdp = _deduction_row_amount(lines, "sgdp", worker_x, width, height)
    worker_unemployment = _deduction_row_amount(
        lines, "issizlik", worker_x, width, height
    )
    employer_unemployment = _deduction_row_amount(
        lines, "issizlik", employer_x, width, height
    )
    employer_extra_unemployment = _deduction_row_amount(
        lines, "ek issizlik", employer_x, width, height
    )
    contribution_values = {
        "İşçi SGK": worker_sgk,
        "İşçi SGDP": worker_sgdp,
        "İşçi İşsizlik": worker_unemployment,
        "İşveren İşsizlik": employer_unemployment,
        "İşveren Ek İşsizlik": employer_extra_unemployment,
    }
    missing = [label for label, value in contribution_values.items() if value is None]
    if missing:
        raise ValueError(f"Okunamayan tutarlar: {', '.join(missing)}.")

    employer_legal_total = _named_amount(
        lines,
        "yasal kesintiler toplami",
        width,
        (0.43, 0.78),
        (0.84, 0.99),
        0,
        height * 0.35,
    )
    if employer_legal_total is None:
        raise ValueError("İşveren yasal kesintiler toplamı okunamadı.")

    special_y = _find_section_y(lines, "ozel kesintiler", width, 0.43, 0.99)
    tax_discount_y = _find_section_y(lines, "vergi indirimi", width, 0.43, 0.99)
    incentives_y = _find_section_y(lines, "tesvikler", width, 0.43, 0.99)
    employer_cost_y = _find_section_y(
        lines, "isveren maliyeti", width, 0.43, 0.99
    )
    if None in (special_y, tax_discount_y, incentives_y, employer_cost_y):
        raise ValueError("Özel Kesintiler, Vergi İndirimi veya Teşvikler bölümü okunamadı.")

    advance = Decimal("0")
    other_advances = Decimal("0")
    for line in lines:
        line_y = _line_y(line)
        if not special_y < line_y < tax_discount_y:
            continue
        label_text = _plain(_regional_text(line, width, 0.43, 0.70))
        if label_text != "avans" and not label_text.startswith("avans "):
            continue
        amounts = _regional_amounts(line, width, 0.70, 0.99)
        if not amounts:
            raise ValueError(f"{label_text.title()} satırındaki tutar okunamadı.")
        amount = max(amounts, key=lambda pair: pair[0])[1]
        if label_text == "avans":
            advance += amount
        else:
            other_advances += amount

    special_total = _section_total(
        lines, width, special_y, tax_discount_y, (0.43, 0.67), (0.72, 0.99)
    )
    if special_total is None:
        bottom_special_total = _named_amount(
            lines,
            "ozel kesintiler toplami",
            width,
            (0.43, 0.78),
            (0.78, 0.99),
            employer_cost_y,
            height,
        )
        if bottom_special_total is not None:
            special_total = bottom_special_total + other_advances
    if special_total is None:
        raise ValueError("Özel Kesintiler bölüm toplamı okunamadı.")

    income_tax_incentive = _named_amount(
        lines,
        "gelir vergisi indirimi",
        width,
        (0.43, 0.78),
        (0.78, 0.99),
        tax_discount_y,
        incentives_y,
    )
    stamp_tax_incentive = _named_amount(
        lines,
        "damga vergisi indirimi",
        width,
        (0.43, 0.78),
        (0.78, 0.99),
        tax_discount_y,
        incentives_y,
    )
    net_paid = _named_amount(
        lines,
        "net odenen",
        width,
        (0.43, 0.76),
        (0.76, 0.99),
        employer_cost_y,
        height,
    )
    employer_cost = _named_amount(
        lines,
        "isveren maliyeti",
        width,
        (0.43, 0.78),
        (0.78, 0.99),
        incentives_y,
        height,
        allow_next_line=True,
    )
    if net_paid is None:
        net_paid = _reread_named_amount(
            page, lines, "net odenen", width, (0.43, 0.76), (0.76, 0.99),
            employer_cost_y, height, ocr_scale,
        )
    required_amounts = {
        "Gelir Vergisi İndirimi": income_tax_incentive,
        "Damga Vergisi İndirimi": stamp_tax_incentive,
        "Net Ödenen": net_paid,
        "İşveren Maliyeti": employer_cost,
    }
    missing = [label for label, value in required_amounts.items() if value is None]
    if missing:
        raise ValueError(f"Okunamayan tutarlar: {', '.join(missing)}.")

    automatic_bes = _named_amount(
        lines,
        "otomatik",
        width,
        (0.43, 0.78),
        (0.78, 0.99),
        employer_cost_y,
        height,
    )
    if automatic_bes is None:
        for line in lines:
            if not employer_cost_y < _line_y(line) < height:
                continue
            if not _text_matches(
                _regional_text(line, width, 0.43, 0.78), "otomatik"
            ):
                continue
            rotated_candidates = [
                value
                for item in line
                if width * 0.78 <= item.center_x <= width * 0.99
                and (value := _parse_upside_down_money(item.text)) is not None
            ]
            if rotated_candidates:
                automatic_bes = rotated_candidates[-1]
                break
    if automatic_bes is None:
        raise ValueError("Otomatik BES kesintisi okunamadı.")

    incentives_total = Decimal("0")
    for line in lines:
        if not incentives_y < _line_y(line) < employer_cost_y:
            continue
        for item in line:
            if not width * 0.78 <= item.center_x <= width * 0.99:
                continue
            amount = _parse_money(item.text)
            if amount is None:
                continue
            if page is not None:
                amount = _reread_money_with_consensus(
                    page, item, amount, ocr_scale
                )
            incentives_total += amount

    return {
        "overtime_total": overtime_total,
        "bonus": bonus,
        "other_extra_payments": extra_total - bonus - yk_fee,
        "extra_payments": extra_payments,
        "extra_payment_meta": extra_payment_meta,
        "extra_payment_warning": extra_payment_warning,
        "worker_sgk": worker_sgk + worker_sgdp,
        "worker_unemployment": worker_unemployment,
        "employer_sgk": (
            employer_legal_total
            - employer_unemployment
            - employer_extra_unemployment
        ),
        "employer_unemployment": employer_unemployment + employer_extra_unemployment,
        "income_tax_incentive": income_tax_incentive,
        "stamp_tax_incentive": stamp_tax_incentive,
        "advance": advance,
        "other_advances": other_advances,
        "account_369_deductions": (
            special_total - advance - other_advances + automatic_bes
        ),
        "net_paid": net_paid,
        "incentives_total": incentives_total,
        "employer_cost": employer_cost,
    }


def _extract_personnel_subarea(
    lines: list[list[TextBox]], width: float, height: float,
) -> str | None:
    for line in lines:
        if _line_y(line) > height * 0.12:
            continue
        text = _regional_text(line, width, 0.00, 0.45)
        match = re.search(r"\bpersonel(?:\s*alt|\s*t)?\s*alan[ıi]\s*:?\s*(.*)", text, re.IGNORECASE)
        if match:
            value = re.sub(r"\s+", " ", match.group(1)).strip(" :")
            return value or None
    return None


def _extract_employee_count(
    lines: list[list[TextBox]], width: float, height: float
) -> int | None:
    for line in lines:
        if _line_y(line) > height * 0.12:
            continue
        text = _regional_text(line, width, 0, 0.30).casefold().replace("ı", "i")
        text = "".join(
            char for char in unicodedata.normalize("NFKD", text)
            if not unicodedata.combining(char)
        )
        match = re.search(
            r"\bcalisan\s+sayisi\s*:?\s*(\d{1,3}(?:[.,]\d{3})+|\d+)(?![\d.,])",
            text,
        )
        if match:
            return int(re.sub(r"[.,]", "", match.group(1)))
    return None


def _extract_page_values_from_boxes(
    page,
    boxes: list[TextBox],
    width: float,
    height: float,
    ocr_scale: float,
) -> dict:
    lines = _group_lines(boxes, height)
    gross, yk_fee = _extract_gross_amount(lines, width, height)
    values = _extract_additional_values(
        lines,
        width,
        height,
        yk_fee=yk_fee,
        page=page,
        ocr_scale=ocr_scale,
    )
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
    values.update({
        "employee_count": _extract_employee_count(lines, width, height),
        "personnel_subarea": _extract_personnel_subarea(lines, width, height),
        "gross": gross,
        "yk_fee": yk_fee,
        "income": income,
        "stamp": stamp,
    })
    if values["employee_count"] is None or values["personnel_subarea"] is None:
        # Retry only the header so metadata OCR does not change financial extraction.
        try:
            header_boxes = _ocr_region_boxes(
                page,
                fitz.Rect(0, 0, page.rect.width * 0.45, page.rect.height * 0.12),
                3.0,
            )
            header_lines = _group_lines(header_boxes, page.rect.height * 3.0)
            if values["employee_count"] is None:
                values["employee_count"] = _extract_employee_count(
                    header_lines, page.rect.width * 3.0, page.rect.height * 3.0,
                )
            if values["personnel_subarea"] is None:
                values["personnel_subarea"] = _extract_personnel_subarea(
                    header_lines, page.rect.width * 3.0, page.rect.height * 3.0,
                )
        except (RuntimeError, ValueError):
            pass
    return values


def extract_page_values(page) -> dict:
    failures = []

    def attempt(name, reader, scale, source_page=None):
        try:
            boxes, width, height = reader()
            return _extract_page_values_from_boxes(
                source_page if source_page is not None else page,
                boxes, width, height, scale,
            )
        except OCRUnavailableError:
            raise
        except Exception as exc:
            failures.append(f"{name}: {exc}")
            return None

    # Do not discard readable native text because one heading has accents,
    # joined words or a line break. The parser itself validates the fields.
    values = attempt(
        "PDF metni",
        lambda: (_native_boxes(page), float(page.rect.width), float(page.rect.height)),
        1.0,
    )
    if values is not None:
        return values
    if getattr(page, "rotation", 0):
        # Some exports attach a rotation flag to an otherwise upright payroll.
        # Try its native coordinates on a copy; never mutate the uploaded PDF.
        with fitz.open() as copy:
            copy.insert_pdf(page.parent, from_page=page.number, to_page=page.number)
            unrotated = copy[0]
            unrotated.set_rotation(0)
            values = attempt(
                "PDF metni (döndürme bilgisi kaldırılmış kopya)",
                lambda: (
                    _native_boxes(unrotated),
                    float(unrotated.rect.width), float(unrotated.rect.height),
                ),
                1.0, source_page=unrotated,
            )
            if values is not None:
                return values
    try:
        deskewed = _deskew_page_image(page)
    except OCRUnavailableError:
        raise
    except Exception as exc:
        failures.append(f"Tarama eğimi kontrolü: {exc}")
        deskewed = None
    if deskewed is not None:
        png, width, height, angle = deskewed
        with fitz.open() as copy:
            corrected = copy.new_page(width=width, height=height)
            corrected.insert_image(corrected.rect, stream=png)
            for scale in (2.0, 3.0):
                values = attempt(
                    f"Eğimi düzeltilmiş OCR ({angle:.2f}°, {scale:g}x)",
                    lambda: _ocr_boxes(corrected, scale=scale),
                    scale, source_page=corrected,
                )
                if values is not None:
                    return values
            values = attempt(
                "Eğimi düzeltilmiş bölgesel OCR",
                lambda: _ocr_banded_boxes(corrected), 3.0, source_page=corrected,
            )
            if values is not None:
                return values
    for scale in (2.0, 3.0):
        values = attempt(
            f"Tam sayfa OCR ({scale:g}x)",
            lambda: _ocr_boxes(page, scale=scale),
            scale,
        )
        if values is not None:
            return values
    values = attempt("Bölgesel OCR", lambda: _ocr_banded_boxes(page), 3.0)
    if values is not None:
        return values
    raise ValueError("Sayfa okunamadı. " + " | ".join(failures))


def analyze_tax_pdfs(pdf_files: list[tuple[str, str]]) -> dict:
    rows = []
    warnings = []
    failed_pages = []
    totals = {key: Decimal("0") for key, _ in _EXCEL_COLUMNS}
    extra_payment_columns: dict[str, dict] = {}
    extra_payment_totals: dict[str, Decimal] = {}

    for path, original_name in pdf_files:
        try:
            document = fitz.open(path)
        except Exception as exc:
            warnings.append(f"{original_name}: PDF açılamadı ({exc}).")
            continue

        try:
            for page_number, page in enumerate(document, start=1):
                try:
                    values = extract_page_values(page)
                except OCRUnavailableError:
                    raise
                except Exception as exc:
                    failed_pages.append({
                        "filename": original_name, "page": page_number, "reason": str(exc),
                    })
                    warnings.append(
                        f"{original_name} - Sayfa {page_number}: {exc} "
                        "Bu sayfa toplamlara dahil edilmedi; toplamlar eksiktir."
                    )
                    continue
                extra_payment_warning = values.pop("extra_payment_warning", None)
                if values.get("employee_count") is None:
                    warnings.append(
                        f"{original_name} - Sayfa {page_number}: "
                        "Çalışan sayısı okunamadı; ilgili hücre boş bırakıldı."
                    )
                if not values.get("personnel_subarea"):
                    warnings.append(
                        f"{original_name} - Sayfa {page_number}: "
                        "Personel Alt Alanı okunamadı; ilgili hücre boş bırakıldı."
                    )
                if extra_payment_warning:
                    warnings.append(
                        f"{original_name} - Sayfa {page_number}: "
                        f"{extra_payment_warning}"
                    )
                for key in totals:
                    totals[key] += values[key]

                page_payments = values.get("extra_payments", {})
                page_metadata = values.get("extra_payment_meta", {})
                resolved_payments: dict[str, Decimal] = {}
                resolved_metadata: dict[str, dict] = {}
                row_index = len(rows)
                for key, amount in page_payments.items():
                    metadata = dict(page_metadata.get(key, {}))
                    header = metadata.get("header", key)
                    resolved_key = key
                    resolved_payments[resolved_key] = (
                        resolved_payments.get(resolved_key, Decimal("0")) + amount
                    )
                    resolved_metadata[resolved_key] = metadata

                    if resolved_key not in extra_payment_columns:
                        extra_payment_columns[resolved_key] = {
                            "key": resolved_key,
                            "header": header,
                            "estimated_sources": [],
                        }
                    column = extra_payment_columns[resolved_key]
                    extra_payment_totals[resolved_key] = (
                        extra_payment_totals.get(resolved_key, Decimal("0")) + amount
                    )
                    if metadata.get("estimated"):
                        source = {
                            "row_index": row_index,
                            "filename": original_name,
                            "page": page_number,
                            "raw_label": metadata.get("raw_label", ""),
                            "reason": metadata.get("reason") or "Başlık kontrol edilmelidir; ödeme türü tahmin edilmedi.",
                        }
                        column["estimated_sources"].append(source)
                        warnings.append(
                            f"{original_name} - Sayfa {page_number}: "
                            f"'{source['raw_label']}' başlığı "
                            f"'{column['header']}' olarak korundu; kontrol edilmelidir. "
                            f"{source['reason']}"
                        )

                row = {
                    "filename": original_name,
                    "page": page_number,
                }
                row.update(values)
                row["extra_payments"] = resolved_payments
                row["extra_payment_meta"] = resolved_metadata
                rows.append(row)
        finally:
            document.close()

    result = {
        "rows": rows,
        "warnings": warnings,
        "failed_pages": failed_pages,
        "totals": totals,
        "extra_payment_columns": list(extra_payment_columns.values()),
        "extra_payment_totals": extra_payment_totals,
    }
    for key, value in totals.items():
        result[f"{key}_total"] = value
        result[f"{key}_total_text"] = format_tr_money(value)
    return result


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
    estimated_fill = "F4CCCC"
    estimated_font = "9C0006"
    thin_border = Border(
        left=Side(style="thin", color=border_color),
        right=Side(style="thin", color=border_color),
        top=Side(style="thin", color=border_color),
        bottom=Side(style="thin", color=border_color),
    )

    extra_columns = result.get("extra_payment_columns", [])
    last_column = 3 + len(_EXCEL_COLUMNS) + len(extra_columns)
    last_column_letter = get_column_letter(last_column)
    sheet.merge_cells(f"A1:{last_column_letter}1")
    title = sheet["A1"]
    title.value = "Bordro Toplamları"
    title.fill = PatternFill("solid", fgColor=dark_blue)
    title.font = Font(name="Arial", size=15, bold=True, color=white)
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 28

    if result.get("failed_pages"):
        sheet.merge_cells(f"A2:{last_column_letter}2")
        notice = sheet["A2"]
        notice.value = (
            f"DİKKAT: {len(result['failed_pages'])} sayfa okunamadı. "
            "Toplamlar yalnızca okunan sayfaları içerir. Uyarılar sayfasını kontrol edin."
        )
        notice.fill = PatternFill("solid", fgColor=estimated_fill)
        notice.font = Font(name="Arial", bold=True, color=estimated_font)
        notice.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[2].height = 32

    header_row = 3
    first_data_row = header_row + 1
    extra_headers = []
    for extra_column in extra_columns:
        excel_rows = sorted({
            first_data_row + int(source["row_index"])
            for source in extra_column.get("estimated_sources", [])
        })
        header = extra_column["header"]
        if excel_rows:
            row_label = "satırı" if len(excel_rows) == 1 else "satırları"
            row_text = ", ".join(str(row) for row in excel_rows)
            header = f"{header} (Kontrol gerekli - Excel {row_label}: {row_text})"
        extra_headers.append(header)
    headers = (
        ["Personel Alt Alanı", "Sayfa", "Çalışan Sayısı"]
        + [header for _, header in _EXCEL_COLUMNS]
        + extra_headers
    )
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(header_row, column, header)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.font = Font(name="Arial", bold=True, color=white)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    extra_start_column = 4 + len(_EXCEL_COLUMNS)
    for offset, extra_column in enumerate(extra_columns):
        if not extra_column.get("estimated_sources"):
            continue
        cell = sheet.cell(header_row, extra_start_column + offset)
        cell.fill = PatternFill("solid", fgColor=estimated_fill)
        cell.font = Font(name="Arial", bold=True, color=estimated_font)

    for row_number, row in enumerate(result["rows"], start=first_data_row):
        values = (
            [row.get("personnel_subarea"), row["page"], row.get("employee_count")]
            + [float(row[key]) for key, _ in _EXCEL_COLUMNS]
            + [
                float(row.get("extra_payments", {}).get(
                    extra_column["key"], Decimal("0")
                ))
                for extra_column in extra_columns
            ]
        )
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
        sheet.cell(row_number, 3).number_format = "#,##0"
        for column in range(4, last_column + 1):
            sheet.cell(row_number, column).number_format = "#,##0.00"

    last_data_row = first_data_row + len(result["rows"]) - 1
    total_row = last_data_row + 1
    sheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    total_label = sheet.cell(
        total_row, 1,
        "OKUNAN SAYFALAR TOPLAMI" if result.get("failed_pages") else "GENEL TOPLAM",
    )
    total_label.alignment = Alignment(horizontal="right", vertical="center")
    for column in range(3, last_column + 1):
        column_letter = get_column_letter(column)
        sheet.cell(
            total_row,
            column,
            f"=SUM({column_letter}{first_data_row}:{column_letter}{last_data_row})",
        )
    for column in range(1, last_column + 1):
        cell = sheet.cell(total_row, column)
        cell.fill = PatternFill("solid", fgColor=medium_blue)
        cell.font = Font(name="Arial", size=11, bold=True, color=dark_blue)
        cell.border = thin_border
    sheet.cell(total_row, 3).number_format = "#,##0"
    for column in range(4, last_column + 1):
        sheet.cell(total_row, column).number_format = "#,##0.00"

    sheet.auto_filter.ref = f"A{header_row}:{last_column_letter}{last_data_row}"
    sheet.freeze_panes = f"C{first_data_row}"
    sheet.row_dimensions[header_row].height = 42
    widths = [36, 9, 18] + [22] * len(_EXCEL_COLUMNS) + [28] * len(extra_columns)
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    for cell in sheet[header_row]:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
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
        f"bordro_toplamlari_{uuid.uuid4().hex[:8]}.xlsx",
    )
    workbook.save(output_path)
    return output_path
