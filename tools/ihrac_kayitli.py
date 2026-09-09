import os
import re
import uuid
import xml.etree.ElementTree as ET
from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, Side


KDV_ORANI = Decimal("0.20")
M3_BOLEN = Decimal("1000000000")

DETAIL_COLUMNS = {
    "reference": ["Referans"],
    "tax_base": ["KDV Matrahı"],
    "currency": ["Belge para birimi"],
    "quantity": ["Faturalanan miktar"],
    "thickness": ["Kalınlık"],
    "length": ["Boy"],
    "width": ["En"],
    "gtip": ["Gtip", "GTİP", "GTIP"],
}

SUMMARY_COLUMNS = {
    "date": ["Faturalama tarihi"],
    "reference": ["Referans"],
    "buyer": ["Fatura alıcısı adı"],
    "tax_no": ["Vergi Numarası"],
}

RESULT_HEADERS = [
    "Sıra No",
    "İhraç Kayıtlı Satış Faturasının Tarihi ",
    "İhraç Kayıtlı Satış Faturasının Serisi ",
    " İhraç Kayıtlı Satış Faturasının Sıra No’su",
    "Alıcının Adı Soyadı/Ünvanı ",
    "Alıcının Vergi Kimlik Numarası / T.C. Kimlik Numarası ",
    " Malın Cinsi ",
    "Malın Miktarı ",
    "Miktar Kodu",
    "Malın Kdv Hariç Tutarı ",
    " Malın Kdv’si ",
    "İhracatçı Tarafından Yurt Dışına Düzenlenen Satış Faturasının Tarihi (GÇB/ETGB-BGB Üzerindeki Bilgiler)",
    "GÇB/ETGB-BGB Tescil No",
    "GÇB/ETGB-BGB",
]

RESULT_WIDTHS = {
    "A": 6.29,
    "B": 12.29,
    "C": 21.86,
    "D": 20.14,
    "E": 26.00,
    "F": 71.86,
    "G": 26.14,
    "H": 19.71,
    "I": 11.71,
    "J": 17.57,
    "K": 20.00,
    "L": 17.14,
    "M": 15.43,
    "N": 21.29,
    "O": 15.29,
    "P": 9.14,
}

RESULT_FORMATS = {
    "B": "0",
    "C": "mm-dd-yy",
    "D": "@",
    "E": "0",
    "F": "@",
    "G": "@",
    "H": "0",
    "I": "#,##0.00",
    "J": "@",
    "K": "#,##0.00",
    "L": "#,##0.00",
    "M": "mm-dd-yy",
    "N": "@",
    "O": "@",
}


def process_ihrac_kayitli(
    detay_path: str,
    ozet_path: str,
    output_dir: str,
    xml_paths: list[str] | None = None,
) -> tuple:
    """
    Detay ve ozet Excel dosyalarindan ihrac kayitli satis faturasi listesi olusturur.
    Doner: (status, output_path, message, missing_xml_refs)
    """
    try:
        summary_rows = _read_summary_rows(ozet_path)
        detail_groups, detail_refs = _read_detail_groups(detay_path)
        invoices = _read_invoice_xmls(xml_paths or [])

        output_rows = []
        missing_refs = []
        missing_xml_refs = set()
        used_refs = set()
        xml_filled_rows = 0

        for summary in summary_rows:
            reference = summary["reference"]
            groups = detail_groups.get(reference)
            if not groups:
                missing_refs.append(reference)
                continue

            used_refs.add(reference)
            sorted_groups = sorted(groups.items(), key=lambda item: item[0], reverse=True)
            currencies = {
                currency
                for _, group in sorted_groups
                for currency in group["currencies"]
            }
            is_try = bool(currencies) and currencies == {"TRY"}

            allocated_bases = None
            allocated_taxes = None
            if not is_try:
                invoice = invoices.get(_normalize_reference(reference))
                if invoice is None:
                    missing_xml_refs.add(reference)
                    continue
                if currencies and currencies != {invoice["currency"]}:
                    detail_currency = ", ".join(sorted(currencies))
                    raise ValueError(
                        f"{reference}: Detay para birimi ({detail_currency}) ile XML para birimi "
                        f"({invoice['currency']}) uyuşmuyor."
                    )

                weights = [group["tax_base"] for _, group in sorted_groups]
                tl_base = _round_two(invoice["tax_base"] * invoice["exchange_rate"])
                tl_tax = _round_two(invoice["tax_amount"] * invoice["exchange_rate"])
                allocated_bases = _allocate_total(tl_base, weights, reference, "matrah")
                allocated_taxes = _allocate_total(tl_tax, weights, reference, "KDV")

            for group_index, (gtip, group) in enumerate(sorted_groups):
                if group["m3"] == 0:
                    miktar = _truncate_two(group["quantity"])
                    miktar_kodu = "MTK"
                else:
                    miktar = _truncate_two(group["m3"])
                    miktar_kodu = "MTQ"

                if is_try:
                    matrah = _round_two(group["tax_base"])
                    kdv = _round_two(group["tax_base"] * KDV_ORANI)
                else:
                    matrah = allocated_bases[group_index]
                    kdv = allocated_taxes[group_index]
                    xml_filled_rows += 1

                output_rows.append({
                    "date": summary["date"],
                    "reference": reference,
                    "buyer": summary["buyer"],
                    "tax_no": summary["tax_no"],
                    "miktar": miktar,
                    "miktar_kodu": miktar_kodu,
                    "matrah": matrah,
                    "kdv": kdv,
                    "gtip": gtip,
                })

        if missing_xml_refs:
            missing_list = sorted(missing_xml_refs)
            message = (
                f"{len(missing_list)} dövizli fatura için XML gerekiyor: "
                f"{', '.join(missing_list)}"
            )
            return "pending", None, message, missing_list

        if not output_rows:
            return (
                "error",
                None,
                "Özet dosyasındaki referanslar için Detay dosyasında eşleşen kayıt bulunamadı.",
                [],
            )

        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, f"ihrac_kayitli_sonuc_{uuid.uuid4().hex[:8]}.xlsx")
        _write_result(output_rows, output_path)

        extra_refs = sorted(detail_refs - {row["reference"] for row in summary_rows})
        warnings = []
        if missing_refs:
            warnings.append(f"Özet'te olup Detay'da bulunmayan referans: {', '.join(missing_refs)}")
        if extra_refs:
            warnings.append(f"Detay'da olup Özet'te bulunmayan referans: {', '.join(extra_refs)}")
        message = f"Tamamlandı. {len(output_rows)} satır oluşturuldu."
        if xml_filled_rows:
            message = f"{message} {xml_filled_rows} dövizli satır XML'den TL olarak dolduruldu."
        if warnings:
            return "partial", output_path, f"{message} {' | '.join(warnings)}", []
        return "success", output_path, message, []

    except Exception as exc:
        return "error", None, f"Hata oluştu: {exc}", []


def _read_invoice_xmls(paths: list[str]) -> dict:
    invoices = {}
    for path in paths:
        invoice = _read_invoice_xml(path)
        reference_key = _normalize_reference(invoice["reference"])
        if reference_key in invoices:
            raise ValueError(f"{invoice['reference']}: Aynı fatura için birden fazla XML yüklendi.")
        invoices[reference_key] = invoice
    return invoices


def _read_invoice_xml(path: str) -> dict:
    try:
        root = ET.parse(path).getroot()
    except (ET.ParseError, OSError) as exc:
        raise ValueError(f"{Path(path).name}: XML okunamadı ({exc}).") from exc

    if _local_name(root.tag) != "Invoice":
        raise ValueError(f"{Path(path).name}: Dosya bir UBL fatura XML'i değil.")

    reference = _direct_child_text(root, "ID")
    currency = _direct_child_text(root, "DocumentCurrencyCode").upper()
    if not reference or not currency:
        raise ValueError(f"{Path(path).name}: Fatura numarası veya para birimi bulunamadı.")

    tax_base_node = _descendant_in(root, "LegalMonetaryTotal", "TaxExclusiveAmount")
    tax_amount_node = _descendant_in(root, "TaxTotal", "TaxAmount")
    if tax_base_node is None or tax_amount_node is None:
        raise ValueError(f"{reference}: XML'de matrah veya KDV toplamı bulunamadı.")

    tax_base_currency = tax_base_node.attrib.get("currencyID", "").upper()
    tax_currency = tax_amount_node.attrib.get("currencyID", "").upper()
    if tax_base_currency and tax_base_currency != currency:
        raise ValueError(f"{reference}: XML matrah para birimi fatura para birimiyle uyuşmuyor.")
    if tax_currency and tax_currency != currency:
        raise ValueError(f"{reference}: XML KDV para birimi fatura para birimiyle uyuşmuyor.")

    if currency == "TRY":
        exchange_rate = Decimal("1")
    else:
        exchange_rate = _find_try_exchange_rate(root, currency, reference)

    return {
        "reference": reference,
        "currency": currency,
        "exchange_rate": exchange_rate,
        "tax_base": _xml_decimal(tax_base_node.text, reference, "matrah"),
        "tax_amount": _xml_decimal(tax_amount_node.text, reference, "KDV"),
    }


def _find_try_exchange_rate(root, currency: str, reference: str) -> Decimal:
    for element in root.iter():
        if _local_name(element.tag) not in {
            "PaymentExchangeRate",
            "PricingExchangeRate",
            "TaxExchangeRate",
        }:
            continue
        source = _direct_child_text(element, "SourceCurrencyCode").upper()
        target = _direct_child_text(element, "TargetCurrencyCode").upper()
        rate_text = _direct_child_text(element, "CalculationRate")
        if source == currency and target == "TRY" and rate_text:
            return _xml_decimal(rate_text, reference, "döviz kuru")
    raise ValueError(f"{reference}: XML'de {currency}→TRY döviz kuru bulunamadı.")


def _allocate_total(
    total: Decimal,
    weights: list[Decimal],
    reference: str,
    field_name: str,
) -> list[Decimal]:
    weight_total = sum(weights, Decimal("0"))
    if weight_total <= 0:
        raise ValueError(f"{reference}: {field_name} GTİP satırlarına bölüştürülemiyor.")

    allocated = [
        _round_two(total * weight / weight_total)
        for weight in weights
    ]
    difference = total - sum(allocated, Decimal("0"))
    if difference:
        largest_index = max(range(len(weights)), key=lambda index: weights[index])
        allocated[largest_index] += difference
    return allocated


def _descendant_in(root, parent_name: str, child_name: str):
    for parent in root.iter():
        if _local_name(parent.tag) != parent_name:
            continue
        for child in list(parent):
            if _local_name(child.tag) == child_name:
                return child
    return None


def _direct_child_text(element, child_name: str) -> str:
    for child in list(element):
        if _local_name(child.tag) == child_name:
            return (child.text or "").strip()
    return ""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml_decimal(value, reference: str, field_name: str) -> Decimal:
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, AttributeError) as exc:
        raise ValueError(f"{reference}: XML {field_name} değeri sayısal değil.") from exc


def _normalize_reference(value) -> str:
    return _clean_text(value).replace(" ", "").upper()


def _read_summary_rows(path: str) -> list:
    ws = _load_data_sheet(path)
    cols = _resolve_columns(ws, SUMMARY_COLUMNS, "Özet")
    rows = []

    for row_num in range(2, ws.max_row + 1):
        reference = _clean_text(ws.cell(row_num, cols["reference"]).value)
        if not reference:
            continue
        rows.append({
            "date": ws.cell(row_num, cols["date"]).value,
            "reference": reference,
            "buyer": _clean_text(ws.cell(row_num, cols["buyer"]).value),
            "tax_no": _clean_text(ws.cell(row_num, cols["tax_no"]).value),
        })

    if not rows:
        raise ValueError("Özet dosyasında işlenecek referans bulunamadı.")
    return rows


def _read_detail_groups(path: str) -> tuple:
    ws = _load_data_sheet(path)
    cols = _resolve_columns(ws, DETAIL_COLUMNS, "Detay")
    groups = defaultdict(dict)
    detail_refs = set()

    for row_num in range(2, ws.max_row + 1):
        reference = _clean_text(ws.cell(row_num, cols["reference"]).value)
        if not reference:
            continue

        detail_refs.add(reference)
        gtip = _clean_text(ws.cell(row_num, cols["gtip"]).value) or "GTIP_YOK"
        group = groups[reference].setdefault(gtip, {
            "tax_base": Decimal("0"),
            "m3": Decimal("0"),
            "quantity": Decimal("0"),
            "currencies": set(),
        })

        tax_base = _decimal_value(ws.cell(row_num, cols["tax_base"]).value, row_num, "KDV Matrahı")
        quantity = _decimal_value(ws.cell(row_num, cols["quantity"]).value, row_num, "Faturalanan miktar")
        thickness = _decimal_value(ws.cell(row_num, cols["thickness"]).value, row_num, "Kalınlık")
        length = _decimal_value(ws.cell(row_num, cols["length"]).value, row_num, "Boy")
        width = _decimal_value(ws.cell(row_num, cols["width"]).value, row_num, "En")
        currency = _clean_text(ws.cell(row_num, cols["currency"]).value).upper()

        group["tax_base"] += tax_base
        group["quantity"] += quantity
        group["m3"] += quantity * thickness * length * width / M3_BOLEN
        if currency:
            group["currencies"].add(currency)

    if not detail_refs:
        raise ValueError("Detay dosyasında işlenecek referans bulunamadı.")
    return groups, detail_refs


def _write_result(rows: list, output_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sonuc"

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_font = Font(name="Arial", size=14, bold=True)
    header_font = Font(name="Arial", size=10, bold=True)
    data_font = Font(name="Arial", size=9)
    total_font = Font(name="Arial", size=9, bold=True)

    for col, width in RESULT_WIDTHS.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[4].height = 140.25

    ws["H2"] = "İHRAÇ KAYITLI SATIŞ FATURASI LİSTESİ"
    ws["H2"].font = title_font
    ws["H2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["H2"].number_format = "@"

    for offset, header in enumerate(RESULT_HEADERS, start=2):
        cell = ws.cell(4, offset, header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        cell.number_format = RESULT_FORMATS.get(cell.column_letter, "General")

    for index, item in enumerate(rows, start=1):
        row_num = index + 4
        values = [
            index,
            item["date"],
            "",
            item["reference"],
            item["buyer"],
            item["tax_no"],
            _clean_text(item["gtip"]),
            _to_float(item["miktar"]),
            item["miktar_kodu"],
            _to_float(item["matrah"]),
            _to_float(item["kdv"]),
            None,
            None,
            "GÇB",
        ]

        for offset, value in enumerate(values, start=2):
            cell = ws.cell(row_num, offset, value)
            cell.font = data_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            cell.number_format = RESULT_FORMATS.get(cell.column_letter, "General")
            if cell.column_letter == "H" and value:
                cell.quotePrefix = True

    total_row = len(rows) + 5
    ws.cell(total_row, 11, "TOPLAM")
    ws.cell(total_row, 12, f"=SUM(K5:K{total_row - 1})")
    for cell in (ws.cell(total_row, 11), ws.cell(total_row, 12)):
        cell.font = total_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        cell.number_format = "#,##0.00"

    wb.save(output_path)


def _load_data_sheet(path: str):
    workbook = load_workbook(path, data_only=True)
    if "Data" in workbook.sheetnames:
        return workbook["Data"]
    return workbook[workbook.sheetnames[0]]


def _resolve_columns(ws, required: dict, label: str) -> dict:
    normalized_headers = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(1, col).value
        if value is not None:
            normalized_headers[_normalize_header(value)] = col

    resolved = {}
    missing = []
    for key, aliases in required.items():
        column = None
        for alias in aliases:
            column = normalized_headers.get(_normalize_header(alias))
            if column:
                break
        if column:
            resolved[key] = column
        else:
            missing.append(aliases[0])

    if missing:
        raise ValueError(f"{label} dosyasında eksik kolon: {', '.join(missing)}")
    return resolved


def _normalize_header(value) -> str:
    text = str(value).strip().lower()
    text = text.translate(str.maketrans({
        "ı": "i",
        "İ": "i",
        "ğ": "g",
        "Ğ": "g",
        "ü": "u",
        "Ü": "u",
        "ş": "s",
        "Ş": "s",
        "ö": "o",
        "Ö": "o",
        "ç": "c",
        "Ç": "c",
    }))
    return re.sub(r"[^a-z0-9]+", "", text)


def _decimal_value(value, row_num: int, field_name: str) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))

    text = str(value).strip().replace(" ", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"{row_num}. satırda {field_name} sayısal değil: {value}") from exc


def _clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _round_two(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _truncate_two(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_DOWN)


def _to_float(value):
    if value is None:
        return None
    return float(value)
