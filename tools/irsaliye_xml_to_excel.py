import os
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd

NS = {
    'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
    'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
}

UNIT_MAP = {
    'KGM': 'kg', 'KG': 'kg',
    'C62': 'adet', 'NIU': 'adet',
    'LTR': 'lt', 'MTK': 'm2',
    'MTR': 'm', 'TNE': 'ton',
}


def _text(node, xpath):
    el = node.find(xpath, NS)
    return el.text.strip() if el is not None and el.text else ''


def _format_date(raw):
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(raw.strip(), fmt).strftime('%d.%m.%Y')
        except ValueError:
            pass
    return raw.strip().replace('-', '.')


def _normalize_qty(text):
    if not text:
        return ''
    text = text.strip().replace(',', '.')
    try:
        num = Decimal(text)
        return int(num) if num == num.to_integral() else float(num)
    except (InvalidOperation, ValueError):
        return text


def _customer_name(root):
    first = _text(root, './/cac:DeliveryCustomerParty/cac:Party/cac:Person/cbc:FirstName')
    last  = _text(root, './/cac:DeliveryCustomerParty/cac:Party/cac:Person/cbc:FamilyName')
    full  = f'{first} {last}'.strip()
    if full:
        return full

    for xpath in (
        './/cac:DeliveryCustomerParty/cac:Party/cac:PartyName/cbc:Name',
        './/cac:DeliveryCustomerParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName',
        './/cac:BuyerCustomerParty/cac:Party/cac:PartyName/cbc:Name',
        './/cac:BuyerCustomerParty/cac:Party/cac:PartyLegalEntity/cbc:RegistrationName',
    ):
        val = _text(root, xpath)
        if val:
            return val

    for note in root.findall('.//cbc:Note', NS):
        text = (note.text or '').strip()
        if text.startswith('FIR:'):
            return text.replace('FIR:', '', 1).strip()

    return ''


def _parse_file(xml_path):
    root = ET.parse(xml_path).getroot()
    customer  = _customer_name(root)
    issue_date = _format_date(_text(root, './/cbc:IssueDate'))
    despatch_no = _text(root, './/cbc:ID')

    lines = root.findall('.//cac:DespatchLine', NS)
    if not lines:
        raise ValueError('DespatchLine bulunamadı')

    rows = []
    for line in lines:
        product = _text(line, './/cac:Item/cbc:Name')
        qty_el  = line.find('.//cbc:DeliveredQuantity', NS)
        quantity = _normalize_qty(qty_el.text if qty_el is not None else '')
        unit_code = qty_el.attrib.get('unitCode', '') if qty_el is not None else ''
        unit = UNIT_MAP.get(unit_code.upper(), unit_code.lower())

        rows.append({
            'Müşteri Adı': customer,
            'Tarih':        issue_date,
            'İrsaliye No':  despatch_no,
            'Ürün':         product,
            'Miktar':       quantity,
            'Birim':        unit,
        })

    return rows


def _autosize(ws):
    for col_cells in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col_cells), default=0)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 45)


def irsaliye_xml_to_excel(file_paths: list, output_dir: str) -> str:
    all_rows = []
    errors   = []

    for path in file_paths:
        try:
            all_rows.extend(_parse_file(path))
        except Exception as e:
            errors.append(f'{os.path.basename(path)}: {e}')

    if not all_rows:
        msg = 'Hiçbir dosyadan kayıt çıkarılamadı.'
        if errors:
            msg += ' Hatalar: ' + '; '.join(errors)
        raise ValueError(msg)

    cols = ['Müşteri Adı', 'Tarih', 'İrsaliye No', 'Ürün', 'Miktar', 'Birim']
    df   = pd.DataFrame(all_rows, columns=cols)

    output_path = os.path.join(output_dir, f'irsaliye_raporu_{uuid.uuid4().hex[:8]}.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='İrsaliye')
        ws = writer.book['İrsaliye']
        _autosize(ws)
        ws.freeze_panes = 'A2'

    return output_path, errors
