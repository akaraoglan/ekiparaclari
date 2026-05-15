import os
import re
import uuid
import xml.etree.ElementTree as ET
import pandas as pd

NAMESPACES = {
    'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
    'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'
}

# UUID formatındaki değerleri dışarıda bırak (örn: 26b30961-69ec-4b9d-84bc-f9bf4ca13e21)
_UUID_RE = re.compile(
    r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$',
    re.IGNORECASE
)


def _is_uuid(val: str) -> bool:
    return bool(_UUID_RE.match(val))


def extract_waybills(root) -> list:
    """Sadece DespatchDocumentReference'dan irsaliye numaralarını çeker."""
    found = []
    for dr in root.findall('.//cac:DespatchDocumentReference', NAMESPACES):
        el = dr.find('cbc:ID', NAMESPACES)
        if el is not None and el.text:
            val = el.text.strip()
            if val and not _is_uuid(val) and val not in found:
                found.append(val)
    return found


def irsaliye_no_to_excel(file_paths: list, output_dir: str) -> str:
    all_waybills = []

    for file_path in file_paths:
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            waybills = extract_waybills(root)
            all_waybills.extend(waybills)
        except Exception:
            pass

    df = pd.DataFrame({'İrsaliye No': all_waybills})

    output_filename = f"irsaliye_no_{uuid.uuid4().hex[:8]}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    df.to_excel(output_path, index=False)

    return output_path
