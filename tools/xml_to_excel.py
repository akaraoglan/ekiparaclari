import os
import uuid
import xml.etree.ElementTree as ET
import pandas as pd
from tools.common import secure_tr_filename

def xml_to_excel(input_path: str, output_dir: str) -> str:
    """
    XML e-fatura dosyasını Excel'e dönüştürür.
    """
    namespaces = {
        'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
        'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'
    }

    try:
        # XML'i işle
        tree = ET.parse(input_path)
        root = tree.getroot()

        # Fatura satırlarını işle
        invoice_lines = root.findall('.//cac:InvoiceLine', namespaces)
        if not invoice_lines:
            raise Exception("Fatura satırı bulunamadı")

        data = process_invoice_lines(invoice_lines, namespaces)

        # Fatura numarasını dosya adı olarak kullan
        invoice_id_element = root.find('.//cbc:ID', namespaces)
        invoice_number = invoice_id_element.text.strip() if invoice_id_element is not None and invoice_id_element.text else None
        if invoice_number:
            safe_name = secure_tr_filename(invoice_number)
        else:
            safe_name = os.path.splitext(os.path.basename(input_path))[0]

        output_filename = f"{safe_name}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        if os.path.exists(output_path):
            output_filename = f"{safe_name}_{uuid.uuid4().hex[:4]}.xlsx"
            output_path = os.path.join(output_dir, output_filename)

        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)

        return output_path

    except Exception as e:
        raise Exception(f"XML işleme hatası: {str(e)}")

def process_invoice_lines(invoice_lines, namespaces):
    """Fatura satırlarını işler"""
    # Faturadaki sütun sıralaması ve XML yolları
    field_mappings = {
        'Sıra No': './/cbc:ID',
        'Mal Hizmet': './/cac:Item/cbc:Name',
        'Miktar': './/cbc:InvoicedQuantity',
        'Birim Fiyat': './/cac:Price/cbc:PriceAmount',
        'İskonto Oranı': './/cac:AllowanceCharge/cbc:MultiplierFactorNumeric',
        'İskonto Tutarı': './/cac:AllowanceCharge/cbc:Amount',
        'KDV Oranı': './/cac:TaxTotal/cac:TaxSubtotal/cbc:Percent',
        'KDV Tutarı': './/cac:TaxTotal/cac:TaxSubtotal/cbc:TaxAmount',
        'Diğer Vergiler': './/cac:WithholdingTaxTotal/cbc:TaxAmount',
        'Mal Hizmet Tutarı': './/cbc:LineExtensionAmount',
        'Teslim Şartı': './/cac:Delivery/cac:DeliveryTerms/cbc:ID',
        'Eşya Kap Cinsi': './/cac:Delivery/cac:Shipment/cac:TransportHandlingUnit/cbc:PackagingTypeCode',
        'Kap No': './/cac:Delivery/cac:Shipment/cac:TransportHandlingUnit/cbc:ID',
        'Kap Adet': './/cac:Delivery/cac:Shipment/cac:TransportHandlingUnit/cbc:Quantity',
        'Teslim/Bedel Ödeme Yeri': './/cac:Delivery/cac:DeliveryAddress/cbc:BuildingName',
        'Gönderilme Şekli': './/cac:Delivery/cac:Shipment/cac:ShipmentStage/cbc:TransportModeCode',
        'GTİP': './/cac:Delivery/cac:Shipment/cac:GoodsItem/cbc:RequiredCustomsID'
    }

    data = []

    for line in invoice_lines:
        row_data = {}
        for field_name, xpath in field_mappings.items():
            try:
                element = line.find(xpath, namespaces)
                value = element.text.strip() if element is not None and element.text else None

                # Sayısal değerleri dönüştür
                if value:
                    try:
                        if '.' in str(value):
                            value = float(value)
                        elif str(value).isdigit():
                            value = int(value)
                    except:
                        pass

                row_data[field_name] = value

            except Exception as e:
                row_data[field_name] = None

        data.append(row_data)

    return data