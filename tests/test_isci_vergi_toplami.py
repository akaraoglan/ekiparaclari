import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from tools.isci_vergi_toplami import (
    TextBox,
    _EXCEL_COLUMNS,
    _canonical_extra_payment_label,
    _extract_extra_payment_items,
    _find_extra_payment_total,
    _named_amount,
    create_tax_excel,
    extract_page_values,
)


class IsciVergiToplamiTest(unittest.TestCase):
    def test_extracts_each_extra_payment_and_reconciles_total(self):
        lines = [
            self._line(200, "Çocuk Parası", "3.544,92"),
            self._line(240, "Yıllık İzin Harçlığı", "5.730,00"),
            self._line(280, "Yakacak Yardımı (Na", "368.678,04"),
        ]

        payments, metadata = _extract_extra_payment_items(
            lines,
            width=1000,
            extra_y=150,
            total_y=350,
        )

        self.assertEqual(Decimal("3544.92"), payments["cocuk_parasi"])
        self.assertEqual(Decimal("5730.00"), payments["yillik_izin_harcligi"])
        self.assertEqual(Decimal("368678.04"), payments["yakacak_yardimi"])
        self.assertEqual(Decimal("377952.96"), sum(payments.values()))
        self.assertFalse(metadata["yakacak_yardimi"]["estimated"])

    def test_infers_partly_unreadable_yakacak_heading(self):
        key, header, estimated, reason, _ = _canonical_extra_payment_label(
            "a ik Yardimi (Na",
            confidence=0.91,
        )

        self.assertEqual("yakacak_yardimi", key)
        self.assertEqual("Yakacak Yardımı", header)
        self.assertTrue(estimated)
        self.assertIn("okunamadı", reason)

    def test_finds_total_when_first_letter_is_not_read(self):
        lines = [self._line(400, "oplam", "377.952,96")]

        result = _find_extra_payment_total(
            lines,
            width=1000,
            y_min=300,
            y_max=500,
            target_x=400,
        )

        self.assertEqual((405.0, Decimal("377952.96")), result)

    def test_reads_wrapped_amount_from_line_below_heading(self):
        lines = [
            [TextBox("İşveren Maliyeti (Teşvikli)", 450, 700, 760, 710)],
            [TextBox("879.691,50", 830, 716, 950, 726)],
        ]

        amount = _named_amount(
            lines,
            "isveren maliyeti",
            width=1000,
            label_x=(0.43, 0.78),
            amount_x=(0.78, 0.99),
            y_min=600,
            y_max=900,
            allow_next_line=True,
        )

        self.assertEqual(Decimal("879691.50"), amount)

    def test_scanned_page_retries_at_higher_resolution(self):
        expected = {"gross": Decimal("1.00")}
        with patch(
            "tools.isci_vergi_toplami._native_boxes",
            return_value=[],
        ), patch(
            "tools.isci_vergi_toplami._ocr_boxes",
            side_effect=[([], 1000.0, 1500.0), ([], 1500.0, 2250.0)],
        ) as ocr_boxes, patch(
            "tools.isci_vergi_toplami._extract_page_values_from_boxes",
            side_effect=[ValueError("düşük kalite"), expected],
        ):
            result = extract_page_values(object())

        self.assertIs(expected, result)
        self.assertEqual(
            [2.0, 3.0],
            [call.kwargs["scale"] for call in ocr_boxes.call_args_list],
        )

    def test_excel_appends_dynamic_columns_and_keeps_amounts_numeric(self):
        rows = []
        for page in (1, 2):
            row = {
                "filename": "ornek.pdf",
                "page": page,
                **{key: Decimal("1.00") for key, _ in _EXCEL_COLUMNS},
                "extra_payments": {
                    "cocuk_parasi": Decimal("10.00") * page,
                    "yakacak_yardimi": Decimal("20.00") * page,
                },
            }
            rows.append(row)

        result = {
            "rows": rows,
            "warnings": ["Yakacak yardımı başlığı kontrol edilmelidir."],
            "extra_payment_columns": [
                {
                    "key": "cocuk_parasi",
                    "header": "Çocuk Parası",
                    "estimated_sources": [],
                },
                {
                    "key": "yakacak_yardimi",
                    "header": "Yakacak Yardımı",
                    "estimated_sources": [{"row_index": 1}],
                },
            ],
        }

        output_root = Path(__file__).resolve().parents[1] / "outputs"
        output_root.mkdir(parents=True, exist_ok=True)
        output_path = create_tax_excel(result, str(output_root))
        try:
            workbook = load_workbook(output_path, data_only=False)
            try:
                sheet = workbook["Vergi Toplamları"]
                self.assertEqual("Brüt Ücret (YK Dahil)", sheet["C3"].value)
                self.assertEqual("İşveren Maliyeti (Teşvikli)", sheet["T3"].value)
                self.assertEqual("Çocuk Parası", sheet["U3"].value)
                self.assertIn("Tahmini - Excel satırı: 5", sheet["V3"].value)
                self.assertEqual("F4CCCC", sheet["V3"].fill.fgColor.rgb[-6:])
                self.assertEqual(40.0, sheet["V5"].value)
                self.assertEqual("n", sheet["V5"].data_type)
                self.assertEqual("#,##0.00", sheet["V5"].number_format)
                self.assertEqual("=SUM(V4:V5)", sheet["V6"].value)
                self.assertIn("Uyarılar", workbook.sheetnames)
            finally:
                workbook.close()
        finally:
            Path(output_path).unlink(missing_ok=True)

    @staticmethod
    def _line(y: float, label: str, amount: str) -> list[TextBox]:
        return [
            TextBox(label, 10, y, 260, y + 10),
            TextBox(amount, 350, y, 450, y + 10),
        ]


if __name__ == "__main__":
    unittest.main()
