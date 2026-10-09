import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import fitz
from openpyxl import load_workbook

from tools.isci_vergi_toplami import (
    TextBox,
    _EXCEL_COLUMNS,
    _extract_employee_count,
    _canonical_extra_payment_label,
    _extract_extra_payment_items,
    _find_extra_payment_total,
    _named_amount,
    analyze_tax_pdfs,
    create_tax_excel,
    extract_page_values,
)


class IsciVergiToplamiTest(unittest.TestCase):
    def test_reads_employee_count_without_using_neighboring_counts(self):
        lines = [[
            TextBox("ÇALIŞAN SAYISI", 10, 60, 130, 70),
            TextBox(":", 140, 60, 145, 70),
            TextBox("21", 180, 60, 200, 70),
            TextBox("İşe Giren : 7", 310, 60, 400, 70),
        ]]
        self.assertEqual(21, _extract_employee_count(lines, 1000, 1500))
        self.assertEqual(1109, _extract_employee_count(
            [[TextBox("ÇALIŞAN SAYISI : 1.109", 10, 60, 250, 70)]], 1000, 1500
        ))
        self.assertEqual(0, _extract_employee_count(
            [[TextBox("CALISAN SAYISI : 0", 10, 60, 200, 70)]], 1000, 1500
        ))
        self.assertIsNone(_extract_employee_count(
            [[TextBox("Kadın Sayısı : 8", 10, 60, 200, 70)]], 1000, 1500
        ))

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

        self.assertEqual(Decimal("3544.92"), payments["ocr_çocuk parası"])
        self.assertEqual(Decimal("5730.00"), payments["ocr_yıllık izin harçlığı"])
        self.assertEqual(Decimal("368678.04"), payments["ocr_yakacak yardımı (na"])
        self.assertEqual(Decimal("377952.96"), sum(payments.values()))
        self.assertFalse(metadata["ocr_yakacak yardımı (na"]["estimated"])

    def test_preserves_payment_headings_without_guessing_categories(self):
        for label in (
            "Bayram Yardımı", "Yemek Yardımı", "Yemek Ödemesi",
            "a ik Yardimi (Na", "Yakacak Yardımı (Na",
            "Doğum Yardımı", "Doğum Yardımı (Nakdi)", "Do um Yardimi",
            "� Yardımı",
        ):
            with self.subTest(label=label):
                _, header, _, _, _ = _canonical_extra_payment_label(label)
                self.assertEqual(label, header)
        _, header, needs_review, reason, _ = _canonical_extra_payment_label(
            "Bayram Yardımı", confidence=0.4,
        )
        self.assertEqual("Bayram Yardımı", header)
        self.assertTrue(needs_review)
        self.assertIn("tahmin edilmedi", reason)

    def test_missing_heading_keeps_amount_with_review_warning(self):
        payments, metadata = _extract_extra_payment_items(
            [self._line(200, "", "125,00")], 1000, 150, 250,
        )
        self.assertEqual(Decimal("125.00"), sum(payments.values()))
        self.assertTrue(next(iter(metadata.values()))["estimated"])
        self.assertIn("Başlık okunamadı", next(iter(metadata.values()))["header"])

    def test_similar_headings_stay_separate_across_pdfs_and_excel(self):
        output_root = Path(__file__).resolve().parents[1] / "outputs"
        output_root.mkdir(parents=True, exist_ok=True)
        pdf_path = output_root / "test_extra_payments.pdf"
        with fitz.open() as document:
            document.new_page()
            document.new_page()
            document.save(pdf_path)
        values = []
        expected = {
            "Bayram Yardımı": Decimal("100.00"),
            "Yemek Yardımı": Decimal("200.00"),
            "Doğum Yardımı": Decimal("300.00"),
            "Do um Yardimi": Decimal("400.00"),
            "Yemek Yardımı (Nakdi)": Decimal("500.00"),
        }
        for labels in (list(expected)[:3], list(expected)[3:]):
            payments, metadata = _extract_extra_payment_items(
                [self._line(200 + i * 20, label, str(expected[label]))
                 for i, label in enumerate(labels)], 1000, 150, 350,
            )
            values.append({
                **{key: Decimal("0") for key, _ in _EXCEL_COLUMNS},
                "employee_count": 1,
                "extra_payments": payments, "extra_payment_meta": metadata,
            })
        excel_path = None
        try:
            with patch("tools.isci_vergi_toplami.extract_page_values", side_effect=values):
                result = analyze_tax_pdfs([(str(pdf_path), "ornek.pdf")])
            self.assertEqual(list(expected), [c["header"] for c in result["extra_payment_columns"]])
            self.assertEqual(sum(expected.values()), sum(result["extra_payment_totals"].values()))
            self.assertEqual([], result["warnings"])
            excel_path = create_tax_excel(result, str(output_root))
            workbook = load_workbook(excel_path)
            try:
                sheet = workbook["Vergi Toplamları"]
                start = 4 + len(_EXCEL_COLUMNS)
                for offset, (label, amount) in enumerate(expected.items()):
                    column = start + offset
                    self.assertEqual(label, sheet.cell(3, column).value)
                    row = 4 if offset < 3 else 5
                    self.assertEqual(float(amount), sheet.cell(row, column).value)
                    self.assertEqual(0, sheet.cell(9 - row, column).value)
            finally:
                workbook.close()
        finally:
            pdf_path.unlink(missing_ok=True)
            if excel_path:
                Path(excel_path).unlink(missing_ok=True)

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
                "employee_count": 21 if page == 1 else None,
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
                self.assertEqual("Çalışan Sayısı", sheet["C3"].value)
                self.assertEqual(21, sheet["C4"].value)
                self.assertEqual("n", sheet["C4"].data_type)
                self.assertEqual("#,##0", sheet["C4"].number_format)
                self.assertIsNone(sheet["C5"].value)
                self.assertEqual("=SUM(C4:C5)", sheet["C6"].value)
                self.assertEqual("#,##0", sheet["C6"].number_format)
                self.assertEqual("Brüt Ücret (YK Dahil)", sheet["D3"].value)
                self.assertEqual("İşveren Maliyeti (Teşvikli)", sheet["U3"].value)
                self.assertEqual("Çocuk Parası", sheet["V3"].value)
                self.assertIn("Kontrol gerekli - Excel satırı: 5", sheet["W3"].value)
                self.assertEqual("F4CCCC", sheet["W3"].fill.fgColor.rgb[-6:])
                self.assertEqual(40.0, sheet["W5"].value)
                self.assertEqual("n", sheet["W5"].data_type)
                self.assertEqual("#,##0.00", sheet["W5"].number_format)
                self.assertEqual("=SUM(W4:W5)", sheet["W6"].value)
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
