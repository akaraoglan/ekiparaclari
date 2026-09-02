import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
from openpyxl import Workbook

from tools.beyanname_pdf_doldur import (
    annotate_pdf,
    format_tr_amount,
    format_tr_quantity,
    load_import_rows,
    _extract_declaration,
)


class BeyannamePdfDoldurTest(unittest.TestCase):
    def test_formats_and_annotates_sorted_rows_with_invoice_totals(self):
        folder = Path(__file__).resolve().parents[1] / "outputs" / "beyanname_pdf_unit_test"
        folder.mkdir(parents=True, exist_ok=True)
        excel_path = folder / "data.xlsx"
        pdf_path = folder / "source.pdf"
        output_path = folder / "output.pdf"
        self._write_excel(excel_path)
        self._write_pdf(pdf_path)

        grouped = load_import_rows(str(excel_path))
        matched, warnings = annotate_pdf(str(pdf_path), str(output_path), grouped)

        self.assertEqual(1, matched)
        self.assertEqual([], warnings)
        self.assertEqual(
            ["2026002414", "2026002415", "2026002416"],
            [row.delivery for row in grouped["26161700IM00000505"]],
        )

        result = fitz.open(output_path)
        try:
            text = "\n".join(page.get_text("text") for page in result)
        finally:
            result.close()

        self.assertIn("GENEL TOPLAM: 10.860,00 EUR", text)
        self.assertIn("FATURA 30000017: 7.740,00 EUR", text)
        self.assertIn("FATURA 30000018: 3.120,00 EUR", text)
        self.assertIn("1.200", text)
        self.assertIn("950", text)
        self.assertIn("55.666,55 TL", text)
        self.assertLess(text.index("2026002414"), text.index("2026002415"))
        self.assertLess(text.index("2026002415"), text.index("2026002416"))

    def test_turkish_amount_format(self):
        from decimal import Decimal

        self.assertEqual("3.120,00", format_tr_amount(Decimal("3120")))
        self.assertEqual("3.258,50", format_tr_amount(Decimal("3258.5")))
        self.assertEqual("1.200", format_tr_quantity(Decimal("1200")))
        self.assertEqual("1.200,5", format_tr_quantity(Decimal("1200.50")))

    def test_scanned_page_uses_bundled_ocr_fallback(self):
        document = fitz.open()
        page = document.new_page()
        known = {"26161700IM00000505"}
        with patch(
            "tools.beyanname_pdf_doldur._rapidocr_text",
            return_value="GUMRUK BEYANNAME NO 26161700IM00000505",
        ):
            self.assertEqual("26161700IM00000505", _extract_declaration(page, known))
        document.close()

    def test_scanned_page_retries_at_higher_resolution(self):
        document = fitz.open()
        page = document.new_page()
        known = {"26161700IM00000505"}
        with patch(
            "tools.beyanname_pdf_doldur._rapidocr_text",
            side_effect=["okunamadi", "26161700IM00000505"],
        ) as rapidocr:
            self.assertEqual("26161700IM00000505", _extract_declaration(page, known))
            self.assertEqual([2.0, 3.0], [call.kwargs["scale"] for call in rapidocr.call_args_list])
        document.close()

    def test_linux_ocr_dependency_error_is_not_hidden(self):
        document = fitz.open()
        page = document.new_page()
        with patch(
            "tools.beyanname_pdf_doldur._rapidocr_text",
            side_effect=ImportError("libGL.so.1: cannot open shared object file"),
        ), patch.object(
            fitz.Page,
            "get_textpage_ocr",
            side_effect=RuntimeError("Tesseract is not installed"),
        ):
            with self.assertRaisesRegex(ValueError, "libGL1 eksik"):
                _extract_declaration(page, {"26161700IM00000505"})
        document.close()

    def test_annotate_pdf_preserves_ocr_dependency_error(self):
        folder = Path(__file__).resolve().parents[1] / "outputs" / "beyanname_pdf_unit_test"
        folder.mkdir(parents=True, exist_ok=True)
        source_path = folder / "ocr_error_source.pdf"
        output_path = folder / "ocr_error_output.pdf"
        self._write_pdf(source_path)

        with patch(
            "tools.beyanname_pdf_doldur._extract_declaration",
            side_effect=ValueError("RapidOCR başlatılamadı: Linux sunucuda libGL1 eksik."),
        ):
            with self.assertRaisesRegex(ValueError, "libGL1 eksik"):
                annotate_pdf(
                    str(source_path),
                    str(output_path),
                    {"26161700IM00000505": []},
                )

    @staticmethod
    def _write_excel(path):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet.append(
            [
                "Teslimat",
                "Belge tarihi",
                "Gümrük beyanname tarihi",
                "Gümrük beyanname numarası",
                "Intrastat grubu",
                "Intrastat grubu",
                "Para birimi",
                "Tsl.mkt.",
                "Satınalma blg",
                "Fatura",
                None,
            ]
        )
        worksheet.append([2026002416, None, None, "26161700IM00000505", 4620, None, "EUR", 1200, 5500004393, 30000017, 15058.88])
        worksheet.append([2026002414, None, None, "26161700IM00000505", 3120, None, "EUR", 950, 5500004393, 30000018, 55666.55])
        worksheet.append([2026002415, None, None, "26161700IM00000505", 3120, None, "EUR", 2000, 5500004393, 30000017, 99655.66])
        workbook.save(path)

    @staticmethod
    def _write_pdf(path):
        document = fitz.open()
        page = document.new_page(width=595, height=841)
        page.insert_text((80, 200), "GUMRUK BEYANNAME NO: 26161700IM00000505")
        document.save(path)
        document.close()


if __name__ == "__main__":
    unittest.main()
