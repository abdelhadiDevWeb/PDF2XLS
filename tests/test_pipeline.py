"""Run with:  .venv\\Scripts\\python -m unittest discover -s tests -v
Requires fpdf2 (dev only) to generate sample PDFs."""
from __future__ import annotations

import io
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import converter  # noqa: E402
from app import app  # noqa: E402

ROWS = [
    ["Product", "Region", "Units", "Revenue"],
    ["Apples", "North", "1200", "3,450.50"],
    ["Bananas", "South", "850", "1,200.00"],
    ["Cherries", "East", "430", "2,980.75"],
    ["Dates", "West", "99", "(150.25)"],
]


def _table_page(pdf, bordered: bool) -> None:
    pdf.add_page()
    pdf.set_font("Helvetica", size=16)
    pdf.cell(0, 12, "Quarterly sales report", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=12)
    pdf.ln(6)
    for row in ROWS:
        for value in row:
            pdf.cell(42, 10, value, border=1 if bordered else 0)
        pdf.ln(10)


def make_text_pdf(path: Path, bordered: bool = True) -> None:
    from fpdf import FPDF

    pdf = FPDF()
    _table_page(pdf, bordered)
    pdf.output(str(path))


INVOICE = [
    ["Code", "Item", "Qty", "Price"],
    ["007", "Bolt M6", "1,200", "$3,450.50"],
    ["010", "Nut M6", "25", "(150.25)"],
    ["010", "Nut M6", "25", "(150.25)"],
    ["112", "Washer", "5", "12.5%"],
]


def make_invoice_pdf(path: Path, bordered: bool = True) -> None:
    """Title + paragraph + table (with a repeated row) + footnote, then a text-only page."""
    from fpdf import FPDF

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=16)
    pdf.cell(0, 12, "Invoice INV-0042", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 6, "Thank you for your order. Below are the items shipped this month.")
    pdf.ln(4)
    for row in INVOICE:
        for value in row:
            pdf.cell(42, 9, value, border=1 if bordered else 0)
        pdf.ln(9)
    pdf.ln(4)
    pdf.cell(0, 6, "* Prices include VAT.", new_x="LMARGIN", new_y="NEXT")
    pdf.add_page()
    pdf.multi_cell(0, 6, "Terms and conditions. Payment is due within 30 days of the invoice date. "
                         "Late payments may incur a fee of 2 percent per month on the balance.")
    pdf.output(str(path))


PROFORMA_HEADER = ["N°", "REF", "Désignation", "Qty", "PU HT"]
PROFORMA_ITEMS = [
    ["1", "30308", ["BIOMERIEUX VIDAS ANTI", "HCV 60 TESTS"], "1", "33 852.32"],
    ["2", "", ["BIOMERIEUX VIDAS TSH 3", "60 TESTS"], "1", "28 792.06"],
    ["3", "30315", ["BIOMERIEUX VIDAS HBS", "AG ULTRA 60 TESTS"], "2", "29 787.70"],
    ["4", "417011", ["BIOMERIEUX VIDAS AMH", "30 TESTS"], "1", "56 773.22"],
    ["5", "30431", ["BIOMERIEUX VIDAS", "ESTRADIOL II 60 TESTS"], "1", "29 106.11"],
]


def make_proforma_pdf(path: Path) -> None:
    """Invoice laid out like a real proforma: letterhead, rows separated by
    horizontal rules drawn as one segment per column (no vertical lines),
    wrapped cells, page 2 continuing without a rule above its first row,
    then amount in words and a totals box."""
    from fpdf import FPDF

    xs = [10, 22, 42, 102, 122, 160]
    pdf = FPDF()
    pdf.set_line_width(0.2)

    def rule(y):
        for a, b in zip(xs, xs[1:]):
            pdf.line(a, y, b, y)

    def item_row(y, item):
        number, ref, lines, qty, price = item
        for x, value in ((xs[0], number), (xs[1], ref), (xs[3], qty), (xs[4], price)):
            pdf.text(x + 1, y + 6, value)
        for i, line in enumerate(lines):
            pdf.text(xs[2] + 1, y + 4 + i * 4, line)

    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.text(10, 15, "SARL TECHNICLAB NORD")
    pdf.set_font("Helvetica", size=9)
    for i, (left, right) in enumerate([("Tel: 023-85-01-58", "RC: 19 B 1047389"),
                                       ("Fax: 023-85-01-58", "NIF: 001916104738992")]):
        pdf.text(10, 22 + i * 5, left)
        pdf.text(120, 22 + i * 5, right)
    pdf.set_font("Helvetica", "B", 11)
    pdf.text(10, 40, "FACTURE PROFORMA FP21/6942")
    pdf.line(10, 42, 160, 42)
    pdf.set_font("Helvetica", size=8)
    pdf.text(10, 52, "ALGER le: 28-12-25")
    pdf.text(60, 52, "LABORATOIRE D'ANALYSES DE BIOLOGIE MEDICALE")
    y = 56
    rule(y)
    for x, title in zip(xs, PROFORMA_HEADER):
        pdf.text(x + 1, y + 5, title)
    y += 7
    rule(y)
    for item in PROFORMA_ITEMS[:3]:
        item_row(y, item)
        y += 10
        rule(y)
    pdf.text(10, 285, "FACTURE PROFORMA FP21/6942 du 28-12-25    Page 1/2")

    pdf.add_page()
    pdf.text(120, 12, "=>Montant HT: 92 431.08")
    y = 16
    for i, item in enumerate(PROFORMA_ITEMS[3:]):
        item_row(y, item)
        y += 10
        rule(y)
    pdf.text(10, y + 5, "RIST: 0.00 Nb.U: 5")
    pdf.text(10, y + 10, "Arrete le present Facture proforma a la somme de: CENT QUATRE-VINGT")
    pdf.rect(110, y + 3, 50, 18)
    for i, (label, value) in enumerate([("Montant HT", "178 311.41"), ("TVA Total", "1 698.01"),
                                        ("NET A PAYER", "180 009.42")]):
        pdf.text(111, y + 7 + i * 6, label)
        pdf.text(140, y + 7 + i * 6, value)
        pdf.line(110, y + 9 + i * 6, 138, y + 9 + i * 6)
        pdf.line(138, y + 9 + i * 6, 160, y + 9 + i * 6)
    pdf.output(str(path))


def sheet_names(path: Path) -> list[str]:
    with pd.ExcelFile(path) as book:
        return book.sheet_names


def sheet_cells(path: Path) -> list[list]:
    import openpyxl

    sheet = openpyxl.load_workbook(path).active
    return [[(c.value, c.number_format) for c in row] for row in sheet.iter_rows()]


def make_scanned_pdf(source: Path, path: Path) -> None:
    import pypdfium2 as pdfium

    doc = pdfium.PdfDocument(str(source))
    image = doc[0].render(scale=200 / 72).to_pil().convert("RGB")
    doc.close()
    image.save(path, "PDF", resolution=200)


def make_mixed_pdf(text_pdf: Path, scanned_pdf: Path, path: Path) -> None:
    import pypdfium2 as pdfium

    out = pdfium.PdfDocument.new()
    for src in (text_pdf, scanned_pdf):
        doc = pdfium.PdfDocument(str(src))
        out.import_pages(doc)
        doc.close()
    out.save(str(path))
    out.close()


class PipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        d = Path(cls.tmp.name)
        cls.bordered = d / "bordered.pdf"
        cls.borderless = d / "borderless.pdf"
        cls.scanned = d / "scanned.pdf"
        cls.mixed = d / "mixed.pdf"
        make_text_pdf(cls.bordered, bordered=True)
        make_text_pdf(cls.borderless, bordered=False)
        make_scanned_pdf(cls.bordered, cls.scanned)
        make_mixed_pdf(cls.bordered, cls.scanned, cls.mixed)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _convert(self, pdf: Path):
        out = Path(self.tmp.name) / (pdf.stem + ".xlsx")
        result = converter.convert_pdf(pdf, out)
        self.assertTrue(out.exists())
        return result, out

    def _assert_sales_table(self, df: pd.DataFrame):
        """Cells are exactly the text printed in the PDF table."""
        self.assertEqual(list(df.columns), ROWS[0])
        self.assertEqual(df.values.tolist(), ROWS[1:])

    def _assert_sales_sheet(self, out: Path, repeat: int = 1):
        """Excel holds real numbers whose display matches the PDF text."""
        expected = [
            [("Apples", "General"), ("North", "General"), (1200, "0"), (3450.5, "#,##0.00")],
            [("Bananas", "General"), ("South", "General"), (850, "0"), (1200.0, "#,##0.00")],
            [("Cherries", "General"), ("East", "General"), (430, "0"), (2980.75, "#,##0.00")],
            [("Dates", "General"), ("West", "General"), (99, "0"), (-150.25, "0.00;(0.00)")],
        ]
        cells = sheet_cells(out)
        self.assertEqual([v for v, _ in cells[0]], ROWS[0])
        self.assertEqual(cells[1:], expected * repeat)

    def test_classify_pages(self):
        self.assertEqual(converter.classify_pages(self.bordered), {1: True})
        self.assertEqual(converter.classify_pages(self.scanned), {1: False})
        self.assertEqual(converter.classify_pages(self.mixed), {1: True, 2: False})

    def test_bordered_text_pdf_uses_camelot_lattice(self):
        result, out = self._convert(self.bordered)
        self.assertEqual(len(result.tables), 1)
        self.assertEqual(result.tables[0].method, "camelot-lattice")
        self._assert_sales_table(result.tables[0].df)
        self.assertEqual(sheet_names(out), ["Tables"])
        self._assert_sales_sheet(out)

    def test_borderless_text_pdf_falls_back(self):
        result, _ = self._convert(self.borderless)
        self.assertGreaterEqual(len(result.tables), 1)
        self.assertIn(result.tables[0].method, {"pdfplumber", "camelot-stream"})
        self._assert_sales_table(result.tables[0].df)

    @unittest.skipUnless(converter.TESSERACT_AVAILABLE, "Tesseract not installed")
    def test_scanned_pdf_uses_ocr(self):
        result, _ = self._convert(self.scanned)
        self.assertEqual(result.pages[0].has_text, False)
        self.assertTrue(result.tables)
        self.assertEqual(result.tables[0].method, "ocr-table")
        self._assert_sales_table(result.tables[0].df)

    @unittest.skipUnless(converter.TESSERACT_AVAILABLE, "Tesseract not installed")
    def test_mixed_pdf_routes_each_page(self):
        result, out = self._convert(self.mixed)
        methods = {p.page: p.method for p in result.pages}
        self.assertEqual(methods[1], "camelot-lattice")
        self.assertTrue(methods[2].startswith("ocr"))

        # Same table on page 1 and 2 -> one continuous table on one sheet.
        self.assertEqual(len(result.blocks), 1)
        self.assertEqual(result.blocks[0].pages, [1, 2])
        self.assertEqual(sheet_names(out), ["Tables"])
        self._assert_sales_sheet(out, repeat=2)

    def test_only_table_data_exactly_as_written(self):
        for bordered in (True, False):
            with self.subTest(bordered=bordered):
                pdf = Path(self.tmp.name) / f"invoice_{bordered}.pdf"
                make_invoice_pdf(pdf, bordered)
                result, out = self._convert(pdf)
                self.assertEqual(len(result.blocks), 1)
                df = result.blocks[0].df
                # no title, paragraph, footnote or terms page; repeated row kept
                self.assertEqual(list(df.columns), INVOICE[0])
                self.assertEqual(df.values.tolist(), INVOICE[1:])
                self.assertEqual({p.page: p.tables for p in result.pages}, {1: 1, 2: 0})

                cells = sheet_cells(out)
                self.assertEqual(len(cells), 5)
                self.assertEqual(cells[1], [("007", "General"), ("Bolt M6", "General"),
                                            (1200, "#,##0"), (3450.5, '"$"#,##0.00')])
                self.assertEqual(cells[2], [("010", "General"), ("Nut M6", "General"),
                                            (25, "0"), (-150.25, "0.00;(0.00)")])
                self.assertEqual(cells[2], cells[3])
                self.assertEqual(cells[4][3], (0.125, "0.0%"))

    def test_ruled_rows_invoice_gives_only_the_table(self):
        pdf = Path(self.tmp.name) / "proforma.pdf"
        make_proforma_pdf(pdf)
        result, out = self._convert(pdf)

        self.assertEqual([p.method for p in result.pages], ["ruled-rows", "ruled-rows"])
        self.assertEqual(len(result.blocks), 1)
        self.assertEqual(result.blocks[0].pages, [1, 2])
        expected = [[n, ref or None, " ".join(lines), qty, price] for n, ref, lines, qty, price in PROFORMA_ITEMS]
        df = result.blocks[0].df
        self.assertEqual(list(df.columns), PROFORMA_HEADER)
        self.assertEqual(df.astype(object).where(df.notna(), None).values.tolist(), expected)

        rows = [[v for v, _ in row] for row in sheet_cells(out)]
        self.assertEqual(rows[0], PROFORMA_HEADER)
        self.assertEqual(len(rows), 1 + len(PROFORMA_ITEMS))  # no letterhead, totals or footer
        text = str(rows)
        for outside in ("TECHNICLAB", "ALGER", "LABORATOIRE", "Montant", "RIST", "NET A PAYER", "Page 1/2"):
            self.assertNotIn(outside, text)

    def test_all_tables_option_keeps_totals_box(self):
        pdf = Path(self.tmp.name) / "proforma_all.pdf"
        make_proforma_pdf(pdf)
        out = Path(self.tmp.name) / "proforma_all.xlsx"
        result = converter.convert_pdf(pdf, out, main_table_only=False)
        self.assertEqual(len(result.blocks), 2)
        self.assertIn("NET A PAYER", str(result.blocks[1].df.values.tolist()))

    def test_clean_table_keeps_cells_as_written(self):
        raw = pd.DataFrame(
            [
                ["", "", "", ""],
                [" Code ", "Name\nof item", "Qty", ""],
                ["007", "Widget", "1,000", None],
                ["010", "Gadget  pro", "25", ""],
                ["010", "Gadget  pro", "25", ""],
            ]
        )
        df = converter.clean_table(raw)
        self.assertEqual(list(df.columns), ["Code", "Name of item", "Qty"])
        self.assertEqual(df.values.tolist(), [
            ["007", "Widget", "1,000"],
            ["010", "Gadget pro", "25"],
            ["010", "Gadget pro", "25"],  # repeated rows are real data
        ])

    def test_clean_table_without_header(self):
        df = converter.clean_table(pd.DataFrame([["1", "2"], ["3", "4"]]))
        self.assertFalse(converter.has_header(df))
        self.assertEqual(df.values.tolist(), [["1", "2"], ["3", "4"]])

    def test_borderless_drops_title_and_footnote(self):
        raw = pd.DataFrame([["Report title", "", ""], ["A", "B", "C"], ["x", "1", "2"], ["* note", "", ""]])
        df = converter.clean_table(raw, borderless=True)
        self.assertEqual(list(df.columns), ["A", "B", "C"])
        self.assertEqual(df.values.tolist(), [["x", "1", "2"]])

    def test_bordered_keeps_merged_row_inside_grid(self):
        raw = pd.DataFrame([["A", "B", "C"], ["x", "1", "2"], ["Total incl. VAT", "", ""]])
        df = converter.clean_table(raw)
        self.assertEqual(df.values.tolist()[-1][0], "Total incl. VAT")

    def test_excel_number(self):
        cases = {
            "1200": (1200, "0"),
            "1,200": (1200, "#,##0"),
            "3,450.50": (3450.5, "#,##0.00"),
            "0.50": (0.5, "0.00"),
            "-7": (-7, "0"),
            "(150.25)": (-150.25, "0.00;(0.00)"),
            "$3,450.50": (3450.5, '"$"#,##0.00'),
            "€ 12": (12, '"€"0'),
            "12.5%": (0.125, "0.0%"),
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(converter.excel_number(text), expected)
        for text in ("007", "1234567890123456", "12 34", "INV-42", "(5", "1.2.3", "(-5)", "(0)", "2024-01-31"):
            with self.subTest(text=text):
                self.assertIsNone(converter.excel_number(text))

    def test_empty_table(self):
        self.assertTrue(converter.clean_table(pd.DataFrame([["", None]])).empty)


class OneSheetTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name) / "out.xlsx"

    def tearDown(self):
        self.tmp.cleanup()

    @staticmethod
    def _t(page, rows, columns):
        return converter.TableResult(page, "test", pd.DataFrame(rows, columns=columns))

    def test_headerless_continuation_is_joined(self):
        blocks = converter.merge_continued_tables([
            self._t(1, [["a", 1]], ["Name", "Qty"]),
            self._t(2, [["b", 2], ["c", 3]], ["Column 1", "Column 2"]),
        ])
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0].pages, [1, 2])
        self.assertEqual(list(blocks[0].df.columns), ["Name", "Qty"])
        self.assertEqual(blocks[0].df["Name"].tolist(), ["a", "b", "c"])

    def test_different_tables_stacked_on_one_sheet(self):
        blocks = converter.merge_continued_tables([
            self._t(1, [["a", 1], ["b", 2]], ["Name", "Qty"]),
            self._t(1, [["x", "y", "z"]], ["A", "B", "C"]),
            self._t(3, [["c", 3]], ["Name", "Qty"]),  # page gap -> not a continuation
        ])
        self.assertEqual([b.pages for b in blocks], [[1], [1], [3]])

        converter.write_excel(blocks, self.out)
        sheets = pd.read_excel(self.out, sheet_name=None, header=None)
        self.assertEqual(list(sheets), ["Tables"])
        rows = sheets["Tables"].astype(object).where(sheets["Tables"].notna(), None).values.tolist()
        self.assertEqual(rows, [
            ["Name", "Qty", None], ["a", 1, None], ["b", 2, None],
            [None, None, None],
            ["A", "B", "C"], ["x", "y", "z"],
            [None, None, None],
            ["Name", "Qty", None], ["c", 3, None],
        ])
        self.assertEqual([b.start_row for b in blocks], [1, 5, 8])

    def test_table_without_header_gets_no_invented_header(self):
        blocks = converter.merge_continued_tables([
            converter.TableResult(1, "test", converter.clean_table(pd.DataFrame([["1", "2"], ["3", "4"]]))),
        ])
        converter.write_excel(blocks, self.out)
        self.assertEqual([[v for v, _ in row] for row in sheet_cells(self.out)], [[1, 2], [3, 4]])

    def test_no_tables_still_writes_one_sheet(self):
        converter.write_excel([], self.out)
        sheets = pd.read_excel(self.out, sheet_name=None)
        self.assertEqual(list(sheets), ["Tables"])


class WebTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.pdf = Path(cls.tmp.name) / "report.pdf"
        make_text_pdf(cls.pdf)
        cls.client = app.test_client()

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_index(self):
        res = self.client.get("/")
        self.assertEqual(res.status_code, 200)
        self.assertIn(b"Convert PDF tables to Excel", res.data)
        self.assertIn(b"const MAX_BYTES = 50 * 1024 * 1024;", res.data)

    def test_upload_and_download(self):
        res = self.client.post(
            "/api/convert",
            data={"file": (io.BytesIO(self.pdf.read_bytes()), "My Report.pdf")},
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 200, res.get_json())
        body = res.get_json()
        self.assertEqual(body["filename"], "My_Report.xlsx")
        self.assertEqual(body["sheet"], "Tables")
        self.assertEqual(body["blocks"][0]["preview"]["columns"], ROWS[0])
        self.assertEqual(body["blocks"][0]["start_row"], 1)
        self.assertTrue(body["blocks"][0]["has_header"])
        self.assertEqual(body["blocks"][0]["preview"]["rows"], ROWS[1:])

        dl = self.client.get(body["download_url"])
        self.assertEqual(dl.status_code, 200)
        self.assertIn("My_Report.xlsx", dl.headers["Content-Disposition"])
        sheets = pd.read_excel(io.BytesIO(dl.data), sheet_name=None)
        dl.close()
        self.assertEqual(list(sheets), ["Tables"])

    def test_rejects_non_pdf(self):
        res = self.client.post(
            "/api/convert",
            data={"file": (io.BytesIO(b"hello"), "notes.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 400)

    def test_rejects_fake_pdf(self):
        res = self.client.post(
            "/api/convert",
            data={"file": (io.BytesIO(b"not really a pdf"), "fake.pdf")},
            content_type="multipart/form-data",
        )
        self.assertEqual(res.status_code, 400)

    def test_healthz(self):
        res = self.client.get("/healthz")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.get_json()["status"], "ok")

    def test_optional_password(self):
        import base64
        from unittest import mock

        def basic(password):
            return {"Authorization": "Basic " + base64.b64encode(f"user:{password}".encode()).decode()}

        with mock.patch("app.APP_PASSWORD", "s3cret"):
            self.assertEqual(self.client.get("/").status_code, 401)
            self.assertEqual(self.client.get("/", headers=basic("wrong")).status_code, 401)
            self.assertEqual(self.client.get("/", headers=basic("s3cret")).status_code, 200)
            self.assertEqual(self.client.post("/api/convert").status_code, 401)
            self.assertEqual(self.client.get("/healthz").status_code, 200)

    def test_download_unknown_job(self):
        self.assertEqual(self.client.get("/api/download/" + "0" * 32).status_code, 404)
        self.assertEqual(self.client.get("/api/download/..%2F..%2Fsecret").status_code, 404)


if __name__ == "__main__":
    unittest.main()
