import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest
from excel_automation.pdf_po_parser import PDFPOData, PDFPOParser


class TestPDFPOData:

    def test_create_pdf_po_data(self):
        data = PDFPOData(
            raw_po="0009013330-1",
            po_number="9013330",
            color_code="3104",
            size_quantities={"046": 60, "048": 140},
            total_quantity=200,
            source_file="Test.pdf"
        )
        assert data.po_number == "9013330"
        assert data.color_code == "3104"
        assert data.size_quantities["046"] == 60
        assert data.total_quantity == 200

    def test_default_values(self):
        data = PDFPOData(raw_po="", po_number="", color_code="")
        assert data.size_quantities == {}
        assert data.total_quantity == 0
        assert data.source_file == ""
        assert data.ordertotal_from_pdf is None
        assert data.quantity_mismatch is False


class TestPOExtraction:

    def test_extract_po_standard(self):
        text = "P. O. No. Your reference\n0009013330-1 Marina Scholander"
        result = PDFPOParser._extract_po_number(text)
        assert result == ("0009013330-1", "9013330")

    def test_extract_po_strip_leading_zeros(self):
        text = "P. O. No. Your reference\n0009013330-1 Someone"
        raw, cleaned = PDFPOParser._extract_po_number(text)
        assert raw == "0009013330-1"
        assert cleaned == "9013330"

    def test_extract_po_no_leading_zeros(self):
        text = "P. O. No. Your reference\n9013330-2 Someone"
        raw, cleaned = PDFPOParser._extract_po_number(text)
        assert raw == "9013330-2"
        assert cleaned == "9013330"

    def test_extract_po_not_found(self):
        text = "This text has no PO number"
        with pytest.raises(RuntimeError, match="Không tìm thấy PO Number"):
            PDFPOParser._extract_po_number(text)


class TestColorExtraction:

    def test_extract_color_from_article_no(self):
        text = "000010 62183104046 AW Stretch Trousers 60 20.290 1217.40 USD"
        result = PDFPOParser._extract_color_code(text)
        assert result == "3104"

    def test_extract_color_different_article(self):
        text = "000010 62189999046 Some Product 60 20.290 1217.40 USD"
        result = PDFPOParser._extract_color_code(text)
        assert result == "9999"

    def test_extract_color_not_found(self):
        text = "No article numbers here"
        with pytest.raises(RuntimeError, match="Không tìm thấy Article Number"):
            PDFPOParser._extract_color_code(text)


class TestSizeQuantityExtraction:

    def test_extract_single_size_qty(self):
        text = "000010 62183104046 AW Stretch Trousers 60 20.290 1217.40 USD\nSize:46"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"046": 60}

    def test_extract_multiple_sizes(self):
        text = (
            "000010 62183104046 AW Stretch Trousers 60 20.290 1217.40 USD\nSize:46\n"
            "000020 62183104048 AW Stretch Trousers 140 20.290 2840.60 USD\nSize:48\n"
            "000030 62183104050 AW Stretch Trousers 200 20.290 4058.00 USD\nSize:50"
        )
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"046": 60, "048": 140, "050": 200}

    def test_normalize_size_below_100(self):
        text = "000010 62183104096 AW Stretch Trousers 20 20.290 405.80 USD\nSize:96"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"096": 20}

    def test_normalize_size_100_and_above(self):
        text = "000010 62183104100 AW Stretch Trousers 20 20.290 405.80 USD\nSize:100"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"100": 20}

    def test_normalize_size_large(self):
        text = "000010 62183104148 AW Stretch Trousers 20 20.290 405.80 USD\nSize:148"
        result = PDFPOParser._extract_size_quantities(text)
        assert result == {"148": 20}

    def test_no_sizes_found(self):
        text = "No sizes here"
        with pytest.raises(RuntimeError, match="Không tìm thấy dữ liệu Size"):
            PDFPOParser._extract_size_quantities(text)


from pathlib import Path as PDFPath


class TestOrdertotalExtraction:

    def test_extract_ordertotal_standard(self):
        text = "Ordertotal 1030 20898.70 USD\n"
        result = PDFPOParser._extract_ordertotal(text)
        assert result == 1030

    def test_extract_ordertotal_not_found(self):
        text = "Some random text without ordertotal\n"
        result = PDFPOParser._extract_ordertotal(text)
        assert result is None

    def test_extract_ordertotal_different_number(self):
        text = "Ordertotal 500 10000.00 USD\n"
        result = PDFPOParser._extract_ordertotal(text)
        assert result == 500


class TestQuantityMismatchDetection:

    def test_no_mismatch_when_ordertotal_matches_sum(self):
        data = PDFPOData(
            raw_po="0009013330-1",
            po_number="9013330",
            color_code="3104",
            size_quantities={"046": 60, "048": 140},
            total_quantity=200,
            source_file="Test.pdf",
            ordertotal_from_pdf=200,
            quantity_mismatch=False
        )
        assert data.quantity_mismatch is False

    def test_mismatch_flag_true_when_differ(self):
        data = PDFPOData(
            raw_po="0009013330-1",
            po_number="9013330",
            color_code="3104",
            size_quantities={"046": 60, "048": 140},
            total_quantity=200,
            source_file="Test.pdf",
            ordertotal_from_pdf=250,
            quantity_mismatch=True
        )
        assert data.quantity_mismatch is True

    def test_no_mismatch_when_ordertotal_is_none(self):
        data = PDFPOData(
            raw_po="0009013330-1",
            po_number="9013330",
            color_code="3104",
            size_quantities={"046": 60, "048": 140},
            total_quantity=200,
            source_file="Test.pdf",
            ordertotal_from_pdf=None,
            quantity_mismatch=False
        )
        assert data.ordertotal_from_pdf is None
        assert data.quantity_mismatch is False


class TestFullParse:

    def test_parse_test_pdf(self):
        pdf_path = PDFPath(__file__).parent.parent / "Test.pdf"
        if not pdf_path.exists():
            pytest.skip("Test.pdf không tồn tại")

        result = PDFPOParser.parse(str(pdf_path))

        assert isinstance(result, PDFPOData)
        assert result.raw_po == "0009013330-1"
        assert result.po_number == "9013330"
        assert result.color_code == "3104"
        assert result.total_quantity == 1030
        assert result.size_quantities["046"] == 60
        assert result.size_quantities["048"] == 140
        assert result.size_quantities["050"] == 200
        assert result.size_quantities["052"] == 200
        assert result.size_quantities["054"] == 160
        assert result.size_quantities["056"] == 100
        assert result.size_quantities["058"] == 20
        assert result.size_quantities["096"] == 20
        assert result.size_quantities["100"] == 20
        assert result.size_quantities["104"] == 20
        assert result.size_quantities["108"] == 20
        assert result.size_quantities["120"] == 10
        assert result.size_quantities["148"] == 20
        assert result.size_quantities["150"] == 20
        assert result.size_quantities["152"] == 20
        assert len(result.size_quantities) == 15
        assert result.source_file == str(pdf_path)
        assert result.ordertotal_from_pdf == 1030
        assert result.quantity_mismatch is False

    def test_parse_nonexistent_file(self):
        with pytest.raises(RuntimeError, match="Không thể đọc file PDF"):
            PDFPOParser.parse("nonexistent.pdf")
