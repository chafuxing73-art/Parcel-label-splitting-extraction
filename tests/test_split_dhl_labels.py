import pytest
import re
from split_dhl_labels import (
    detect_layout,
    extract_master_no,
    is_fedex,
)


class TestDetectLayout:
    def test_grid_layout(self):
        assert detect_layout(600, 800) == "grid"
        assert detect_layout(501, 400) == "grid"
        assert detect_layout(400, 501) == "grid"

    def test_single_layout(self):
        assert detect_layout(400, 400) == "single"
        assert detect_layout(500, 500) == "single"
        assert detect_layout(300, 450) == "single"


class TestExtractMasterNo:
    def test_extract_j_number(self):
        text = "J123456789012"
        assert extract_master_no(text) == "J123456789012"

    def test_extract_als_number(self):
        text = "ALS12345678901"
        assert extract_master_no(text) == "ALS12345678901"

    def test_extract_case_insensitive(self):
        text = "als12345678901"
        assert extract_master_no(text) == "ALS12345678901"

    def test_no_match(self):
        text = "No tracking number here"
        assert extract_master_no(text) is None

    def test_multiple_numbers(self):
        text = "J123456789012 and ALS12345678901"
        assert extract_master_no(text) == "J123456789012"


class TestIsFedex:
    def test_fedex_format(self):
        assert is_fedex("1234 5678 9012") is True
        assert is_fedex("1234 5678 9012/5678 1234 5678") is True

    def test_dhl_format(self):
        assert is_fedex("(00)123456789012345678") is False

    def test_empty_string(self):
        assert is_fedex("") is False

    def test_none(self):
        assert is_fedex(None) is False
