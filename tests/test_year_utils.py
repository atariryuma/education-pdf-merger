"""
year_utils のユニットテスト

年度計算・和暦変換をテスト
"""
import pytest
from unittest.mock import patch
import datetime

from infrastructure.year_utils import calculate_next_fiscal_year, calculate_year_short


@pytest.mark.unit
class TestCalculateYearShort:
    """calculate_year_short のテスト"""

    def test_reiwa_2019(self):
        """令和元年（2019）"""
        assert calculate_year_short("2019") == "R1"

    def test_reiwa_2025(self):
        """令和7年（2025）"""
        assert calculate_year_short("2025") == "R7"

    def test_reiwa_2026(self):
        """令和8年（2026）"""
        assert calculate_year_short("2026") == "R8"

    def test_reiwa_2030(self):
        """令和12年（2030）"""
        assert calculate_year_short("2030") == "R12"

    def test_heisei_1989(self):
        """平成元年（1989）"""
        assert calculate_year_short("1989") == "H1"

    def test_heisei_2018(self):
        """平成30年（2018）"""
        assert calculate_year_short("2018") == "H30"

    def test_showa_1926(self):
        """昭和元年（1926）"""
        assert calculate_year_short("1926") == "S1"

    def test_showa_1988(self):
        """昭和63年（1988）"""
        assert calculate_year_short("1988") == "S63"

    def test_pre_showa(self):
        """昭和以前はそのまま"""
        assert calculate_year_short("1925") == "1925"

    def test_invalid_string(self):
        """無効な文字列はValueError"""
        with pytest.raises(ValueError):
            calculate_year_short("invalid")

    def test_empty_string(self):
        """空文字列はValueError"""
        with pytest.raises(ValueError):
            calculate_year_short("")


@pytest.mark.unit
class TestCalculateNextFiscalYear:
    """calculate_next_fiscal_year のテスト"""

    @patch("infrastructure.year_utils.datetime")
    def test_april_returns_next_year(self, mock_dt):
        """4月は翌年度"""
        mock_dt.datetime.now.return_value = datetime.datetime(2025, 4, 1)
        year, year_short = calculate_next_fiscal_year()
        assert year == "2026"
        assert year_short == "R8"

    @patch("infrastructure.year_utils.datetime")
    def test_december_returns_next_year(self, mock_dt):
        """12月は翌年度"""
        mock_dt.datetime.now.return_value = datetime.datetime(2025, 12, 15)
        year, year_short = calculate_next_fiscal_year()
        assert year == "2026"
        assert year_short == "R8"

    @patch("infrastructure.year_utils.datetime")
    def test_january_returns_current_year(self, mock_dt):
        """1月は現年度"""
        mock_dt.datetime.now.return_value = datetime.datetime(2026, 1, 10)
        year, year_short = calculate_next_fiscal_year()
        assert year == "2026"
        assert year_short == "R8"

    @patch("infrastructure.year_utils.datetime")
    def test_march_returns_current_year(self, mock_dt):
        """3月は現年度"""
        mock_dt.datetime.now.return_value = datetime.datetime(2026, 3, 31)
        year, year_short = calculate_next_fiscal_year()
        assert year == "2026"
        assert year_short == "R8"

    @patch("infrastructure.year_utils.datetime")
    def test_boundary_april_first(self, mock_dt):
        """4月1日の境界値"""
        mock_dt.datetime.now.return_value = datetime.datetime(2025, 4, 1)
        year, _ = calculate_next_fiscal_year()
        assert year == "2026"

    @patch("infrastructure.year_utils.datetime")
    def test_boundary_march_31(self, mock_dt):
        """3月31日の境界値"""
        mock_dt.datetime.now.return_value = datetime.datetime(2025, 3, 31)
        year, _ = calculate_next_fiscal_year()
        assert year == "2025"
