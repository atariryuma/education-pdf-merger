"""
年度計算ユーティリティ

教育計画は次年度のものを作成する運用に対応した年度計算機能を提供
"""
import datetime
from typing import Tuple


def calculate_next_fiscal_year() -> Tuple[str, str]:
    """
    次年度の年度（西暦・和暦）を自動計算

    教育計画は次年度のものを作成するため、現在の月に応じて自動判定:
    - 1月～3月: 現在年度のまま（例: 2026年2月 → 2026年度＝令和8年度）
    - 4月～12月: 翌年度（例: 2025年4月 → 2026年度＝令和8年度）

    Returns:
        Tuple[str, str]: (西暦4桁の年度, 和暦短縮形) 例: ("2026", "R8")

    Examples:
        >>> # 2025年4月の場合
        >>> calculate_next_fiscal_year()
        ('2026', 'R8')  # 次年度

        >>> # 2026年2月の場合
        >>> calculate_next_fiscal_year()
        ('2026', 'R8')  # 現年度
    """
    now = datetime.datetime.now()
    current_year = now.year
    current_month = now.month

    # 教育計画の年度判定（4月始まり）
    if current_month >= 4:
        # 4月以降 → 翌年度の計画を作成
        fiscal_year = current_year + 1
    else:
        # 1月～3月 → 現年度の計画を作成
        fiscal_year = current_year

    year = str(fiscal_year)
    year_short = calculate_year_short(year)

    return year, year_short


def calculate_year_short(year: str) -> str:
    """
    西暦から和暦短縮形を計算

    Args:
        year: 西暦4桁の文字列（例: "2026"）

    Returns:
        str: 和暦短縮形（例: "R8"）

    Examples:
        >>> calculate_year_short("2019")
        'R1'
        >>> calculate_year_short("2025")
        'R7'
        >>> calculate_year_short("2026")
        'R8'
    """
    try:
        year_int = int(year)

        # 令和元年は2019年
        if year_int >= 2019:
            reiwa_year = year_int - 2018
            return f"R{reiwa_year}"
        # 平成31年（2019年4月30日まで）は考慮しない
        # 平成（1989年～2019年4月）
        elif year_int >= 1989:
            heisei_year = year_int - 1988
            return f"H{heisei_year}"
        # 昭和（1926年～1989年1月）
        elif year_int >= 1926:
            showa_year = year_int - 1925
            return f"S{showa_year}"
        else:
            # それ以前はそのまま返す
            return year

    except ValueError:
        # 変換失敗時はデフォルト値
        return "R8"
