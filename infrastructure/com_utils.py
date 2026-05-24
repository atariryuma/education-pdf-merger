"""
COM初期化ユーティリティ

スレッド単位でCOMを初期化・解放するコンテキストマネージャ。
Office/Excelの自動化やpywin32経由の操作で利用する。

Note:
    tkinterと併用するときは GUI スレッドが STA で初期化されている必要があるため、
    バックグラウンドスレッドで COM を扱う場合はこのモジュールを使う。
"""
import logging
from contextlib import contextmanager
from typing import Generator

import pythoncom

logger = logging.getLogger(__name__)


@contextmanager
def com_apartment(sta: bool = True) -> Generator[None, None, None]:
    """
    スレッド単位で COM を初期化し、終了時に解放するコンテキストマネージャ

    Args:
        sta: True なら Single-Threaded Apartment (STA)、False なら MTA で初期化する。
             tkinter のバックグラウンドスレッドやExcel/Word操作は通常 STA を使う。
    """
    if sta:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    else:
        pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            # 既に解放済み等は無視（多重Uninitializeは例外を投げる場合がある）
            logger.debug("CoUninitialize で例外（無視）: %s", e)
