# auto_fill_diary.py
from __future__ import annotations

from datetime import datetime
from typing import Optional

from openpyxl.workbook.workbook import Workbook


def process_workbook(wb: Workbook, *, source_filename: str) -> Optional[str]:
    """
    ここに「エクセル自動入力」の中身を書きます。
    返値でファイル名を差し替えたい場合は新しいファイル名（str）を返す。
    差し替え不要なら None を返す。

    ---- サンプル実装（最小） ----
    - 最初のシートの A1 に処理日時を入れるだけ
    - 既存値がある場合は上書きしたくない等の条件は自由に追加できます
    """
    ws = wb.worksheets[0]  # 最初のワークシート
    ws["A1"].value = f"Processed at {datetime.now():%Y-%m-%d %H:%M:%S}"

    # 返却ファイル名を「元名の末尾に _processed」を付ける例
    # （日本語名でもOK。Content-Disposition 側で UTF-8 エンコードします）
    if source_filename.lower().endswith(".xlsx"):
        return source_filename[:-5] + "_processed.xlsx"
    return None
