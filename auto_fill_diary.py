# auto_fill_diary.py
from openpyxl import Workbook
from fastapi import HTTPException
import re

def _find_target_sheet(wb):
    """
    シート検出の例：
    - 現在の月から推定（ここでは簡易： '日誌' を含む先頭シート）
    - 本番はあなたのルール（例：'日誌\s*0?{M}月' の正規表現→なければフォールバック）を実装
    """
    # まずは '日誌' を含む最も左のシート
    for name in wb.sheetnames:
        if "日誌" in name:
            return wb[name]
    # 見つからなければ一番左
    return wb[wb.sheetnames[0]]

def process_workbook(wb: Workbook) -> None:
    """
    ここに “V/X/Y/AA 列の自動入力” などのロジックを書く。
    いまは動作確認のため ZZ1 に『処理OK』を書くだけ。
    """
    ws = _find_target_sheet(wb)

    # --- ここから本番ロジックを書いていく ---------------------------------
    # 例：
    # 1) 今日の月のシートを選ぶ
    # 2) 休日判定（U列=50 の行は V/X/Y/AA を空欄のまま）
    # 3) 既存値があるセルは上書き禁止、空欄のみ補完
    # 4) 参照優先：前月→同シート上方既知値
    # 5) 実習内容コード（1〜6）は 4→3→2→5→6 の順で配分
    # 6) 3日以上の同カテゴリ連続を回避
    #
    #  ※ 実シートの列位置・行範囲はテンプレート仕様に合わせて調整
    # ---------------------------------------------------------------------

    # ★動作確認の軽い書き込み（不要なら削除OK）
    ws["ZZ1"].value = "処理OK"

    # 例）もし保護されていて書けない場合はエラー
    # try:
    #     ws["ZZ1"].value = "処理OK"
    # except Exception as e:
    #     raise HTTPException(status_code=400, detail=f"シートに書き込めません: {e}")

    # ここで return は不要。呼び出し側が wb を保存して返却します。

