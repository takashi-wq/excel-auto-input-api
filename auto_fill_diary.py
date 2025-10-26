# auto_fill_diary.py
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
import tempfile
from datetime import datetime

def _make_download_name(source_filename: str | None) -> str:
    """
    返却用ファイル名を生成。
    元名が .xlsx ならその手前に _updated を挿入、なければ末尾に付与。
    日本語名もそのまま返す（送信側で filename*= でエンコード）
    """
    base = source_filename or "output.xlsx"
    if base.lower().endswith(".xlsx"):
        return base[:-5] + "_updated.xlsx"
    return base + "_updated.xlsx"

def process_workbook(
    src_path: str,
    *,
    source_filename: str | None = None,   # ★ app.py から渡される
) -> tuple[str, str]:
    """
    入力:  src_path（一時保存されたアップロード xlsx）
    出力: (out_path, download_name)

    out_path は API が読み出してクライアントへ返す
    download_name は Content-Disposition 用のダウンロード名
    """
    # 読み込み（データ検証拡張の Warning は仕様上出ることがあります）
    wb = load_workbook(filename=src_path, data_only=False)
    ws = wb.active

    # ---- ここから「実処理」：テンプレートに合わせて置き換えてください ----
    # いまは処理が走った目印として B1 にスタンプするだけ
    try:
        stamp = f"自動処理済み {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        # 既存の値を極力壊さないように空セルに書く（B1が空ならB1、埋まってたらC1）
        target_cell = "B1" if (ws["B1"].value in (None, "")) else "C1"
        ws[target_cell] = stamp
    except Exception:
        # テンプレートによっては保護などで書けない場合があるので握りつぶす
        pass
    # ---- 実処理ここまで ----

    # 出力一時ファイルに保存
    out_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    out_tmp.close()
    wb.save(out_tmp.name)

    # 返却名
    download_name = _make_download_name(source_filename)

    return out_tmp.name, download_name
