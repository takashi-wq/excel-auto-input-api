# app.py
import os
import io
import re
from datetime import datetime
from urllib.parse import quote

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
from typing import Optional, List

app = FastAPI(title="Excel Auto Input API")

# === ユーティリティ ===
def sanitize_filename(name: str) -> str:
    """
    Windows/Excel が苦手な文字を除去。拡張子は維持。
    """
    s = name.replace("\\", "_").replace("/", "_").replace(":", "_") \
            .replace("*", "_").replace("?", "_").replace('"', "_") \
            .replace("<", "_").replace(">", "_").replace("|", "_")
    # 長すぎ対応（任意）
    if len(s) > 150:
        root, ext = os.path.splitext(s)
        s = root[:140] + ext
    return s


def pick_target_sheet(wb) -> str:
    """
    実際に書き込むシートを選ぶ。
    - 候補に合致すればそれを、無ければアクティブシート。
    """
    candidates: List[str] = [
        "実習日誌", "実習 日誌", "日誌", "Sheet1", "Sheet", "シート1"
    ]
    sheetnames = wb.sheetnames

    # 完全一致優先
    for cand in candidates:
        if cand in sheetnames:
            return cand

    # 部分一致（例: '日誌（10月）' など）
    for cand in candidates:
        for s in sheetnames:
            if cand in s:
                return s

    # フォールバック
    return wb.active.title


def add_api_result_sheet(wb, original_name: str):
    """
    変更の痕跡が必ず残るよう、API_RESULT シートを（無ければ）追加して
    簡単なメタ情報を書き込む。
    """
    sheet_name = "API_RESULT"
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "processed", original_name])
    else:
        ws = wb.create_sheet(sheet_name)
        ws["A1"] = "timestamp"
        ws["B1"] = "status"
        ws["C1"] = "source_filename"
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "processed", original_name])


def process_workbook(in_bytes: bytes, original_filename: str) -> bytes:
    """
    ここに本番ロジックを実装していく。まずは確実に変更が入る形。
    """
    # 読み込み
    bio = io.BytesIO(in_bytes)
    wb = load_workbook(bio, data_only=False)

    # 変更が目視できるように、必ず1つのセルに印を書き込む
    target_sheet = pick_target_sheet(wb)
    ws = wb[target_sheet]

    # 目立つ位置にスタンプ（必要なら座標は調整してください）
    ws["B2"] = f"✅ Processed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

    # 変更履歴シートも作る/追記する
    add_api_result_sheet(wb, original_filename)

    # --- ここから下に “本命の書き込み処理” を追加していけばOK ---
    # 例:
    # ws["E7"] = "秋元 寛子"     # 指導者氏名（例）
    # ws["C12"] = "パターン記号の見かた指導"  # 指導内容（例）
    # ※ 実際のセル番地に合わせて書き換えてください。

    # 保存して戻す
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


def content_disposition_for(filename: str) -> str:
    """
    日本語やスペースを含むファイル名に対応した Content-Disposition を作る。
    例: attachment; filename="updated.xlsx"; filename*=UTF-8''%E6%9B%B...
    """
    safe = sanitize_filename(filename)
    ascii_fallback = re.sub(r"[^\x20-\x7E]", "_", safe)  # ASCII 以外を _
    quoted = quote(safe)
    return f'attachment; filename="{ascii_fallback}"; filename*=UTF-8\'\'{quoted}'


# === エンドポイント ===
@app.post("/process")
async def process(
    file: UploadFile = File(...),
    password: Optional[str] = Form(None)
):
    # 認証（Render の環境変数 UPLOAD_PASSWORD と一致必須）
    expected = os.getenv("UPLOAD_PASSWORD")
    if expected and password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # 拡張子チェック（xlsx のみ）
    if not (file.filename.lower().endswith(".xlsx")):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported")

    # 元ファイル名
    original_name = file.filename or "uploaded.xlsx"

    # バイト取得
    in_bytes = await file.read()

    try:
        out_bytes = process_workbook(in_bytes, original_name)
    except Exception as e:
        # 失敗したら理由を返す（ログも併用してください）
        raise HTTPException(status_code=500, detail=f"Processing failed: {e}")

    # ダウンロード用ファイル名
    root, ext = os.path.splitext(sanitize_filename(original_name))
    updated_name = f"{root}_updated{ext}"

    headers = {
        "Content-Disposition": content_disposition_for(updated_name)
    }
    return StreamingResponse(
        io.BytesIO(out_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
