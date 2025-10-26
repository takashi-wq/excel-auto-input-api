# app.py
from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Tuple
from urllib.parse import quote

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from openpyxl import load_workbook
from zipfile import BadZipFile

from auto_fill_diary import process_workbook  # ← ここで中身の加工を行う

app = FastAPI(title="Excel Auto Input API", version="1.0.0")


def _get_password_from_env() -> str:
    """
    Render/Workers どちらでも拾えるように、よく使われるキー名を順に探す。
    例: PASSWORD / UPLOAD_PASSWORD / API_PASSWORD
    """
    for key in ("PASSWORD", "UPLOAD_PASSWORD", "API_PASSWORD"):
        val = os.getenv(key)
        if val:
            return val
    # 何も設定されていない場合のデフォルト（明示的に 5124 を使いたいケース用）
    return "5124"


def _build_content_disposition(filename: str) -> str:
    """
    日本語ファイル名を含む Content-Disposition を安全に生成。
    - ASCII 版 filename は Windows 等との互換性のために「可能なら」落とし、
      ダメそうならプレースホルダ名にする
    - UTF-8 版は filename*=UTF-8''<percent-encoded> を付ける
    """
    # 不正なパス文字を除去
    safe = filename.replace("\\", "_").replace("/", "_").replace("\n", "_").replace("\r", "_")
    # ASCII だけの見かけのファイル名（fallback）
    try:
        ascii_name = safe.encode("ascii", "strict").decode("ascii")
    except UnicodeError:
        # どうしても ASCII にならない場合のフォールバック
        stem, _, ext = safe.rpartition(".")
        if not ext:
            ext = "xlsx"
        ascii_name = "download." + ext

    utf8_name = quote(safe, safe="")  # RFC 5987 形式にパーセントエンコード
    # ダブルクォートを避けるためにエスケープ
    ascii_name = ascii_name.replace('"', "'")

    return f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{utf8_name}'


@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/process")
async def process(
    file: UploadFile = File(..., description="xlsx ファイル"),
    password: str = Form(..., description="アップロード用パスワード"),
):
    # パスワードチェック
    expected = _get_password_from_env()
    if password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # 拡張子の軽いバリデーション（必須ではないが明示的に）
    original_name = file.filename or "upload.xlsx"
    if not original_name.lower().endswith(".xlsx"):
        raise HTTPException(status_code=422, detail="Only .xlsx is supported")

    # ファイル読込 → openpyxl
    try:
        raw = await file.read()
        if not raw:
            raise HTTPException(status_code=400, detail="Empty file")

        bio_in = io.BytesIO(raw)
        wb = load_workbook(bio_in, data_only=False)  # 数式は data_only では評価されないが編集は可能
    except BadZipFile:
        raise HTTPException(status_code=400, detail="The file is not a valid .xlsx")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to open workbook: {e}")

    # ここで帳票の自動入力ロジックを実行
    try:
        # 返り値として（必要なら）出力ファイル名を上書きできる
        new_name = process_workbook(wb, source_filename=original_name)
        if new_name:
            original_name = new_name
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing failed: {e}")

    # 書き出し（ストリーミングで返却）
    bio_out = io.BytesIO()
    wb.save(bio_out)
    bio_out.seek(0)

    headers = {
        "Content-Disposition": _build_content_disposition(original_name),
        # ダウンロードの明示
        "X-Download-Options": "noopen",
    }

    return StreamingResponse(
        bio_out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
