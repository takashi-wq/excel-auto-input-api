# app.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import os, io, tempfile
from openpyxl import load_workbook

# ★ 本処理をインポート
from auto_fill_diary import process_workbook

app = FastAPI()

@app.get("/")
def health():
    return {"status": "ok"}

@app.post("/process")
async def process(file: UploadFile = File(...), password: str = Form(None)):
    # パスワードチェック
    expected = os.getenv("PASSWORD")
    if expected and password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # ファイルを一時保存して openpyxl で開く
    content = await file.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path, data_only=False)  # 書式・数式は openpyxl の仕様で保持
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid xlsx: {e}")

    # ★ ここで本処理を実行（この関数の中身にルールを書いていく）
    try:
        process_workbook(wb)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing error: {e}")

    # 返却用ストリームに保存
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    # ダウンロード時のファイル名（RFC 5987 で日本語も安全）
    filename = "updated.xlsx"
    disposition = (
        f"attachment; filename*=UTF-8''{filename}"
    )

    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": disposition},
    )
