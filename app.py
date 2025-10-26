# app.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import os, io, tempfile, shutil
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

app = FastAPI()

@app.get("/")
def health():
    return {"status": "ok"}

@app.post("/process")
async def process(file: UploadFile = File(...), password: str = Form(None)):
    # 認証（Render の環境変数 PASSWORD と一致必須。未設定ならスキップ）
    expected = os.getenv("PASSWORD")
    if expected and password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # 受け取り → 一時ファイルへ安全にストリームコピー
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            shutil.copyfileobj(file.file, tmp)         # ← これが一番堅い
            tmp_path = tmp.name
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"upload write error: {type(e).__name__}")

    # XLSX 読み込み（典型的な失敗を 400 で返す）
    try:
        wb = load_workbook(filename=tmp_path, data_only=False)
    except (BadZipFile, InvalidFileException, KeyError) as e:
        # xlsxでない/壊れている/ZIP破損など
        raise HTTPException(status_code=400, detail=f"xlsx load error: {type(e).__name__}")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"load_workbook error: {type(e).__name__}")
    finally:
        # 一時ファイルはもう不要
        try:
            os.remove(tmp_path)
        except Exception:
            pass

    # TODO: ここで wb をルールに従って更新する
    # 例: ws = wb.active など

    # バイトとして返却
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    # ダウンロードさせる
    filename = getattr(file, "filename", None) or "updated.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
