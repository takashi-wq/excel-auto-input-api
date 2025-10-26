# app.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import os, io, tempfile, urllib.parse

from auto_fill_diary import process_workbook

app = FastAPI(title="Excel Auto Input API")

# CORS（必要に応じて調整）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def health():
    return {"status": "ok"}

@app.post("/process")
async def process(
    file: UploadFile = File(...),
    password: str = Form(None),
):
    # パスワードチェック（Render側 環境変数 UPLOAD_PASSWORD）
    expected = os.getenv("PASSWORD") or os.getenv("UPLOAD_PASSWORD")
    if expected and password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # アップロードファイルを一時保存
    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    try:
        content = await file.read()
        tmp_in.write(content)
        tmp_in.flush()
        tmp_in.close()

        # 変換処理（ここでテンプレートに必要な編集を実行）
        out_path, download_name = process_workbook(
            tmp_in.name,
            source_filename=file.filename  # ★ 受け渡し
        )

        # 日本語ファイル名で返す（RFC 5987 / 6266）
        quoted = urllib.parse.quote(download_name)
        # 英字フォールバック（日本語を含む場合は無難な ascii 名に）
        ascii_fallback = "updated.xlsx" if any(ord(c) > 127 for c in download_name) else download_name

        headers = {
            "Content-Disposition": f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{quoted}"
        }

        with open(out_path, "rb") as f:
            data = f.read()
        return StreamingResponse(
            io.BytesIO(data),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
    except HTTPException:
        raise
    except Exception as e:
        # 例外内容をクライアントにも出す（デバッグ後はログにのみ出力へ）
        raise HTTPException(status_code=500, detail=f"Processing failed: {e}")
    finally:
        # 入力一時ファイル削除
        try:
            os.remove(tmp_in.name)
        except Exception:
            pass
