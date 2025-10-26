# app.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
import os
import io
import tempfile
import re
from urllib.parse import quote

from openpyxl import load_workbook

# もし独自の処理モジュールを分けたい場合はここで import
# from auto_fill_diary import process_workbook  # <- 使うならコメント解除

app = FastAPI(title="Excel Auto Input API")

# ===== ユーティリティ =====
def get_password() -> str:
    # Render/Workers 環境変数（設定: UPLOAD_PASSWORD=5124）
    # なければ 5124 をデフォルトとして扱う
    return os.getenv("UPLOAD_PASSWORD", "5124")


def make_content_disposition(original_name: str) -> str:
    """
    日本語ファイル名でもエラーにならない Content-Disposition を生成。
    - filename= は ASCII の安全名
    - filename*= は RFC5987 (UTF-8 + URL エンコード)
    """
    base = os.path.splitext(os.path.basename(original_name or "updated"))[0]
    safe_base = re.sub(r"[^A-Za-z0-9_.-]", "_", base)[:60]  # 非ASCIIを'_'に、長すぎるのもカット
    fname = f"{safe_base}.xlsx"
    # RFC5987 形式（日本語名でもOK）
    return f'attachment; filename="{fname}"; filename*=UTF-8\'\'{quote(fname)}'


def identity_process_workbook(wb):
    """
    ひとまず何もしない（ワークブックを開いて閉じるだけ）の安全なダミー処理。
    実処理を作り込む場合は auto_fill_diary.py などに切り出して呼び替えてください。
    """
    return wb


# ===== エンドポイント =====
@app.get("/")
def health():
    return {"status": "ok"}


@app.post("/process")
async def process(file: UploadFile = File(...), password: str = Form(None)):
    # パスワードチェック
    expected = get_password()
    if password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # 一時ファイルに保存して openpyxl で開く
    try:
        content = await file.read()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        # Excel をロード
        wb = load_workbook(tmp_path)

        # ▼処理本体（現状はダミー処理）
        # wb = process_workbook(wb)  # ←独自処理を使うならこちらに変更
        wb = identity_process_workbook(wb)

        # バイナリに書き出して返す
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

        headers = {"Content-Disposition": make_content_disposition(file.filename or "updated.xlsx")}
        return StreamingResponse(
            bio,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except HTTPException:
        # 上で投げたものはそのまま
        raise
    except Exception as e:
        # 例外は 500 で返す（ログに出したい場合は print/ログ基盤へ）
        return JSONResponse(status_code=500, content={"detail": f"Internal Server Error: {type(e).__name__}"})
    finally:
        # 一時ファイルの掃除
        try:
            if "tmp_path" in locals() and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
