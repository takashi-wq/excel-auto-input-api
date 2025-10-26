# app.py (debug-temp)
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
import os, io, tempfile, shutil, logging, traceback
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("app")

app = FastAPI()

@app.get("/")
def health():
    return {"status": "ok"}

@app.post("/process")
async def process(file: UploadFile = File(...), password: str = Form(None)):
    # 認証（Render の環境変数 PASSWORD と一致）
    expected = os.getenv("PASSWORD")
    if expected and password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    # ---- まずはファイルを確実に受け取る ----
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            shutil.copyfileobj(file.file, tmp)
            tmp_path = tmp.name
        size = os.path.getsize(tmp_path)
        log.info(f"uploaded: name={getattr(file, 'filename', None)} size={size} path={tmp_path}")
    except Exception as e:
        log.error("upload write error", exc_info=True)
        raise HTTPException(status_code=400, detail=f"upload write error: {type(e).__name__}")

    try:
        # ---- ECHO_ONLY=1 の時は openpyxl を通さず そのまま返す（切り分け用）----
        if os.getenv("ECHO_ONLY") == "1":
            with open(tmp_path, "rb") as f:
                data = f.read()
            os.remove(tmp_path)
            name = getattr(file, "filename", None) or "uploaded.xlsx"
            return StreamingResponse(
                io.BytesIO(data),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f'attachment; filename="{name}"'}
            )

        # ---- ここから本来の処理：openpyxl でロード ----
        try:
            wb = load_workbook(filename=tmp_path, data_only=False)
        except (BadZipFile, InvalidFileException, KeyError) as e:
            log.warning("xlsx load error (known)", exc_info=True)
            raise HTTPException(status_code=400, detail=f"xlsx load error: {type(e).__name__}")
        except Exception as e:
            log.error("load_workbook error", exc_info=True)
            raise HTTPException(status_code=400, detail=f"load_workbook error: {type(e).__name__}")
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

        # TODO: ここで wb をルールに従って更新する
        # ws = wb.active など

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        name = getattr(file, "filename", None) or "updated.xlsx"
        return StreamingResponse(
            bio,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{name}"'}
        )

    except HTTPException:
        raise
    except Exception as e:
        # 最後の砦：トレースバックをログし、500 ではなく 400 で詳細を返す
        tb = traceback.format_exc()
        log.error("unhandled error", exc_info=True)
        return JSONResponse(status_code=400, content={"detail": f"unhandled: {type(e).__name__}", "trace": tb})
