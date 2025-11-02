
# app.py (with /inspect endpoint for debugging)
from fastapi import FastAPI, UploadFile, File, Response
from tempfile import NamedTemporaryFile
from pathlib import Path
import uvicorn
import json
from auto_fill_diary import process

FIXED_MSG = "処理に失敗しました。もう一度ファイルをアップロードしてください。"

app = FastAPI()

@app.post("/process")
async def process_xlsx(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        return Response(FIXED_MSG, status_code=400, media_type="text/plain; charset=utf-8")
    try:
        with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = Path(tmp.name)

        modified, logs = process(str(tmp_path))

        # Minimal server-side logs (stdout)
        print(json.dumps({"level":"INFO","event":"process_done","logs":logs}, ensure_ascii=False))

        if modified == 0:
            # Zero-update guard => return fixed message with 422 (Worker向けの既定動作)
            return Response(FIXED_MSG, status_code=422, media_type="text/plain; charset=utf-8")

        out = tmp_path.read_bytes()
        return Response(
            content=out,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="processed.xlsx"'}
        )
    except Exception as e:
        print(json.dumps({"level":"ERROR","event":"process_error","msg":str(e)}, ensure_ascii=False))
        return Response(FIXED_MSG, status_code=500, media_type="text/plain; charset=utf-8")
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass

@app.post("/inspect")
async def inspect_xlsx(file: UploadFile = File(...)):
    """
    Debug endpoint: runs the same analysis but always returns JSON with counters.
    Does not return the Excel. Intended for direct debugging, not for Worker.
    """
    if not file.filename.lower().endswith(".xlsx"):
        return Response(
            json.dumps({"error":"xlsx_only"} , ensure_ascii=False),
            status_code=400, media_type="application/json; charset=utf-8"
        )
    try:
        with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = Path(tmp.name)

        modified, logs = process(str(tmp_path))
        resp = {"modified_count": modified, **logs}
        return Response(json.dumps(resp, ensure_ascii=False), media_type="application/json; charset=utf-8")
    except Exception as e:
        return Response(
            json.dumps({"error":"exception","message":str(e)}, ensure_ascii=False),
            status_code=500, media_type="application/json; charset=utf-8"
        )
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8080)
