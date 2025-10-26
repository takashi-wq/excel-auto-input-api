# app.py
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import os, io, tempfile
import openpyxl

# Uvicorn が参照する ASGI アプリ本体（この名前が必須）
app = FastAPI()

@app.get("/")
def health():
    return {"status": "ok"}

@app.post("/process")
async def process(file: UploadFile = File(...), password: str = Form(None)):
    expected = os.getenv("PASSWORD")
    if expected and password != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")

    content = await file.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    wb = openpyxl.load_workbook(tmp_path)
    # 将来ここで本処理を呼ぶ想定：
    # from auto_fill_diary import process_workbook
    # wb = process_workbook(wb)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{file.filename or "updated.xlsx"}"'}
    )

