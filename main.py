"""
PDF to Word/Text 변환 웹 앱 - FastAPI 메인 애플리케이션
특허 문서 분석용
"""
import os
import uuid
import asyncio
import shutil
import zipfile
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, Dict, List
from contextlib import asynccontextmanager
from io import BytesIO

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request, BackgroundTasks
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from pdf_processor import (
    convert_pdf,
    sanitize_filename,
    generate_file_id,
    PDFProcessingError,
    EncryptedPDFError,
    CorruptedPDFError
)


# 상수
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
ALLOWED_EXTENSIONS = {".pdf"}
TEMP_DIR = Path("temp")
FILE_EXPIRY_MINUTES = 10

# 파일 저장소 (메모리)
# 구조: {file_id: {"filename": str, "text": str, "docx": bytes, "txt": bytes, "created_at": datetime, "session_id": str}}
file_storage: Dict[str, dict] = {}

# 세션 저장소
# 구조: {session_id: [file_id, ...]}
session_storage: Dict[str, List[str]] = {}


async def cleanup_old_files():
    """오래된 파일 정리 (10분 이상)"""
    while True:
        await asyncio.sleep(60)  # 1분마다 체크
        now = datetime.now()
        expired_ids = []

        for file_id, data in file_storage.items():
            if now - data["created_at"] > timedelta(minutes=FILE_EXPIRY_MINUTES):
                expired_ids.append(file_id)

        for file_id in expired_ids:
            session_id = file_storage[file_id].get("session_id")
            del file_storage[file_id]

            # 세션에서도 제거
            if session_id and session_id in session_storage:
                if file_id in session_storage[session_id]:
                    session_storage[session_id].remove(file_id)
                if not session_storage[session_id]:
                    del session_storage[session_id]

        # temp 디렉토리 정리
        for temp_file in TEMP_DIR.glob("*"):
            if temp_file.name != ".gitkeep":
                try:
                    if temp_file.is_file():
                        file_age = datetime.now() - datetime.fromtimestamp(temp_file.stat().st_mtime)
                        if file_age > timedelta(minutes=FILE_EXPIRY_MINUTES):
                            temp_file.unlink()
                except Exception:
                    pass


@asynccontextmanager
async def lifespan(app: FastAPI):
    """앱 시작/종료 시 실행"""
    # 시작
    TEMP_DIR.mkdir(exist_ok=True)
    cleanup_task = asyncio.create_task(cleanup_old_files())
    yield
    # 종료
    cleanup_task.cancel()
    try:
        await cleanup_task
    except asyncio.CancelledError:
        pass


app = FastAPI(
    title="PDF to Word/Text 변환기",
    description="특허 문서 PDF를 Word 또는 텍스트 파일로 변환합니다.",
    version="1.0.0",
    lifespan=lifespan
)

# 정적 파일 및 템플릿 설정
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """메인 페이지"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/health")
async def health_check():
    """헬스 체크"""
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}


@app.post("/upload")
async def upload_files(
    files: List[UploadFile] = File(...),
    output_format: str = Form("both"),
    ocr_language: str = Form("auto")
):
    """
    PDF 파일 업로드 및 변환

    Args:
        files: PDF 파일 목록
        output_format: 출력 형식 (docx, txt, both)
        ocr_language: OCR 언어 (kor, eng, chi_sim, chi_tra, jpn, auto)

    Returns:
        변환 결과 정보
    """
    if not files:
        raise HTTPException(status_code=400, detail="파일을 선택해주세요.")

    session_id = str(uuid.uuid4())[:12]
    results = []
    errors = []

    for file in files:
        # 파일 검증
        if not file.filename:
            errors.append({"filename": "unknown", "error": "파일명이 없습니다."})
            continue

        # 확장자 검사
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            errors.append({"filename": file.filename, "error": "PDF 파일만 업로드 가능합니다."})
            continue

        # 파일 크기 검사
        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            errors.append({"filename": file.filename, "error": "50MB 이하 파일만 업로드 가능합니다."})
            continue

        # 임시 파일 저장
        file_id = generate_file_id()
        temp_path = TEMP_DIR / f"{file_id}.pdf"

        try:
            with open(temp_path, "wb") as f:
                f.write(content)

            # PDF 변환
            result = await convert_pdf(
                str(temp_path),
                file.filename,
                output_format,
                ocr_language
            )

            # 결과 저장
            safe_filename = sanitize_filename(file.filename)
            base_name = os.path.splitext(safe_filename)[0]

            storage_data = {
                "filename": base_name,
                "original_filename": file.filename,
                "text": result["text"],
                "total_pages": result["total_pages"],
                "docx": result["docx_buffer"].read() if result["docx_buffer"] else None,
                "txt": result["txt_buffer"].read() if result["txt_buffer"] else None,
                "created_at": datetime.now(),
                "session_id": session_id
            }
            file_storage[file_id] = storage_data

            # 세션에 추가
            if session_id not in session_storage:
                session_storage[session_id] = []
            session_storage[session_id].append(file_id)

            results.append({
                "file_id": file_id,
                "filename": file.filename,
                "safe_filename": base_name,
                "total_pages": result["total_pages"],
                "text_length": len(result["text"]),
                "has_docx": result["docx_buffer"] is not None,
                "has_txt": result["txt_buffer"] is not None
            })

        except EncryptedPDFError as e:
            errors.append({"filename": file.filename, "error": str(e)})
        except CorruptedPDFError as e:
            errors.append({"filename": file.filename, "error": str(e)})
        except PDFProcessingError as e:
            errors.append({"filename": file.filename, "error": str(e)})
        except Exception as e:
            errors.append({"filename": file.filename, "error": f"처리 중 오류가 발생했습니다: {str(e)}"})
        finally:
            # 임시 파일 삭제
            if temp_path.exists():
                temp_path.unlink()

    if not results and errors:
        raise HTTPException(status_code=400, detail={"errors": errors})

    return {
        "session_id": session_id,
        "results": results,
        "errors": errors,
        "total_files": len(results),
        "total_errors": len(errors)
    }


@app.get("/preview/{file_id}")
async def preview_file(file_id: str, full: bool = False):
    """
    변환된 텍스트 미리보기

    Args:
        file_id: 파일 ID
        full: 전체 텍스트 반환 여부 (기본값: False, 5000자까지만)
    """
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다. 파일이 만료되었을 수 있습니다.")

    data = file_storage[file_id]
    text = data["text"]

    if not full and len(text) > 5000:
        preview_text = text[:5000]
        has_more = True
    else:
        preview_text = text
        has_more = False

    return {
        "file_id": file_id,
        "filename": data["original_filename"],
        "text": preview_text,
        "total_length": len(text),
        "has_more": has_more,
        "total_pages": data["total_pages"]
    }


@app.get("/download/{file_id}")
async def download_file(file_id: str, format: str = "docx"):
    """
    개별 파일 다운로드

    Args:
        file_id: 파일 ID
        format: 파일 형식 (docx 또는 txt)
    """
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다. 파일이 만료되었을 수 있습니다.")

    data = file_storage[file_id]
    filename = data["filename"]

    if format == "docx":
        if not data["docx"]:
            raise HTTPException(status_code=404, detail="Word 파일이 생성되지 않았습니다.")
        content = data["docx"]
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        extension = ".docx"
    elif format == "txt":
        if not data["txt"]:
            raise HTTPException(status_code=404, detail="텍스트 파일이 생성되지 않았습니다.")
        content = data["txt"]
        media_type = "text/plain; charset=utf-8"
        extension = ".txt"
    else:
        raise HTTPException(status_code=400, detail="지원하지 않는 형식입니다. (docx 또는 txt)")

    return StreamingResponse(
        BytesIO(content),
        media_type=media_type,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{filename}{extension}"
        }
    )


@app.get("/download-all/{session_id}")
async def download_all(session_id: str):
    """
    세션의 모든 파일 ZIP으로 다운로드
    """
    if session_id not in session_storage or not session_storage[session_id]:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다. 파일이 만료되었을 수 있습니다.")

    file_ids = session_storage[session_id]

    # ZIP 파일 생성
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_id in file_ids:
            if file_id in file_storage:
                data = file_storage[file_id]
                filename = data["filename"]

                if data["docx"]:
                    zip_file.writestr(f"{filename}.docx", data["docx"])
                if data["txt"]:
                    zip_file.writestr(f"{filename}.txt", data["txt"])

    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f"attachment; filename=converted_files_{session_id}.zip"
        }
    )


@app.get("/file-info/{file_id}")
async def get_file_info(file_id: str):
    """파일 정보 조회"""
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")

    data = file_storage[file_id]
    created_at = data["created_at"]
    expires_at = created_at + timedelta(minutes=FILE_EXPIRY_MINUTES)
    remaining_seconds = max(0, (expires_at - datetime.now()).total_seconds())

    return {
        "file_id": file_id,
        "filename": data["original_filename"],
        "total_pages": data["total_pages"],
        "text_length": len(data["text"]),
        "has_docx": data["docx"] is not None,
        "has_txt": data["txt"] is not None,
        "expires_in_seconds": int(remaining_seconds)
    }


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
