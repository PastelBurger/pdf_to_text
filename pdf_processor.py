"""
PDF 처리 모듈 - PDF에서 텍스트 추출, OCR, Word 변환
"""
import os
import re
import uuid
import asyncio
from pathlib import Path
from typing import Optional, Tuple, Callable
from io import BytesIO

import pdfplumber
from PyPDF2 import PdfReader
from PyPDF2.errors import PdfReadError
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pytesseract
from pdf2image import convert_from_path
from PIL import Image


# 지원 언어 매핑
OCR_LANGUAGES = {
    "kor": "kor",
    "eng": "eng",
    "auto": "kor+eng"  # 자동 감지 시 한국어+영어 사용
}

# 페이지 구분자
PAGE_SEPARATOR = "\n\n--- 페이지 {page_num} ---\n\n"

# 임시 파일 디렉토리
TEMP_DIR = Path("temp")
TEMP_DIR.mkdir(exist_ok=True)


def clean_text(text: str) -> str:
    """
    추출된 텍스트 정리
    - 연속 줄바꿈 제거
    - 문장 중간 줄바꿈 연결
    """
    # 페이지 구분자 임시 보존
    page_markers = re.findall(r'--- 페이지 \d+ ---', text)
    text = re.sub(r'--- 페이지 (\d+) ---', r'<<<PAGE_\1>>>', text)

    # 연속 줄바꿈을 단일 줄바꿈으로
    text = re.sub(r'\n{3,}', '\n\n', text)

    # 문장 중간 줄바꿈 연결 (문장 종결 부호가 아닌 경우)
    # 단, 페이지 마커 주변은 유지
    lines = text.split('\n')
    cleaned_lines = []

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            cleaned_lines.append('')
            continue

        # 페이지 마커는 그대로 유지
        if re.match(r'<<<PAGE_\d+>>>', line):
            cleaned_lines.append(line)
            continue

        # 이전 줄과 연결 가능한지 확인
        if cleaned_lines and cleaned_lines[-1]:
            prev_line = cleaned_lines[-1]
            # 이전 줄이 페이지 마커가 아니고, 문장 종결 부호로 끝나지 않으면 연결
            if not re.match(r'<<<PAGE_\d+>>>', prev_line):
                if not re.search(r'[.!?。！？:：;\n]$', prev_line):
                    cleaned_lines[-1] = prev_line + ' ' + line
                    continue

        cleaned_lines.append(line)

    text = '\n'.join(cleaned_lines)

    # 페이지 마커 복원
    text = re.sub(r'<<<PAGE_(\d+)>>>', r'--- 페이지 \1 ---', text)

    # 연속 공백 정리
    text = re.sub(r' +', ' ', text)

    return text.strip()


class PDFProcessingError(Exception):
    """PDF 처리 관련 에러"""
    pass


class EncryptedPDFError(PDFProcessingError):
    """암호화된 PDF 에러"""
    pass


class CorruptedPDFError(PDFProcessingError):
    """손상된 PDF 에러"""
    pass


def sanitize_filename(filename: str) -> str:
    """파일명에서 특수문자 제거"""
    # 확장자 분리
    name, ext = os.path.splitext(filename)
    # 특수문자 제거 (한글, 영문, 숫자, 공백, 언더스코어, 하이픈만 허용)
    name = re.sub(r'[^\w\s가-힣一-龥ぁ-んァ-ン\-]', '', name)
    # 연속 공백 제거
    name = re.sub(r'\s+', ' ', name).strip()
    # 빈 이름이면 기본값
    if not name:
        name = "converted"
    return name + ext


def generate_file_id() -> str:
    """고유 파일 ID 생성"""
    return str(uuid.uuid4())[:8]


def check_pdf_validity(pdf_path: str) -> None:
    """PDF 파일 유효성 검사"""
    try:
        reader = PdfReader(pdf_path)
        if reader.is_encrypted:
            raise EncryptedPDFError("암호화된 PDF입니다. 암호 해제 후 다시 시도해주세요.")
    except PdfReadError:
        raise CorruptedPDFError("PDF 파일이 손상되었습니다.")
    except EncryptedPDFError:
        raise
    except Exception as e:
        raise CorruptedPDFError(f"PDF 파일을 읽을 수 없습니다: {str(e)}")


def extract_text_with_pdfplumber(pdf_path: str, progress_callback: Optional[Callable] = None) -> Tuple[str, int]:
    """pdfplumber를 사용하여 텍스트 추출"""
    texts = []
    total_pages = 0

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""
            texts.append(PAGE_SEPARATOR.format(page_num=i+1) + page_text)
            if progress_callback and callable(progress_callback):
                progress_callback(i + 1, total_pages, "텍스트 추출 중")

    return "".join(texts), total_pages


def extract_text_with_pypdf2(pdf_path: str, progress_callback: Optional[Callable] = None) -> Tuple[str, int]:
    """PyPDF2를 사용하여 텍스트 추출 (폴백)"""
    texts = []
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    for i, page in enumerate(reader.pages):
        page_text = page.extract_text() or ""
        texts.append(PAGE_SEPARATOR.format(page_num=i+1) + page_text)
        if progress_callback and callable(progress_callback):
            progress_callback(i + 1, total_pages, "텍스트 추출 중 (PyPDF2)")

    return "".join(texts), total_pages


def extract_text_with_ocr(
    pdf_path: str,
    language: str = "auto",
    progress_callback: Optional[Callable] = None,
    dpi: int = 200  # 메모리 절약을 위해 200 DPI 사용
) -> Tuple[str, int]:
    """OCR을 사용하여 스캔본 PDF에서 텍스트 추출"""
    lang = OCR_LANGUAGES.get(language, OCR_LANGUAGES["auto"])
    texts = []

    # PDF를 이미지로 변환 (페이지별로 처리하여 메모리 절약)
    try:
        # 먼저 전체 페이지 수 확인
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)

        # 페이지별로 처리 (메모리 절약)
        for page_num in range(1, total_pages + 1):
            if progress_callback and callable(progress_callback):
                progress_callback(page_num, total_pages, f"OCR 처리 중 (페이지 {page_num}/{total_pages})")

            # 단일 페이지만 변환
            images = convert_from_path(
                pdf_path,
                dpi=dpi,
                first_page=page_num,
                last_page=page_num,
                thread_count=1  # 메모리 절약
            )

            if images:
                page_text = pytesseract.image_to_string(images[0], lang=lang)
                texts.append(PAGE_SEPARATOR.format(page_num=page_num) + page_text)
                # 메모리 해제
                del images

        return "".join(texts), total_pages

    except Exception as e:
        raise PDFProcessingError(f"OCR 처리 중 오류가 발생했습니다: {str(e)}")


async def process_pdf(
    pdf_path: str,
    language: str = "auto",
    progress_callback: Optional[Callable] = None
) -> Tuple[str, int]:
    """
    PDF에서 텍스트 추출 (통합 함수)
    1. pdfplumber로 시도
    2. 텍스트가 적으면 OCR 시도
    3. 실패하면 PyPDF2로 시도
    """
    # PDF 유효성 검사
    check_pdf_validity(pdf_path)

    # 1차: pdfplumber로 추출
    text, total_pages = await asyncio.to_thread(
        extract_text_with_pdfplumber, pdf_path, progress_callback
    )

    # 추출된 텍스트가 너무 적으면 (페이지당 평균 100자 미만) 스캔본으로 판단
    stripped_text = re.sub(r'--- 페이지 \d+ ---', '', text).strip()
    avg_chars_per_page = len(stripped_text) / max(total_pages, 1)

    if avg_chars_per_page < 100:
        if progress_callback and callable(progress_callback):
            progress_callback(0, total_pages, "스캔본 PDF 감지됨, OCR 시작...")

        try:
            # OCR 시도
            text, total_pages = await asyncio.to_thread(
                extract_text_with_ocr, pdf_path, language, progress_callback
            )
        except Exception as ocr_error:
            # OCR 실패 시 PyPDF2로 최종 시도
            if progress_callback and callable(progress_callback):
                progress_callback(0, total_pages, "PyPDF2로 재시도 중...")

            try:
                text, total_pages = await asyncio.to_thread(
                    extract_text_with_pypdf2, pdf_path, progress_callback
                )
            except Exception:
                # 모든 방법 실패
                raise PDFProcessingError(
                    "텍스트를 추출할 수 없습니다. 이미지 PDF일 수 있습니다."
                )

    # 텍스트 정리 (줄바꿈 제거)
    text = clean_text(text)

    return text, total_pages


def add_horizontal_line(paragraph):
    """Word 문서에 가로선 추가"""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)


def create_word_document(text: str, original_filename: str) -> BytesIO:
    """텍스트를 Word 문서로 변환"""
    doc = Document()

    # 기본 스타일 설정
    style = doc.styles['Normal']
    font = style.font
    font.name = '맑은 고딕'
    font.size = Pt(11)
    # 한글 폰트 설정
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

    # 제목 추가
    title = doc.add_heading(f'변환된 문서: {original_filename}', level=1)

    # 텍스트 파싱 및 추가
    lines = text.split('\n')
    current_para = None

    for line in lines:
        # 페이지 구분자 처리
        if re.match(r'^--- 페이지 \d+ ---$', line.strip()):
            # 가로선과 페이지 번호 추가
            if current_para:
                doc.add_paragraph()  # 빈 줄
            para = doc.add_paragraph()
            add_horizontal_line(para)
            page_para = doc.add_paragraph(line.strip())
            page_para.alignment = 1  # 중앙 정렬
            run = page_para.runs[0]
            run.bold = True
            run.font.size = Pt(10)
            run.font.color.rgb = None  # 기본 색상
            current_para = None
        elif line.strip():
            # 일반 텍스트
            if current_para is None:
                current_para = doc.add_paragraph()
            else:
                current_para.add_run('\n')
            current_para.add_run(line)
        else:
            # 빈 줄 - 새 단락
            current_para = None

    # BytesIO로 저장
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def create_text_file(text: str) -> BytesIO:
    """텍스트 파일 생성"""
    buffer = BytesIO()
    buffer.write(text.encode('utf-8'))
    buffer.seek(0)
    return buffer


async def convert_pdf(
    pdf_path: str,
    original_filename: str,
    output_format: str = "both",  # "docx", "txt", "both"
    language: str = "auto",
    progress_callback: Optional[Callable] = None
) -> dict:
    """
    PDF 변환 메인 함수

    Returns:
        dict: {
            "text": str,  # 추출된 텍스트
            "total_pages": int,
            "docx_buffer": BytesIO or None,
            "txt_buffer": BytesIO or None
        }
    """
    # 텍스트 추출
    text, total_pages = await process_pdf(pdf_path, language, progress_callback)

    result = {
        "text": text,
        "total_pages": total_pages,
        "docx_buffer": None,
        "txt_buffer": None
    }

    # 출력 형식에 따라 파일 생성
    clean_filename = sanitize_filename(original_filename)
    base_name = os.path.splitext(clean_filename)[0]

    if output_format in ("docx", "both"):
        if progress_callback and callable(progress_callback):
            progress_callback(total_pages, total_pages, "Word 문서 생성 중...")
        result["docx_buffer"] = await asyncio.to_thread(
            create_word_document, text, base_name
        )

    if output_format in ("txt", "both"):
        if progress_callback and callable(progress_callback):
            progress_callback(total_pages, total_pages, "텍스트 파일 생성 중...")
        result["txt_buffer"] = await asyncio.to_thread(
            create_text_file, text
        )

    return result
