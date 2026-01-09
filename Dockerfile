FROM python:3.11-slim

# 시스템 패키지 설치
RUN apt-get update && apt-get install -y \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-kor \
    tesseract-ocr-eng \
    tesseract-ocr-chi-sim \
    tesseract-ocr-chi-tra \
    tesseract-ocr-jpn \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 의존성 먼저 설치 (캐시 활용)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 소스 코드 복사
COPY . .

# temp 디렉토리 생성
RUN mkdir -p temp

# 환경변수
ENV PYTHONUNBUFFERED=1

# 포트 설정 (Railway에서 PORT 환경변수 사용)
EXPOSE 8000

# 실행
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
