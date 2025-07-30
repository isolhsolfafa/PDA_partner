FROM python:3.9-slim

# 시스템 패키지 설치
RUN apt-get update && apt-get install -y \
    fonts-nanum \
    fonts-nanum-coding \
    fonts-nanum-extra \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# 작업 디렉토리 설정
WORKDIR /app

# Python 의존성 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 애플리케이션 코드 복사
COPY . .

# 환경변수 설정
ENV PYTHONPATH=/app
ENV PYTHONUNBUFFERED=1

# 실행 권한 설정
RUN chmod +x PDA_partner.py

# 기본 명령어
CMD ["python", "PDA_partner.py"] 