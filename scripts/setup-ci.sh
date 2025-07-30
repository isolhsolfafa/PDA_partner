#!/bin/bash

# PDA Partner CI/CD 설정 스크립트

set -e

echo "🚀 PDA Partner CI/CD 환경 설정을 시작합니다..."

# 1. GitHub Secrets 설정 확인
echo "📋 필요한 GitHub Secrets 목록:"
echo "=================================="
echo "필수 Secrets:"
echo "- SHEETS_SERVICE_ACCOUNT_KEY (base64 인코딩된 서비스 계정 키)"
echo "- DRIVE_SERVICE_ACCOUNT_KEY (base64 인코딩된 서비스 계정 키)"
echo "- SPREADSHEET_ID"
echo "- EMAIL_ADDRESS"
echo "- EMAIL_PASS"
echo "- RECEIVER_EMAIL"
echo ""
echo "선택적 Secrets:"
echo "- KAKAO_REST_API_KEY"
echo "- KAKAO_ACCESS_TOKEN"
echo "- KAKAO_REFRESH_TOKEN"
echo "- GITHUB_USERNAME"
echo "- GITHUB_REPO"
echo "- DRIVE_FOLDER_ID"
echo "- JSON_DRIVE_FOLDER_ID"
echo ""

# 2. 서비스 계정 키 파일 base64 인코딩 도우미
if [ -f "config/gst-manegemnet-70faf8ce1bff.json" ]; then
    echo "📄 Sheets 서비스 계정 키 파일을 base64로 인코딩합니다..."
    echo "다음 값을 SHEETS_SERVICE_ACCOUNT_KEY Secret에 설정하세요:"
    echo "=================================="
    base64 -i config/gst-manegemnet-70faf8ce1bff.json
    echo "=================================="
    echo ""
fi

if [ -f "config/gst-manegemnet-ab8788a05cff.json" ]; then
    echo "📄 Drive 서비스 계정 키 파일을 base64로 인코딩합니다..."
    echo "다음 값을 DRIVE_SERVICE_ACCOUNT_KEY Secret에 설정하세요:"
    echo "=================================="
    base64 -i config/gst-manegemnet-ab8788a05cff.json
    echo "=================================="
    echo ""
fi

# 3. GitHub Actions 워크플로우 파일 확인
if [ -f ".github/workflows/pda-partner.yml" ]; then
    echo "✅ GitHub Actions 워크플로우 파일이 존재합니다."
else
    echo "❌ GitHub Actions 워크플로우 파일이 없습니다."
    echo "   .github/workflows/pda-partner.yml 파일을 생성해야 합니다."
fi

# 4. Docker 파일 확인
if [ -f "docker/Dockerfile" ]; then
    echo "✅ Dockerfile이 존재합니다."
else
    echo "❌ Dockerfile이 없습니다."
fi

# 5. requirements.txt 확인
if [ -f "requirements.txt" ]; then
    echo "✅ requirements.txt가 존재합니다."
else
    echo "❌ requirements.txt가 없습니다."
fi

echo ""
echo "🔧 CI/CD 설정 가이드:"
echo "=================================="
echo "1. GitHub Repository Settings > Secrets and variables > Actions에서 위의 Secrets를 설정하세요."
echo "2. 서비스 계정 키 파일은 base64로 인코딩하여 설정하세요."
echo "3. 워크플로우는 다음 일정으로 실행됩니다:"
echo "   - 매일 오후 2시 (KST)"
echo "   - 수동 실행도 가능합니다."
echo "4. Docker를 사용하려면 'docker-compose up' 명령을 사용하세요."
echo ""
echo "✅ 설정 완료!" 