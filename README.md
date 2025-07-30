# 🚀 PDA Partner - Production Data Analysis System

**생산 데이터 분석 및 자동화 시스템**

## 📋 개요

PDA Partner는 Google Sheets에서 생산 데이터를 자동으로 수집하고 분석하여 다음과 같은 기능을 제공합니다:

- 📊 **실시간 데이터 분석**: NaN 값, 오버타임 발생률 분석
- 📈 **히트맵 생성**: 주간/월간 트렌드 시각화
- 📧 **자동 알림**: 이메일 및 카카오톡 알림
- 🌐 **웹 리포트**: HTML 대시보드 자동 생성
- ⚡ **GitHub Pages**: 자동 배포

## 🏗️ 시스템 구조

```
📁 PDA_partner/
├── 🚀 PDA_partner.py          # 메인 실행 파일 (All-in-One)
├── 📦 requirements.txt        # Python 의존성
├── 📁 config/                 # Google 서비스 계정 키들
├── 📁 .github/workflows/      # CI/CD 파이프라인
├── 📁 output/                 # 생성된 파일들 (자동 생성)
└── 📖 README.md              # 이 문서
```

## 🚀 설치 및 실행

### 로컬 실행

```bash
# 1. 저장소 클론
git clone https://github.com/isolhsolfafa/PDA_partner.git
cd PDA_partner

# 2. 의존성 설치
pip install -r requirements.txt

# 3. 환경변수 설정
cp .env.example .env
# .env 파일을 편집하여 필요한 설정값 입력

# 4. Google 서비스 계정 키 설정
# config/ 폴더에 서비스 계정 JSON 키 파일들 배치

# 5. 실행
python PDA_partner.py
```

### GitHub Actions (자동 실행)

- **매일 오후 2시 (KST)** 자동 실행
- **수동 실행** 가능 (GitHub Actions 탭에서)

## ⚙️ 환경변수 설정

### 필수 환경변수

```env
# Google API
SPREADSHEET_ID="your_spreadsheet_id"
TARGET_SHEET_NAME="시트이름"
DRIVE_FOLDER_ID="drive_folder_id"
JSON_DRIVE_FOLDER_ID="json_folder_id"

# GitHub 설정
GITHUB_TOKEN="your_github_token"
GITHUB_USERNAME="your_username"
GITHUB_REPO="your_repo"

# 카카오톡 API
KAKAO_REST_API_KEY="your_kakao_api_key"
KAKAO_ACCESS_TOKEN="your_access_token"
KAKAO_REFRESH_TOKEN="your_refresh_token"

# 이메일 설정
EMAIL_ADDRESS="your_email@gmail.com"
EMAIL_PASS="your_app_password"
RECEIVER_EMAIL="receiver@email.com"
```

### 선택적 환경변수

```env
# 실행 옵션
LIMIT="200"                    # 처리할 항목 수
GENERATE_GRAPHS="auto"         # 그래프 생성 (auto/true/false)
GITHUB_UPLOAD="auto"           # GitHub 업로드 (auto/true/false)
TEST_MODE="false"              # 테스트 모드
```

## 📊 주요 기능

### 1. 데이터 분석
- **NaN 값 감지**: 누락된 데이터 식별
- **오버타임 계산**: 작업 시간 초과 분석
- **협력사별 분석**: 파트너사별 성과 측정

### 2. 시각화
- **주간 히트맵**: 일별 트렌드 분석
- **월간 히트맵**: 월별 패턴 분석 (마지막 금요일 생성)
- **대시보드**: HTML 기반 통합 리포트

### 3. 알림 시스템
- **이메일**: 상세 리포트 및 첨부파일
- **카카오톡**: 요약 알림
- **자동 토큰 갱신**: 카카오톡 API 토큰 자동 관리

## 🔧 GitHub Secrets 설정

GitHub 저장소의 **Settings > Secrets and variables > Actions**에서 다음 값들을 설정:

### Google API
- `SPREADSHEET_ID`
- `DRIVE_FOLDER_ID`
- `JSON_DRIVE_FOLDER_ID`
- `TARGET_SHEET_NAME`
- `SHEETS_SERVICE_ACCOUNT_KEY` (base64 인코딩)
- `DRIVE_SERVICE_ACCOUNT_KEY` (base64 인코딩)

### GitHub 설정
- `GH_TOKEN`
- `GH_USERNAME`
- `GH_REPO`
- `GH_BRANCH`

### 카카오톡 API
- `KAKAO_REST_API_KEY`
- `KAKAO_ACCESS_TOKEN`
- `KAKAO_REFRESH_TOKEN`

### 이메일 설정
- `EMAIL_ADDRESS`
- `EMAIL_PASS`
- `RECEIVER_EMAIL`

## 🕐 스케줄링

### 자동 실행
- **매일 오후 2시 (KST)**: 정기 실행
- **월요일/금요일**: 그래프 생성 활성화
- **마지막 금요일**: 월간 히트맵 생성

### 수동 실행
GitHub Actions 탭에서 **"Run workflow"** 버튼으로 수동 실행 가능

## 📁 출력 파일들

### 로컬 생성
```
output/
├── nan_ot_results_YYYYMMDD_HHMMSS_요일_회차.json
├── weekly_partner_nan_heatmap_YYYYMMDD.png
├── monthly_partner_nan_heatmap_YYYYMMDD.png
└── monthly_model_nan_heatmap_YYYYMMDD.png
```

### 웹 배포
- **partner.html**: GitHub Pages로 자동 배포
- **히트맵 이미지들**: Google Drive 자동 업로드

## 🚨 문제 해결

### 일반적인 문제들

1. **Google API 권한 오류**
   - 서비스 계정에 Sheets/Drive 권한 확인
   - JSON 키 파일 형식 확인

2. **카카오톡 알림 실패**
   - 토큰 만료 확인 (자동 갱신됨)
   - REST API 키 확인

3. **GitHub 업로드 실패**
   - Personal Access Token 권한 확인
   - 저장소 접근 권한 확인

### 로그 확인
```bash
# 실행 로그 확인
tail -f pda_partner.log

# GitHub Actions 로그는 Actions 탭에서 확인
```

## 🔄 업데이트 및 유지보수

### 의존성 업데이트
```bash
pip install --upgrade -r requirements.txt
```

### 토큰 갱신
- **카카오톡**: 자동 갱신 (코드에서 처리)
- **GitHub**: 필요시 새 Personal Access Token 발급
- **Google**: 서비스 계정은 만료 없음

## 📞 지원

문제가 발생하면 GitHub Issues에서 문의해주세요.

---

**⚡ 매일 자동으로 실행되는 안정적인 생산 데이터 분석 시스템입니다.** 