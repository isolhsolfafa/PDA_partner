# 🤖 PDA Partner - 생산 데이터 자동화 시스템

Google Sheets 기반의 생산 데이터 분석 및 자동화 시스템입니다. 작업시간 분석, NaN/오버타임 분석, 시각화, HTML 리포트 생성, 이메일/카카오톡 알림 기능을 제공합니다.

## ✨ 주요 기능

- 📊 Google Sheets 데이터 자동 추출 및 분석
- ⏰ 작업시간 및 진행률 계산
- 🔍 NaN 값 및 오버타임 발생률 분석
- 📈 다양한 시각화 그래프 생성 (작업시간, 범례, 주간 분석)
- 🗺️ 히트맵 생성 (주간/월간, 파트너별/모델별)
- 📧 이메일 자동 발송 (Gmail SMTP)
- 💬 카카오톡 알림 발송
- 🌐 HTML 대시보드 생성 및 Google Drive 업로드
- 📂 GitHub 자동 업로드
- 📋 JSON 데이터 저장 및 관리

## 🚀 CI/CD 설정

### GitHub Actions 사용

이 프로젝트는 GitHub Actions를 통한 자동화를 지원합니다.

#### 1. GitHub Secrets 설정

Repository Settings > Secrets and variables > Actions에서 다음 Secrets를 설정하세요:

**필수 Secrets:**
- `SHEETS_SERVICE_ACCOUNT_KEY`: Google Sheets API 서비스 계정 키 (base64 인코딩)
- `DRIVE_SERVICE_ACCOUNT_KEY`: Google Drive API 서비스 계정 키 (base64 인코딩)
- `SPREADSHEET_ID`: 메인 스프레드시트 ID
- `EMAIL_ADDRESS`: 발송자 이메일 주소
- `EMAIL_PASS`: 이메일 앱 비밀번호
- `RECEIVER_EMAIL`: 수신자 이메일 주소

**선택적 Secrets:**
- `KAKAO_REST_API_KEY`: 카카오 REST API 키
- `KAKAO_ACCESS_TOKEN`: 카카오 액세스 토큰
- `KAKAO_REFRESH_TOKEN`: 카카오 리프레시 토큰
- `GITHUB_USERNAME`: GitHub 사용자명
- `GITHUB_REPO`: GitHub 저장소명
- `DRIVE_FOLDER_ID`: Google Drive 폴더 ID
- `JSON_DRIVE_FOLDER_ID`: JSON 저장용 Drive 폴더 ID

#### 2. 서비스 계정 키 base64 인코딩

```bash
# CI/CD 설정 도우미 스크립트 실행
./scripts/setup-ci.sh
```

또는 수동으로:

```bash
# Sheets 서비스 계정 키
base64 -i config/gst-manegemnet-70faf8ce1bff.json

# Drive 서비스 계정 키  
base64 -i config/gst-manegemnet-ab8788a05cff.json
```

#### 3. 실행 스케줄

- **자동 실행**: 매일 오후 2시 (KST)
- **수동 실행**: GitHub Actions 탭에서 "Run workflow" 버튼 클릭

### Docker 사용

#### Docker Compose로 실행

```bash
# 환경변수 설정 (.env 파일 생성)
cp .env.example .env
# .env 파일을 편집하여 필요한 값들 설정

# 서비스 계정 키 파일을 config/ 디렉토리에 배치
mkdir -p config
cp /path/to/sheets-service-account.json config/
cp /path/to/drive-service-account.json config/

# Docker Compose 실행
cd docker
docker-compose up -d
```

#### 단일 Docker 컨테이너 실행

```bash
# Docker 이미지 빌드
docker build -f docker/Dockerfile -t pda-partner .

# 컨테이너 실행
docker run -d \
  --name pda-partner \
  -v $(pwd)/config:/app/config:ro \
  -v $(pwd)/output:/app/output \
  -e SHEETS_KEY_PATH=/app/config/sheets-service-account.json \
  -e DRIVE_KEY_PATH=/app/config/drive-service-account.json \
  -e SPREADSHEET_ID=your_spreadsheet_id \
  -e EMAIL_ADDRESS=your_email@gmail.com \
  -e EMAIL_PASS=your_app_password \
  -e RECEIVER_EMAIL=receiver@gmail.com \
  pda-partner
```

## 🛠️ 로컬 개발 환경 설정

### 1. 저장소 클론

```bash
git clone https://github.com/your-username/PDA_partner.git
cd PDA_partner
```

### 2. Python 환경 설정

```bash
# Python 3.9+ 권장
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt
```

### 3. 환경변수 설정

```bash
# .env 파일 생성
cp .env.example .env

# .env 파일 편집하여 필요한 값들 설정
nano .env
```

### 4. Google API 서비스 계정 키 설정

1. Google Cloud Console에서 서비스 계정 생성
2. Google Sheets API 및 Google Drive API 활성화
3. 서비스 계정 키 JSON 파일 다운로드
4. `config/` 디렉토리에 키 파일 배치

### 5. 실행

```bash
python PDA_patner.py
```

## 📋 환경변수 설명

### 필수 환경변수

| 변수명 | 설명 | 예시 |
|--------|------|------|
| `SHEETS_KEY_PATH` | Google Sheets API 서비스 계정 키 파일 경로 | `./config/sheets-key.json` |
| `DRIVE_KEY_PATH` | Google Drive API 서비스 계정 키 파일 경로 | `./config/drive-key.json` |
| `SPREADSHEET_ID` | 메인 스프레드시트 ID | `19dkwKNW6VshCg3wTemzmbbQlbATfq6brAWluaps1Rm0` |
| `EMAIL_ADDRESS` | 발송자 이메일 주소 | `sender@gmail.com` |
| `EMAIL_PASS` | 이메일 앱 비밀번호 | `your_app_password` |
| `RECEIVER_EMAIL` | 수신자 이메일 주소 | `receiver@gmail.com` |

### 선택적 환경변수

| 변수명 | 기본값 | 설명 |
|--------|--------|------|
| `TEST_MODE` | `true` | 테스트 모드 (true/false) |
| `LIMIT` | `1` | 처리할 스프레드시트 수 제한 |
| `GENERATE_GRAPHS` | `auto` | 그래프 생성 옵션 (auto/true/false) |
| `GITHUB_UPLOAD` | `auto` | GitHub 업로드 옵션 (auto/true/false) |
| `KAKAO_REST_API_KEY` | - | 카카오 REST API 키 |
| `GITHUB_TOKEN` | - | GitHub 개인 액세스 토큰 |

## 🔧 기능별 설정

### 그래프 생성 옵션 (`GENERATE_GRAPHS`)

- `auto`: 월요일과 금요일에만 그래프 생성 (기본값)
- `true`: 항상 그래프 생성
- `false`: 그래프 생성 안함

### GitHub 업로드 옵션 (`GITHUB_UPLOAD`)

- `auto`: TEST_MODE가 false일 때만 업로드 (기본값)
- `true`: 항상 업로드
- `false`: 업로드 안함

## 📊 출력 파일

- `output/`: 생성된 그래프 및 히트맵 이미지
- `*.html`: HTML 대시보드 파일
- `*.json`: 분석 결과 JSON 데이터
- `pda_partner.log`: 실행 로그

## 🐛 문제 해결

### 1. 폰트 관련 오류

```bash
# macOS
brew install font-nanum-gothic

# Ubuntu/Debian
sudo apt-get install fonts-nanum fonts-nanum-coding fonts-nanum-extra
```

### 2. Google API 인증 오류

- 서비스 계정 키 파일 경로 확인
- Google Sheets/Drive API 활성화 확인
- 서비스 계정에 스프레드시트 및 드라이브 폴더 권한 부여

### 3. 이메일 발송 실패

- Gmail 2단계 인증 활성화
- 앱 비밀번호 생성 및 사용
- SMTP 설정 확인

## 📝 라이센스

이 프로젝트는 MIT 라이센스 하에 배포됩니다.

## 🤝 기여하기

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 지원

문제가 발생하거나 질문이 있으시면 이슈를 생성해 주세요.

## 🎯 주요 기능

### 📊 데이터 분석
- Google Sheets에서 생산 작업 데이터 자동 추출
- 작업시간 분석 (워킹데이, 오버타임 계산)
- NaN (결측값) 발생률 분석
- 협력사별 성과 분석

### 📈 시각화
- 작업별 소요시간 그래프 생성
- 범례 차트 및 WD 차트 생성
- **주간 협력사 NaN 히트맵** (실시간 생성)
- **월간 협력사 NaN 히트맵** (월말 금요일 자동 생성)
- **월간 모델별 NaN 히트맵** (월말 금요일 자동 생성)
- 진행률 시각화 (진행률 바 및 완료 상태 표시)

### 🔔 알림 시스템
- **이메일 리포트 자동 발송** (HTML 형식, 히트맵 첨부)
- **카카오톡 알림 메시지** (NaN 발생 시 즉시 알림)
- **HTML 대시보드 생성** (partner.html)
- **📊트렌드 지표** 섹션 (주간/월간 히트맵 링크 통합)

### ☁️ 클라우드 연동
- **Google Drive 자동 업로드** (히트맵, JSON 파일)
- **GitHub Pages 배포** (partner.html 자동 배포)
- **JSON 데이터 백업** (날짜별 실행 결과 저장)
- **Google Drive 링크 통합** (히트맵 직접 접근 URL 제공)

## 🛠️ 설치 및 설정

### 1. 환경 요구사항
```bash
Python 3.8+
```

### 2. 패키지 설치
```bash
pip install -r requirements.txt
```

### 3. 환경변수 설정
`.env` 파일을 생성하고 다음 변수들을 설정하세요:

```env
# 시스템 설정
TEST_MODE=true

# Google Drive 설정 (폴더 분리)
DRIVE_FOLDER_ID=1Gylm36vhtrl_yCHurZYGgeMlt5U0CliE          # 그래프/히트맵/HTML 저장용
JSON_DRIVE_FOLDER_ID=13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1     # JSON 데이터 저장용
NOVA_FOLDER_ID=13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1           # 호환성 지원

# 이메일 설정 (필수) - 기존 .env 호환성 지원
EMAIL_ADDRESS=your_email@gmail.com     # 또는 SMTP_USER
EMAIL_PASS=your_app_password           # 또는 SMTP_PASSWORD
RECEIVER_EMAIL=receiver@example.com
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Google Sheets 설정
SPREADSHEET_ID=your_main_spreadsheet_id
SHEETS_KEY_PATH=/path/to/sheets_service_account.json
DRIVE_KEY_PATH=/path/to/drive_service_account.json

# GitHub 설정 (public 폴더 업로드)
GITHUB_USERNAME=isolhsolfafa
GITHUB_REPO=gst-factory
GITHUB_BRANCH=main
GITHUB_TOKEN=your_github_token_with_repo_permissions

# 카카오톡 API 설정
KAKAO_REST_API_KEY=your_kakao_api_key
KAKAO_ACCESS_TOKEN=your_kakao_access_token
KAKAO_REFRESH_TOKEN=your_kakao_refresh_token

# 처리 제한 설정
LIMIT=1                                # 테스트용, 운영시 제거 또는 큰 값
WORKSHEET_RANGE=WORKSHEET!A1:Z100
INFO_RANGE=정보판!A1:Z100

# 그래프 생성 옵션 (auto, true, false)
GENERATE_GRAPHS=auto                   # auto: 월/금요일만, true: 항상, false: 안함

# GitHub 업로드 옵션 (auto, true, false)
GITHUB_UPLOAD=auto                     # auto: 운영모드만, true: 항상, false: 안함
```

### 4. Google API 설정
1. Google Cloud Console에서 프로젝트 생성
2. Sheets API 및 Drive API 활성화
3. 서비스 계정 생성 및 키 파일 다운로드
4. 환경변수에 키 파일 경로 설정

## 🚀 사용법

### 기본 실행
```bash
python PDA_patner.py
```

### 주요 모드
- **TEST_MODE=true**: 모든 기능 실행 (개발/테스트용)
- **TEST_MODE=false**: 운영 모드 (자동 스케줄링)

### 그래프 생성 옵션
- **GENERATE_GRAPHS=auto**: 월요일/금요일에만 생성 (기본값)
- **GENERATE_GRAPHS=true**: 항상 생성
- **GENERATE_GRAPHS=false**: 생성 안함

### GitHub 업로드 옵션
- **GITHUB_UPLOAD=auto**: TEST_MODE가 아닐 때만 업로드 (기본값)
- **GITHUB_UPLOAD=true**: 항상 업로드
- **GITHUB_UPLOAD=false**: 업로드 안함

## 📁 프로젝트 구조

```
PDA_partner/
├── PDA_patner.py          # 메인 스크립트
├── requirements.txt       # 의존성 패키지
├── README.md             # 프로젝트 문서
├── .env                  # 환경변수 (생성 필요)
├── config/               # 설정 파일들
├── output/               # 출력 파일들
│   ├── *.png            # 생성된 그래프들
│   ├── *.json           # 분석 결과 JSON
│   └── *.html           # HTML 리포트
└── pda_partner.log       # 로그 파일
```

## 🔧 주요 함수들

### 데이터 처리
- `collect_and_process_data()`: 메인 데이터 수집 및 처리
- `fetch_data_from_sheets()`: Google Sheets 데이터 추출
- `process_data()`: 작업시간 계산 및 분류
- `compute_occurrence_rates()`: NaN/오버타임 발생률 계산

### 시각화
- `generate_and_save_graph()`: 작업시간 그래프 생성
- `generate_legend_chart()`: 범례 차트 생성
- `generate_weekly_report_heatmap()`: 주간 히트맵 생성
- `generate_heatmap()`: 월간 히트맵 생성 (협력사별/모델별)
- `load_json_files_from_drive()`: Google Drive에서 JSON 데이터 로드

### 알림 및 리포트
- `send_occurrence_email()`: 이메일 리포트 발송
- `send_nan_alert_to_kakao()`: 카카오톡 알림
- `build_combined_email_body()`: HTML 리포트 생성 (히트맵 링크 포함)
- `upload_to_github()`: GitHub Pages 자동 배포

### 클라우드 연동
- `upload_to_drive()`: Google Drive 파일 업로드
- `get_drive_file_id()`: 업로드된 파일 ID 추출
- `create_drive_link()`: 공유 가능한 Google Drive 링크 생성

## 📊 작업 분류 시스템

### 기본 분류
- **기구**: 하드웨어 조립 작업
- **전장**: 전기/전자 작업
- **TMS_반제품**: TMS 관련 반제품 작업
- **검사**: 품질 검사 작업
- **마무리**: 최종 마무리 작업

### 모델별 작업 정의
각 제품 모델(GAIA-I, DRAGON, SWS-I 등)에 따라 작업 분류가 자동으로 조정됩니다.

## 🔍 모니터링 지표

### NaN 발생률 분석
- **작업 시간 누락**: 시작/완료 시간이 기록되지 않은 경우
- **진행률 누락**: 작업 진행률이 기록되지 않은 경우
- **협력사별 발생 비율**: 주간/월간 히트맵으로 시각화
- **모델별 발생 비율**: 제품 모델별 NaN 발생 패턴 분석

### 오버타임 분석
- 평균 작업시간 대비 초과 시간 계산
- 허용 오차: 2시간 (환경변수로 설정 가능)
- 작업별/협력사별 오버타임 발생률 추적
- 오버타임 발생 시 자동 알림

### 진행률 및 완료도 추적
- **작업 분류별 완료율**: 기구, 전장, 검사 등 단계별 진행률
- **실시간 모니터링**: HTML 대시보드를 통한 실시간 현황 확인
- **완료 상태 시각화**: 진행률 바와 상태 아이콘으로 직관적 표시

## 🔒 보안 고려사항

### 환경변수 사용
- 모든 민감한 정보는 환경변수로 관리
- `.env` 파일을 통한 로컬 설정 지원 (python-dotenv)
- 하드코딩된 크리덴셜 완전 제거
- 필수 환경변수 검증 로직 추가

### API 키 관리
- **GitHub 토큰**: repo 권한 필요 (public 폴더 업로드용)
- **Google 서비스 계정**: Sheets/Drive API 최소 권한
- **카카오 토큰**: 자동 리프레시 토큰 갱신 지원
- **이메일**: Gmail App Password 사용 권장

### 드라이브 폴더 분리
- **그래프/히트맵**: `1Gylm36vhtrl_yCHurZYGgeMlt5U0CliE`
- **JSON 데이터**: `13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1`
- 데이터 타입별 접근 권한 분리 관리

## 📝 로그 및 디버깅

### 로그 파일
- `pda_partner.log`: 모든 실행 로그 기록 (파일 + 콘솔)
- 에러 및 경고 메시지 추적
- 성능 지표 기록
- 임시 파일 자동 정리 로그

### 디버깅 모드
```bash
# 상세 로그와 함께 실행
python -v PDA_patner.py

# 특정 개수만 처리 (개발/테스트용)
export LIMIT=5
python PDA_patner.py

# 그래프 생성 강제 실행
export GENERATE_GRAPHS=true
python PDA_patner.py

# GitHub 업로드 비활성화
export GITHUB_UPLOAD=false
python PDA_patner.py
```

## 🤝 기여 및 개발

### 개발 환경 설정
1. 저장소 클론
2. 가상환경 생성 및 활성화
3. 의존성 설치: `pip install -r requirements.txt`
4. 환경변수 설정
5. 테스트 실행

### 코드 품질
- PEP 8 스타일 가이드 준수
- 함수별 독립적인 단위 테스트
- 에러 처리 및 로깅 강화

## 📊 히트맵 시스템

### 주간 히트맵 (실시간 생성)
- 매주 실행 시 자동으로 생성되어 Google Drive에 업로드
- 협력사별 NaN 발생률을 색상으로 표시
- HTML 리포트에 자동으로 링크 포함

### 월간 히트맵 (월말 자동 생성)
- **협력사별 월간 히트맵**: 매월 마지막 금요일에 자동 생성
- **모델별 월간 히트맵**: 제품 모델별 NaN 발생 패턴 분석
- Google Drive에 자동 업로드 및 공유 링크 생성

### 📊트렌드 지표 대시보드
HTML 리포트에 통합된 히트맵 링크들:
- 📅 주간 협력사 NaN 히트맵 (동적 생성)
- 🗓️ 월간 협력사 NaN 히트맵 (기존 최신)
- 📈 월간 모델별 NaN 히트맵 (기존 최신)

## 🔄 자동화 스케줄링

### 권장 실행 주기
- **월요일**: 주간 리포트 및 히트맵 생성
- **금요일**: 주간 마감 리포트 및 월말 시 월간 히트맵 생성
- **실시간**: NaN 발생 시 즉시 카카오톡 알림

### 환경별 설정
```bash
# 운영 환경 (cron 등 자동 실행)
TEST_MODE=false
GENERATE_GRAPHS=auto
GITHUB_UPLOAD=auto

# 개발 환경 (수동 테스트)
TEST_MODE=true
GENERATE_GRAPHS=true
GITHUB_UPLOAD=false
LIMIT=10
```

## 📞 지원 및 문의

문제가 발생하거나 개선 제안이 있으시면 이슈를 등록해 주세요.

### 주요 업데이트 내역
- **v2.1**: 보안 강화 및 환경변수 시스템 개선 (2025-06-23)
  - `.env` 파일 지원 추가 (python-dotenv)
  - 드라이브 폴더 분리 (JSON/그래프 별도 관리)
  - GitHub 업로드 경로 수정 (public 폴더)
  - 카카오톡 토큰 자동 갱신 지원
  - 하이퍼링크 문제 해결 (=HYPERLINK 공식 사용)
- **v2.0**: 월간 히트맵 시스템 추가
- **v1.9**: 📊트렌드 지표 대시보드 통합
- **v1.8**: Google Drive 링크 자동 생성
- **v1.7**: GitHub Pages 자동 배포

---

**⚡ 효율적인 생산 관리를 위한 데이터 기반 의사결정 도구** 