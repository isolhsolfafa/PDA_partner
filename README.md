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

## 🏗️ 시스템 구조

```
📁 PDA_partner/
├── 🚀 PDA_partner.py          # 메인 실행 파일 (All-in-One 구조)
├── 📦 requirements.txt        # Python 의존성
├── 📁 config/                 # Google 서비스 계정 키들
│   ├── gst-manegemnet-70faf8ce1bff.json    # Sheets API 키
│   └── gst-manegemnet-ab8788a05cff.json    # Drive API 키
├── 📁 .github/workflows/      # CI/CD 파이프라인
│   └── pda-partner.yml        # GitHub Actions 워크플로우
├── 📁 output/                 # 생성된 파일들 (자동 생성)
│   ├── *.png                  # 히트맵 이미지들
│   ├── *.html                 # HTML 리포트
│   └── *.json                 # JSON 데이터
├── 📁 backup_*/               # 백업 폴더들
├── 📁 archive_*/              # 아카이브 폴더들
└── 📖 README.md              # 이 문서
```

## 🚀 CI/CD 설정

### GitHub Actions 사용

이 프로젝트는 GitHub Actions를 통한 자동화를 지원합니다.

#### 1. GitHub Secrets 설정

Repository Settings > Secrets and variables > Actions에서 다음 Secrets를 설정하세요:

**필수 Secrets:**
- `SHEETS_SERVICE_ACCOUNT_KEY`: Google Sheets API 서비스 계정 키 (base64 인코딩)
- `DRIVE_SERVICE_ACCOUNT_KEY`: Google Drive API 서비스 계정 키 (base64 인코딩)
- `SPREADSHEET_ID`: 메인 스프레드시트 ID
- `TARGET_SHEET_NAME`: 타겟 시트 이름 (예: "1월", "2월" 등)
- `EMAIL_ADDRESS`: 발송자 이메일 주소
- `EMAIL_PASS`: 이메일 앱 비밀번호
- `RECEIVER_EMAIL`: 수신자 이메일 주소
- `DRIVE_FOLDER_ID`: Google Drive 메인 폴더 ID
- `JSON_DRIVE_FOLDER_ID`: JSON 저장용 Drive 폴더 ID

**선택적 Secrets:**
- `KAKAO_REST_API_KEY`: 카카오 REST API 키
- `KAKAO_ACCESS_TOKEN`: 카카오 액세스 토큰
- `KAKAO_REFRESH_TOKEN`: 카카오 리프레시 토큰
- `GH_USERNAME`: GitHub 사용자명
- `GH_REPO`: GitHub 저장소명
- `GH_TOKEN`: GitHub Personal Access Token
- `GH_BRANCH`: GitHub 업로드 브랜치 (기본: main)

#### 2. 서비스 계정 키 base64 인코딩

```bash
# macOS/Linux에서
base64 -i config/gst-manegemnet-70faf8ce1bff.json

# 또는 온라인 base64 인코더 사용
# 결과를 SHEETS_SERVICE_ACCOUNT_KEY에 입력

base64 -i config/gst-manegemnet-ab8788a05cff.json
# 결과를 DRIVE_SERVICE_ACCOUNT_KEY에 입력
```

#### 3. 실행 스케줄 및 설정

- **자동 실행**: 매일 오후 12시 20분 (KST)
- **수동 실행**: GitHub Actions 탭에서 "Run workflow" 버튼 클릭

**워크플로우 입력 옵션:**
- `처리할 항목 수`: 기본값 200개
- `그래프 생성`: 기본값 false (필요시에만 체크)
- `GitHub 업로드`: 기본값 true (HTML 파일 자동 업로드)

## 🛠️ 로컬 실행

### 환경 설정

```bash
# 1. 저장소 클론
git clone https://github.com/isolhsolfafa/PDA_partner.git
cd PDA_partner

# 2. Python 의존성 설치
pip install -r requirements.txt

# 3. 환경변수 설정 (.env 파일 생성)
SPREADSHEET_ID=your_spreadsheet_id
TARGET_SHEET_NAME=1월
DRIVE_FOLDER_ID=your_drive_folder_id
JSON_DRIVE_FOLDER_ID=your_json_drive_folder_id
EMAIL_ADDRESS=your_email@gmail.com
EMAIL_PASS=your_app_password
RECEIVER_EMAIL=receiver@gmail.com
KAKAO_REST_API_KEY=your_kakao_api_key
KAKAO_ACCESS_TOKEN=your_access_token
KAKAO_REFRESH_TOKEN=your_refresh_token
SHEETS_KEY_PATH=config/gst-manegemnet-70faf8ce1bff.json
DRIVE_KEY_PATH=config/gst-manegemnet-ab8788a05cff.json

# 4. 서비스 계정 키 파일 배치
mkdir -p config
# Google Cloud Console에서 다운로드한 서비스 계정 키 파일들을 config/ 폴더에 배치

# 5. 실행
python PDA_partner.py
```

### 실행 옵션

```bash
# 기본 실행 (모든 기능)
python PDA_partner.py

# 테스트 모드 (제한된 항목 처리)
export LIMIT=10
python PDA_partner.py

# 그래프 생성 없이 실행
export GENERATE_GRAPHS=false
python PDA_partner.py

# GitHub 업로드 없이 실행
export GITHUB_UPLOAD=false
python PDA_partner.py
```

## 📊 주요 구성 요소

### 1. 데이터 분석
- Google Sheets에서 생산 데이터 실시간 수집
- NaN 값 발생률 및 오버타임 분석
- 협력사별, 모델별 성과 지표 계산
- 작업 진행률 및 시간 분석

### 2. 시각화
- **작업시간 그래프**: 모델별 총 작업시간 차트
- **범례 차트**: 작업 단계별 색상 범례
- **WD 작업시간**: 평일 작업시간 분석
- **히트맵**: 주간/월간 NaN 비율 트렌드 (협력사별/모델별)

### 3. 리포팅
- **HTML 대시보드**: 종합 분석 결과 웹 페이지
- **이메일 리포트**: 요약 정보 및 첨부파일 자동 발송
- **카카오톡 알림**: 실시간 처리 결과 알림
- **JSON 데이터**: 구조화된 분석 결과 저장

### 4. 자동화
- **GitHub Actions**: 매일 자동 실행
- **Google Drive**: 모든 결과물 자동 업로드
- **GitHub Pages**: HTML 리포트 자동 배포
- **에러 처리**: 견고한 예외 처리 및 재시도 로직

## 🔧 설정 가이드

### Google API 설정

1. **Google Cloud Console**에서 프로젝트 생성
2. **Google Sheets API**, **Google Drive API** 활성화
3. **서비스 계정** 생성 및 키 다운로드
4. **Google Drive 폴더** 생성 및 서비스 계정에 편집 권한 부여
5. **Google Sheets** 문서를 서비스 계정에 공유

### 카카오톡 API 설정

1. **Kakao Developers**에서 앱 생성
2. **REST API 키** 발급
3. **카카오 로그인** 활성화
4. **액세스 토큰** 및 **리프레시 토큰** 발급

### 이메일 설정

1. **Gmail**에서 2단계 인증 활성화
2. **앱 비밀번호** 생성
3. SMTP 설정으로 자동 이메일 발송

## 📈 모니터링

### 로그 확인

```bash
# 로컬 실행 시 로그 파일
tail -f pda_partner.log

# GitHub Actions 로그
# Actions 탭에서 워크플로우 실행 로그 확인
```

### 주요 메트릭

- **처리 성공률**: 정상 처리된 모델 수 / 전체 모델 수
- **NaN 발생률**: 데이터 누락 비율
- **오버타임 발생률**: 예상 시간 초과 비율
- **실행 시간**: 전체 파이프라인 실행 소요 시간

## 🤝 기여 방법

1. 이 저장소를 포크합니다
2. 기능 브랜치를 생성합니다 (`git checkout -b feature/amazing-feature`)
3. 변경사항을 커밋합니다 (`git commit -m 'Add amazing feature'`)
4. 브랜치에 푸시합니다 (`git push origin feature/amazing-feature`)
5. Pull Request를 생성합니다

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

## 📞 지원

문제가 발생하거나 질문이 있으시면:

1. **GitHub Issues**: 버그 리포트 및 기능 요청
2. **GitHub Discussions**: 일반적인 질문 및 토론
3. **Wiki**: 상세한 설정 가이드 및 FAQ

---

**⚡ PDA Partner**로 생산 데이터 분석을 자동화하고 업무 효율성을 극대화하세요!