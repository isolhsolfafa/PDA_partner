# 🔑 GitHub Secrets 설정 가이드

**GitHub 저장소 > Settings > Secrets and variables > Actions**에서 설정해주세요.

## 📋 필수 Secrets 목록 (15개)

### 🟢 **Google API 관련 (6개)**

| Secret Name | 설명 | 예시/형식 |
|-------------|------|-----------|
| `SPREADSHEET_ID` | 메인 스프레드시트 ID | `19dkwKNW6VshCg3wTemzmbbQlbATfq6brAWluaps1Rm0` |
| `TARGET_SHEET_NAME` | 처리할 시트 이름 | `출하예정리스트(TEST)` |
| `DRIVE_FOLDER_ID` | 그래프/HTML 저장 폴더 | `1Gylm36vhtrl_yCHurZYGgeMlt5U0CliE` |
| `JSON_DRIVE_FOLDER_ID` | JSON 데이터 저장 폴더 | `13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1` |
| `SHEETS_SERVICE_ACCOUNT_KEY` | Sheets 서비스 계정 키 (base64) | `base64로 인코딩된 JSON 키` |
| `DRIVE_SERVICE_ACCOUNT_KEY` | Drive 서비스 계정 키 (base64) | `base64로 인코딩된 JSON 키` |

### 🟡 **GitHub 설정 (4개)**

| Secret Name | 설명 | 예시/형식 |
|-------------|------|-----------|
| `GH_TOKEN` | GitHub Personal Access Token | `ghp_xxxxxxxxxx` |
| `GH_USERNAME` | GitHub 사용자명 | `isolhsolfafa` |
| `GH_REPO` | 업로드할 저장소명 | `gst-factory` |
| `GH_BRANCH` | 업로드할 브랜치 | `main` |

### 🟣 **카카오톡 API (3개)**

| Secret Name | 설명 | 예시/형식 |
|-------------|------|-----------|
| `KAKAO_REST_API_KEY` | 카카오 REST API 키 | `d24d0956456fb747fefea7129a295a4e` |
| `KAKAO_ACCESS_TOKEN` | 카카오 액세스 토큰 | `xxxxxxxxxxxxxxxxxxxx` |
| `KAKAO_REFRESH_TOKEN` | 카카오 리프레시 토큰 | `xxxxxxxxxxxxxxxxxxxx` |

### 🔵 **이메일 설정 (3개)**

| Secret Name | 설명 | 예시/형식 |
|-------------|------|-----------|
| `EMAIL_ADDRESS` | 발송 이메일 주소 | `your_email@gmail.com` |
| `EMAIL_PASS` | Gmail 앱 패스워드 | `qyts eiwd kjma exya` |
| `RECEIVER_EMAIL` | 수신 이메일 주소 | `receiver@naver.com` |

---

## 🔧 상세 설정 방법

### 1️⃣ **Google 서비스 계정 키 base64 인코딩**

```bash
# macOS/Linux
base64 -i config/gst-manegemnet-70faf8ce1bff.json | pbcopy

# 또는 온라인 base64 인코더 사용
# JSON 파일 내용을 복사해서 base64로 인코딩
```

**⚠️ 주의**: JSON 파일 전체 내용을 base64로 인코딩해서 붙여넣기

### 2️⃣ **GitHub Personal Access Token 생성**

1. GitHub > Settings > Developer settings > Personal access tokens > Tokens (classic)
2. "Generate new token" 클릭
3. **필요한 권한 체크**:
   - ✅ `repo` (전체 저장소 접근)
   - ✅ `workflow` (Actions 워크플로우)
4. 생성된 토큰을 `GH_TOKEN`에 설정

### 3️⃣ **카카오톡 API 설정**

1. [카카오 개발자센터](https://developers.kakao.com/)에서 앱 생성
2. REST API 키를 `KAKAO_REST_API_KEY`에 설정
3. 카카오톡 메시지 권한 설정 후 토큰 발급

### 4️⃣ **Gmail 앱 패스워드 생성**

1. Google 계정 > 보안 > 2단계 인증 활성화
2. 앱 패스워드 생성 (메일 전용)
3. 생성된 패스워드를 `EMAIL_PASS`에 설정

---

## ✅ 설정 체크리스트

### Google API (6개)
- [ ] `SPREADSHEET_ID`
- [ ] `TARGET_SHEET_NAME`
- [ ] `DRIVE_FOLDER_ID`
- [ ] `JSON_DRIVE_FOLDER_ID`
- [ ] `SHEETS_SERVICE_ACCOUNT_KEY`
- [ ] `DRIVE_SERVICE_ACCOUNT_KEY`

### GitHub 설정 (4개)
- [ ] `GH_TOKEN`
- [ ] `GH_USERNAME`
- [ ] `GH_REPO`
- [ ] `GH_BRANCH`

### 카카오톡 API (3개)
- [ ] `KAKAO_REST_API_KEY`
- [ ] `KAKAO_ACCESS_TOKEN`
- [ ] `KAKAO_REFRESH_TOKEN`

### 이메일 설정 (3개)
- [ ] `EMAIL_ADDRESS`
- [ ] `EMAIL_PASS`
- [ ] `RECEIVER_EMAIL`

---

## 🧪 테스트 방법

설정 완료 후:

1. **GitHub Actions 탭** 이동
2. **"PDA Partner Daily Run"** 워크플로우 선택
3. **"Run workflow"** 버튼 클릭
4. 수동 실행으로 테스트

---

## 🔗 현재 설정값 참고

```env
# 현재 .env 파일의 값들을 참고하여 설정
SPREADSHEET_ID="19dkwKNW6VshCg3wTemzmbbQlbATfq6brAWluaps1Rm0"
TARGET_SHEET_NAME="출하예정리스트(TEST)"
DRIVE_FOLDER_ID="1Gylm36vhtrl_yCHurZYGgeMlt5U0CliE"
JSON_DRIVE_FOLDER_ID="13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1"
GITHUB_USERNAME="isolhsolfafa"
GITHUB_REPO="gst-factory"
GITHUB_BRANCH="main"
KAKAO_REST_API_KEY="d24d0956456fb747fefea7129a295a4e"
EMAIL_ADDRESS="angdredong@gmail.com"
RECEIVER_EMAIL="kdkyu311@naver.com"
```

**📝 총 15개 Secrets 설정 완료하면 CI/CD 파이프라인이 완벽하게 작동합니다!** 