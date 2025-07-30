# 🔧 Google 서비스 키 재생성 가이드

## 🚨 문제점
현재 서비스 키에 제어 문자 `^F` (0x06)가 포함되어 JSON 파싱 오류 발생

## ✅ 해결 단계

### 1️⃣ Google Cloud Console에서 새 키 생성

1. **Google Cloud Console** → **IAM 및 관리** → **서비스 계정**
2. 기존 서비스 계정 선택 (또는 새로 생성)
3. **키** 탭 → **키 추가** → **새 키 만들기**
4. **JSON** 선택 → **만들기**
5. 다운로드된 **JSON 파일을 텍스트 에디터로 열기**

### 2️⃣ 키 검증 및 base64 인코딩

```bash
# 1. JSON 파일 유효성 검사
cat downloaded-key.json | jq '.'

# 2. 제어 문자 확인
cat downloaded-key.json | cat -v

# 3. base64 인코딩 (줄바꿈 제거)
cat downloaded-key.json | base64 -w 0 > encoded-key.txt

# 4. 인코딩된 키 복사
cat encoded-key.txt
```

### 3️⃣ GitHub Secrets 업데이트

1. **GitHub 레포지토리** → **Settings** → **Secrets and variables** → **Actions**
2. `SHEETS_SERVICE_ACCOUNT_KEY` **Update**
3. `DRIVE_SERVICE_ACCOUNT_KEY` **Update**
4. 새로 생성한 base64 인코딩된 키 값 붙여넣기

### 4️⃣ 최종 검증

```bash
# GitHub Actions에서 테스트
echo "새로운 키로 워크플로우 실행"
```

## ⚠️ 주의사항

- **반드시 새 키를 생성**하세요 (기존 키 수정 불가)
- **복사-붙여넣기 시 제어 문자 주의**
- **base64 인코딩 시 `-w 0` 옵션으로 줄바꿈 제거**
- **두 키 모두 동일한 방식으로 재생성** 