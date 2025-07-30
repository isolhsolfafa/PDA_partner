import base64
import io
import logging
import os
import re
import smtplib
from datetime import datetime
from email.message import EmailMessage

import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from tenacity import retry, stop_after_attempt, wait_fixed

# 🔍 로깅 설정
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)
handler = logging.FileHandler("script.log")
handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
logger.addHandler(handler)


# 로그 파일 플러시 보장
def flush_log():
    handler.flush()
    os.fsync(handler.stream.fileno())


# 종속성 확인
try:
    import googleapiclient
    import requests
except ImportError as e:
    logger.error(f"필수 라이브러리 누락: {e}")
    flush_log()
    raise

# 🔐 하드코딩된 설정
# 첫 번째 저장소
GITHUB_USERNAME_1 = "isolhsolfafa"
GITHUB_REPO_1 = "GST_Factory_Dashboard"
GITHUB_BRANCH_1 = "main"
GITHUB_TOKEN_1 = os.getenv("GITHUB_TOKEN", "")
GITHUB_UPLOAD_FILENAME_1 = "partner.html"
# 두 번째 저장소
GITHUB_USERNAME_2 = "isolhsolfafa"
GITHUB_REPO_2 = "gst-factory"
GITHUB_BRANCH_2 = "main"
GITHUB_TOKEN_2 = os.getenv("GITHUB_TOKEN", "")
GITHUB_UPLOAD_FILENAME_2 = "public/partner.html"
# 기타 설정
SERVICE_ACCOUNT_FILE = "/Users/kdkyu311/Downloads/gst-manegemnet-70faf8ce1bff.json"
INDEX_FOLDER_ID = "1Gylm36vhtrl_yCHurZYGgeMlt5U0CliE"
NOVA_FOLDER_ID = "13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "angdredong@gmail.com"
EMAIL_PASS = "qyts eiwd kjma exya"
RECEIVER_EMAIL = "kdkyu311@naver.com"

# 📁 파일 경로 변수 (실행 날짜 동적 삽입)
EXECUTION_DATE = datetime.now().strftime("%Y%m%d")
LOCAL_HTML_PATH = f"partner_{EXECUTION_DATE}.html"
DRIVE_UPLOAD_FILENAME = f"partner_{EXECUTION_DATE}.html"

# 📁 구글 드라이브 설정
SCOPES = ["https://www.googleapis.com/auth/drive"]


# ✅ 구글 인증
@retry(stop=stop_after_attempt(3), wait=wait_fixed(2))
def get_drive_service():
    try:
        credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        drive_service = build("drive", "v3", credentials=credentials)
        logger.info("✅ Google Drive 인증 성공")
        flush_log()
        return drive_service
    except Exception as e:
        logger.error(f"❌ Google Drive 인증 실패: {e}")
        flush_log()
        raise


drive_service = get_drive_service()


# ✅ 최신 index.html ID 가져오기
def get_latest_indexed_html_id():
    try:
        results = (
            drive_service.files()
            .list(
                q=f"'{INDEX_FOLDER_ID}' in parents and name='index.html'",
                orderBy="modifiedTime desc",
                fields="files(id, name, modifiedTime)",
            )
            .execute()
        )
        files = results.get("files", [])
        if not files:
            logger.error("❌ index.html 파일을 찾을 수 없습니다.")
            flush_log()
            return None
        logger.info(f"✅ 최신 index.html 파일 ID: {files[0]['id']}")
        flush_log()
        return files[0]["id"]
    except Exception as e:
        logger.error(f"❌ 최신 index.html ID 가져오기 실패: {e}")
        flush_log()
        return None


# ✅ 파일 다운로드
def download_file(file_id, local_path):
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.FileIO(local_path, "wb")
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        logger.info(f"✅ 파일 다운로드 완료: {local_path}")
        flush_log()
    except Exception as e:
        logger.error(f"❌ 파일 다운로드 실패: {e}")
        flush_log()
        raise


# ✅ 최신 NOVA 그래프 이미지 3종 찾기
def get_latest_nova_images():
    try:
        query = f"'{NOVA_FOLDER_ID}' in parents and name contains 'nan_heatmap'"
        response = drive_service.files().list(q=query, orderBy="modifiedTime desc", fields="files(id, name)").execute()
        latest = {}
        for key in ["weekly_partner", "monthly_partner", "monthly_model"]:
            for f in response.get("files", []):
                if re.match(rf".*{key}.*\.png$", f["name"]):
                    latest[key] = {"name": f["name"], "link": f"https://drive.google.com/uc?export=view&id={f['id']}"}
                    logger.info(f"✅ NOVA 그래프 확인됨: {f['name']}")
                    break
            else:
                logger.warning(f"⚠️ NOVA 그래프 누락: {key}")
        return latest
    except Exception as e:
        logger.error(f"❌ NOVA 그래프 검색 실패: {e}")
        return {}


# ✅ HTML에 트렌드 그래프 삽입
def insert_nova_graphs_to_html(html_path, nova_images):
    try:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as f:
            html = f.read()
        marker = re.search(r"📊 대시보드에서 상세 내용 확인하세요!.*?</a>", html)
        if not marker:
            logger.warning("⚠️ 삽입 위치를 찾지 못했습니다. <body> 끝에 삽입합니다.")
            insert_point = html.rfind("</body>")
            if insert_point == -1:
                logger.error("❌ </body> 태그를 찾을 수 없습니다.")
                flush_log()
                return None
        else:
            insert_point = marker.end()

        def graph_line(label, key):
            if key in nova_images:
                return f'{label}: <a href="{nova_images[key]['link']}" target="_blank">그래프 보기</a>'
            return f'{label}: <span style="color:red;">(링크 없음)</span>'

        insert_html = (
            "<p><strong>📊 NOVA 트렌드 그래프</strong></p>"
            "<ul>"
            f"<li>{graph_line('📅 주간 협력사 NaN 히트맵', 'weekly_partner')}</li>"
            f"<li>{graph_line('🗓️ 월간 협력사 NaN 히트맵', 'monthly_partner')}</li>"
            f"<li>{graph_line('📈 월간 모델별 NaN 히트맵', 'monthly_model')}</li>"
            "</ul>"
        )
        modified = html[:insert_point] + insert_html + html[insert_point:]
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(modified)
        logger.info(f"✅ 수정된 HTML 저장 완료: {html_path}")
        flush_log()
        return html_path
    except Exception as e:
        logger.error(f"❌ HTML 수정 실패: {e}")
        flush_log()
        return None


# ✅ 구글 드라이브 업로드
@retry(stop=stop_after_attempt(3), wait=wait_fixed(2))
def upload_file(local_path, filename=DRIVE_UPLOAD_FILENAME):
    try:
        metadata = {"name": filename, "parents": [INDEX_FOLDER_ID], "mimeType": "text/html"}
        media = MediaFileUpload(local_path, mimetype="text/html", resumable=True)
        res = drive_service.files().create(body=metadata, media_body=media, fields="id").execute()
        logger.info(f"✅ 구글 드라이브 업로드 완료: {filename}")
        flush_log()
        return res.get("id")
    except Exception as e:
        logger.error(f"❌ 구글 드라이브 업로드 실패: {e}")
        flush_log()
        return None


# ✅ GitHub 업로드 개선
def upload_to_github(file_path, username, repo, branch, token, filename):
    logger.info(f"GitHub 업로드 시작 - 사용자: {username}, 레포: {repo}, 브랜치: {branch}, 파일: {filename}")
    flush_log()
    try:
        content = base64.b64encode(open(file_path, "rb").read()).decode()
        url = f"https://api.github.com/repos/{username}/{repo}/contents/{filename}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        logger.info(f"GitHub API 요청 URL: {url}")
        flush_log()

        get_resp = requests.get(url, headers=headers)
        logger.info(f"GitHub GET 응답 상태: {get_resp.status_code}")
        flush_log()
        if get_resp.status_code == 404:
            logger.info("파일이 존재하지 않음 - 새 파일 생성")
        elif get_resp.status_code != 200:
            logger.error(f"❌ GitHub 파일 확인 실패: {get_resp.status_code}, {get_resp.json()}")
            flush_log()
            return False

        sha = get_resp.json().get("sha") if get_resp.status_code == 200 else None

        payload = {
            "message": f"자동 업로드: {filename}",
            "content": content,
            "branch": branch,
            **({"sha": sha} if sha else {}),
        }
        put_resp = requests.put(url, headers=headers, json=payload)
        if put_resp.status_code in (200, 201):
            logger.info(f"✅ GitHub 업로드 성공: {username}/{repo}/{filename}")
            flush_log()
            return True
        logger.error(f"❌ GitHub 업로드 실패: {put_resp.status_code}, {put_resp.json()}")
        flush_log()
        return False
    except Exception as e:
        logger.error(f"❌ GitHub 업로드 예외 발생: {username}/{repo}/{filename} - {e}")
        flush_log()
        return False


# ✅ 이메일 발송
def send_email_with_attachment(file_path):
    try:
        msg = EmailMessage()
        msg["Subject"] = f"📊 최신 {DRIVE_UPLOAD_FILENAME}에 NOVA 그래프 반영 완료"
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = RECEIVER_EMAIL
        msg.set_content(
            f"아래 항목이 반영된 {DRIVE_UPLOAD_FILENAME}을 첨부합니다.\n"
            "- NOVA 주간/월간 그래프 반영\n"
            "- 기존 index.html 유지\n"
            "- 대시보드 안내 아래에 그래프 섹션 삽입"
        )
        with open(file_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="text", subtype="html", filename=os.path.basename(file_path))
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_ADDRESS, EMAIL_PASS)
            smtp.send_message(msg)
        logger.info(f"📩 이메일 전송 완료 → {RECEIVER_EMAIL}")
        flush_log()
    except Exception as e:
        logger.error(f"❌ 이메일 전송 실패: {e}")
        flush_log()


# ✅ 실행 흐름
def main():
    try:
        latest_id = get_latest_indexed_html_id()
        if not latest_id:
            return

        download_file(latest_id, LOCAL_HTML_PATH)
        nova_images = get_latest_nova_images()
        modified = insert_nova_graphs_to_html(LOCAL_HTML_PATH, nova_images)
        if not modified:
            return

        upload_id = upload_file(modified)
        if not upload_id:
            return

        send_email_with_attachment(modified)

        # 첫 번째 저장소 업로드
        if not upload_to_github(
            modified, GITHUB_USERNAME_1, GITHUB_REPO_1, GITHUB_BRANCH_1, GITHUB_TOKEN_1, GITHUB_UPLOAD_FILENAME_1
        ):
            logger.error("⚠️ GST_Factory_Dashboard 업로드 실패")
            flush_log()
        else:
            logger.info("✅ GST_Factory_Dashboard 업로드 성공")
            flush_log()

        # 두 번째 저장소 업로드
        if not upload_to_github(
            modified, GITHUB_USERNAME_2, GITHUB_REPO_2, GITHUB_BRANCH_2, GITHUB_TOKEN_2, GITHUB_UPLOAD_FILENAME_2
        ):
            logger.error("⚠️ gst-factory 업로드 실패")
            flush_log()
        else:
            logger.info("✅ gst-factory 업로드 성공")
            flush_log()

        logger.info("✅ 모든 작업이 성공적으로 완료되었습니다!")
        flush_log()
    except Exception as e:
        logger.error(f"❌ 전체 작업 실패: {e}")
        flush_log()


if __name__ == "__main__":
    main()
