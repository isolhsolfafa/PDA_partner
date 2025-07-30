"""
파일 업로드 및 관리 기능
"""

import base64
import json
import os
from datetime import datetime, timedelta

import requests
from googleapiclient.http import MediaFileUpload

from .settings import GITHUB_BRANCH, GITHUB_REPO, GITHUB_TOKEN, GITHUB_USERNAME, NOVA_FOLDER_ID, logger


class GitHubUploader:
    """GitHub 업로드 클래스"""

    def __init__(self):
        self.username = GITHUB_USERNAME
        self.repo = GITHUB_REPO
        self.branch = GITHUB_BRANCH
        self.token = GITHUB_TOKEN
        self.headers = {"Authorization": f"token {self.token}", "Accept": "application/vnd.github.v3+json"}

    def upload_file(self, file_path, target_filename=None):
        """파일을 GitHub에 업로드"""
        if target_filename is None:
            target_filename = os.path.basename(file_path)

        logger.info(f"GitHub 업로드 시작 - 사용자: {self.username}, 레포: {self.repo}, 파일: {target_filename}")

        try:
            with open(file_path, "rb") as f:
                content = base64.b64encode(f.read()).decode()

            url = f"https://api.github.com/repos/{self.username}/{self.repo}/contents/{target_filename}"

            # 기존 파일 확인
            get_resp = requests.get(url, headers=self.headers)
            if get_resp.status_code == 404:
                logger.info("파일이 존재하지 않음 - 새 파일 생성")
                sha = None
            elif get_resp.status_code != 200:
                logger.error(f"❌ GitHub 파일 확인 실패: {get_resp.status_code}, {get_resp.json()}")
                return False
            else:
                sha = get_resp.json().get("sha")

            # 파일 업로드
            payload = {"message": f"자동 업로드: {target_filename}", "content": content, "branch": self.branch}
            if sha:
                payload["sha"] = sha

            put_resp = requests.put(url, headers=self.headers, json=payload)

            if put_resp.status_code in (200, 201):
                logger.info(f"✅ GitHub 업로드 성공: {self.username}/{self.repo}/{target_filename}")
                return True

            logger.error(f"❌ GitHub 업로드 실패: {put_resp.status_code}, {put_resp.json()}")
            return False

        except Exception as e:
            logger.error(f"❌ GitHub 업로드 예외 발생: {self.username}/{self.repo}/{target_filename} - {e}")
            return False

    def delete_file(self, filename):
        """GitHub에서 파일 삭제"""
        try:
            url = f"https://api.github.com/repos/{self.username}/{self.repo}/contents/{filename}"

            # 파일 정보 가져오기
            get_resp = requests.get(url, headers=self.headers)
            if get_resp.status_code == 404:
                logger.warning(f"⚠️ 삭제할 파일이 존재하지 않습니다: {filename}")
                return True
            elif get_resp.status_code != 200:
                logger.error(f"❌ GitHub 파일 정보 가져오기 실패: {get_resp.status_code}")
                return False

            sha = get_resp.json()["sha"]

            # 파일 삭제
            delete_payload = {"message": f"자동 삭제: {filename}", "sha": sha, "branch": self.branch}

            delete_resp = requests.delete(url, headers=self.headers, json=delete_payload)

            if delete_resp.status_code == 200:
                logger.info(f"✅ GitHub 파일 삭제 성공: {filename}")
                return True

            logger.error(f"❌ GitHub 파일 삭제 실패: {delete_resp.status_code}")
            return False

        except Exception as e:
            logger.error(f"❌ GitHub 파일 삭제 중 오류 발생: {e}")
            return False


class DriveUploader:
    """Google Drive 업로드 클래스"""

    def __init__(self, drive_service):
        self.drive_service = drive_service

    def upload_file(self, file_path, mime_type=None):
        """파일을 Google Drive에 업로드"""
        try:
            if mime_type is None:
                mime_type = self._guess_mime_type(file_path)

            file_metadata = {"name": os.path.basename(file_path), "parents": [NOVA_FOLDER_ID]}

            media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)

            file = self.drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()

            logger.info(f"✅ Drive 업로드 완료: {file.get('id')}")
            return file.get("id")

        except Exception as e:
            logger.error(f"❌ Drive 업로드 실패: {e}")
            return None

    def delete_old_files(self, query, days_to_keep=30):
        """오래된 파일 삭제"""
        try:
            files = (
                self.drive_service.files()
                .list(q=query, fields="files(id, name, createdTime)")
                .execute()
                .get("files", [])
            )

            if not files:
                logger.info("삭제할 파일이 없습니다.")
                return

            now = datetime.utcnow()
            cutoff = now - timedelta(days=days_to_keep)

            for file in files:
                created_time = datetime.strptime(file["createdTime"], "%Y-%m-%dT%H:%M:%S.%fZ")
                if created_time < cutoff:
                    try:
                        self.drive_service.files().delete(fileId=file["id"]).execute()
                        logger.info(f"✅ 오래된 파일 삭제 완료: {file['name']}")
                    except Exception as e:
                        logger.error(f"❌ 파일 삭제 실패: {file['name']} - {e}")

        except Exception as e:
            logger.error(f"❌ 오래된 파일 삭제 중 오류 발생: {e}")

    def _guess_mime_type(self, file_path):
        """파일 확장자에 따른 MIME 타입 추측"""
        ext = os.path.splitext(file_path)[1].lower()
        mime_types = {
            ".json": "application/json",
            ".html": "text/html",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
            ".pdf": "application/pdf",
            ".txt": "text/plain",
            ".csv": "text/csv",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
        return mime_types.get(ext, "application/octet-stream")
