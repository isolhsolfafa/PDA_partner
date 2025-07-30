"""
알림 기능 (이메일, 카카오톡)
"""

import json
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests

from .settings import KAKAO_TOKEN, SMTP_PASSWORD, SMTP_PORT, SMTP_SERVER, SMTP_USER, logger


class EmailNotifier:
    """이메일 알림 클래스"""

    def __init__(self):
        self.server = SMTP_SERVER
        self.port = SMTP_PORT
        self.username = SMTP_USER
        self.password = SMTP_PASSWORD

    def send_email(self, to_emails, subject, content, image_paths=None):
        """이메일 전송"""
        try:
            # 메시지 생성
            msg = MIMEMultipart()
            msg["Subject"] = subject
            msg["From"] = self.username
            msg["To"] = ", ".join(to_emails)

            # 본문 추가
            msg.attach(MIMEText(content, "html"))

            # 이미지 첨부
            if image_paths:
                for i, img_path in enumerate(image_paths):
                    with open(img_path, "rb") as f:
                        img = MIMEImage(f.read())
                        img.add_header("Content-ID", f"<image{i}>")
                        img.add_header("Content-Disposition", "inline", filename=img_path.split("/")[-1])
                        msg.attach(img)

            # SMTP 서버 연결 및 전송
            with smtplib.SMTP(self.server, self.port) as server:
                server.starttls()
                server.login(self.username, self.password)
                server.send_message(msg)

            logger.info(f"✅ 이메일 전송 완료: {to_emails}")
            return True

        except Exception as e:
            logger.error(f"❌ 이메일 전송 실패: {e}")
            return False


class KakaoNotifier:
    """카카오톡 알림 클래스"""

    def __init__(self):
        self.token = KAKAO_TOKEN
        self.headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/x-www-form-urlencoded"}

    def send_message(self, content):
        """카카오톡 메시지 전송"""
        try:
            url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
            data = {
                "template_object": json.dumps(
                    {"object_type": "text", "text": content, "link": {"web_url": "", "mobile_web_url": ""}}
                )
            }

            response = requests.post(url, headers=self.headers, data=data)

            if response.status_code == 200:
                logger.info("✅ 카카오톡 메시지 전송 완료")
                return True

            logger.error(f"❌ 카카오톡 메시지 전송 실패: {response.status_code}")
            return False

        except Exception as e:
            logger.error(f"❌ 카카오톡 메시지 전송 중 오류 발생: {e}")
            return False

    def send_image(self, image_path, content=None):
        """카카오톡 이미지 전송"""
        try:
            # 이미지 업로드
            upload_url = "https://kapi.kakao.com/v2/api/talk/memo/upload"
            with open(image_path, "rb") as f:
                response = requests.post(upload_url, headers=self.headers, files={"file": f})

            if response.status_code != 200:
                logger.error(f"❌ 카카오톡 이미지 업로드 실패: {response.status_code}")
                return False

            # 이미지 URL 가져오기
            image_url = response.json()["image_url"]

            # 메시지 전송
            send_url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
            template_object = {
                "object_type": "image",
                "image_url": image_url,
                "image_width": 800,
                "image_height": 600,
                "link": {"web_url": "", "mobile_web_url": ""},
            }

            if content:
                template_object["text"] = content

            data = {"template_object": json.dumps(template_object)}

            response = requests.post(send_url, headers=self.headers, data=data)

            if response.status_code == 200:
                logger.info("✅ 카카오톡 이미지 전송 완료")
                return True

            logger.error(f"❌ 카카오톡 이미지 전송 실패: {response.status_code}")
            return False

        except Exception as e:
            logger.error(f"❌ 카카오톡 이미지 전송 중 오류 발생: {e}")
            return False
