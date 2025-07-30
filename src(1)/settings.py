"""
설정 및 상수 정의
"""

import logging
import os
from datetime import date
from pathlib import Path

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# .env 파일 로드
load_dotenv()

# 기본 경로 설정
BASE_DIR = Path(__file__).resolve().parent.parent

# Credentials 파일 경로 설정
SHEETS_CREDENTIALS_PATH = BASE_DIR / "config" / "gst-manegemnet-70faf8ce1bff.json"
DRIVE_CREDENTIALS_PATH = BASE_DIR / "config" / "gst-manegemnet-ab8788a05cff.json"

# API 스코프
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

# 환경 변수
NOVA_FOLDER_ID = os.getenv("NOVA_FOLDER_ID")
GITHUB_USERNAME = os.getenv("GITHUB_USERNAME")
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "main")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
KAKAO_TOKEN = os.getenv("KAKAO_TOKEN")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
TEST_MODE = os.getenv("TEST_MODE", "false").lower() == "true"

# 스프레드시트 설정
WORKSHEET_RANGE = os.getenv("WORKSHEET_RANGE", "'WORKSHEET'!A1:Z100")
INFO_RANGE = os.getenv("INFO_RANGE", "'정보판'!A1:Z100")
PROCESS_LIMIT = int(os.getenv("LIMIT", "1"))  # 처리할 스프레드시트 수 제한

# 로깅 설정
logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)


def get_sheets_credentials():
    """Google Sheets API 인증 정보 가져오기"""
    try:
        return Credentials.from_service_account_file(SHEETS_CREDENTIALS_PATH, scopes=SHEETS_SCOPES)
    except Exception as e:
        logger.error(f"❌ Google Sheets API 인증 실패: {e}")
        raise


def get_drive_credentials():
    """Google Drive API 인증 정보 가져오기"""
    try:
        return Credentials.from_service_account_file(DRIVE_CREDENTIALS_PATH, scopes=DRIVE_SCOPES)
    except Exception as e:
        logger.error(f"❌ Google Drive API 인증 실패: {e}")
        raise


def get_google_services():
    """Google 서비스 객체 가져오기"""
    try:
        sheets_creds = get_sheets_credentials()
        drive_creds = get_drive_credentials()

        sheets = build("sheets", "v4", credentials=sheets_creds)
        drive = build("drive", "v3", credentials=drive_creds)

        return sheets, drive
    except Exception as e:
        logger.error(f"❌ Google 서비스 초기화 실패: {e}")
        raise


# 작업 시간 설정
WORK_SCHEDULE = {
    "WORK_START": (8, 0),  # 8:00 AM
    "WORK_END": (20, 0),  # 8:00 PM
    "MAX_DAILY_HOURS": 12,
    "LUNCH_START": (11, 20),  # 11:20 AM
    "LUNCH_END": (12, 20),  # 12:20 PM
    "DINNER_START": (17, 0),  # 5:00 PM
    "DINNER_END": (18, 0),  # 6:00 PM
    "BREAK_1_START": (10, 0),  # 10:00 AM
    "BREAK_1_END": (10, 20),  # 10:20 AM
    "BREAK_2_START": (15, 0),  # 3:00 PM
    "BREAK_2_END": (15, 20),  # 3:20 PM
}

# 2025년 공휴일 설정
HOLIDAYS = [
    date(2025, 1, 1),  # 신정
    date(2025, 1, 27),  # 설날
    date(2025, 1, 28),  # 설날
    date(2025, 1, 29),  # 설날
    date(2025, 3, 1),  # 삼일절
    date(2025, 5, 5),  # 어린이날
    date(2025, 5, 6),  # 대체공휴일
    date(2025, 6, 6),  # 현충일
    date(2025, 8, 15),  # 광복절
    date(2025, 10, 3),  # 개천절
    date(2025, 10, 6),  # 추석
    date(2025, 10, 7),  # 추석
    date(2025, 10, 8),  # 추석
    date(2025, 10, 9),  # 한글날
    date(2025, 12, 25),  # 성탄절
]

# 모델별 기본 작업
DEFAULT_TASKS = {
    "MECHANICAL": [
        "CABINET ASSY",
        "BURNER ASSY(TMS)",
        "WET TANK ASSY(TMS)",
        "3-WAY VALVE ASSY",
        "N2 LINE ASSY",
        "N2 TUBE ASSY",
        "CDA LINE ASSY",
        "CDA TUBE ASSY",
        "BCW LINE ASSY",
        "PCW LINE ASSY",
        "O2 LINE ASSY",
        "LNG LINE ASSY",
        "WASTE GAS LINE ASSY",
        "COOLING UNIT(TMS)",
        "REACTOR ASSY(TMS)",
        "HEATING JACKET",
        "CIR LINE TUBING",
        "설비 CLEANING",
        "자주검사",
    ],
    "ELECTRICAL": [
        "AC 백 판넬 작업",
        "DC 백 판넬 작업",
        "케비넷 준비 작업(덕트, 철거작업)",
        "판넬 취부 및 선분리",
        "내, 외부 작업",
        "탱크 작업",
        "판넬 작업",
        "탱크 도킹 후 결선 작업",
    ],
    "INSPECTION": ["LNG/Util", "Chamber", "I/O 체크, 가동 검사, 전장 마무리"],
    "FINISHING": ["캐비넷 커버 장착 및 포장", "상부 마무리"],
}

# 모델별 작업 시간 계수
MODEL_TIME_FACTORS = {
    "GAIA-I DUAL": 1.8,
    "DRAGON DUAL": 1.8,
    "GAIA-I": 1.0,
    "DRAGON": 1.0,
    "GAIA-II": 1.2,
    "SWS-I": 0.8,
    "GAIA-P DUAL": 0.9,
    "GAIA-P": 0.9,
}

# 기본 작업 시간 (시간 단위)
BASE_WORK_HOURS = {"기구": 120, "전장": 80, "검사": 40, "마무리": 16, "TMS_반제품": 40}
