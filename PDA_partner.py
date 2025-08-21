import base64
import copy
import json
import os
import re
import smtplib
import ssl
import time as systime
from datetime import date, datetime, time, timedelta
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from itertools import islice

import certifi
import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pytz
import requests
import seaborn as sns
from backoff import expo, on_exception  # 지수 백오프 추가
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from matplotlib.patches import Patch
from oauth2client.service_account import ServiceAccountCredentials

# 환경변수 로딩
try:
    from dotenv import load_dotenv

    load_dotenv()  # .env 파일에서 환경변수 로드
    print("✅ .env 파일에서 환경변수를 로드했습니다.")
except ImportError:
    print("⚠️ python-dotenv가 설치되지 않았습니다. 시스템 환경변수를 사용합니다.")
    print("   설치 방법: pip install python-dotenv")

# 로깅 설정
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("pda_partner.log"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# SSL 환경 변수 설정
os.environ["SSL_CERT_FILE"] = certifi.where()

# Global Variables - 보안 개선: 환경변수 사용
TEST_MODE = os.getenv("TEST_MODE", "True").lower() == "true"
DRIVE_FOLDER_ID = os.getenv(
    "DRIVE_FOLDER_ID", "1Gylm36vhtrl_yCHurZYGgeMlt5U0CliE"
)  # 최종 리포트(HTML), 그래프 등 저장용
JSON_DRIVE_FOLDER_ID = os.getenv(
    "JSON_DRIVE_FOLDER_ID", "13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1"
)  # JSON 데이터 저장용

# 환경변수 디버깅 로그
print(f"🔍 [DEBUG] DRIVE_FOLDER_ID 환경변수: '{DRIVE_FOLDER_ID}'")
print(f"🔍 [DEBUG] JSON_DRIVE_FOLDER_ID 환경변수: '{JSON_DRIVE_FOLDER_ID}'")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
# 이메일 설정 - 기존 .env 파일 호환성 지원
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS") or os.getenv("SMTP_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS") or os.getenv("SMTP_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

# 이메일 설정 검증 (선택사항)
email_configured = EMAIL_ADDRESS and EMAIL_PASS and RECEIVER_EMAIL
if not email_configured:
    print(f"⚠️ 이메일 설정 확인 (선택사항):")
    print(f"   EMAIL_ADDRESS/SMTP_USER: {'✅' if EMAIL_ADDRESS else '❌'}")
    print(f"   EMAIL_PASS/SMTP_PASSWORD: {'✅' if EMAIL_PASS else '❌'}")
    print(f"   RECEIVER_EMAIL: {'✅' if RECEIVER_EMAIL else '❌'}")
    print(f"   📧 이메일 기능이 비활성화됩니다.")
else:
    print(f"✅ 이메일 설정이 완료되었습니다.")

# Sheet Range Settings
WORKSHEET_RANGE = os.getenv("WORKSHEET_RANGE", "'WORKSHEET'!A1:Z100")
INFO_RANGE = os.getenv("INFO_RANGE", "정보판!A1:Z100")

# API 키들 - 환경변수에서 로드
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
REST_API_KEY = os.getenv("KAKAO_REST_API_KEY")
KAKAO_ACCESS_TOKEN = os.getenv("KAKAO_ACCESS_TOKEN")
REFRESH_TOKEN = os.getenv("KAKAO_REFRESH_TOKEN")

if not GITHUB_TOKEN or not REST_API_KEY:
    print("⚠️ 일부 API 키가 설정되지 않았습니다. 해당 기능이 제한될 수 있습니다.")

if not KAKAO_ACCESS_TOKEN or not REFRESH_TOKEN:
    print(
        "⚠️ 카카오톡 토큰이 설정되지 않았습니다. 카카오톡 알림 기능이 제한될 수 있습니다."
    )

# 서비스 계정 키 파일 경로
sheets_json_key_path = os.getenv("SHEETS_KEY_PATH")
drive_json_key_path = os.getenv("DRIVE_KEY_PATH")
spreadsheet_id = os.getenv(
    "SPREADSHEET_ID", "19dkwKNW6VshCg3wTemzmbbQlbATfq6brAWluaps1Rm0"
)
folder_id = DRIVE_FOLDER_ID

SCOPES_SHEETS = ["https://www.googleapis.com/auth/spreadsheets"]
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive"]

# 서비스 계정 인증 및 서비스 초기화 (에러 처리 강화)
try:
    if not sheets_json_key_path:
        raise ValueError("SHEETS_KEY_PATH 환경변수가 설정되지 않았습니다.")
    if not drive_json_key_path:
        raise ValueError("DRIVE_KEY_PATH 환경변수가 설정되지 않았습니다.")
    if not os.path.exists(sheets_json_key_path):
        raise FileNotFoundError(
            f"Sheets 서비스 계정 키 파일을 찾을 수 없습니다: {sheets_json_key_path}"
        )
    if not os.path.exists(drive_json_key_path):
        raise FileNotFoundError(
            f"Drive 서비스 계정 키 파일을 찾을 수 없습니다: {drive_json_key_path}"
        )

    sheets_credentials = Credentials.from_service_account_file(
        sheets_json_key_path, scopes=SCOPES_SHEETS
    )
    drive_credentials = Credentials.from_service_account_file(
        drive_json_key_path, scopes=SCOPES_DRIVE
    )

    sheets_service = build("sheets", "v4", credentials=sheets_credentials)
    drive_service = build("drive", "v3", credentials=drive_credentials)

    print("✅ Google API 서비스 초기화 완료")
except Exception as e:
    print(f"❌ Google API 서비스 초기화 실패: {e}")
    raise

# Font Setting
font_paths = [
    # macOS 경로들
    "/Users/kdkyu311/Library/Fonts/NanumGothic.ttf",
    "/System/Library/AssetsV2/com_apple_MobileAsset_Font7/bad9b4bf17cf1669dde54184ba4431c22dcad27b.asset/AssetData/NanumGothic.ttc",
    "/Library/Fonts/NanumGothic.ttf",
    "/System/Library/Fonts/Supplemental/NanumGothic.ttf",
    # Ubuntu/Linux 경로들 (GitHub Actions용)
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
    "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
    "/usr/share/fonts/opentype/nanum/NanumGothic.ttf",
    # 일반적인 Linux 시스템 경로들
    "/usr/share/fonts/TTF/NanumGothic.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",  # 대체 폰트
]
font_path = next((path for path in font_paths if os.path.exists(path)), None)

if font_path:
    print(f"✅ NanumGothic 폰트 적용 완료: {font_path}")
    font_prop = fm.FontProperties(fname=font_path)
    plt.rc("font", family=font_prop.get_name())
else:
    # 시스템에서 설치된 한글 폰트를 동적으로 찾기
    print("🔍 시스템에서 한글 폰트를 검색 중...")
    available_fonts = [f.name for f in fm.fontManager.ttflist]
    korean_fonts = [
        font
        for font in available_fonts
        if any(
            keyword in font.lower()
            for keyword in ["nanum", "malgun", "dotum", "gulim", "batang"]
        )
    ]

    if korean_fonts:
        selected_font = korean_fonts[0]
        print(f"✅ 한글 폰트 발견 및 적용: {selected_font}")
        font_prop = fm.FontProperties(family=selected_font)
        plt.rc("font", family=selected_font)
    else:
        print("🚨 한글 폰트를 찾을 수 없습니다. 기본 폰트로 진행합니다.")
        font_prop = None

# 마이너스 기호 깨짐 방지
plt.rcParams["axes.unicode_minus"] = False


# 백오프 데코레이터 설정 - 429 (Rate Limit) 에러도 재시도하도록 수정
@on_exception(expo, HttpError, max_tries=10, max_time=300, giveup=lambda e: getattr(e, "response", None) and e.response.status_code not in [429, 503, 500, 502, 504])  # type: ignore
def api_call_with_backoff(func, *args, **kwargs):
    try:
        return func(*args, **kwargs)
    except HttpError as e:
        if getattr(e, "response", None) and e.response.status_code == 429:
            print(f"⚠️ [Rate Limit] API 할당량 초과, 백오프 재시도: {e}")
        else:
            print(f"⚠️ [Retrying] API 호출 실패: {e}")
        raise


# --------------------------
# 함수: 시트 이름으로 sheetId 가져오기
def get_sheet_id_by_name(spreadsheet_id, sheet_name):
    metadata = api_call_with_backoff(
        sheets_service.spreadsheets().get, spreadsheetId=spreadsheet_id
    ).execute()
    for sheet in metadata.get("sheets", []):
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    raise ValueError(f"시트 '{sheet_name}'을(를) 찾을 수 없습니다.")


# 환경변수에서 TARGET_SHEET_NAME 읽기
TARGET_SHEET_NAME = os.getenv("TARGET_SHEET_NAME", "출하예정리스트(TEST)")
print(f"📋 사용할 시트 이름: {TARGET_SHEET_NAME}")
TARGET_SHEET_ID = get_sheet_id_by_name(spreadsheet_id, TARGET_SHEET_NAME)
# --------------------------


# 함수: 전체 시트 데이터 한 번만 가져오기 (TARGET_SHEET_NAME!A:AA)
def fetch_entire_sheet_values(spreadsheet_id, sheet_range=None):
    if sheet_range is None:
        sheet_range = f"'{TARGET_SHEET_NAME}'!A:AA"
    result = api_call_with_backoff(
        sheets_service.spreadsheets().values().get,
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
    ).execute()
    return result.get("values", [])


# HTML & Drive Upload Functions
def generate_html_from_content(html_content, output_filename="index.html"):
    styled_html = f"""
    <html>
    <head>
      <meta charset="UTF-8">
      <title>PDA Dashboard</title>
      <style>
        body {{ font-family: 'NanumGothic', sans-serif; font-size: 12px; margin: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        details {{ margin-bottom: 10px; }}
        summary {{ cursor: pointer; font-weight: bold; color: #333; }}
        details[open] summary {{ color: #0056b3; }}
        details > *:not(summary) {{ display: block; margin-left: 20px; }}
        ul {{ margin: 5px 0; padding-left: 20px; }}
        p {{ margin: 5px 0; }}
        hr {{ margin: 20px 0; }}
        a {{ color: #0056b3; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        img {{ max-width: 100%; height: auto; }}
      </style>
    </head>
    <body>{html_content}</body></html>
    """
    try:
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(styled_html)
        print(f"📄 HTML 파일 생성 완료: {output_filename}")
        html_link = upload_to_drive(output_filename)
        if html_link:
            print(f"✅ HTML 업로드 완료, 링크: {html_link}")
        return output_filename, html_link
    except Exception as e:
        print(f"❌ HTML 생성 오류: {e}")
        return None, None


def upload_to_drive(file_path, drive_service_param=None):
    import os

    from googleapiclient.http import MediaFileUpload

    try:
        # 파일명만 추출 (경로 제거)
        file_name = os.path.basename(file_path)

        # 전역 변수 가져오기
        global DRIVE_FOLDER_ID, drive_service

        print(
            f"🔍 [DEBUG] Drive 업로드 시도 - 파일: {file_name}, 폴더 ID: {DRIVE_FOLDER_ID}"
        )

        # 폴더 ID 검증
        if not DRIVE_FOLDER_ID or DRIVE_FOLDER_ID.strip() == "":
            print(f"❌ [Drive 업로드 오류] 폴더 ID가 비어있습니다: '{DRIVE_FOLDER_ID}'")
            return None

        file_metadata = {"name": file_name, "parents": [DRIVE_FOLDER_ID]}
        mime_type = (
            "text/html"
            if file_path.endswith(".html")
            else (
                "image/png"
                if file_path.endswith(".png")
                else "application/octet-stream"
            )
        )
        media = MediaFileUpload(file_path, mimetype=mime_type)

        # drive_service 파라미터 우선 사용, 없으면 전역 변수 사용
        service = drive_service_param if drive_service_param else drive_service

        file = api_call_with_backoff(
            service.files().create, body=file_metadata, media_body=media, fields="id"
        ).execute()
        file_id = file.get("id")
        api_call_with_backoff(
            service.permissions().create,
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
        ).execute()
        image_url = f"https://drive.google.com/uc?export=view&id={file_id}"
        print(f"✅ Drive 업로드 완료: {file_name} -> {image_url}")
        return image_url
    except Exception as e:
        print(f"❌ [Drive 업로드 오류] {file_path}: {e}")
        return None


# Spreadsheet Functions with Batch Processing and reduced read calls
def get_spreadsheet_title(spreadsheet_id):
    try:
        info = api_call_with_backoff(
            sheets_service.spreadsheets().get,
            spreadsheetId=spreadsheet_id,
            fields="properties.title",
        ).execute()
        return info["properties"]["title"]
    except Exception as e:
        print(f"❌ [오류] 스프레드시트 제목 가져오기 실패: {spreadsheet_id} -> {e}")
        return f"Unknown_{spreadsheet_id}"


def get_order_no(spreadsheet_id):
    return get_spreadsheet_title(spreadsheet_id)


def get_linked_spreadsheet_ids(spreadsheet_id):
    """하이퍼링크에서 스프레드시트 ID 추출 (Rate Limit 방지)"""
    import time

    pmmd_hyperlink_range = f"'{TARGET_SHEET_NAME}'!A:A"
    print(f"🔍 스프레드시트 ID 추출 중... (Rate Limit 방지를 위해 천천히 진행)")

    # Rate Limit 방지를 위한 지연
    time.sleep(2)

    result = api_call_with_backoff(
        sheets_service.spreadsheets().values().get,
        spreadsheetId=spreadsheet_id,
        range=pmmd_hyperlink_range,
        valueRenderOption="FORMULA",
    ).execute()

    formulas = result.get("values", [])
    linked_spreadsheet_ids = [
        re.search(r"/d/([a-zA-Z0-9-_]+)", cell).group(1)
        for row in formulas
        for cell in row
        if cell.startswith("=HYPERLINK(")
    ]

    print(
        f"추출된 스프레드시트 ID들: {linked_spreadsheet_ids[:5]}{'...' if len(linked_spreadsheet_ids) > 5 else ''} (총 {len(linked_spreadsheet_ids)}개)"
    )
    return linked_spreadsheet_ids


# 스프레드시트 ID 추출 실행
linked_spreadsheet_ids = get_linked_spreadsheet_ids(spreadsheet_id)

# Work Schedule Variables
holidays = [
    date(2025, 1, 1),
    date(2025, 1, 27),
    date(2025, 1, 28),
    date(2025, 1, 29),
    date(2025, 3, 1),
    date(2025, 5, 5),
    date(2025, 5, 6),
    date(2025, 6, 6),
    date(2025, 8, 15),
    date(2025, 10, 3),
    date(2025, 10, 6),
    date(2025, 10, 7),
    date(2025, 10, 8),
    date(2025, 10, 9),
    date(2025, 12, 25),
]
WORK_START, WORK_END, MAX_DAILY_HOURS = time(8, 0, 0), time(20, 0, 0), 12
LUNCH_START, LUNCH_END = time(11, 20, 0), time(12, 20, 0)
DINNER_START, DINNER_END = time(17, 0, 0), time(18, 0, 0)
BREAK_1_START, BREAK_1_END = time(10, 0, 0), time(10, 20, 0)
BREAK_2_START, BREAK_2_END = time(15, 0, 0), time(15, 20, 0)


# ====================================
# Utility Functions
# ====================================
def parse_korean_datetime(dt_input):
    if isinstance(dt_input, (int, float)):
        try:
            return pd.to_datetime(dt_input, unit="D", origin="1899-12-30")
        except:
            return pd.NaT
    elif isinstance(dt_input, str):
        s = dt_input.strip()
        if not s:
            return pd.NaT
        if re.match(r"^\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\s*$", s):
            s += " 00:00:00"
        s = s.replace("오전", "AM").replace("오후", "PM")
        fmt = (
            "%Y. %m. %d %p %I:%M:%S"
            if "AM" in s or "PM" in s
            else "%Y. %m. %d %H:%M:%S"
        )
        return pd.to_datetime(s, format=fmt, errors="coerce")
    return pd.NaT


# 중복 함수 제거됨 - 아래의 개선된 버전 사용


def calculate_working_hours_with_holidays(start_time, end_time):
    if pd.isna(start_time) or pd.isna(end_time):
        return 0
    total_hours = 0
    current_time = start_time
    while current_time < end_time:
        day_of_week = current_time.weekday()
        work_start = datetime.combine(
            current_time.date(), time(8, 0, 0) if day_of_week in [5, 6] else WORK_START
        )
        work_end = datetime.combine(
            current_time.date(), time(17, 0, 0) if day_of_week in [5, 6] else WORK_END
        )
        work_start = max(work_start, current_time)
        work_end = min(work_end, end_time)
        daily_hours = (work_end - work_start).total_seconds() / 3600
        breaks = [
            (LUNCH_START, LUNCH_END),
            (BREAK_1_START, BREAK_1_END),
            (BREAK_2_START, BREAK_2_END),
            (DINNER_START, DINNER_END),
        ]
        for b_start, b_end in breaks:
            b_start_dt = datetime.combine(current_time.date(), b_start)
            b_end_dt = datetime.combine(current_time.date(), b_end)
            if work_start < b_end_dt and b_start_dt < work_end:
                daily_hours -= (
                    min(work_end, b_end_dt) - max(work_start, b_start_dt)
                ).total_seconds() / 3600
        total_hours += min(daily_hours, 9 if day_of_week in [5, 6] else MAX_DAILY_HOURS)
        current_time = datetime.combine(
            current_time.date() + timedelta(days=1), WORK_START
        )
    return total_hours


def process_data(df_use, model_name):
    df_complete = df_use.dropna(subset=["시작 시간", "완료 시간"]).copy()
    df_complete["워킹데이 소요 시간"] = df_complete.apply(
        lambda row: calculate_working_hours_with_holidays(
            row["시작 시간"], row["완료 시간"]
        ),
        axis=1,
    )
    df_complete["작업 분류"] = df_complete["내용"].apply(
        lambda x: classify_task(x, model_name)
    )
    task_total_time = (
        df_complete.groupby("내용")["워킹데이 소요 시간"].sum().reset_index()
    )
    task_total_time["총 워킹 소요 시간 (시간:분)"] = task_total_time[
        "워킹데이 소요 시간"
    ].apply(format_hours)
    return task_total_time.sort_values("워킹데이 소요 시간", ascending=True)


def format_hours(decimal_hours):
    hours = int(decimal_hours)
    minutes = int(round((decimal_hours - hours) * 60))
    return f"{hours}h {minutes}m"


# ====================================
# Task Classification
# ====================================
default_mechanical_tasks = [
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
]
default_electrical_tasks = [
    "AC 백 판넬 작업",
    "DC 백 판넬 작업",
    "케비넷 준비 작업(덕트, 철거작업)",
    "판넬 취부 및 선분리",
    "내, 외부 작업",
    "탱크 작업",
    "판넬 작업",
    "탱크 도킹 후 결선 작업",
]
default_inspection_tasks = ["LNG/Util", "Chamber", "I/O 체크, 가동 검사, 전장 마무리"]
default_finishing_tasks = ["캐비넷 커버 장착 및 포장", "상부 마무리"]

model_mechanical_tasks = {
    "GAIA-I DUAL": default_mechanical_tasks,
    "GAIA-I": default_mechanical_tasks,
    "DRAGON": default_mechanical_tasks,
    "DRAGON DUAL": default_mechanical_tasks,
    "GAIA-II": default_mechanical_tasks,
    "SWS-I": [
        "CABINET ASSY",
        "BURNER ASSY(TMS)",
        "WET TANK ASSY(TMS)",
        "3-WAY VALVE ASSY",
        "N2 LINE ASSY",
        "N2 TUBE ASSY",
        "BCW LINE ASSY",
        "WASTE GAS LINE ASSY",
        "COOLING UNIT(TMS)",
        "REACTOR ASSY(TMS)",
        "HEATING JACKET",
        "CIR LINE TUBING",
        "설비 CLEANING",
        "자주검사",
    ],
    "GAIA-P DUAL": [
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
        "WASTE GAS LINE ASSY",
        "COOLING UNIT(TMS)",
        "REACTOR ASSY(TMS)",
        "HEATING JACKET",
        "CIR LINE TUBING",
        "설비 CLEANING",
        "자주검사",
    ],
    "GAIA-P": [
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
        "WASTE GAS LINE ASSY",
        "COOLING UNIT(TMS)",
        "REACTOR ASSY(TMS)",
        "HEATING JACKET",
        "CIR LINE TUBING",
        "설비 CLEANING",
        "자주검사",
    ],
    "IVAS": [
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
}


def get_mechanical_tasks(model_name):
    return model_mechanical_tasks.get(model_name.upper(), default_mechanical_tasks)


def classify_task(content, model_name):
    model_name = model_name.upper()
    tms_tasks = [
        "BURNER ASSY(TMS)",
        "WET TANK ASSY(TMS)",
        "COOLING UNIT(TMS)",
        "REACTOR ASSY(TMS)",
    ]

    # DRAGON, DRAGON DUAL, SWS-I 모델에서는 tms_tasks도 기구로 분류
    if model_name in [
        "DRAGON",
        "DRAGON DUAL",
        "SWS-I",
    ] and content in get_mechanical_tasks(model_name):
        return "기구"

    if content in tms_tasks and model_name not in ["DRAGON", "DRAGON DUAL", "SWS-I"]:
        return "TMS_반제품"
    elif content in get_mechanical_tasks(model_name):
        return "기구"
    elif content in default_electrical_tasks:
        return "전장"
    elif content in default_inspection_tasks:
        return "검사"
    elif content in default_finishing_tasks:
        return "마무리"
    return "기타"


def calculate_progress_by_category(df, model_name):
    df = df.copy()
    df["작업 분류"] = df["내용"].apply(lambda x: classify_task(x, model_name))
    df["진행율"] = df.apply(
        lambda row: (
            100.0
            if pd.isna(row["진행율"])
            and pd.notna(row["시작 시간"])
            and pd.notna(row["완료 시간"])
            else row["진행율"]
        ),
        axis=1,
    )
    df_valid = df.dropna(subset=["내용"])
    df_max = df_valid.groupby(["내용", "작업 분류"])["진행율"].max().reset_index()
    progress_summary = {}
    for category in ["기구", "전장", "TMS_반제품"]:
        df_cat = df_max[df_max["작업 분류"] == category]
        total_tasks = len(df_cat)
        completed = df_cat[df_cat["진행율"] == 100.0]
        progress = (len(completed) / total_tasks * 100) if total_tasks > 0 else 0
        progress_summary[category] = round(progress, 1)
    return progress_summary


def parse_avg_time_string(s):
    s = s.lower().strip()
    match = re.match(r"(?:(\d+)\s*h)?\s*(?:(\d+)\s*m)?", s)
    hours = int(match.group(1)) if match and match.group(1) else 0
    minutes = int(match.group(2)) if match and match.group(2) else 0
    return hours + minutes / 60.0


def get_avg_time_mapping(model_name):
    avg_spreadsheet_id = "1PHKsQ-3kcyaB9HdJqdaLN4siqnHRE8FC7XzGR2pmoLc"
    sheet_range = f"'{model_name.strip()}'!A:B"
    try:
        avg_values = (
            sheets_service.spreadsheets()
            .values()
            .get(
                spreadsheetId=avg_spreadsheet_id,
                range=sheet_range,
                valueRenderOption="FORMATTED_VALUE",
            )
            .execute()
            .get("values", [])
        )
        if len(avg_values) <= 1:
            return {}
        return {
            row[0].strip(): (
                parse_avg_time_string(row[1])
                if "h" in row[1].lower() or "m" in row[1].lower()
                else float(row[1])
            )
            for row in avg_values[1:]
            if len(row) >= 2
        }
    except Exception as e:
        print(f"[오류] AVDATA에서 '{model_name}' 시트를 읽는 중 오류 발생: {e}")
        return {}


# ====================================
# Graph Functions
# ====================================
def generate_and_save_graph(task_total_time, order_no, model_name):
    avg_mapping = get_avg_time_mapping(model_name)
    fig, ax = plt.subplots(figsize=(12, 8))
    bars = ax.barh(
        task_total_time["내용"], task_total_time["워킹데이 소요 시간"], color="skyblue"
    )
    ax.set_yticks(range(len(task_total_time)))
    ax.set_yticklabels(
        [
            (
                f"{task} (평균: {format_hours(avg_mapping[task])})"
                if task in avg_mapping
                else task
            )
            for task in task_total_time["내용"]
        ],
        fontsize=10,
        fontweight="bold",
        color="blue",
    )
    for bar in bars:
        width = bar.get_width()
        ax.text(
            width + 0.2,
            bar.get_y() + bar.get_height() / 2,
            f"{int(width)}h {int(round((width - int(width)) * 60))}m",
            va="center",
            ha="left",
            color="black",
        )
    ax.set_xlim(0, max(bar.get_width() for bar in bars) + 1)
    ax.set_xlabel("Working Hours")
    ax.set_title(f"Total Working Hours per Task for {order_no}")
    plt.tight_layout()
    file_name = f"Total_Working_Hours_{order_no.replace('/', '_')}_{model_name}.png"
    plt.savefig(file_name)
    plt.close()
    return file_name


def generate_legend_chart(task_total_time, order_no, model_name):
    avg_mapping = get_avg_time_mapping(model_name)
    task_total_time["작업 분류"] = task_total_time["내용"].apply(
        lambda x: classify_task(x, model_name)
    )
    task_total_time_sorted = task_total_time.sort_values(
        by=["작업 분류", "워킹데이 소요 시간"], ascending=[True, False]
    )
    category_totals = task_total_time_sorted.groupby("작업 분류")[
        "워킹데이 소요 시간"
    ].sum()
    total_time = category_totals.sum()
    category_colors = {
        "기구": "blue",
        "TMS_반제품": "cyan",
        "전장": "orange",
        "검사": "green",
        "마무리": "red",
        "기타": "gray",
    }
    legend_elements = [
        Patch(
            facecolor="black",
            edgecolor="black",
            label=f"총 소요시간: {format_hours(total_time)}",
        )
    ]
    for category, color in category_colors.items():
        if category in category_totals:
            category_time = category_totals[category]
            legend_elements.append(
                Patch(
                    facecolor=color,
                    edgecolor="black",
                    label=f"{category} (총 {format_hours(category_time)})",
                )
            )
            for _, row in task_total_time_sorted[
                task_total_time_sorted["작업 분류"] == category
            ].iterrows():
                avg_str = (
                    f" (평균: {format_hours(avg_mapping[row['내용']])})"
                    if row["내용"] in avg_mapping
                    else ""
                )
                legend_elements.append(
                    Patch(
                        facecolor="white",
                        edgecolor=color,
                        label=f"  {row['내용']}: {row['총 워킹 소요 시간 (시간:분)']}{avg_str}",
                    )
                )
    plt.figure(figsize=(8, len(legend_elements) * 0.3))
    legend = plt.legend(
        handles=legend_elements,
        loc="center",
        fontsize=10,
        title="작업 분류 및 작업별 소요 시간 (내림차순)",
        frameon=False,
    )
    plt.axis("off")
    plt.title(f"{order_no}", fontsize=12, loc="center", color="black")
    if legend.get_texts():
        legend.get_texts()[0].set_color("red")
    file_name = f"Legend_Chart_{order_no.replace('/', '_')}_{model_name}.png"
    plt.savefig(file_name)
    plt.close()
    return file_name


def generate_and_save_graph_wd(task_total_time, df, order_no, model_name):
    df_valid = df.dropna(subset=["시작 시간", "완료 시간"])
    plt.figure(figsize=(16, 10))
    colors = plt.cm.tab20.colors
    time_offset = pd.Timedelta(hours=1)
    for index, task in enumerate(task_total_time["내용"]):
        group = df_valid[df_valid["내용"] == task].sort_values("시작 시간")
        if group.empty:
            continue
        total_duration_text = task_total_time.loc[
            task_total_time["내용"] == task, "총 워킹 소요 시간 (시간:분)"
        ].values[0]
        for i, (_, row) in enumerate(group.iterrows()):
            start_time = row["시작 시간"] + (i * time_offset)
            end_time = row["완료 시간"] + (i * time_offset)
            plt.plot(
                [start_time, end_time],
                [index, index],
                color=colors[index % len(colors)],
                linewidth=3,
                marker="o",
            )
        plt.text(
            group["완료 시간"].max() + pd.Timedelta(hours=2),
            index,
            total_duration_text,
            va="center",
            fontsize=10,
            color="black",
        )
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))
    plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=1))
    plt.xticks(rotation=45)
    plt.yticks(range(len(task_total_time["내용"])), task_total_time["내용"])
    plt.xlabel("Date")
    plt.ylabel("Tasks")
    plt.title(f"Task Time Chart with Total WD Duration - {order_no}")
    plt.grid(axis="x", linestyle="--", alpha=0.7)
    plt.tight_layout()
    file_name = f"WD_Working_Hours_{order_no.replace('/', '_')}_{model_name}.png"
    plt.savefig(file_name)
    plt.close()
    return file_name


# Utility Functions
def fetch_data_from_sheets(spreadsheet_id, sheet_range):
    result = api_call_with_backoff(
        sheets_service.spreadsheets().values().get,
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
    ).execute()
    values = result.get("values", [])
    if not values or len(values) <= 7:
        raise ValueError(f"'{sheet_range}'에 충분한 데이터가 없습니다.")
    header = values[6]
    data = values[7:]
    max_cols = max(len(header), max((len(row) for row in data), default=0))
    if max_cols > len(header):
        header.extend([""] * (max_cols - len(header)))
    adjusted_data = [
        row + [""] * (max_cols - len(row)) if len(row) < max_cols else row[:max_cols]
        for row in data
    ]
    df_raw = pd.DataFrame(adjusted_data, columns=header[:max_cols])
    needed_cols = ["내용", "시작 시간", "완료 시간", "진행율"]
    if not all(col in df_raw.columns for col in needed_cols):
        raise ValueError(f"필요한 컬럼 {needed_cols}이(가) '{sheet_range}'에 없습니다.")
    df_use = df_raw[needed_cols].copy()
    df_use["시작 시간"] = df_use["시작 시간"].apply(parse_korean_datetime)
    df_use["완료 시간"] = df_use["완료 시간"].apply(parse_korean_datetime)
    # 진행율 float 변환에 방어코드 및 로깅 추가
    try:
        df_use["진행율"] = (
            df_use["진행율"]
            .astype(str)
            .str.replace("%", "")
            .str.strip()
            .replace("", np.nan)
            .astype(float)
        )
    except Exception as e:
        print(f"[진행율 변환 오류] {e}")
        print("진행율 값 목록:", df_use["진행율"].unique())
        raise
    return df_use[df_use["내용"].notna()]


def fetch_info_board_extended(spreadsheet_id):
    ranges = [
        ("정보판!D4", "model_name"),
        ("정보판!B5", "mech_partner"),
        ("정보판!D5", "elec_partner"),
    ]
    batch_request = (
        sheets_service.spreadsheets()
        .values()
        .batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=[rng for rng, _ in ranges],
            valueRenderOption="FORMATTED_VALUE",
        )
    )
    result = api_call_with_backoff(batch_request.execute)
    results = {}
    for (rng, key), response in zip(ranges, result.get("valueRanges", [])):
        values = response.get("values", [[]])
        results[key] = values[0][0].strip() if values else "미정"
    print(
        f"📌 [디버깅] 모델명: {results['model_name']}, 기구협력사: {results['mech_partner']}, 전장협력사: {results['elec_partner']}"
    )
    return (
        results["model_name"] or "NoValue",
        results["mech_partner"],
        results["elec_partner"],
    )


def batch_update_spreadsheet(spreadsheet_id, requests):
    body = {"requests": requests}
    api_call_with_backoff(
        sheets_service.spreadsheets().batchUpdate,
        spreadsheetId=spreadsheet_id,
        body=body,
    ).execute()


# --- 수정된 업데이트 함수 (sheet_values 전달) ---
def update_spreadsheet_with_product_name(
    spreadsheet_id, order_no, product_name, sheet_values
):
    if not product_name or product_name == "NoValue":
        print(f"⚠️ 제품명이 비어 있음 (Row: {order_no})")
        return
    if not sheet_values:
        print("🚨 스프레드시트 데이터가 없습니다.")
        return
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            requests.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": TARGET_SHEET_ID,
                            "startRowIndex": i - 1,
                            "endRowIndex": i,
                            "startColumnIndex": 3,
                            "endColumnIndex": 4,
                        },
                        "rows": [
                            {
                                "values": [
                                    {"userEnteredValue": {"stringValue": product_name}}
                                ]
                            }
                        ],
                        "fields": "userEnteredValue",
                    }
                }
            )
    if requests:
        batch_update_spreadsheet(spreadsheet_id, requests)
        print(
            f"✅ {TARGET_SHEET_NAME}에서 Order No '{order_no}'의 제품명이 업데이트되었습니다."
        )


def update_spreadsheet_with_total_time(
    spreadsheet_id, order_no, total_time, sheet_values
):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            requests.append(
                (
                    {
                        "sheetId": TARGET_SHEET_ID,
                        "startRowIndex": i - 1,
                        "endRowIndex": i,
                        "startColumnIndex": 22,
                        "endColumnIndex": 23,
                    },
                    total_time,
                )
            )
    batch_update_spreadsheet_values(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 총 소요시간이 업데이트되었습니다.")


def update_spreadsheet_with_mechanical_time(
    spreadsheet_id, order_no, mechanical_time, sheet_values
):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            requests.append(
                (
                    {
                        "sheetId": TARGET_SHEET_ID,
                        "startRowIndex": i - 1,
                        "endRowIndex": i,
                        "startColumnIndex": 23,
                        "endColumnIndex": 24,
                    },
                    mechanical_time,
                )
            )
    batch_update_spreadsheet_values(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 기구작업 소요시간이 업데이트되었습니다.")


def update_spreadsheet_with_electrical_time(
    spreadsheet_id, order_no, electrical_time, sheet_values
):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            requests.append(
                (
                    {
                        "sheetId": TARGET_SHEET_ID,
                        "startRowIndex": i - 1,
                        "endRowIndex": i,
                        "startColumnIndex": 24,
                        "endColumnIndex": 25,
                    },
                    electrical_time,
                )
            )
    batch_update_spreadsheet_values(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 전장 작업 시간이 업데이트되었습니다.")


def update_spreadsheet_with_inspection_time(
    spreadsheet_id, order_no, inspection_time, sheet_values
):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            requests.append(
                (
                    {
                        "sheetId": TARGET_SHEET_ID,
                        "startRowIndex": i - 1,
                        "endRowIndex": i,
                        "startColumnIndex": 25,
                        "endColumnIndex": 26,
                    },
                    inspection_time,
                )
            )
    batch_update_spreadsheet_values(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 검사 작업 시간이 업데이트되었습니다.")


def update_spreadsheet_with_finishing_time(
    spreadsheet_id, order_no, finishing_time, sheet_values
):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            requests.append(
                (
                    {
                        "sheetId": TARGET_SHEET_ID,
                        "startRowIndex": i - 1,
                        "endRowIndex": i,
                        "startColumnIndex": 26,
                        "endColumnIndex": 27,
                    },
                    finishing_time,
                )
            )
    batch_update_spreadsheet_values(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 마무리 작업 시간이 업데이트되었습니다.")


def update_spreadsheet_with_working_hours(spreadsheet_id, order_no, link, sheet_values):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return link
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            # 하이퍼링크 공식을 직접 입력 (앞의 ' 방지)
            requests.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": TARGET_SHEET_ID,
                            "startRowIndex": i - 1,
                            "endRowIndex": i,
                            "startColumnIndex": 21,
                            "endColumnIndex": 22,
                        },
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "formulaValue": f'=HYPERLINK("{link}", "Working Hours")'
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue",
                    }
                }
            )
    if requests:
        batch_update_spreadsheet(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 WORKING HOURS 그래프 링크가 업데이트되었습니다.")
    return link


def update_spreadsheet_with_legend(spreadsheet_id, order_no, link, sheet_values):
    if not sheet_values:
        print("스프레드시트 데이터를 가져올 수 없습니다.")
        return link
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            # 하이퍼링크 공식을 직접 입력 (앞의 ' 방지)
            requests.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": TARGET_SHEET_ID,
                            "startRowIndex": i - 1,
                            "endRowIndex": i,
                            "startColumnIndex": 20,
                            "endColumnIndex": 21,
                        },
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "formulaValue": f'=HYPERLINK("{link}", "Legend Chart")'
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue",
                    }
                }
            )
    if requests:
        batch_update_spreadsheet(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 범례 차트 링크가 업데이트되었습니다.")
    return link


def update_spreadsheet_with_wd_graph(spreadsheet_id, order_no, link, sheet_values):
    if not sheet_values:
        print("🚨 [오류] 스프레드시트 데이터를 가져올 수 없습니다.")
        return link
    requests = []
    for i, row in enumerate(sheet_values[1:], 2):
        if row and row[0].strip().lower() == order_no.strip().lower():
            # 하이퍼링크 공식을 직접 입력 (앞의 ' 방지)
            requests.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": TARGET_SHEET_ID,
                            "startRowIndex": i - 1,
                            "endRowIndex": i,
                            "startColumnIndex": 28,
                            "endColumnIndex": 29,
                        },
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "formulaValue": f'=HYPERLINK("{link}", "WD Chart")'
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue",
                    }
                }
            )
    if requests:
        batch_update_spreadsheet(spreadsheet_id, requests)
    print(f"모델 '{order_no}'의 WD 작업시간 그래프 링크가 업데이트되었습니다.")
    return link


# Helper: Batch update for values (used by update functions)
def batch_update_spreadsheet_values(spreadsheet_id, data):
    requests = []
    for range_spec, value in data:
        requests.append(
            {
                "updateCells": {
                    "range": range_spec,
                    "rows": [
                        {"values": [{"userEnteredValue": {"stringValue": str(value)}}]}
                    ],
                    "fields": "userEnteredValue",
                }
            }
        )
    if requests:
        batch_update_spreadsheet(spreadsheet_id, requests)


# NaN & Overtime Stats
def compute_occurrence_rates(
    df,
    task_total_time,
    avg_mapping,
    model_name,
    tolerance=2,
    mech_partner=None,
    elec_partner=None,
):
    categories = ["기구", "TMS_반제품", "전장", "검사", "마무리", "기타"]
    occurrence_stats = {
        cat: {
            "total_count": 0,
            "nan_count": 0,
            "ot_count": 0,
            "nan_tasks": [],
            "ot_task_details": [],
        }
        for cat in categories
    }
    partner_stats = {
        "mech": {"nan_count": 0, "ot_count": 0},
        "elec": {"nan_count": 0, "ot_count": 0},
    }
    df["진행율"] = pd.to_numeric(df["진행율"], errors="coerce")
    completed_tasks = set(
        df[
            (pd.to_numeric(df["진행율"], errors="coerce") >= 100)
            | (df["시작 시간"].notna() & df["완료 시간"].notna())
        ]["내용"]
    )
    nan_task_checked = set()
    for _, row in df.iterrows():
        task_name = row["내용"]
        category = classify_task(task_name, model_name)
        occurrence_stats[category]["total_count"] += 1
        is_nan = (
            pd.isna(row["시작 시간"])
            or pd.isna(row["완료 시간"])
            or pd.isna(row["진행율"])
        )
        if is_nan:
            if task_name in completed_tasks:
                continue
            if (task_name, category) in nan_task_checked:
                continue
            occurrence_stats[category]["nan_count"] += 1
            occurrence_stats[category]["nan_tasks"].append(task_name)
            nan_task_checked.add((task_name, category))
            if category == "기구":
                partner_stats["mech"]["nan_count"] += 1
            elif category == "전장":
                partner_stats["elec"]["nan_count"] += 1
    for _, row in task_total_time.iterrows():
        task_name = row["내용"]
        actual_hours = row["워킹데이 소요 시간"]
        category = classify_task(task_name, model_name)
        if (
            task_name in avg_mapping
            and actual_hours > avg_mapping[task_name] + tolerance
        ):
            occurrence_stats[category]["ot_count"] += 1
            occurrence_stats[category]["ot_task_details"].append(
                (task_name, actual_hours)
            )
            if category == "기구":
                partner_stats["mech"]["ot_count"] += 1
            elif category == "전장":
                partner_stats["elec"]["ot_count"] += 1
    return occurrence_stats, partner_stats


# Bar Chart Generation
def generate_nan_bar_charts(all_results):
    partner_stats = {}
    for (
        _,
        _,
        mech_partner_raw,
        elec_partner_raw,
        occurrence_stats,
        partner_stats_individual,
        _,
        _,
        _,
    ) in all_results:
        mech_partner = "TMS(m)" if mech_partner_raw == "TMS" else mech_partner_raw
        elec_partner = "TMS(e)" if elec_partner_raw == "TMS" else elec_partner_raw
        mech_nan = partner_stats_individual.get("mech", {}).get("nan_count", 0)
        mech_total = occurrence_stats.get("기구", {}).get("total_count", 0)
        if mech_partner:
            partner_stats[mech_partner] = partner_stats.get(
                mech_partner, {"nan_count": 0, "total_tasks": 0}
            )
            partner_stats[mech_partner]["nan_count"] += mech_nan
            partner_stats[mech_partner]["total_tasks"] += mech_total
        tms_nan = occurrence_stats.get("TMS_반제품", {}).get("nan_count", 0)
        tms_total = occurrence_stats.get("TMS_반제품", {}).get("total_count", 0)
        if tms_total > 0:
            partner_stats["TMS_반제품"] = partner_stats.get(
                "TMS_반제품", {"nan_count": 0, "total_tasks": 0}
            )
            partner_stats["TMS_반제품"]["nan_count"] += tms_nan
            partner_stats["TMS_반제품"]["total_tasks"] += tms_total
        elec_nan = partner_stats_individual.get("elec", {}).get("nan_count", 0)
        elec_total = occurrence_stats.get("전장", {}).get("total_count", 0)
        if elec_partner:
            partner_stats[elec_partner] = partner_stats.get(
                elec_partner, {"nan_count": 0, "total_tasks": 0}
            )
            partner_stats[elec_partner]["nan_count"] += elec_nan
            partner_stats[elec_partner]["total_tasks"] += elec_total
    nan_counts = [stats["nan_count"] for stats in partner_stats.values()]
    total_nan = sum(nan_counts)
    if total_nan == 0:
        print("⚠️ NaN 건수가 없어 그래프를 생성하지 않습니다.")
        return None, None
    labels = list(partner_stats.keys())
    nan_ratios = [
        (
            stats["nan_count"] / stats["total_tasks"] * 100
            if stats["total_tasks"] > 0
            else 0
        )
        for stats in partner_stats.values()
    ]
    plt.figure(figsize=(10, 6))
    bars = plt.bar(labels, nan_ratios, color=plt.cm.Paired.colors[: len(labels)])
    plt.title("협력사별 NaN 발생 비율 (작업 수 대비)", fontsize=14, pad=20)
    plt.xlabel("협력사", fontsize=12)
    plt.ylabel("NaN 발생 비율 (%)", fontsize=12)
    plt.xticks(rotation=45, ha="right")
    for i, bar in enumerate(bars):
        height = bar.get_height()
        count = nan_counts[i]
        total = partner_stats[labels[i]]["total_tasks"]
        plt.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            f"{height:.1f}%\n({count}/{total})",
            ha="center",
            va="bottom",
            fontsize=10,
        )
    plt.tight_layout()
    tasks_file = "NaN_Summary_by_Tasks.png"
    plt.savefig(tasks_file, bbox_inches="tight")
    plt.close()
    plt.figure(figsize=(8, 8))
    plt.pie(
        nan_counts,
        labels=labels,
        autopct=lambda pct: f"{pct:.1f}% ({int(total_nan * pct / 100)})",
        startangle=90,
        colors=plt.cm.Paired.colors[: len(labels)],
    )
    plt.axis("equal")
    plt.title("협력사별 NaN 발생 비율 (전체 NaN 건수 대비)", fontsize=14, pad=20)
    total_file = "NaN_Summary_by_Total.png"
    plt.savefig(total_file, bbox_inches="tight")
    plt.close()
    print(f"📊 막대그래프 생성 완료: {tasks_file}, 파이차트 생성 완료: {total_file}")
    return tasks_file, total_file


# Email & Notification Functions
def render_progress_bar(percent, total, label):
    completed = round((percent / 100) * total)
    tooltip = f"{label} 완료: {completed} / {total}건"
    if percent == 100:
        return f'<span title="{tooltip}" style="font-size: 16px;">✅</span>'
    else:
        return (
            f'<div style="width: 100%; background-color: #e0e0e0; height: 12px; border-radius: 3px;">'
            f'<div style="width: {percent}%; background-color: orange; height: 100%; border-radius: 3px;" title="{tooltip}"></div></div>'
            f'<span style="font-size: 12px;">{percent:.1f}%</span>'
        )


def build_combined_email_body(
    all_results,
    nan_tasks_link=None,
    nan_total_link=None,
    heatmap_url=None,
    monthly_partner_url=None,
    monthly_model_url=None,
):
    kst = pytz.timezone("Asia/Seoul")
    execution_time = datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S")
    year, week_num, _ = date.today().isocalendar()
    dashboard_link = os.getenv("DASHBOARD_URL", "https://gst-factory.netlify.app")

    # 고유 값 수집 (필터 드롭다운용)
    unique_values = {
        "Order": set(),
        "모델명": set(),
        "기구협력사": set(),
        "전장협력사": set(),
        "총 작업 수": set(),
        "기구 NaN": set(),
        "기구 OT": set(),
        "기구 진행률": set(),
        "전장 NaN": set(),
        "전장 OT": set(),
        "전장 진행률": set(),
        "TMS NaN": set(),
        "TMS OT": set(),
        "TMS 진행률": set(),
    }
    for (
        order_no,
        model_name,
        mech_partner,
        elec_partner,
        occurrence_stats,
        partner_stats,
        _,
        spreadsheet_url,
        progress_summary,
    ) in all_results:
        total_tasks = sum(stats["total_count"] for stats in occurrence_stats.values())
        mech_stats = partner_stats.get("mech", {})
        elec_stats = partner_stats.get("elec", {})
        tms_stats = occurrence_stats.get("TMS_반제품", {})
        unique_values["Order"].add(order_no)
        unique_values["모델명"].add(model_name)
        unique_values["기구협력사"].add(mech_partner)
        unique_values["전장협력사"].add(elec_partner)
        unique_values["총 작업 수"].add(str(total_tasks))
        unique_values["기구 NaN"].add(str(mech_stats.get("nan_count", 0)))
        unique_values["기구 OT"].add(str(mech_stats.get("ot_count", 0)))
        unique_values["전장 NaN"].add(str(elec_stats.get("nan_count", 0)))
        unique_values["전장 OT"].add(str(elec_stats.get("ot_count", 0)))
        unique_values["TMS NaN"].add(str(tms_stats.get("nan_count", 0)))
        unique_values["TMS OT"].add(str(tms_stats.get("ot_count", 0)))
        # 진행률 고유 값 추가 (소수점 1자리 문자열)
        prog = progress_summary or {"기구": 0, "전장": 0, "TMS_반제품": 0}
        unique_values["기구 진행률"].add(f"{prog.get('기구', 0):.1f}")
        unique_values["전장 진행률"].add(f"{prog.get('전장', 0):.1f}")
        unique_values["TMS 진행률"].add(f"{prog.get('TMS_반제품', 0):.1f}")

    lines = [
        '<div style="text-align: center; margin-bottom: 20px;">' "</div>",
        '<div class="chart-section" style="margin-top: 30px;">',
        '<iframe src="partner_entry_chart.html" width="100%" height="1200" frameborder="0"></iframe>',
        "</div>",
        f"<h1>PDA Dashboard - {year}년 {week_num}주차</h1>",
        f"<h3>📌 [알림] PDA Overtime 및 NaN 체크 결과 (총 {len(all_results)}건 처리)</h3>",
        f"<p>📅 실행 시간: {execution_time} (KST)</p>",
        f'<p>📊 대시보드에서 상세 내용 확인하세요! (<a href="{dashboard_link}">대시보드 바로가기</a>',
    ]

    # NOVA 트렌드 그래프 섹션을 대시보드 링크 바로 뒤에 추가
    nova_links = []
    if heatmap_url:
        nova_links.append(
            f'📅 주간 협력사 NaN 히트맵: <a href="{heatmap_url}" target="_blank">그래프 보기</a>'
        )

    # 월간 히트맵 링크들 (Google Drive에서 최신 파일 검색)
    if not monthly_partner_url:
        # Google Drive에서 최신 월간 협력사 히트맵 파일 검색
        try:
            query = f"'{DRIVE_FOLDER_ID}' in parents and name contains 'monthly_partner_nan_heatmap_'"
            files = (
                drive_service.files()
                .list(q=query, fields="files(id, name)")
                .execute()
                .get("files", [])
            )

            if files:
                # 파일명에서 날짜 추출하여 최신 파일 선택
                latest_partner_file = max(
                    files, key=lambda x: x["name"].split("_")[-1].replace(".png", "")
                )
                monthly_partner_url = f"https://drive.google.com/uc?export=view&id={latest_partner_file['id']}"
                print(
                    f"📁 Drive에서 최신 월간 협력사 히트맵 발견: {latest_partner_file['name']}"
                )
                print(f"✅ 월간 협력사 히트맵 URL: {monthly_partner_url}")
            else:
                print("⚠️ Drive에서 월간 협력사 히트맵을 찾을 수 없습니다.")
        except Exception as e:
            print(f"❌ Drive 검색 오류 (협력사): {e}")

        # Drive에서 찾지 못하면 환경변수 기본값 사용
        if not monthly_partner_url:
            monthly_partner_url = os.getenv(
                "MONTHLY_PARTNER_HEATMAP_URL",
                "https://drive.google.com/uc?export=view&id=1Bh1iUvPIQfsQ_wUTs_DOln0cZGY_hHL7",
            )
            print(f"⚠️ Drive 검색 실패, 환경변수 URL 사용: {monthly_partner_url}")

    if not monthly_model_url:
        # Google Drive에서 최신 월간 모델 히트맵 파일 검색
        try:
            query = f"'{DRIVE_FOLDER_ID}' in parents and name contains 'monthly_model_nan_heatmap_'"
            files = (
                drive_service.files()
                .list(q=query, fields="files(id, name)")
                .execute()
                .get("files", [])
            )

            if files:
                # 파일명에서 날짜 추출하여 최신 파일 선택
                latest_model_file = max(
                    files, key=lambda x: x["name"].split("_")[-1].replace(".png", "")
                )
                monthly_model_url = f"https://drive.google.com/uc?export=view&id={latest_model_file['id']}"
                print(
                    f"📁 Drive에서 최신 월간 모델 히트맵 발견: {latest_model_file['name']}"
                )
                print(f"✅ 월간 모델 히트맵 URL: {monthly_model_url}")
            else:
                print("⚠️ Drive에서 월간 모델 히트맵을 찾을 수 없습니다.")
        except Exception as e:
            print(f"❌ Drive 검색 오류 (모델): {e}")

        # Drive에서 찾지 못하면 환경변수 기본값 사용
        if not monthly_model_url:
            monthly_model_url = os.getenv(
                "MONTHLY_MODEL_HEATMAP_URL",
                "https://drive.google.com/uc?export=view&id=1DGOJCR5Ie5VGgMMcgIEQc0D45z8-uuIG",
            )
            print(f"⚠️ Drive 검색 실패, 환경변수 URL 사용: {monthly_model_url}")

    # 링크가 있는 경우에만 추가
    if monthly_partner_url:
        nova_links.append(
            f'🗓️ 월간 협력사 NaN 히트맵: <a href="{monthly_partner_url}" target="_blank">그래프 보기</a>'
        )
    if monthly_model_url:
        nova_links.append(
            f'📈 월간 모델별 NaN 히트맵: <a href="{monthly_model_url}" target="_blank">그래프 보기</a>'
        )

    if nova_links:
        lines.append("<p><strong>📊트렌드 지표</strong></p><ul>")
        for link in nova_links:
            lines.append(f"<li>{link}</li>")
        lines.append("</ul>")

    lines.append(")</p>")
    lines.extend(
        [
            "<h4>요약 테이블</h4>",
            '<table id="summaryTable" border="1" style="border-collapse: collapse; width: 95%; font-size: 13px;">',
            '<tr style="background-color: #f2f2f2;">',
        ]
    )

    # 헤더 및 필터 드롭다운 추가
    headers = [
        "Order",
        "모델명",
        "기구협력사",
        "전장협력사",
        "총 작업 수",
        "기구 NaN",
        "기구 OT",
        "기구 진행률",
        "전장 NaN",
        "전장 OT",
        "전장 진행률",
        "TMS NaN",
        "TMS OT",
        "TMS 진행률",
    ]
    for header in headers:
        lines.append(
            f'<th data-column="{header}">{header}<br>'
            f'<select onchange="filterTable(\'{header}\')" style="width: 100%; font-size: 12px;">'
            f'<option value="">전체</option>'
        )
        for value in sorted(unique_values[header]):
            lines.append(f'<option value="{value}">{value}</option>')
        lines.append("</select></th>")
    lines.append("</tr>")

    # 테이블 행 생성 (Order 열에 하이퍼링크 추가)
    for (
        order_no,
        model_name,
        mech_partner,
        elec_partner,
        occurrence_stats,
        partner_stats,
        links,
        spreadsheet_url,
        progress_summary,
    ) in all_results:
        total_tasks = sum(stats["total_count"] for stats in occurrence_stats.values())
        prog = progress_summary or {"기구": 0, "전장": 0, "TMS_반제품": 0}
        prog_mech = prog.get("기구", 0)
        prog_elec = prog.get("전장", 0)
        prog_tms = prog.get("TMS_반제품", 0)
        mech_stats = partner_stats.get("mech", {})
        elec_stats = partner_stats.get("elec", {})
        tms_stats = occurrence_stats.get("TMS_반제품", {})
        mech_nan = mech_stats.get("nan_count", 0)
        mech_ot = mech_stats.get("ot_count", 0)
        mech_total = mech_stats.get("total_count", 0)
        elec_nan = elec_stats.get("nan_count", 0)
        elec_ot = elec_stats.get("ot_count", 0)
        elec_total = elec_stats.get("total_count", 0)
        tms_nan = tms_stats.get("nan_count", 0)
        tms_ot = tms_stats.get("ot_count", 0)
        tms_total = tms_stats.get("total_count", 0)
        lines.append(
            f'<tr><td><a href="{spreadsheet_url}">{order_no}</a></td><td>{model_name}</td><td>{mech_partner}</td><td>{elec_partner}</td><td>{total_tasks}</td>'
            f'<td{" style=" + chr(34) + "color: red; font-weight: bold;" + chr(34) if mech_nan > 0 else ""}>{mech_nan}</td>'
            f'<td{" style=" + chr(34) + "color: red; font-weight: bold;" + chr(34) if mech_ot > 0 else ""}>{mech_ot}</td>'
            f'<td>{render_progress_bar(prog_mech, mech_total, "기구")}</td>'
            f'<td{" style=" + chr(34) + "color: red; font-weight: bold;" + chr(34) if elec_nan > 0 else ""}>{elec_nan}</td>'
            f'<td{" style=" + chr(34) + "color: red; font-weight: bold;" + chr(34) if elec_ot > 0 else ""}>{elec_ot}</td>'
            f'<td>{render_progress_bar(prog_elec, elec_total, "전장")}</td>'
            f'<td{" style=" + chr(34) + "color: red; font-weight: bold;" + chr(34) if tms_nan > 0 else ""}>{tms_nan}</td>'
            f'<td{" style=" + chr(34) + "color: red; font-weight: bold;" + chr(34) if tms_ot > 0 else ""}>{tms_ot}</td>'
            f'<td>{render_progress_bar(prog_tms, tms_total, "TMS")}</td></tr>'
        )
    lines.append("</table><br>")

    # JavaScript 필터링 로직 (진행률 처리 포함)
    lines.append(
        """
    <script>
    function filterTable(column) {
        const table = document.getElementById("summaryTable");
        const select = table.querySelector(`th[data-column="${column}"] select`);
        const filterValue = select.value.toLowerCase();
        const rows = table.getElementsByTagName("tr");

        for (let i = 1; i < rows.length; i++) { // Skip header row
            const cells = rows[i].getElementsByTagName("td");
            const cellIndex = Array.from(table.getElementsByTagName("th")).findIndex(th => th.dataset.column === column);
            let cellValue = cells[cellIndex].innerText.toLowerCase();
            // 진행률 열의 경우 숫자 부분만 추출 (예: "85.7%" → "85.7")
            if (column.includes("진행률")) {
                const match = cellValue.match(/\\d+\\.\\d/); // 숫자.소수점 패턴
                cellValue = match ? match[0] : "0.0";
            }
            // Order 열의 경우 하이퍼링크 텍스트 추출
            if (column === "Order") {
                cellValue = cells[cellIndex].querySelector("a") ? cells[cellIndex].querySelector("a").innerText.toLowerCase() : cellValue;
            }
            let showRow = true;

            // 모든 필터 확인
            table.querySelectorAll("th select").forEach(sel => {
                const col = sel.parentElement.dataset.column;
                const val = sel.value.toLowerCase();
                if (val) {
                    const idx = Array.from(table.getElementsByTagName("th")).findIndex(th => th.dataset.column === col);
                    let compareValue = cells[idx].innerText.toLowerCase();
                    if (col.includes("진행률")) {
                        const match = compareValue.match(/\\d+\\.\\d/);
                        compareValue = match ? match[0] : "0.0";
                    }
                    if (col === "Order") {
                        compareValue = cells[idx].querySelector("a") ? cells[idx].querySelector("a").innerText.toLowerCase() : compareValue;
                    }
                    if (compareValue !== val) {
                        showRow = false;
                    }
                }
            });

            rows[i].style.display = showRow ? "" : "none";
        }
    }
    </script>
    """
    )

    for (
        order_no,
        model_name,
        mech_partner,
        elec_partner,
        occurrence_stats,
        partner_stats,
        links,
        spreadsheet_url,
        progress_summary,
    ) in all_results:
        lines.append(
            f"<details><summary><strong>📍 Order: {order_no}, 모델명: {model_name}</strong></summary>"
        )
        lines.append(
            f"<p>🏭 기구협력사: {mech_partner}, ⚡ 전장협력사: {elec_partner}</p>"
        )
        lines.append(
            f'<p>📋 <strong>모델 스프레드시트</strong>: <a href="{spreadsheet_url}">바로가기</a></p>'
        )
        lines.append(
            f'<p>📊 대시보드 링크: <a href="{dashboard_link}">대시보드 바로가기</a></p>'
        )
        lines.append(
            f"<p>📊 그래프 링크:</p><ul>"
            f'<li>Working Hours: <a href="{links["working_hours"]}">바로가기</a></li>'
            f'<li>Legend Chart: <a href="{links["legend"]}">바로가기</a></li>'
            f'<li>WD Chart: <a href="{links["wd"]}">바로가기</a></li></ul>'
        )
        for category in ["기구", "TMS_반제품", "전장", "검사", "마무리", "기타"]:
            stats = occurrence_stats.get(
                category,
                {
                    "total_count": 0,
                    "nan_count": 0,
                    "ot_count": 0,
                    "nan_tasks": [],
                    "ot_task_details": [],
                },
            )
            total_count = stats["total_count"]
            lines.append(
                f"<p><b>🔹 {category} 작업</b><br> - 전체 작업 수: {total_count} 건<br>"
            )
            nan_count = stats["nan_count"]
            nan_ratio = (nan_count / total_count) * 100 if total_count > 0 else 0
            lines.append(
                f' <span{" style=" + chr(34) + "color: red;" + chr(34) if nan_count > 0 else ""}>⚠️ 누락(NaN): {nan_count} 건 (비율: {nan_ratio:.2f}%)</span><br>'
            )
            if nan_count > 0:
                lines.append("".join(f"   - {task}<br>" for task in stats["nan_tasks"]))
            ot_count = stats["ot_count"]
            ot_ratio = (ot_count / total_count) * 100 if total_count > 0 else 0
            lines.append(
                f' <span{" style=" + chr(34) + "color: red;" + chr(34) if ot_count > 0 else ""}>⏳ 오버타임: {ot_count} 건 (비율: {ot_ratio:.2f}%)</span><br>'
            )
            if ot_count > 0:
                lines.append(
                    "".join(
                        f"   - {task} {format_hours(hours)}<br>"
                        for task, hours in stats["ot_task_details"]
                    )
                )
            lines.append("</p>")
        lines.append("</details><hr>")

    return "\n".join(lines)


def send_occurrence_email(subject, body_text, graph_files=None, dashboard_file=None):
    # 이메일 설정이 없으면 건너뛰기
    if not email_configured:
        print(f"⚠️ 이메일 설정이 없어 이메일 전송을 건너뜁니다.")
        print(f"📧 제목: {subject}")
        return

    context = ssl.create_default_context()
    msg = MIMEMultipart("mixed")
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = RECEIVER_EMAIL
    msg["Subject"] = subject
    msg.attach(MIMEText(body_text, "html", _charset="utf-8"))
    if graph_files:
        for graph_file in graph_files:
            try:
                with open(graph_file, "rb") as f:
                    img = MIMEImage(f.read())
                    img.add_header(
                        "Content-Disposition", "attachment", filename=graph_file
                    )
                    msg.attach(img)
                print(f"✅ 그래프 파일 첨부 완료: {graph_file}")
            except Exception as e:
                print(f"❌ [첨부 오류] 그래프 파일 {graph_file} 첨부 실패: {e}")
    if dashboard_file:
        try:
            with open(dashboard_file, "rb") as f:
                html_attachment = MIMEApplication(f.read(), _subtype="html")
                html_attachment.add_header(
                    "Content-Disposition", "attachment", filename=dashboard_file
                )
                msg.attach(html_attachment)
            print(f"📄 [이메일 첨부] HTML 대시보드 파일 {dashboard_file} 추가 완료")
        except Exception as e:
            print(
                f"❌ [이메일 첨부 오류] HTML 대시보드 파일 {dashboard_file} 추가 실패: {e}"
            )
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(EMAIL_ADDRESS, EMAIL_PASS)
            server.send_message(msg)
        print(
            f"📧 [이메일 발송] {RECEIVER_EMAIL}로 통합 HTML 알림 메일 및 첨부파일 전송 완료"
        )
    except Exception as e:
        print(f"❌ [이메일 발송 실패]: {e}")


def cross_check_data_integrity(all_results):
    """
    기존 결과값들을 크로스 체크하여 데이터 정합성 확인

    Args:
        all_results: 처리된 모든 결과 데이터

    Returns:
        dict: 크로스 체크 결과 및 경고 메시지
    """
    check_report = {
        "total_models": len(all_results),
        "warnings": [],
        "summary": {
            "total_nan_by_category": 0,
            "total_nan_by_partner": 0,
            "total_ot_by_category": 0,
            "total_ot_by_partner": 0,
            "category_breakdown": {},
            "partner_breakdown": {},
        },
    }

    # 카테고리별 총합 계산 (기존 방식)
    category_nan_total = 0
    category_ot_total = 0
    category_breakdown = {}

    # 협력사별 총합 계산
    partner_nan_total = 0
    partner_ot_total = 0
    partner_breakdown = {"mech": 0, "elec": 0}

    for result in all_results:
        (
            order_no,
            model_name,
            mech_partner,
            elec_partner,
            occurrence_stats,
            partner_stats,
            _,
            _,
            _,
        ) = result

        # 카테고리별 통계 집계
        for category, stats in occurrence_stats.items():
            nan_count = stats.get("nan_count", 0)
            ot_count = stats.get("ot_count", 0)

            category_nan_total += nan_count
            category_ot_total += ot_count

            if category not in category_breakdown:
                category_breakdown[category] = {"nan": 0, "ot": 0}
            category_breakdown[category]["nan"] += nan_count
            category_breakdown[category]["ot"] += ot_count

        # 협력사별 통계 집계
        for partner_type, stats in partner_stats.items():
            nan_count = stats.get("nan_count", 0)
            ot_count = stats.get("ot_count", 0)

            partner_nan_total += nan_count
            partner_ot_total += ot_count

            if partner_type in partner_breakdown:
                partner_breakdown[partner_type] += nan_count

    # 크로스 체크 1: 기구/전장 카테고리와 협력사 통계 비교
    mech_category_nan = category_breakdown.get("기구", {}).get("nan", 0)
    elec_category_nan = category_breakdown.get("전장", {}).get("nan", 0)

    mech_partner_nan = partner_breakdown.get("mech", 0)
    elec_partner_nan = partner_breakdown.get("elec", 0)

    if mech_category_nan != mech_partner_nan:
        check_report["warnings"].append(
            f"⚠️ 기구 NaN 불일치: 카테고리({mech_category_nan}) ≠ 협력사({mech_partner_nan})"
        )

    if elec_category_nan != elec_partner_nan:
        check_report["warnings"].append(
            f"⚠️ 전장 NaN 불일치: 카테고리({elec_category_nan}) ≠ 협력사({elec_partner_nan})"
        )

    # 크로스 체크 2: 카테고리별 주요 통계 확인
    major_categories = ["기구", "전장", "TMS_반제품"]
    for category in major_categories:
        if category in category_breakdown:
            nan_count = category_breakdown[category]["nan"]
            total_count = sum(
                stats.get("total_count", 0)
                for result in all_results
                for stats in [result[4].get(category, {})]
            )

            if total_count > 0:
                nan_ratio = (nan_count / total_count) * 100
                if nan_ratio > 80:  # 80% 이상이면 경고
                    check_report["warnings"].append(
                        f"⚠️ {category} NaN 비율 높음: {nan_ratio:.1f}% ({nan_count}/{total_count})"
                    )

    # 크로스 체크 3: 전체 모델 수와 실제 처리된 데이터 수 비교
    processed_models = len(
        [
            r
            for r in all_results
            if any(
                stats.get("nan_count", 0) > 0 or stats.get("ot_count", 0) > 0
                for stats in r[4].values()
            )
        ]
    )

    if processed_models != len(all_results):
        check_report["warnings"].append(
            f"ℹ️ 처리된 모델: {processed_models}/{len(all_results)} (정상 범위 내 모델 제외)"
        )

    # 요약 정보 업데이트
    check_report["summary"]["total_nan_by_category"] = category_nan_total
    check_report["summary"]["total_nan_by_partner"] = partner_nan_total
    check_report["summary"]["total_ot_by_category"] = category_ot_total
    check_report["summary"]["total_ot_by_partner"] = partner_ot_total
    check_report["summary"]["category_breakdown"] = category_breakdown
    check_report["summary"]["partner_breakdown"] = partner_breakdown

    return check_report


def send_nan_alert_to_kakao(all_results):
    if not all_results:
        print("⚠️ [알림] 전송할 데이터가 없습니다.")
        return

    # 기존 결과값들을 크로스 체크
    print("🔍 기존 결과값 크로스 체크 중...")
    check_report = cross_check_data_integrity(all_results)

    # 크로스 체크 결과 출력
    if check_report["warnings"]:
        print("⚠️ 데이터 크로스 체크 결과:")
        for warning in check_report["warnings"]:
            print(f"  {warning}")
    else:
        print("✅ 데이터 크로스 체크 통과")

    # 기존 방식으로 총합 계산 (검증된 값 사용)
    total_nan = check_report["summary"]["total_nan_by_category"]
    total_ot = check_report["summary"]["total_ot_by_category"]

    # 기존 방식과 비교 (디버깅용)
    original_nan = sum(
        stats["nan_count"] for result in all_results for stats in result[4].values()
    )
    original_ot = sum(
        stats["ot_count"] for result in all_results for stats in result[4].values()
    )

    if total_nan != original_nan or total_ot != original_ot:
        print(
            f"⚠️ 계산 방식 차이 발견: NaN({total_nan}vs{original_nan}), OT({total_ot}vs{original_ot})"
        )
        # 기존 방식 사용 (안전)
        total_nan = original_nan
        total_ot = original_ot

    kst = pytz.timezone("Asia/Seoul")
    execution_time = datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S")

    # 카카오톡 메시지 구성
    text = f"📢 PDA Overtime 및 NaN 체크 결과\n📅 실행 시간: {execution_time} (KST)\n📊 총 {len(all_results)}건 처리\n⚠️ 누락(NaN): {total_nan} 건\n⏳ 오버타임: {total_ot} 건"

    # 크로스 체크 경고가 있는 경우 추가 정보
    if check_report["warnings"]:
        text += f"\n🔍 데이터 체크: {len(check_report['warnings'])}건 확인 필요"

    # 주요 카테고리 요약 추가
    category_summary = []
    for category in ["기구", "전장", "TMS_반제품"]:
        if category in check_report["summary"]["category_breakdown"]:
            nan_count = check_report["summary"]["category_breakdown"][category]["nan"]
            if nan_count > 0:
                category_summary.append(f"{category}: {nan_count}건")

    if category_summary:
        text += f"\n📋 주요 누락: {', '.join(category_summary)}"

    text += "\n👇 대시보드에서 상세 내용 확인하세요!"

    # 메시지 전송
    access_token = refresh_access_token()
    if not access_token:
        print("❌ [카카오톡 발송 실패] 액세스 토큰이 없어 메시지 발송 불가.")
        return

    success = send_kakao_message(text, access_token)

    if success:
        print("✅ 크로스 체크 완료 및 카카오톡 메시지 전송 성공!")
        print(f"📊 검증된 통계: NaN {total_nan}건, OT {total_ot}건")
        if check_report["warnings"]:
            print(f"⚠️ 확인 필요 항목: {len(check_report['warnings'])}건")
    else:
        print("❌ 카카오톡 메시지 전송 실패")


def refresh_access_token():
    url = "https://kauth.kakao.com/oauth/token"
    data = {
        "grant_type": "refresh_token",
        "client_id": REST_API_KEY,
        "refresh_token": REFRESH_TOKEN,
    }
    try:
        response = requests.post(url, data=data)
        response.raise_for_status()
        token_info = response.json()
        if "access_token" in token_info:
            new_access_token = token_info["access_token"]
            print(f"✅ 새 액세스 토큰: {new_access_token}")
            return new_access_token
        else:
            print(f"❌ 토큰 갱신 실패: {token_info}")
            return None
    except requests.exceptions.RequestException as e:
        print(f"❌ [네트워크 오류] 토큰 갱신 실패: {e}")
        return None


def send_kakao_message(text, access_token=None):
    if access_token is None:
        try:
            access_token = KAKAO_ACCESS_TOKEN
        except NameError:
            print("❌ [오류] KAKAO_ACCESS_TOKEN이 정의되지 않았습니다.")
            return False
    url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    data = {
        "template_object": json.dumps(
            {
                "object_type": "text",
                "text": text,
                "link": {"web_url": "", "mobile_web_url": ""},
            },
            ensure_ascii=False,
        )
    }
    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        print(
            f"✅ 카카오톡 메시지 전송 성공! (시간: {datetime.now().strftime('%H:%M:%S')})"
        )
        return True
    except Exception as e:
        print(f"❌ [카카오톡 메시지 전송 실패]: {e}")
        return False


def upload_to_github(file_path):
    GITHUB_USERNAME = os.getenv("GITHUB_USERNAME", "isolhsolfafa")
    GITHUB_REPO = os.getenv("GITHUB_REPO", "gst-factory")
    GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "main")
    GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
    if not GITHUB_TOKEN:
        print("❌ [오류] GITHUB_TOKEN 환경변수가 설정되지 않았습니다.")
        return
    with open(file_path, "rb") as file:
        content = base64.b64encode(file.read()).decode("utf-8")
    file_name = os.path.basename(file_path)
    url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{GITHUB_REPO}/contents/public/{file_name}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }
    response = requests.get(url, headers=headers)
    sha = response.json().get("sha", "")
    data = {
        "message": f"자동 업로드: {file_name}",
        "content": content,
        "branch": GITHUB_BRANCH,
    }
    if sha:
        data["sha"] = sha
    response = requests.put(url, headers=headers, json=data)
    if response.status_code in [200, 201]:
        print(f"✅ {file_name} GitHub 업로드 성공!")
    else:
        print(f"❌ {file_name} GitHub 업로드 실패! {response.text}")


def get_mech_start_date(spreadsheet_url, sheets_service):
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url)
        if not match:
            return pd.NaT
        spreadsheet_id = match.group(1)
        result = api_call_with_backoff(
            sheets_service.spreadsheets().values().get,
            spreadsheetId=spreadsheet_id,
            range="정보판!B6",
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
        raw_date = result.get("values", [[]])[0][0]
        return pd.to_datetime(raw_date, errors="coerce")
    except Exception as e:
        print(f"❌ [오류] 기구 시작일 가져오기 실패: {e}")
        return pd.NaT


def sort_all_results_by_mech_start(all_results, sheets_service):
    def extract_start_date(entry):
        spreadsheet_url = entry[7]
        return get_mech_start_date(spreadsheet_url, sheets_service)

    return sorted(all_results, key=extract_start_date)


def collect_and_process_data():
    limit = int(os.getenv("LIMIT", "1"))
    batch_size = 10
    if not linked_spreadsheet_ids:
        print("🚨 [오류] 추출된 스프레드시트 ID가 없습니다.")
        return []
    target_ids = linked_spreadsheet_ids[:limit] if limit > 0 else linked_spreadsheet_ids
    print(
        f"총 {len(linked_spreadsheet_ids)}개 중 처음 {len(target_ids)}개만 처리합니다."
    )
    sheet_values = fetch_entire_sheet_values(
        spreadsheet_id, f"'{TARGET_SHEET_NAME}'!A:AA"
    )
    all_results = []
    current_weekday = datetime.today().weekday()

    # 그래프 생성 옵션 확인 (환경변수에서 제어)
    GENERATE_GRAPHS = os.getenv("GENERATE_GRAPHS", "auto").lower()

    if GENERATE_GRAPHS == "true":
        # 강제로 그래프 생성
        generate_graphs_today = True
        print(
            "✅ GENERATE_GRAPHS=true: 강제로 그래프 생성 및 스프레드시트 업데이트 진행"
        )
    elif GENERATE_GRAPHS == "false":
        # 그래프 생성 비활성화
        generate_graphs_today = False
        print("⛔ GENERATE_GRAPHS=false: 그래프 생성 및 업데이트 비활성화")
    else:
        # 자동 모드: 월요일(0), 금요일(4)에만 그래프 생성 - 또는 TEST_MODE에서는 항상 생성
        generate_graphs_today = current_weekday in [0, 4] or TEST_MODE
    if generate_graphs_today:
        if TEST_MODE:
            print("✅ TEST_MODE: 그래프 생성 및 스프레드시트 업데이트 진행")
        else:
            print("✅ 오늘은 월요일/금요일: 그래프 생성 및 스프레드시트 업데이트 진행")
    else:
        print(
            "⛔ 오늘은 월요일/금요일이 아니므로 그래프 생성 및 업데이트는 생략됩니다."
        )

    print(
        f"📊 그래프 생성 설정: GENERATE_GRAPHS={GENERATE_GRAPHS}, 실제 생성 여부: {generate_graphs_today}"
    )

    def process_batch(batch_ids):
        import time

        for idx, target_spreadsheet_id in enumerate(batch_ids, 1):
            try:
                print(f"--- 🚀 처리 중: {idx}/{len(batch_ids)} (Batch) ---")

                # Rate Limit 방지를 위한 지연 (첫 번째가 아닌 경우)
                if idx > 1:
                    print("⏱️ Rate Limit 방지를 위해 3초 대기...")
                    time.sleep(3)

                df = fetch_data_from_sheets(target_spreadsheet_id, WORKSHEET_RANGE)
                product_name, mech_partner, elec_partner = fetch_info_board_extended(
                    target_spreadsheet_id
                )
                print(f"📌 Processing Model: {product_name}")
                task_total_time = process_data(df, product_name)
                task_total_time["작업 분류"] = task_total_time["내용"].apply(
                    lambda x: classify_task(x, product_name)
                )
                order_no = get_spreadsheet_title(target_spreadsheet_id)
                print(f"📌 Processing Order No: {order_no}")
                spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{target_spreadsheet_id}/edit"
                update_spreadsheet_with_product_name(
                    spreadsheet_id, order_no, product_name, sheet_values
                )
                if generate_graphs_today:
                    # 그래프 파일들을 먼저 생성
                    working_hours_file = generate_and_save_graph(
                        task_total_time, order_no, product_name
                    )
                    legend_file = generate_legend_chart(
                        task_total_time, order_no, product_name
                    )
                    wd_file = generate_and_save_graph_wd(
                        task_total_time, df, order_no, product_name
                    )

                    # Drive에 업로드하고 링크 업데이트
                    links = {
                        "working_hours": update_spreadsheet_with_working_hours(
                            spreadsheet_id,
                            order_no,
                            upload_to_drive(working_hours_file),
                            sheet_values,
                        ),
                        "legend": update_spreadsheet_with_legend(
                            spreadsheet_id,
                            order_no,
                            upload_to_drive(legend_file),
                            sheet_values,
                        ),
                        "wd": update_spreadsheet_with_wd_graph(
                            spreadsheet_id,
                            order_no,
                            upload_to_drive(wd_file),
                            sheet_values,
                        ),
                    }

                    # 임시 파일 정리
                    for temp_file in [working_hours_file, legend_file, wd_file]:
                        try:
                            if temp_file and os.path.exists(temp_file):
                                os.remove(temp_file)
                                logger.info(f"임시 파일 삭제: {temp_file}")
                        except Exception as e:
                            logger.warning(f"임시 파일 삭제 실패 {temp_file}: {e}")
                else:
                    links = {"working_hours": None, "legend": None, "wd": None}
                    print("⛔ 그래프 생성 및 링크 업데이트 생략됨")
                progress_summary = calculate_progress_by_category(df, product_name)
                total_time_decimal = task_total_time["워킹데이 소요 시간"].sum()
                total_time_formatted = format_hours(total_time_decimal)
                update_spreadsheet_with_total_time(
                    spreadsheet_id, order_no, total_time_formatted, sheet_values
                )
                print(
                    f"🎯 총 소요시간 {total_time_formatted}이 W열에 업데이트되었습니다."
                )
                mechanical_time_decimal = task_total_time[
                    task_total_time["작업 분류"] == "기구"
                ]["워킹데이 소요 시간"].sum()
                electrical_time_decimal = task_total_time[
                    task_total_time["작업 분류"] == "전장"
                ]["워킹데이 소요 시간"].sum()
                inspection_time_decimal = task_total_time[
                    task_total_time["작업 분류"] == "검사"
                ]["워킹데이 소요 시간"].sum()
                finishing_time_decimal = task_total_time[
                    task_total_time["작업 분류"] == "마무리"
                ]["워킹데이 소요 시간"].sum()
                update_spreadsheet_with_mechanical_time(
                    spreadsheet_id,
                    order_no,
                    format_hours(mechanical_time_decimal),
                    sheet_values,
                )
                update_spreadsheet_with_electrical_time(
                    spreadsheet_id,
                    order_no,
                    format_hours(electrical_time_decimal),
                    sheet_values,
                )
                update_spreadsheet_with_inspection_time(
                    spreadsheet_id,
                    order_no,
                    format_hours(inspection_time_decimal),
                    sheet_values,
                )
                update_spreadsheet_with_finishing_time(
                    spreadsheet_id,
                    order_no,
                    format_hours(finishing_time_decimal),
                    sheet_values,
                )
                print(f"🎯 모델 '{order_no}'의 작업별 소요시간이 업데이트되었습니다.")
                avg_mapping = get_avg_time_mapping(product_name)
                occurrence_stats, partner_stats = compute_occurrence_rates(
                    df,
                    task_total_time,
                    avg_mapping,
                    product_name,
                    tolerance=2,
                    mech_partner=mech_partner,
                    elec_partner=elec_partner,
                )
                if any(
                    stats["nan_count"] > 0 or stats["ot_count"] > 0
                    for stats in occurrence_stats.values()
                ):
                    all_results.append(
                        (
                            order_no,
                            product_name,
                            mech_partner,
                            elec_partner,
                            occurrence_stats,
                            partner_stats,
                            links,
                            spreadsheet_url,
                            progress_summary,
                        )
                    )
                else:
                    print("✅ [알림] 모든 작업이 정상 범위 내에 있습니다.")
                print(f"✅ 모델 '{order_no}' 처리 완료.\n")
                systime.sleep(1)
            except Exception as e:
                print(
                    f"❌ [오류 발생: 스프레드시트 ID {target_spreadsheet_id}] -> {e}\n"
                )
                systime.sleep(5)

    iterator = iter(target_ids)
    while batch := list(islice(iterator, batch_size)):
        process_batch(batch)
        systime.sleep(10)
    return all_results


def save_results_to_json(all_results, drive_service):
    """
    주요 처리 결과를 JSON 파일로 저장하고 'JSON 데이터 저장용' 구글 드라이브에 업로드합니다.
    (데이터 -> JSON)
    업로드된 파일의 Google Drive ID를 반환합니다.
    """
    if not all_results:
        print("ℹ️ 처리할 결과가 없어 JSON 파일을 생성하지 않습니다.")
        return None

    now_kst = datetime.now(pytz.timezone("Asia/Seoul"))
    execution_time_str = now_kst.strftime("%Y%m%d_%H%M%S")
    execution_time_for_json = execution_time_str  # 기존 형식 유지: "20250618_231207"
    weekday_kor = ["월", "화", "수", "목", "금", "토", "일"][now_kst.weekday()]
    session = now_kst.weekday() + 1
    filename = (
        f"output/nan_ot_results_{execution_time_str}_{weekday_kor}_{session}회차.json"
    )
    os.makedirs("output", exist_ok=True)

    results_list = []
    for res in all_results:
        (
            order_no,
            model_name,
            mech_partner,
            elec_partner,
            occurrence_stats,
            partner_stats,
            links,
            spreadsheet_url,
            progress_summary,
        ) = res

        total_tasks = sum(
            stats.get("total_count", 0) for stats in occurrence_stats.values()
        )

        def calc_ratio(cat_stats):
            total_count = cat_stats.get("total_count", 0)
            if total_count == 0:
                return 0.0, 0.0
            nan_ratio = (cat_stats.get("nan_count", 0) / total_count) * 100
            ot_ratio = (cat_stats.get("ot_count", 0) / total_count) * 100
            return nan_ratio, ot_ratio

        ratios = {}
        if "기구" in occurrence_stats:
            ratios["mech_nan_ratio"], ratios["mech_ot_ratio"] = calc_ratio(
                occurrence_stats["기구"]
            )
        if "전장" in occurrence_stats:
            ratios["elec_nan_ratio"], ratios["elec_ot_ratio"] = calc_ratio(
                occurrence_stats["전장"]
            )
        if "TMS_반제품" in occurrence_stats:
            ratios["tms_nan_ratio"], ratios["tms_ot_ratio"] = calc_ratio(
                occurrence_stats["TMS_반제품"]
            )

        # 기존 JSON 구조에 맞게 occurrence_stats에서 세부 정보 제거
        cleaned_occurrence_stats = {}
        for category, stats in occurrence_stats.items():
            cleaned_occurrence_stats[category] = {
                "total_count": stats.get("total_count", 0),
                "nan_count": stats.get("nan_count", 0),
                "ot_count": stats.get("ot_count", 0),
            }

        # 기존 JSON 구조에 맞게 links를 order_href 하나로 통합
        legacy_links = {"order_href": spreadsheet_url}

        results_list.append(
            {
                "order_no": order_no,
                "model_name": model_name,
                "mech_partner": mech_partner,
                "elec_partner": elec_partner,
                "total_tasks": total_tasks,
                "ratios": ratios,
                "links": legacy_links,  # 기존 구조 사용
                "occurrence_stats": cleaned_occurrence_stats,  # 세부 정보 제거
                "partner_stats": partner_stats,
                "spreadsheet_url": "",  # 기존 구조에서는 빈 문자열
                # progress_summary 제거 (기존 구조에 없음)
            }
        )

    json_data = {"execution_time": execution_time_for_json, "results": results_list}

    with open(filename, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2, default=str)
    print(f"✅ JSON 저장 완료: {filename}")

    # 업로드 시 JSON 전용 폴더 ID 사용
    file_metadata = {
        "name": os.path.basename(filename),
        "parents": [JSON_DRIVE_FOLDER_ID],
    }
    media = MediaFileUpload(filename, mimetype="application/json")
    uploaded = (
        drive_service.files()
        .create(body=file_metadata, media_body=media, fields="id, name")
        .execute()
    )
    print(
        f"✅ JSON 구글 드라이브 업로드 완료: {uploaded.get('name')} (id: {uploaded.get('id')})"
    )

    return uploaded.get("id")


def load_json_files_from_drive(
    drive_service, period="weekly", week_number=None, target_day=None
):
    """
    Google Drive 폴더 내 JSON 파일 로드
    period: "weekly" (주간), "monthly" (월간)
    week_number: 특정 주차 필터링 (주간용)
    target_day: 특정 요일 파일만 로드 ("friday", "sunday", None=모든 요일, "mixed"=주차별 혼합)
    """

    # 월간 히트맵용 스마트 target_day 자동 설정 (효율성 개선)
    if period == "monthly" and target_day is None:
        target_day = "mixed"  # 32주 이전=금요일, 33주 이후=일요일 혼합
        print("📊 월간 히트맵: 32주 이전=금요일, 33주 이후=일요일 JSON 혼합 로드")

    query = f"'{JSON_DRIVE_FOLDER_ID}' in parents and name contains 'nan_ot_results_'"
    if target_day == "friday":
        query += " and name contains '_금_'"
    elif target_day == "sunday":
        query += " and name contains '_일_'"
    elif target_day == "mixed":
        # 금요일과 일요일 모두 포함 (나중에 주차별로 필터링)
        query += " and (name contains '_금_' or name contains '_일_')"
    # target_day가 None이면 모든 요일 포함

    # 최대 3번까지 재시도 (Drive 파일 처리 지연 대응)
    for attempt in range(3):
        files = (
            drive_service.files()
            .list(q=query, fields="files(id, name)")
            .execute()
            .get("files", [])
        )

        if files:
            break

        if attempt < 2:  # 마지막 시도가 아닌 경우에만 대기
            print(f"⏳ 파일을 찾을 수 없습니다. 3초 후 재시도... ({attempt + 1}/3)")
            systime.sleep(3)
        else:
            print("⚠️ 로드할 JSON 파일이 없습니다.")
            return []

    data_list = []

    for file in files:
        file_name = file["name"]
        match = re.search(r"nan_ot_results_(\d{8})", file_name)
        if not match:
            continue

        file_date = match.group(1)  # YYYYMMDD
        try:
            file_datetime = pd.to_datetime(file_date, format="%Y%m%d")
            file_week = file_datetime.isocalendar().week
        except ValueError:
            continue

        # week_number 필터링
        if week_number is not None and file_week != week_number:
            continue

        # mixed 모드에서 주차별 요일 필터링 (32주 이전=금요일, 33주 이후=일요일)
        if target_day == "mixed":
            if file_week < 33:
                # 32주 이전: 금요일만
                if "_금_" not in file_name:
                    continue
            else:
                # 33주 이후: 일요일만
                if "_일_" not in file_name:
                    continue

        print(f"📁 JSON 파일 로드 중: {file_name}")
        file_id = file["id"]
        request = drive_service.files().get_media(fileId=file_id)
        content = request.execute().decode("utf-8")
        data = json.loads(content)
        for result in data["results"]:
            result["execution_time"] = data["execution_time"]
        data_list.extend(data["results"])

    print(f"📂 총 {len(data_list)}개의 로그 데이터를 로드했습니다.")
    return data_list


def ratio_calc(stats):
    """NaN 비율 계산"""
    total = stats.get("total_count", 0)
    nan_count = stats.get("nan_count", 0)
    return (nan_count / total * 100) if total > 0 else 0


def generate_heatmap(
    drive_service,
    period="weekly",
    group_by="partner",
    week_number=None,
    target_day=None,
):
    """
    히트맵 생성 함수
    period: "weekly" (주간), "monthly" (월간)
    group_by: "partner" (협력사), "model" (모델)
    target_day: 월간 히트맵용 특정 요일 ("friday", "sunday", None=auto)
    """
    # 폰트 설정 추가
    global font_prop
    try:
        # font_prop가 정의되어 있는지 확인
        font_prop
    except NameError:
        # font_prop가 정의되지 않은 경우 기본값으로 설정
        font_prop = None
    # 데이터 로드 (월간 히트맵의 경우 load_json_files_from_drive에서 스마트 target_day 자동 설정)
    all_data = load_json_files_from_drive(
        drive_service, period, week_number, target_day
    )

    if not all_data:
        print("⚠️ 데이터를 로드할 수 없습니다.")
        return None

    # DataFrame 생성 및 데이터 변환
    df_data = []
    for d in all_data:
        execution_time = d["execution_time"]
        occurrence_stats = d.get("occurrence_stats", {})
        mech_partner = d.get("mech_partner", "").strip().upper()
        elec_partner = d.get("elec_partner", "").strip().upper()

        entry = {
            "date": execution_time,
            "model_name": d["model_name"],
            "mech_partner": mech_partner,
            "elec_partner": elec_partner,
            "bat_nan_ratio": 0.0,
            "fni_nan_ratio": 0.0,
            "tms_m_nan_ratio": 0.0,
            "cna_nan_ratio": 0.0,
            "pns_nan_ratio": 0.0,
            "tms_e_nan_ratio": 0.0,
            "tms_semi_nan_ratio": 0.0,
        }

        # 기구 협력사별 NaN 비율 계산
        if mech_partner == "BAT":
            target = occurrence_stats.get("기구", {})
            entry["bat_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "FNI":
            target = occurrence_stats.get("기구", {})
            entry["fni_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "TMS":
            target = occurrence_stats.get("기구", {})
            entry["tms_m_nan_ratio"] = ratio_calc(target)

        # 전장 협력사별 NaN 비율 계산
        if elec_partner == "C&A":
            target = occurrence_stats.get("전장", {})
            entry["cna_nan_ratio"] = ratio_calc(target)
        elif elec_partner == "P&S":
            target = occurrence_stats.get("전장", {})
            entry["pns_nan_ratio"] = ratio_calc(target)
        elif elec_partner == "TMS":
            target = occurrence_stats.get("전장", {})
            entry["tms_e_nan_ratio"] = ratio_calc(target)

        # TMS 반제품 NaN 비율
        tms_semi_stats = occurrence_stats.get("TMS_반제품", {})
        entry["tms_semi_nan_ratio"] = ratio_calc(tms_semi_stats)

        df_data.append(entry)

    df = pd.DataFrame(df_data)
    df["date"] = pd.to_datetime(df["date"], format="%Y%m%d_%H%M%S", errors="coerce")
    df = df.dropna(subset=["date"])

    # 협력사 카테고리 정의
    partner_categories = [
        ("bat_nan_ratio", "BAT", "blue"),
        ("fni_nan_ratio", "FNI", "cyan"),
        ("tms_m_nan_ratio", "TMS(m)", "orange"),
        ("cna_nan_ratio", "C&A", "green"),
        ("pns_nan_ratio", "P&S", "red"),
        ("tms_e_nan_ratio", "TMS(e)", "purple"),
        ("tms_semi_nan_ratio", "TMS_반제품", "magenta"),
    ]

    # 주간/월간별 그룹핑
    if period == "weekly":
        df = df.sort_values("date")
        df["day"] = df["date"].dt.strftime("%m월%d일")  # 날짜 형식 변경

        if group_by == "partner":
            df_grouped = df.groupby("day").mean(numeric_only=True)
            categories = partner_categories
            labels = list(df_grouped.index)  # 이미 포맷된 날짜 사용
            title = "주간 NaN 비율 추이 (mixed)"  # 기존 제목과 일치
            y_label = "협력사"
        elif group_by == "model":
            # 주간 모델별 데이터 처리
            model_daily_data = []
            for (day, model), group in df.groupby(["day", "model_name"]):
                total_nan_ratio = 0
                count = 0
                for col, _, _ in partner_categories:
                    if col in group.columns and group[col].sum() > 0:
                        total_nan_ratio += group[col].mean()
                        count += 1
                avg_nan_ratio = total_nan_ratio / count if count > 0 else 0
                model_daily_data.append(
                    {"day": day, "model_name": model, "avg_nan_ratio": avg_nan_ratio}
                )

            model_df = pd.DataFrame(model_daily_data)
            if not model_df.empty:
                df_grouped = model_df.pivot(
                    index="model_name", columns="day", values="avg_nan_ratio"
                ).fillna(0)
                labels = [f"{d.month}월{d.day}일" for d in df_grouped.columns]
            else:
                print("⚠️ 모델별 주간 데이터가 없습니다.")
                return None
            title = "주간 모델별 NaN 비율 히트맵"
            y_label = "모델"

    elif period == "monthly":
        if group_by == "partner":
            df_grouped = df.groupby(df["date"].dt.to_period("M")).mean(
                numeric_only=True
            )
            df_grouped.index = df_grouped.index.to_timestamp()
            categories = partner_categories
            labels = [d.strftime("%Y-%m") for d in df_grouped.index]
            title = "월간 협력사별 NaN 비율 히트맵 (금요일 기준)"
            y_label = "협력사"
        elif group_by == "model":
            # 기존 방식과 완전히 동일: 월간 모델별 처리
            df_grouped = (
                df.groupby([df["date"].dt.to_period("M"), "model_name"])
                .mean(numeric_only=True)
                .reset_index()
            )
            df_grouped["date"] = df_grouped["date"].apply(lambda x: x.to_timestamp())

            # 모델명 리스트 생성 (categories 변수)
            categories = [
                (row["model_name"], row["model_name"], "blue")
                for _, row in df_grouped[["model_name"]].drop_duplicates().iterrows()
            ]

            # Pivot 테이블 생성: 모든 협력사 비율 컬럼 유지
            df_grouped = df_grouped.pivot(
                index="date",
                columns="model_name",
                values=[col[0] for col in partner_categories],
            )

            # 핵심! 기존 방식: 컬럼명 단순화
            df_grouped.columns = [col[1] for col in df_grouped.columns]

            labels = [d.strftime("%Y-%m") for d in df_grouped.index]
            title = "월간 NaN 비율 추이 (금요일 기준)"
            y_label = "모델"

    # 히트맵 데이터 준비
    if group_by == "partner":
        heatmap_data = df_grouped[[cat[0] for cat in partner_categories]].T
        heatmap_data.index = [cat[1] for cat in partner_categories]
    else:
        # 모델별 히트맵: 기존 방식과 동일하게 groupby로 평균 계산 후 transpose
        heatmap_data = df_grouped.groupby(axis=1, level=0).mean().T

    # 히트맵 생성
    plt.figure(figsize=(12, max(6, len(heatmap_data.index) * 0.6)))
    sns.heatmap(
        heatmap_data,
        annot=True,
        fmt=".1f",
        cmap="YlOrRd",
        cbar_kws={"label": "NaN 비율 (%)"},
        linewidths=0.5,
    )

    if font_prop:
        plt.title(title, fontproperties=font_prop, fontsize=16)
        plt.xlabel("측정 날짜", fontproperties=font_prop)
        plt.ylabel(y_label, fontproperties=font_prop)
        plt.xticks(
            ticks=np.arange(len(labels)) + 0.5,
            labels=labels,
            rotation=45,
            ha="right",
            fontproperties=font_prop,
        )
        plt.yticks(rotation=0, fontproperties=font_prop)
    else:
        plt.title(title, fontsize=16)
        plt.xlabel("측정 날짜")
        plt.ylabel(y_label)
        plt.xticks(
            ticks=np.arange(len(labels)) + 0.5, labels=labels, rotation=45, ha="right"
        )
        plt.yticks(rotation=0)
    plt.tight_layout()

    # 파일명 및 저장
    filename = f"output/{period}_{group_by}_nan_heatmap_{datetime.now().strftime('%Y%m%d')}.png"
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    plt.savefig(filename, bbox_inches="tight")
    plt.close()

    print(f"✅ {title} 생성 완료: {filename}")
    return filename


def generate_weekly_report_heatmap(drive_service, output_path=None):
    """
    이번 주(월~금)의 모든 JSON을 Drive에서 읽어, 협력사별/날짜별 NaN 비율 히트맵을 생성합니다.
    (Drive JSONs -> Weekly Heatmap)
    """
    # 폰트 설정 추가
    global font_prop
    try:
        # font_prop가 정의되어 있는지 확인
        font_prop
    except NameError:
        # font_prop가 정의되지 않은 경우 기본값으로 설정
        font_prop = None
    # 1. 이번 주 날짜 및 주차 계산 (월요일 ~ 금요일)
    today = datetime.now(pytz.timezone("Asia/Seoul"))
    start_of_week = today - timedelta(days=today.weekday())
    end_of_week = start_of_week + timedelta(days=4)
    current_week = today.isocalendar().week
    print(
        f"\n--- 📊 주간 히트맵 생성을 위해 {current_week}주차 데이터를 로드합니다 ---"
    )
    print(
        f"({start_of_week.strftime('%Y-%m-%d')} ~ {end_of_week.strftime('%Y-%m-%d')})"
    )

    # 2. 이번 주차 데이터만 로드 (효율성 개선)
    all_data = load_json_files_from_drive(
        drive_service, period="weekly", week_number=current_week
    )

    if not all_data:
        print("⚠️ 이번 주 데이터가 없어 주간 히트맵을 생성할 수 없습니다.")
        return None

    # 3. DataFrame 생성 및 데이터 변환
    df_data = []
    for d in all_data:
        execution_time = d["execution_time"]
        occurrence_stats = d.get("occurrence_stats", {})
        mech_partner = d.get("mech_partner", "").strip().upper()
        elec_partner = d.get("elec_partner", "").strip().upper()

        # 날짜 파싱
        try:
            if isinstance(execution_time, str):
                if "_" in execution_time:  # 기존 형식: 20250616_132845
                    date_obj = pd.to_datetime(execution_time, format="%Y%m%d_%H%M%S")
                else:  # 새 형식: 2025-06-18 23:12:07
                    date_obj = pd.to_datetime(execution_time)
            else:
                date_obj = pd.to_datetime(execution_time)
        except:
            continue

        # 이번 주에 해당하는 데이터만 필터링
        if not (start_of_week.date() <= date_obj.date() <= end_of_week.date()):
            continue

        entry = {
            "date": date_obj,
            "model_name": d["model_name"],
            "mech_partner": mech_partner,
            "elec_partner": elec_partner,
            "bat_nan_ratio": 0.0,
            "fni_nan_ratio": 0.0,
            "tms_m_nan_ratio": 0.0,
            "cna_nan_ratio": 0.0,
            "pns_nan_ratio": 0.0,
            "tms_e_nan_ratio": 0.0,
            "tms_semi_nan_ratio": 0.0,
        }

        # 기구 협력사별 NaN 비율 계산
        if mech_partner == "BAT":
            target = occurrence_stats.get("기구", {})
            entry["bat_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "FNI":
            target = occurrence_stats.get("기구", {})
            entry["fni_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "TMS":
            target = occurrence_stats.get("기구", {})
            entry["tms_m_nan_ratio"] = ratio_calc(target)

        # 전장 협력사별 NaN 비율 계산
        if elec_partner == "C&A":
            target = occurrence_stats.get("전장", {})
            entry["cna_nan_ratio"] = ratio_calc(target)
        elif elec_partner == "P&S":
            target = occurrence_stats.get("전장", {})
            entry["pns_nan_ratio"] = ratio_calc(target)
        elif elec_partner == "TMS":
            target = occurrence_stats.get("전장", {})
            entry["tms_e_nan_ratio"] = ratio_calc(target)

        # TMS 반제품 NaN 비율
        tms_semi_stats = occurrence_stats.get("TMS_반제품", {})
        entry["tms_semi_nan_ratio"] = ratio_calc(tms_semi_stats)

        df_data.append(entry)

    if not df_data:
        print("⚠️ 이번 주 데이터가 없어 주간 히트맵을 생성할 수 없습니다.")
        return None

    df = pd.DataFrame(df_data)
    df = df.sort_values("date")
    df["day"] = df["date"].dt.strftime("%m월%d일")

    # 협력사 카테고리 정의
    partner_categories = [
        ("bat_nan_ratio", "BAT"),
        ("fni_nan_ratio", "FNI"),
        ("tms_m_nan_ratio", "TMS(m)"),
        ("cna_nan_ratio", "C&A"),
        ("pns_nan_ratio", "P&S"),
        ("tms_e_nan_ratio", "TMS(e)"),
        ("tms_semi_nan_ratio", "TMS_반제품"),
    ]

    # 날짜별 그룹화 및 평균 계산
    df_grouped = df.groupby("day").mean(numeric_only=True)

    if df_grouped.empty:
        print("⚠️ 그룹화된 데이터가 없어 히트맵을 생성할 수 없습니다.")
        return None

    # 히트맵 데이터 준비 (협력사 x 날짜)
    heatmap_data = []
    for col, label in partner_categories:
        if col in df_grouped.columns:
            heatmap_data.append(df_grouped[col].values)
        else:
            heatmap_data.append([0] * len(df_grouped))

    heatmap_array = np.array(heatmap_data)
    partner_labels = [label for _, label in partner_categories]
    date_labels = list(df_grouped.index)

    # 히트맵 생성
    plt.figure(
        figsize=(max(8, len(date_labels) * 1.5), max(6, len(partner_labels) * 0.8))
    )
    sns.heatmap(
        heatmap_array,
        annot=True,
        fmt=".1f",
        cmap="Reds",
        linewidths=0.5,
        xticklabels=date_labels,
        yticklabels=partner_labels,
        cbar_kws={"label": "NaN 비율 (%)"},
    )

    if font_prop:
        plt.title("주간 NaN 비율 추이 (mixed)", fontproperties=font_prop, fontsize=16)
        plt.xlabel("측정 날짜", fontproperties=font_prop)
        plt.ylabel("협력사", fontproperties=font_prop)
        plt.xticks(rotation=0, fontproperties=font_prop)
        plt.yticks(rotation=0, fontproperties=font_prop)
    else:
        plt.title("주간 NaN 비율 추이 (mixed)", fontsize=16)
        plt.xlabel("측정 날짜")
        plt.ylabel("협력사")
        plt.xticks(rotation=0)
        plt.yticks(rotation=0)
    plt.tight_layout()

    # 날짜가 포함된 파일명 생성
    if output_path is None:
        output_path = (
            f"output/weekly_partner_nan_heatmap_{datetime.now().strftime('%Y%m%d')}.png"
        )

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    plt.savefig(output_path, bbox_inches="tight")
    plt.close()

    print(f"✅ 주간 리포트용 히트맵 저장 완료: {output_path}")
    return output_path


def should_generate_monthly_heatmap():
    """월간 히트맵 생성 조건 확인 (33주부터 일요일로 변경)"""
    today = datetime.now(pytz.timezone("Asia/Seoul"))
    current_week = today.isocalendar().week

    if current_week < 33:
        # 32주 이전: 금요일 기준
        if today.weekday() != 4:  # 금요일이 아니면
            return False, "friday"

        # 다음 주 금요일이 다음 달인지 확인
        next_friday = today + timedelta(days=7)
        return next_friday.month != today.month, "friday"
    else:
        # 33주부터: 일요일 기준
        if today.weekday() != 6:  # 일요일이 아니면
            return False, "sunday"

        # 다음 주 일요일이 다음 달인지 확인
        next_sunday = today + timedelta(days=7)
        return next_sunday.month != today.month, "sunday"


def generate_final_html(
    all_results,
    heatmap_path,
    output_filename="partner.html",
    monthly_partner_link=None,
    monthly_model_link=None,
):
    """
    처리된 데이터와 히트맵을 바탕으로 최종 HTML 파일을 생성합니다.
    """
    heatmap_url_for_html = upload_to_drive(heatmap_path) if heatmap_path else None

    html_body = build_combined_email_body(
        all_results,
        heatmap_url=heatmap_url_for_html,
        monthly_partner_url=monthly_partner_link,
        monthly_model_url=monthly_model_link,
    )

    # generate_html_from_content 함수가 파일 저장과 업로드를 같이 하므로 분리할 필요가 있음
    # 지금은 기존 함수를 재활용하여 경로만 반환
    styled_html = f"""
    <html>
    <head>
      <meta charset="UTF-8">
      <title>PDA Dashboard</title>
      <style>
        body {{ font-family: 'NanumGothic', sans-serif; font-size: 12px; margin: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        details {{ margin-bottom: 10px; }}
        summary {{ cursor: pointer; font-weight: bold; color: #333; }}
        details[open] summary {{ color: #0056b3; }}
        details > *:not(summary) {{ display: block; margin-left: 20px; }}
        ul {{ margin: 5px 0; padding-left: 20px; }}
        p {{ margin: 5px 0; }}
        hr {{ margin: 20px 0; }}
        a {{ color: #0056b3; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        img {{ max-width: 100%; height: auto; }}
      </style>
    </head>
    <body>{html_body}</body></html>
    """
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(styled_html)
    print(f"📄 최종 HTML 파일 생성 완료: {output_filename}")
    return output_filename


# ====================================
# MAIN EXECUTION BLOCK (REFACTORED)
# ====================================
if __name__ == "__main__":
    # 1. 데이터 추출 및 가공
    print("--- 1. 데이터 추출 및 가공 시작 ---")
    all_results = collect_and_process_data()

    if all_results:
        # 2. 결과 JSON으로 저장 및 Drive 업로드
        print("\n--- 2. 결과 JSON으로 저장 및 업로드 시작 ---")
        save_results_to_json(all_results, drive_service)  # ID 반환값은 더이상 필요 없음

        # JSON 업로드 후 Drive에서 파일이 완전히 처리될 때까지 잠시 대기
        print("⏱️ Google Drive 파일 처리 대기 중... (5초)")
        systime.sleep(5)

        # 3. 주간/월간 히트맵 생성
        print("\n--- 3. 히트맵 생성 시작 ---")

        # 3-1. 주간 리포트용 히트맵 (이번 주 모든 JSON 취합)
        heatmap_path = generate_weekly_report_heatmap(drive_service)

        # 3-2. 월간 히트맵 생성 (33주부터 일요일 기준으로 변경)
        monthly_partner_link = None
        monthly_model_link = None

        should_generate, target_day = should_generate_monthly_heatmap()
        if should_generate:
            if target_day == "friday":
                print("\n--- 📊 월의 마지막 금요일: 월간 히트맵 생성 시작 ---")
            else:
                print("\n--- 📊 월의 마지막 일요일: 월간 히트맵 생성 시작 ---")

            # 월간 협력사 히트맵
            monthly_partner_heatmap = generate_heatmap(
                drive_service,
                period="monthly",
                group_by="partner",
                target_day=target_day,
            )

            # 월간 모델 히트맵
            monthly_model_heatmap = generate_heatmap(
                drive_service, period="monthly", group_by="model", target_day=target_day
            )

            print(f"✅ 월간 히트맵 생성 완료:")
            if monthly_partner_heatmap:
                print(f"   - 협력사별: {monthly_partner_heatmap}")
                # 드라이브 업로드
                monthly_partner_link = upload_to_drive(monthly_partner_heatmap)
                if monthly_partner_link:
                    print(f"   - 협력사별 드라이브 업로드 완료: {monthly_partner_link}")
            if monthly_model_heatmap:
                print(f"   - 모델별: {monthly_model_heatmap}")
                # 드라이브 업로드
                monthly_model_link = upload_to_drive(monthly_model_heatmap)
                if monthly_model_link:
                    print(f"   - 모델별 드라이브 업로드 완료: {monthly_model_link}")
        else:
            current_week = datetime.now(pytz.timezone("Asia/Seoul")).isocalendar().week
            if current_week < 33:
                print(
                    "ℹ️ 월의 마지막 금요일이 아니므로 월간 히트맵을 생성하지 않습니다."
                )
            else:
                print(
                    "ℹ️ 월의 마지막 일요일이 아니므로 월간 히트맵을 생성하지 않습니다."
                )

        # 4. 최종 HTML 생성
        print("\n--- 4. 최종 HTML 리포트 생성 시작 ---")
        # generate_nan_bar_charts는 all_results를 사용하므로 여기서 호출
        tasks_file, total_file = generate_nan_bar_charts(all_results)
        final_html_path = generate_final_html(
            all_results,
            heatmap_path,
            output_filename="partner.html",
            monthly_partner_link=monthly_partner_link,
            monthly_model_link=monthly_model_link,
        )

        # 5. 알림 및 업로드
        print("\n--- 5. 알림 및 업로드 시작 ---")

        # GitHub 업로드 옵션 확인 (환경변수에서 제어)
        GITHUB_UPLOAD = os.getenv("GITHUB_UPLOAD", "auto").lower()

        should_upload_github = False
        if GITHUB_UPLOAD == "true":
            should_upload_github = True
            print("✅ GITHUB_UPLOAD=true: 강제로 GitHub 업로드 진행")
        elif GITHUB_UPLOAD == "false":
            should_upload_github = False
            print("⛔ GITHUB_UPLOAD=false: GitHub 업로드 비활성화")
        else:
            # 자동 모드: TEST_MODE가 아닐 때만 업로드
            should_upload_github = not TEST_MODE
            if should_upload_github:
                print("✅ 운영 모드: GitHub 업로드 진행")
            else:
                print("🧪 TEST_MODE: GitHub 업로드 생략")

        print(
            f"📤 GitHub 업로드 설정: GITHUB_UPLOAD={GITHUB_UPLOAD}, 실제 업로드 여부: {should_upload_github}"
        )

        if should_upload_github:
            upload_to_github(final_html_path)
            print(f"✅ GitHub 업로드 완료: {final_html_path}")
        else:
            print(f"⛔ GitHub 업로드 생략됨: {final_html_path}")

        # 이메일 발송
        email_body = open(final_html_path, "r", encoding="utf-8").read()
        attachment_files = [f for f in [tasks_file, total_file, heatmap_path] if f]
        send_occurrence_email(
            f"[알림] PDA Overtime 및 NaN 체크 결과 - 총 {len(all_results)}건",
            email_body,
            graph_files=attachment_files,
        )

        # 카카오톡 발송
        send_nan_alert_to_kakao(all_results)

        print("\n✅ [종료] 데이터 중심 파이프라인 처리 완료.")

    else:
        print("\n✅ [종료] 처리할 데이터가 없어 작업을 마칩니다.")
