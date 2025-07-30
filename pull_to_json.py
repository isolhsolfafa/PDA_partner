# type: ignore
import json
import os
import re
from datetime import datetime

from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# 🔧 설정 영역
sheets_json_key_path = "/Users/kdkyu311/Downloads/gst-manegemnet-70faf8ce1bff.json"
html_file_path = "/Users/kdkyu311/Desktop/GST/협력사/index_06.13.html"
DRIVE_FOLDER_ID = "13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# 파일 존재 여부 확인
if not os.path.exists(sheets_json_key_path):
    raise FileNotFoundError(f"JSON 키 파일을 찾을 수 없습니다: {sheets_json_key_path}")
if not os.path.exists(html_file_path):
    raise FileNotFoundError(f"HTML 파일을 찾을 수 없습니다: {html_file_path}")

# 🔐 인증 및 서비스 초기화
credentials = Credentials.from_service_account_file(sheets_json_key_path, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)

# 🧾 HTML 로드
with open(html_file_path, "r", encoding="utf-8") as f:
    soup = BeautifulSoup(f.read(), "html.parser")

# 🕒 실행 시간 파싱
execution_tag = soup.find("p", string=re.compile("📅 실행 시간"))
if not execution_tag:
    raise ValueError("실행 시간 태그를 찾을 수 없습니다.")
exec_time_match = re.search(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", execution_tag.text)
if not exec_time_match:
    raise ValueError(f"실행 시간 형식이 예상과 다릅니다: {execution_tag.text}")
exec_time = exec_time_match.group()
execution_time = exec_time.replace("-", "").replace(" ", "_").replace(":", "")
exec_dt = datetime.strptime(exec_time, "%Y-%m-%d %H:%M:%S")
weekday_kor = ["월", "화", "수", "목", "금", "토", "일"][exec_dt.weekday()]
session = {0: 1, 1: 2, 2: 3, 3: 4, 4: 5}.get(exec_dt.weekday(), 0)
filename = f"nan_ot_results_{execution_time}_{weekday_kor}_{session}회차.json"

# 📦 JSON 구조 초기화
json_data = {"execution_time": execution_time, "results": []}

# 📊 요약 테이블 추출
table = soup.find("table", id="summaryTable")
if not table:
    raise ValueError("테이블을 찾을 수 없습니다.")
rows = table.find_all("tr")
if len(rows) < 2:
    raise ValueError("테이블에 데이터 행이 없습니다.")

# 헤더 찾기: 첫 번째 행을 헤더로 간주
header_row = rows[0]
header_cells = header_row.find_all("th")
if not header_cells:
    print("테이블 첫 5행 구조:")
    for i, row in enumerate(rows[:5]):
        cells = row.find_all(["th", "td"])
        cell_texts = [cell.get_text().strip() for cell in cells]
        print(f"행 {i}: {cell_texts}")
    raise ValueError("헤더 행을 찾을 수 없습니다. 테이블 구조를 확인하세요.")

# 헤더 텍스트 추출: data-column 속성 우선 사용
headers = []
for cell in header_cells:
    # data-column 속성 사용
    header_name = cell.get("data-column", "")
    if not header_name:
        # 대체로 텍스트 사용, <select> 태그 제외
        for select in cell.find_all("select"):
            select.decompose()
        header_name = cell.get_text(separator=" ").strip().split(" ")[0]
    headers.append(header_name)

# 필드 매핑
field_map = {
    "order_no": "Order",
    "model_name": "모델명",
    "mech_partner": "기구협력사",
    "elec_partner": "전장협력사",
    "total_tasks": ["총 작업 수", "총"],  # "총" 허용
}
field_indices = {}
for key, expected in field_map.items():
    if isinstance(expected, list):
        for idx, h in enumerate(headers):
            if any(h.lower().strip() == e.lower().strip() for e in expected):
                field_indices[key] = idx
                break
    else:
        for idx, h in enumerate(headers):
            if h.lower().strip() == expected.lower().strip():
                field_indices[key] = idx
                break
if len(field_indices) < len(field_map):
    print(f"헤더: {headers}")
    raise ValueError(f"일부 필요한 필드가 헤더에 없습니다. 발견된 필드: {field_indices}")

# 데이터 행 처리: 헤더 행 이후
data_rows = []
for row in rows[1:]:
    cells = row.find_all("td")
    if cells and len(cells) >= len(headers):
        order_no_cell = cells[field_indices["order_no"]]
        order_no = order_no_cell.get_text().strip()
        # 주문 번호 형식(/ 포함) 확인
        if order_no and "/" in order_no:
            data_rows.append(row)

for row in data_rows:
    cols = row.find_all("td")
    order_no_cell = cols[field_indices["order_no"]]
    order_no = order_no_cell.get_text().strip()
    # 하이퍼링크 추출
    order_link = order_no_cell.find("a")
    order_href = order_link["href"] if order_link else ""

    try:
        total_tasks = int(cols[field_indices["total_tasks"]].get_text().strip())
    except ValueError:
        continue  # 숫자로 변환 불가 시 건너뛰기

    json_data["results"].append(
        {
            "order_no": order_no,
            "model_name": cols[field_indices["model_name"]].get_text().strip(),
            "mech_partner": cols[field_indices["mech_partner"]].get_text().strip(),
            "elec_partner": cols[field_indices["elec_partner"]].get_text().strip(),
            "total_tasks": total_tasks,
            "ratios": {},
            "links": {"order_href": order_href},
        }
    )

# 📋 상세 데이터 파싱
categories = ["기구", "TMS_반제품", "전장", "검사", "마무리", "기타"]
details_tags = soup.find_all("details")

for detail in details_tags:
    summary = detail.find("summary")
    if not summary:
        continue

    # order_no 추출
    order_match = re.search(r"Order: (.+?),", summary.get_text())
    if not order_match:
        continue
    order_no = order_match.group(1).strip()

    result = next((r for r in json_data["results"] if r["order_no"] == order_no), None)
    if not result:
        continue

    occurrence_stats = {}
    partner_stats = {"mech": {"nan_count": 0, "ot_count": 0}, "elec": {"nan_count": 0, "ot_count": 0}}
    ratios = {}

    for category in categories:
        p_tag = next((p for p in detail.find_all("p") if f"🔹 {category} 작업" in p.get_text()), None)
        if not p_tag:
            occurrence_stats[category] = {"total_count": 0, "nan_count": 0, "ot_count": 0}
            continue
        text = p_tag.get_text(separator="\n")
        total_match = re.search(r"전체 작업 수:\s*(\d+)", text)
        nan_match = re.search(r"누락\(NaN\):\s*(\d+)", text)
        ot_match = re.search(r"오버타임:\s*(\d+)", text)
        total = int(total_match.group(1)) if total_match else 0
        nan = int(nan_match.group(1)) if nan_match else 0
        ot = int(ot_match.group(1)) if ot_match else 0
        occurrence_stats[category] = {"total_count": total, "nan_count": nan, "ot_count": ot}

        if category == "기구":
            partner_stats["mech"]["nan_count"] = nan
            partner_stats["mech"]["ot_count"] = ot
            ratios["mech_nan_ratio"] = nan / total * 100 if total else 0
            ratios["mech_ot_ratio"] = ot / total * 100 if total else 0
        elif category == "전장":
            partner_stats["elec"]["nan_count"] = nan
            partner_stats["elec"]["ot_count"] = ot
            ratios["elec_nan_ratio"] = nan / total * 100 if total else 0
            ratios["elec_ot_ratio"] = ot / total * 100 if total else 0
        elif category == "TMS_반제품":
            ratios["tms_nan_ratio"] = nan / total * 100 if total else 0
            ratios["tms_ot_ratio"] = ot / total * 100 if total else 0

    result["occurrence_stats"] = occurrence_stats
    result["partner_stats"] = partner_stats
    result["ratios"] = ratios
    result["spreadsheet_url"] = ""

# 💾 JSON 저장
with open(filename, "w", encoding="utf-8") as f:
    json.dump(json_data, f, ensure_ascii=False, indent=2)
print(f"✅ JSON 저장 완료: {filename}")

# ☁️ 구글 드라이브 업로드
file_metadata = {"name": filename, "parents": [DRIVE_FOLDER_ID]}
media = MediaFileUpload(filename, mimetype="application/json")
uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
print(f"✅ 드라이브 업로드 완료: ID = {uploaded['id']}")

# 🗑️ 로컬 파일 삭제
if os.path.exists(filename):
    os.remove(filename)
    print(f"🗑️ 로컬 파일 삭제 완료: {filename}")
else:
    print(f"⚠️ 삭제하려는 파일이 존재하지 않습니다: {filename}")
