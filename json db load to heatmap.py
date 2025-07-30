import json
import os
import re
from datetime import datetime, timedelta

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ---------------------------
# 1. Google Drive API 설정
# ---------------------------
sheets_json_key_path = "/Users/kdkyu311/Downloads/gst-manegemnet-70faf8ce1bff.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
DRIVE_FOLDER_ID = "13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1"
credentials = Credentials.from_service_account_file(sheets_json_key_path, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)

# ---------------------------
# 2. 폰트 설정 (NanumGothic)
# ---------------------------
font_paths = [
    "/Users/kdkyu311/Library/Fonts/NanumGothic.ttf",
    "/System/Library/AssetsV2/com_apple_MobileAsset_Font7/bad9b4bf17cf1669dde54184ba4431c22dcad27b.asset/AssetData/NanumGothic.ttc",
    "/Library/Fonts/NanumGothic.ttf",
    "/System/Library/Fonts/Supplemental/NanumGothic.ttf",
]
font_path = next((path for path in font_paths if os.path.exists(path)), None)
if font_path:
    print(f"✅ NanumGothic 폰트 적용 완료: {font_path}")
    font_prop = fm.FontProperties(fname=font_path)
    plt.rc("font", family=font_prop.get_name())
else:
    print("🚨 NanumGothic 폰트를 찾을 수 없습니다. 기본 폰트로 진행합니다.")
plt.rcParams["axes.unicode_minus"] = False


# ---------------------------
# 3. JSON 파일 로드 함수
# ---------------------------
def load_json_files_from_drive(week_number=None, friday_only=False):
    """Google Drive 폴더 내 JSON 파일 로드, week_number 또는 friday_only로 필터링."""
    query = f"'{DRIVE_FOLDER_ID}' in parents and name contains 'nan_ot_results_'"
    if friday_only:
        query += " and name contains '_금_'"
    files = drive_service.files().list(q=query, fields="files(id, name)").execute().get("files", [])
    data_list = []
    if not files:
        print("⚠️ 로드할 JSON 파일이 없습니다.")
        return data_list

    for file in files:
        file_name = file["name"]
        match = re.search(r"nan_ot_results_(\d{8})_", file_name)
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
    total = stats.get("total_count", 0)
    nan_count = stats.get("nan_count", 0)
    return (nan_count / total * 100) if total > 0 else 0


# ---------------------------
# 4. 자동 주간 정보 계산 함수
# ---------------------------
def auto_week_info():
    all_data = load_json_files_from_drive()  # 모든 JSON으로 주간 정보 계산
    if not all_data:
        print("⚠️ JSON 데이터를 로드할 수 없습니다.")
        return None, None
    times = [pd.to_datetime(item["execution_time"], format="%Y%m%d_%H%M%S") for item in all_data]
    latest_date = max(times)
    week_number = latest_date.isocalendar().week
    start_date = latest_date - timedelta(days=latest_date.weekday())  # 월요일
    end_date = start_date + timedelta(days=4)  # 금요일
    print(f"🧠 자동 계산 주간 정보: week_number={week_number}, date_range=({start_date.date()}, {end_date.date()})")
    return week_number, (start_date, end_date)


# ---------------------------
# 5. 월간 그래프 생성 조건: 마지막 주 금요일 감지 함수
# ---------------------------
def is_last_friday_in_month():
    all_data = load_json_files_from_drive(friday_only=True)
    if not all_data:
        return False
    df = pd.DataFrame(
        [
            {"execution_time": d["execution_time"], "date": pd.to_datetime(d["execution_time"], format="%Y%m%d_%H%M%S")}
            for d in all_data
        ]
    )
    df["month"] = df["date"].dt.month
    df["week"] = df["date"].dt.isocalendar().week
    df["weekday"] = df["date"].dt.weekday
    result = {}
    for month, group in df.groupby("month"):
        max_week = group["week"].max()
        last_week_data = group[group["week"] == max_week]
        result[month] = 4 in last_week_data["weekday"].values
    current_month = datetime.now().month
    return result.get(current_month, False)


# ---------------------------
# 6. 그래프 생성 함수
# ---------------------------
def generate_nan_trend_graph(
    period="weekly",
    chart_type="stacked_bar",
    quarterly_targets=None,
    week_number=None,
    date_range=None,
    days_per_week="auto",
    group_by="partner",
):
    if quarterly_targets is None:
        quarterly_targets = {1: 40, 2: 30, 3: 20, 4: 0}

    # 주간: 이번 주 JSON, 월간: 금요일 JSON
    load_week = week_number if period == "weekly" else None
    friday_only = period == "monthly"
    all_data = load_json_files_from_drive(week_number=load_week, friday_only=friday_only)
    if not all_data:
        print("⚠️ 데이터를 로드할 수 없습니다.")
        return None

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
        if mech_partner == "BAT":
            target = occurrence_stats.get("기구", {})
            entry["bat_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "FNI":
            target = occurrence_stats.get("기구", {})
            entry["fni_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "TMS":
            target = occurrence_stats.get("기구", {})
            entry["tms_m_nan_ratio"] = ratio_calc(target)
        if elec_partner == "C&A":
            target = occurrence_stats.get("전장", {})
            entry["cna_nan_ratio"] = ratio_calc(target)
        elif elec_partner == "P&S":
            target = occurrence_stats.get("전장", {})
            entry["pns_nan_ratio"] = ratio_calc(target)
        elif elec_partner == "TMS":
            target = occurrence_stats.get("전장", {})
            entry["tms_e_nan_ratio"] = ratio_calc(target)
        tms_semi_stats = occurrence_stats.get("TMS_반제품", {})
        entry["tms_semi_nan_ratio"] = ratio_calc(tms_semi_stats)
        df_data.append(entry)

    df = pd.DataFrame(df_data)
    df["date"] = pd.to_datetime(df["date"], format="%Y%m%d_%H%M%S", errors="coerce")
    df = df.dropna(subset=["date"])
    df["quarter"] = df["date"].dt.quarter
    df["week"] = df["date"].dt.isocalendar().week

    if week_number is not None:
        df = df[df["week"] == week_number]
        if df.empty:
            print(f"⚠️ {week_number}주차 데이터가 없습니다.")
            return None

    if date_range is not None:
        start_date, end_date = date_range
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        df = df[(df["date"] >= start_date) & (df["date"] <= end_date)]
        if df.empty:
            print(f"⚠️ {start_date} ~ {end_date} 사이의 데이터가 없습니다.")
            return None
        else:
            print("📅 포함된 날짜:", df["date"].dt.date.unique())

    if period == "weekly":
        if days_per_week == "auto":
            unique_weekdays = set(df["date"].dt.weekday.unique())
            if set(range(5)).issubset(unique_weekdays):
                days_per_week = "5days"
            elif {0, 2, 4}.issubset(unique_weekdays) and len(unique_weekdays) == 3:
                days_per_week = "3days"
            else:
                days_per_week = "mixed"
            print(f"🧠 자동 감지된 days_per_week: {days_per_week}")
        if days_per_week == "3days":
            df = df[df["date"].dt.weekday.isin([0, 2, 4])]
        elif days_per_week == "5days":
            df = df[df["date"].dt.weekday.isin([0, 1, 2, 3, 4])]
        else:
            print("🧠 mixed days_per_week: 필터 미적용")

    if df.empty:
        print("⚠️ 필터링 후 데이터가 없습니다.")
        return None

    partner_categories = [
        ("bat_nan_ratio", "BAT", "blue"),
        ("fni_nan_ratio", "FNI", "cyan"),
        ("tms_m_nan_ratio", "TMS(m)", "orange"),
        ("cna_nan_ratio", "C&A", "green"),
        ("pns_nan_ratio", "P&S", "red"),
        ("tms_e_nan_ratio", "TMS(e)", "purple"),
        ("tms_semi_nan_ratio", "TMS_반제품", "magenta"),
    ]

    if period == "weekly":
        df = df.sort_values("date")
        df["day"] = df["date"].dt.date
        if group_by == "partner":
            df_grouped = df.groupby("day").mean(numeric_only=True)
        elif group_by == "model":
            df_grouped = df.groupby(["day", "model_name"]).mean(numeric_only=True).reset_index()
            categories = [
                (row["model_name"], row["model_name"], "blue")
                for _, row in df_grouped[["model_name"]].drop_duplicates().iterrows()
            ]
            df_grouped = df_grouped.pivot(
                index="day", columns="model_name", values=[col[0] for col in partner_categories]
            )
            df_grouped.columns = [col[1] for col in df_grouped.columns]
        else:
            raise ValueError("group_by는 'partner' 또는 'model'이어야 합니다.")
        labels = [f"{d.month}월{d.day}일" for d in df_grouped.index]
        title = f"주간 NaN 비율 추이 ({days_per_week})"
        xlabel = "측정 날짜"
    elif period == "monthly":
        if group_by == "partner":
            df_grouped = df.groupby(df["date"].dt.to_period("M")).mean(numeric_only=True)
            df_grouped.index = df_grouped.index.to_timestamp()
            categories = partner_categories
            labels = [d.strftime("%Y-%m") for d in df_grouped.index]
        elif group_by == "model":
            df_grouped = df.groupby([df["date"].dt.to_period("M"), "model_name"]).mean(numeric_only=True).reset_index()
            df_grouped["date"] = df_grouped["date"].apply(lambda x: x.to_timestamp())
            categories = [
                (row["model_name"], row["model_name"], "blue")
                for _, row in df_grouped[["model_name"]].drop_duplicates().iterrows()
            ]
            df_grouped = df_grouped.pivot(
                index="date", columns="model_name", values=[col[0] for col in partner_categories]
            )
            df_grouped.columns = [col[1] for col in df_grouped.columns]
            labels = [d.strftime("%Y-%m") for d in df_grouped.index]
        else:
            raise ValueError("group_by는 'partner' 또는 'model'이어야 합니다.")
        title = "월간 NaN 비율 추이 (금요일 기준)"
        xlabel = "월"

    df_grouped["quarter"] = df_grouped.index.map(lambda x: pd.to_datetime(x).quarter)

    plt.figure(figsize=(12, 6))
    if chart_type == "heatmap":
        if group_by == "partner":
            heatmap_data = df_grouped[[cat[0] for cat in partner_categories]].T
            heatmap_data.index = [cat[1] for cat in partner_categories]
            y_label = "협력사"
        else:
            heatmap_data = df_grouped.groupby(axis=1, level=0).mean().T
            y_label = "모델"
        sns.heatmap(heatmap_data, annot=True, fmt=".1f", cmap="YlOrRd", cbar_kws={"label": "NaN 비율 (%)"})
        plt.title(title)
        plt.xlabel(xlabel)
        plt.ylabel(y_label)
        plt.xticks(ticks=np.arange(len(labels)) + 0.5, labels=labels, rotation=45, ha="right")
        filename = f"{period}_{group_by}_nan_heatmap_{datetime.now().strftime('%Y%m%d')}.png"
        plt.savefig(filename, bbox_inches="tight")
        plt.close()
        print(f"✅ NaN 히트맵 생성 완료: {filename}")
        return filename
    elif chart_type == "stacked_bar":
        bottom = np.zeros(len(df_grouped))
        for category, label, color in partner_categories:
            if category in df_grouped.columns:
                plt.bar(df_grouped.index, df_grouped[category], bottom=bottom, label=label, color=color)
                for i, (x, y) in enumerate(zip(df_grouped.index, df_grouped[category])):
                    if y >= 30:
                        plt.text(x, bottom[i] + y / 2, f"{y:.1f}%", ha="center", va="center", color="black")
                bottom += df_grouped[category].fillna(0)
        for quarter, target in quarterly_targets.items():
            valid_quarter = df_grouped["quarter"].dropna().unique()
            if quarter in valid_quarter:
                plt.axhline(y=target, color="black", linestyle="--", label=f"{quarter}분기 목표 ({target}%)", alpha=0.5)
    else:
        for category, label, color in partner_categories:
            if category in df_grouped.columns:
                plt.plot(df_grouped.index, df_grouped[category], label=label, marker="o", color=color)
                for x, y in zip(df_grouped.index, df_grouped[category]):
                    if y >= 30:
                        plt.text(x, y, f"{y:.1f}%", ha="center", va="bottom", color=color)

    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel("NaN 비율 (%)")
    plt.legend()
    plt.grid(True)
    plt.ylim(0, 100)
    plt.xticks(ticks=np.arange(len(labels)) + 0.5, labels=labels, rotation=45, ha="right")
    filename = f"{period}_{group_by}_nan_trend_{chart_type}_{datetime.now().strftime('%Y%m%d')}.png"
    plt.savefig(filename, bbox_inches="tight")
    plt.close()
    print(f"✅ NaN 추이 그래프 생성 완료: {filename}")
    return filename


# ---------------------------
# 7. 구글 드라이브 업로드 함수
# ---------------------------
def upload_to_drive(filename):
    file_metadata = {"name": filename, "parents": [DRIVE_FOLDER_ID]}
    media = MediaFileUpload(filename, mimetype="image/png")
    file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    file_id = file.get("id")
    drive_service.permissions().create(fileId=file_id, body={"type": "anyone", "role": "reader"}).execute()
    image_url = f"https://drive.google.com/uc?export=view&id={file_id}"
    print(f"✅ 그래프 파일 구글 드라이브 업로드 완료: {filename}, 파일 ID: {file_id}")
    os.remove(filename)
    print(f"🗑️ 로컬 파일 삭제 완료: {filename}")
    return image_url


# ---------------------------
# 8. 최종 실행
# ---------------------------
if __name__ == "__main__":
    quarterly_targets = {1: 40, 2: 30, 3: 20, 4: 0}

    # 자동으로 주간 정보 계산
    auto_week, auto_range = auto_week_info()
    if auto_week is None or auto_range is None:
        print("⚠️ 주간 정보 자동 계산 실패")
    else:
        week_number = auto_week
        date_range = (auto_range[0].strftime("%Y-%m-%d"), auto_range[1].strftime("%Y-%m-%d"))

    # 주간 그래프 (이번 주 JSON, 협력사 기준, heatmap)
    weekly_partner_heatmap = generate_nan_trend_graph(
        period="weekly",
        chart_type="heatmap",
        quarterly_targets=quarterly_targets,
        week_number=week_number,
        date_range=date_range,
        days_per_week="auto",
        group_by="partner",
    )
    if weekly_partner_heatmap:
        weekly_partner_heatmap_link = upload_to_drive(weekly_partner_heatmap)

    # 월간 그래프 (금요일 JSON, 마지막 주 금요일에만 생성)
    if is_last_friday_in_month():
        monthly_partner_heatmap = generate_nan_trend_graph(
            period="monthly", chart_type="heatmap", quarterly_targets=quarterly_targets, group_by="partner"
        )
        if monthly_partner_heatmap:
            monthly_partner_heatmap_link = upload_to_drive(monthly_partner_heatmap)

        monthly_model_heatmap = generate_nan_trend_graph(
            period="monthly", chart_type="heatmap", quarterly_targets=quarterly_targets, group_by="model"
        )
        if monthly_model_heatmap:
            monthly_model_heatmap_link = upload_to_drive(monthly_model_heatmap)
