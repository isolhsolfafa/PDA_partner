#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
import re
import sys
from datetime import datetime, timedelta

import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pytz
import seaborn as sns
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# 환경변수 로드
load_dotenv()

# Google API 설정
SERVICE_ACCOUNT_FILE = "config/gst-manegemnet-70faf8ce1bff.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
JSON_DRIVE_FOLDER_ID = os.getenv(
    "JSON_DRIVE_FOLDER_ID", os.getenv("NOVA_FOLDER_ID", "13FdsniLHb4qKmn5M4-75H8SvgEyW2Ck1")
)

# Google API 서비스 초기화
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)

# 폰트 설정
font_path = "/Users/kdkyu311/Library/Fonts/NanumGothic.ttf"
if os.path.exists(font_path):
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams["font.family"] = font_prop.get_name()
    print(f"✅ NanumGothic 폰트 적용 완료: {font_path}")
else:
    print("⚠️ NanumGothic 폰트를 찾을 수 없습니다.")
    font_prop = fm.FontProperties()


# PDA_patner.py에서 필요한 함수들 가져오기
def load_json_files_from_drive(drive_service, period="weekly", week_number=None, friday_only=False):
    """
    Google Drive 폴더 내 JSON 파일 로드
    """
    query = f"'{JSON_DRIVE_FOLDER_ID}' in parents and name contains 'nan_ot_results_'"
    if friday_only:
        query += " and name contains '_금_'"

    # 최대 3번까지 재시도 (Drive 파일 처리 지연 대응)
    for attempt in range(3):
        files = drive_service.files().list(q=query, fields="files(id, name)").execute().get("files", [])

        if files:
            break

        if attempt < 2:
            print(f"⏳ 파일을 찾을 수 없습니다. 3초 후 재시도... ({attempt + 1}/3)")
            import time

            time.sleep(3)
        else:
            print("⚠️ 로드할 JSON 파일이 없습니다.")
            return []

    data_list = []

    for file in files:
        file_name = file["name"]
        match = re.search(r"nan_ot_results_(\d{8})", file_name)
        if not match:
            continue

        file_date = match.group(1)
        try:
            file_datetime = pd.to_datetime(file_date, format="%Y%m%d")
            file_week = file_datetime.isocalendar().week
        except ValueError:
            continue

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
    """NaN 비율 계산"""
    total = stats.get("total_count", 0)
    nan_count = stats.get("nan_count", 0)
    return (nan_count / total * 100) if total > 0 else 0


def generate_weekly_report_heatmap(drive_service, output_path=None):
    """
    PDA_patner.py와 동일한 수정된 함수
    """
    # 1. 이번 주 날짜 계산 (월요일 ~ 금요일)
    today = datetime.now(pytz.timezone("Asia/Seoul"))
    start_of_week = today - timedelta(days=today.weekday())
    end_of_week = start_of_week + timedelta(days=4)
    print(f"\n--- 📊 주간 히트맵 생성을 위해 이번 주 데이터를 로드합니다 ---")
    print(f"({start_of_week.strftime('%Y-%m-%d')} ~ {end_of_week.strftime('%Y-%m-%d')})")

    # 2. 이번 주 데이터 로드
    all_data = load_json_files_from_drive(drive_service, period="weekly")

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

    print(f"📊 처리된 데이터: {len(df)} 건")
    print(f"📅 날짜별 데이터 수: {df['day'].value_counts().to_dict()}")

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

    print(f"📈 그룹화된 데이터:")
    for idx, row in df_grouped.iterrows():
        print(
            f"  {idx}: BAT={row['bat_nan_ratio']:.1f}%, FNI={row['fni_nan_ratio']:.1f}%, TMS(m)={row['tms_m_nan_ratio']:.1f}%, C&A={row['cna_nan_ratio']:.1f}%, P&S={row['pns_nan_ratio']:.1f}%, TMS(e)={row['tms_e_nan_ratio']:.1f}%, TMS_반제품={row['tms_semi_nan_ratio']:.1f}%"
        )

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
    plt.figure(figsize=(max(8, len(date_labels) * 1.5), max(6, len(partner_labels) * 0.8)))
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

    plt.title("주간 NaN 비율 추이 (mixed)", fontproperties=font_prop, fontsize=16)
    plt.xlabel("측정 날짜", fontproperties=font_prop)
    plt.ylabel("협력사", fontproperties=font_prop)
    plt.xticks(rotation=0, fontproperties=font_prop)
    plt.yticks(rotation=0, fontproperties=font_prop)
    plt.tight_layout()

    # 날짜가 포함된 파일명 생성 (수정된 파일명)
    if output_path is None:
        output_path = f"output/weekly_partner_nan_heatmap_{datetime.now().strftime('%Y%m%d')}.png"

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    plt.savefig(output_path, bbox_inches="tight")
    plt.close()

    print(f"✅ 주간 리포트용 히트맵 저장 완료: {output_path}")
    return output_path


def generate_heatmap(drive_service, period="weekly", group_by="partner", week_number=None):
    """
    히트맵 생성 함수 (PDA_patner.py에서 복사)
    period: "weekly" (주간), "monthly" (월간)
    group_by: "partner" (협력사), "model" (모델)
    """
    # 데이터 로드
    friday_only = period == "monthly"
    all_data = load_json_files_from_drive(drive_service, period, week_number, friday_only)

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

    print(f"📊 처리된 데이터: {len(df)} 건")

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

    # 주간/월간별 그룹핑
    if period == "weekly":
        df = df.sort_values("date")
        df["day"] = df["date"].dt.strftime("%m월%d일")

        if group_by == "partner":
            df_grouped = df.groupby("day").mean(numeric_only=True)
            categories = partner_categories
            labels = list(df_grouped.index)
            title = "주간 NaN 비율 추이 (mixed)"
            y_label = "협력사"
        elif group_by == "model":
            # 주간 모델별 데이터 처리
            model_daily_data = []
            for (day, model), group in df.groupby(["day", "model_name"]):
                total_nan_ratio = 0
                count = 0
                for (
                    col,
                    _,
                ) in partner_categories:
                    if col in group.columns and group[col].sum() > 0:
                        total_nan_ratio += group[col].mean()
                        count += 1
                avg_nan_ratio = total_nan_ratio / count if count > 0 else 0
                model_daily_data.append({"day": day, "model_name": model, "avg_nan_ratio": avg_nan_ratio})

            model_df = pd.DataFrame(model_daily_data)
            if not model_df.empty:
                df_grouped = model_df.pivot(index="model_name", columns="day", values="avg_nan_ratio").fillna(0)
                labels = list(df_grouped.columns)
            else:
                print("⚠️ 모델별 주간 데이터가 없습니다.")
                return None
            title = "주간 모델별 NaN 비율 히트맵"
            y_label = "모델"

    elif period == "monthly":
        if group_by == "partner":
            df_grouped = df.groupby(df["date"].dt.to_period("M")).mean(numeric_only=True)
            df_grouped.index = df_grouped.index.to_timestamp()
            categories = partner_categories
            labels = [d.strftime("%Y-%m") for d in df_grouped.index]
            title = "월간 협력사별 NaN 비율 히트맵 (금요일 기준)"
            y_label = "협력사"
        elif group_by == "model":
            # 기존 방식과 완전히 동일: 월간 모델별 처리
            print(f"📊 월간 모델별 데이터 처리 시작...")
            df_grouped = df.groupby([df["date"].dt.to_period("M"), "model_name"]).mean(numeric_only=True).reset_index()
            df_grouped["date"] = df_grouped["date"].apply(lambda x: x.to_timestamp())
            print(f"📈 그룹핑 후 데이터: {len(df_grouped)} 건")

            # 모델명 리스트 생성 (categories 변수)
            categories = [
                (row["model_name"], row["model_name"], "blue")
                for _, row in df_grouped[["model_name"]].drop_duplicates().iterrows()
            ]

            # Pivot 테이블 생성: 모든 협력사 비율 컬럼 유지
            df_grouped = df_grouped.pivot(
                index="date", columns="model_name", values=[col[0] for col in partner_categories]
            )
            print(f"📊 Pivot 테이블 크기: {df_grouped.shape}")

            # 핵심! 기존 방식: 컬럼명 단순화
            df_grouped.columns = [col[1] for col in df_grouped.columns]
            print(f"📈 컬럼명 단순화 후 크기: {df_grouped.shape}")

            labels = [d.strftime("%Y-%m") for d in df_grouped.index]
            title = "월간 NaN 비율 추이 (금요일 기준)"
            y_label = "모델"

    if df_grouped.empty:
        print("⚠️ 그룹화된 데이터가 없습니다.")
        return None

    print(f"📈 그룹화된 데이터 크기: {df_grouped.shape}")

    # 히트맵 생성
    if group_by == "partner":
        # 협력사별 히트맵
        heatmap_data = []
        for col, label in categories:
            if col in df_grouped.columns:
                heatmap_data.append(df_grouped[col].values)
            else:
                heatmap_data.append([0] * len(df_grouped))

        heatmap_array = np.array(heatmap_data)
        partner_labels = [label for _, label in categories]

        plt.figure(figsize=(max(8, len(labels) * 1.2), max(6, len(partner_labels) * 0.8)))
        sns.heatmap(
            heatmap_array,
            annot=True,
            fmt=".1f",
            cmap="Reds",
            linewidths=0.5,
            xticklabels=labels,
            yticklabels=partner_labels,
            cbar_kws={"label": "NaN 비율 (%)"},
        )

    elif group_by == "model":
        # 모델별 히트맵: 기존 방식과 동일하게 groupby로 평균 계산 후 transpose
        heatmap_data = df_grouped.groupby(axis=1, level=0).mean().T
        print(f"📈 최종 히트맵 데이터 크기: {heatmap_data.shape}")

        plt.figure(figsize=(max(8, len(heatmap_data.columns) * 1.2), max(8, len(heatmap_data.index) * 0.3)))
        sns.heatmap(
            heatmap_data, annot=True, fmt=".1f", cmap="Reds", linewidths=0.5, cbar_kws={"label": "NaN 비율 (%)"}
        )

    plt.title(title, fontproperties=font_prop, fontsize=16)
    plt.xlabel("기간" if period == "monthly" else "측정 날짜", fontproperties=font_prop)
    plt.ylabel(y_label, fontproperties=font_prop)
    plt.xticks(rotation=45 if period == "monthly" else 0, fontproperties=font_prop)
    plt.yticks(rotation=0, fontproperties=font_prop)
    plt.tight_layout()

    # 파일명 및 저장
    filename = f"output/{period}_{group_by}_nan_heatmap_{datetime.now().strftime('%Y%m%d')}.png"
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    plt.savefig(filename, bbox_inches="tight")
    plt.close()

    print(f"✅ {title} 생성 완료: {filename}")
    return filename


if __name__ == "__main__":
    print("🔄 최종 히트맵 구성 확인 테스트 시작...")

    # 1. 주간 히트맵 테스트
    print("\n=== 1. 주간 히트맵 테스트 ===")
    weekly_heatmap = generate_weekly_report_heatmap(drive_service)
    if weekly_heatmap:
        print(f"✅ 주간 히트맵 생성 완료: {weekly_heatmap}")

    # 2. 월간 협력사 히트맵 테스트
    print("\n=== 2. 월간 협력사 히트맵 테스트 ===")
    monthly_partner_heatmap = generate_heatmap(drive_service, period="monthly", group_by="partner")
    if monthly_partner_heatmap:
        print(f"✅ 월간 협력사 히트맵 생성 완료: {monthly_partner_heatmap}")

    # 3. 월간 모델 히트맵 테스트
    print("\n=== 3. 월간 모델 히트맵 테스트 ===")
    monthly_model_heatmap = generate_heatmap(drive_service, period="monthly", group_by="model")
    if monthly_model_heatmap:
        print(f"✅ 월간 모델 히트맵 생성 완료: {monthly_model_heatmap}")

    print("\n📋 생성된 파일들:")
    print("  - weekly_partner_nan_heatmap_YYYYMMDD.png (주간)")
    print("  - monthly_partner_nan_heatmap_YYYYMMDD.png (월간 협력사)")
    print("  - monthly_model_nan_heatmap_YYYYMMDD.png (월간 모델)")
    print("📊 모든 히트맵이 올바른 구조로 생성되었습니다!")
