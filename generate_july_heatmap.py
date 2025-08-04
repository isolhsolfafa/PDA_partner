#!/usr/bin/env python3
"""
2025년 7월 월간 히트맵 생성 스크립트
7월 금요일 JSON 데이터만 로드하여 월간 히트맵 생성
"""

import os
import sys
import json
import re
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
import numpy as np
from datetime import datetime, timedelta
import pytz
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# .env 파일에서 환경변수 로드
from dotenv import load_dotenv
load_dotenv()

print("✅ .env 파일에서 환경변수를 로드했습니다.")

# Google Drive 서비스 초기화
def init_drive_service():
    DRIVE_KEY_PATH = os.getenv("DRIVE_KEY_PATH")
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    
    credentials = Credentials.from_service_account_file(DRIVE_KEY_PATH, scopes=SCOPES)
    return build("drive", "v3", credentials=credentials)

# 폰트 설정 (PDA_partner.py와 동일한 로직)
def setup_font():
    try:
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
            plt.rcParams["axes.unicode_minus"] = False
            return font_prop
        else:
            # 시스템에서 설치된 한글 폰트를 동적으로 찾기
            print("🔍 시스템에서 한글 폰트를 검색 중...")
            available_fonts = [f.name for f in fm.fontManager.ttflist]
            korean_fonts = [font for font in available_fonts if any(keyword in font.lower() for keyword in ['nanum', 'malgun', 'dotum', 'gulim', 'batang'])]
            
            if korean_fonts:
                selected_font = korean_fonts[0]
                print(f"✅ 한글 폰트 발견 및 적용: {selected_font}")
                font_prop = fm.FontProperties(family=selected_font)
                plt.rc("font", family=selected_font)
                plt.rcParams["axes.unicode_minus"] = False
                return font_prop
            else:
                print("🚨 한글 폰트를 찾을 수 없습니다. 기본 폰트로 진행합니다.")
                plt.rcParams["axes.unicode_minus"] = False
                return None
                
    except Exception as e:
        print(f"❌ 폰트 설정 실패: {e}")
        plt.rcParams["axes.unicode_minus"] = False
        return None

# 월별 JSON 파일 로드 (3월~7월 트렌드)
def load_monthly_json_files(drive_service, start_month=3, end_month=7):
    """2025년 3월~7월 금요일 JSON 파일 로드 (트렌드 분석용)"""
    JSON_DRIVE_FOLDER_ID = os.getenv("JSON_DRIVE_FOLDER_ID")
    
    all_files = []
    for month in range(start_month, end_month + 1):
        month_str = f"2025{month:02d}"
        # 각 월의 금요일 파일 쿼리
        query = f"'{JSON_DRIVE_FOLDER_ID}' in parents and name contains 'nan_ot_results_{month_str}' and name contains '_금_'"
        
        files = drive_service.files().list(q=query, fields="files(id, name)").execute().get("files", [])
        all_files.extend(files)
        print(f"📁 {month}월 금요일 JSON 파일 {len(files)}개 발견")
    
    print(f"📁 총 {len(all_files)}개 파일 발견:")
    for file in all_files:
        print(f"   - {file['name']}")
    
    data_list = []
    
    for file in all_files:
        file_name = file["name"]
        print(f"📂 로딩 중: {file_name}")
        
        file_id = file["id"]
        request = drive_service.files().get_media(fileId=file_id)
        content = request.execute().decode("utf-8")
        data = json.loads(content)
        
        # execution_time을 각 결과에 추가
        for result in data["results"]:
            result["execution_time"] = data["execution_time"]
        
        data_list.extend(data["results"])
    
    print(f"📊 총 {len(data_list)}개의 월별 데이터 로드 완료")
    return data_list

# ratio_calc 함수 (PDA_partner.py와 동일)
def ratio_calc(stats):
    """NaN 비율 계산"""
    total = stats.get("total_count", 0)
    nan_count = stats.get("nan_count", 0)
    return (nan_count / total * 100) if total > 0 else 0.0

# 월별 트렌드 히트맵 생성 (3월~7월)
def generate_monthly_trend_heatmap(data_list, group_by="partner"):
    """3월~7월 데이터로 월별 트렌드 히트맵 생성 (PDA_partner.py 로직 기반)"""
    
    if not data_list:
        print("❌ 월별 데이터가 없습니다.")
        return None
    
    # DataFrame 생성 (PDA_partner.py와 동일한 로직)
    df_data = []
    for d in data_list:
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
    # execution_time 포맷: "20250725_223100" -> "2025-07-25"
    df["date"] = pd.to_datetime(df["date"], format="%Y%m%d_%H%M%S", errors="coerce")
    df = df.dropna(subset=["date"])

    # 협력사 카테고리 정의 (PDA_partner.py와 동일)
    partner_categories = [
        ("bat_nan_ratio", "BAT", "blue"),
        ("fni_nan_ratio", "FNI", "cyan"),
        ("tms_m_nan_ratio", "TMS(m)", "orange"),
        ("cna_nan_ratio", "C&A", "green"),
        ("pns_nan_ratio", "P&S", "red"),
        ("tms_e_nan_ratio", "TMS(e)", "purple"),
        ("tms_semi_nan_ratio", "TMS_반제품", "magenta"),
    ]

    if group_by == "partner":
        # 월별 그룹화 (3월~7월 트렌드)
        df_grouped = df.groupby(df["date"].dt.to_period("M")).mean(numeric_only=True)
        df_grouped.index = df_grouped.index.to_timestamp()
        
        categories = partner_categories
        labels = [d.strftime("%Y-%m") for d in df_grouped.index]
        title = "월간 협력사별 NaN 비율 추이 (금요일 기준)"
        y_label = "협력사"
        
        # 히트맵 데이터 준비
        heatmap_data = df_grouped[[cat[0] for cat in partner_categories]].T
        heatmap_data.index = [cat[1] for cat in partner_categories]
        
    elif group_by == "model":
        # 모델별 그룹화 (PDA_partner.py와 동일한 로직)
        df_grouped = df.groupby([df["date"].dt.to_period("M"), "model_name"]).mean(numeric_only=True).reset_index()
        df_grouped["date"] = df_grouped["date"].apply(lambda x: x.to_timestamp())

        # 모델명 리스트 생성
        categories = [
            (row["model_name"], row["model_name"], "blue")
            for _, row in df_grouped[["model_name"]].drop_duplicates().iterrows()
        ]

        # Pivot 테이블 생성: 모든 협력사 비율 컬럼 유지
        df_grouped = df_grouped.pivot(
            index="date", columns="model_name", values=[col[0] for col in partner_categories]
        )

        # 컬럼명 단순화
        df_grouped.columns = [col[1] for col in df_grouped.columns]

        labels = [d.strftime("%Y-%m") for d in df_grouped.index]
        title = "월간 NaN 비율 추이 (금요일 기준)"
        y_label = "모델"
        
        # 모델별 히트맵: 기존 방식과 동일하게 groupby로 평균 계산 후 transpose
        heatmap_data = df_grouped.groupby(axis=1, level=0).mean().T

    # 히트맵 생성
    plt.figure(figsize=(12, max(6, len(heatmap_data.index) * 0.6)))
    sns.heatmap(heatmap_data, annot=True, fmt=".1f", cmap="YlOrRd",
                xticklabels=labels, yticklabels=heatmap_data.index,
                cbar_kws={"label": "NaN 비율 (%)"})

    plt.title(title, fontsize=16, fontweight="bold", pad=20)
    plt.xlabel("월", fontsize=12)
    plt.ylabel(y_label, fontsize=12)
    plt.xticks(rotation=45)
    plt.yticks(rotation=0)
    plt.tight_layout()

    # 파일 저장
    filename = f"monthly_trend_{group_by}_nan_heatmap_2025.png"
    plt.savefig(filename, dpi=300, bbox_inches="tight")
    plt.close()

    print(f"✅ {title} 생성 완료: {filename}")
    return filename

def main():
    print("🚀 2025년 3월~7월 월별 트렌드 히트맵 생성 시작...")
    
    # 폰트 설정
    setup_font()
    
    # Drive 서비스 초기화
    drive_service = init_drive_service()
    
    # 3월~7월 JSON 데이터 로드
    monthly_data = load_monthly_json_files(drive_service, start_month=3, end_month=7)
    
    if not monthly_data:
        print("❌ 월별 데이터를 찾을 수 없습니다.")
        return
    
    # 협력사별 히트맵 생성 (완전한 카테고리)
    partner_heatmap = generate_monthly_trend_heatmap(monthly_data, group_by="partner")
    
    # 모델별 히트맵 생성  
    model_heatmap = generate_monthly_trend_heatmap(monthly_data, group_by="model")
    
    print("\n🎉 월별 트렌드 히트맵 생성 완료!")
    if partner_heatmap:
        print(f"   📊 협력사별: {partner_heatmap}")
    if model_heatmap:
        print(f"   📊 모델별: {model_heatmap}")
    
    print("\n📋 생성된 히트맵 특징:")
    print("   ✅ 3월~7월 월별 트렌드 표시")
    print("   ✅ 전체 협력사 카테고리 포함 (BAT, FNI, TMS(m), C&A, P&S, TMS(e), TMS_반제품)")
    print("   ✅ PDA_partner.py와 동일한 로직 적용")
    print("   ✅ 한글 폰트 정상 적용")

if __name__ == "__main__":
    main()