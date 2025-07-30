#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
from datetime import datetime, timedelta

import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

# 한글 폰트 설정
font_path = "/Users/kdkyu311/Library/Fonts/NanumGothic.ttf"
if os.path.exists(font_path):
    fm.fontManager.addfont(font_path)
    plt.rcParams["font.family"] = "NanumGothic"
    print(f"✅ NanumGothic 폰트 적용 완료: {font_path}")
else:
    print("⚠️ NanumGothic 폰트를 찾을 수 없습니다.")


def load_local_json_files():
    """로컬 JSON 파일들을 로드"""
    json_files = [
        "output/nan_ot_results_20250616_132845_월_1회차.json",  # 월요일
        "output/nan_ot_results_20250618_235044_수_3회차.json",  # 수요일
    ]

    all_data = []

    for file_path in json_files:
        if os.path.exists(file_path):
            print(f"📁 로드 중: {file_path}")
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    json_data = json.load(f)
                    execution_time = json_data.get("execution_time", "")
                    results = json_data.get("results", [])

                    # 각 결과에 execution_time 추가
                    for result in results:
                        result["execution_time"] = execution_time

                    all_data.extend(results)
                    print(f"   - {len(results)}개 데이터 로드 (실행시간: {execution_time})")
            except Exception as e:
                print(f"❌ 파일 로드 실패: {file_path} - {e}")
        else:
            print(f"⚠️ 파일이 없습니다: {file_path}")

    print(f"✅ 총 {len(all_data)}개 데이터 로드 완료")
    return all_data


def process_data_for_heatmap(data):
    """히트맵용 데이터 처리 - 기존 구조 참고"""
    df_data = []

    def ratio_calc(stats):
        """NaN 비율 계산"""
        total = stats.get("total_count", 0)
        nan_count = stats.get("nan_count", 0)
        return (nan_count / total * 100) if total > 0 else 0

    for d in data:
        occurrence_stats = d.get("occurrence_stats", {})
        mech_partner = d.get("mech_partner", "").strip().upper()
        elec_partner = d.get("elec_partner", "").strip().upper()
        execution_time = d.get("execution_time", "")

        # execution_time을 날짜로 변환
        try:
            date_obj = pd.to_datetime(execution_time, format="%Y%m%d_%H%M%S")
            date_str = date_obj.strftime("%m월%d일")
        except:
            date_str = execution_time

        entry = {
            "date": date_str,
            "execution_time": execution_time,
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

        # 기구 협력사별 NaN 비율 설정
        if mech_partner == "BAT":
            target = occurrence_stats.get("기구", {})
            entry["bat_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "FNI":
            target = occurrence_stats.get("기구", {})
            entry["fni_nan_ratio"] = ratio_calc(target)
        elif mech_partner == "TMS":
            target = occurrence_stats.get("기구", {})
            entry["tms_m_nan_ratio"] = ratio_calc(target)

        # 전장 협력사별 NaN 비율 설정
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

    return df_data


def generate_weekly_heatmap(data):
    """주간 히트맵 생성 - 기존 구조 참고"""
    if not data:
        print("❌ 히트맵 생성할 데이터가 없습니다.")
        return

    # 데이터프레임 생성
    df = pd.DataFrame(data)

    # 협력사 카테고리 정의 (기존 구조와 동일)
    partner_categories = [
        ("bat_nan_ratio", "BAT", "blue"),
        ("fni_nan_ratio", "FNI", "cyan"),
        ("tms_m_nan_ratio", "TMS(m)", "orange"),
        ("cna_nan_ratio", "C&A", "green"),
        ("pns_nan_ratio", "P&S", "red"),
        ("tms_e_nan_ratio", "TMS(e)", "purple"),
        ("tms_semi_nan_ratio", "TMS_반제품", "magenta"),
    ]

    # 날짜별 평균 계산 (기존 로직과 동일)
    df_grouped = df.groupby("date").mean(numeric_only=True)

    print("📊 날짜별 평균 NaN 비율:")
    for date in df_grouped.index:
        print(f"   {date}:")
        for category, label, _ in partner_categories:
            if category in df_grouped.columns:
                value = df_grouped.loc[date, category]
                if value >= 0:  # 0 이상 모든 값 표시
                    print(f"     {label}: {value:.1f}%")

    # 히트맵 데이터 준비
    heatmap_data = df_grouped[[cat[0] for cat in partner_categories]].T
    heatmap_data.index = [cat[1] for cat in partner_categories]

    # 0인 행도 유지 (기존 구조처럼)
    # heatmap_data = heatmap_data[(heatmap_data != 0).any(axis=1)]  # 이 줄 제거

    print(f"\n📊 히트맵 데이터:")
    print(heatmap_data)

    # 히트맵 생성
    plt.figure(figsize=(12, 8))

    # 컬러맵 설정 (YlOrRd: 노랑->주황->빨강)
    sns.heatmap(
        heatmap_data,
        annot=True,
        fmt=".1f",
        cmap="YlOrRd",
        cbar_kws={"label": "NaN 비율 (%)"},
        square=False,
        linewidths=0.5,
    )

    # 제목 및 레이블 설정 (기존 형식에 맞게)
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)

    plt.title(f"주간 NaN 비율 추이 (mixed)", fontsize=16, fontweight="bold", pad=20)
    plt.xlabel("측정 날짜", fontsize=12, fontweight="bold")
    plt.ylabel("협력사", fontsize=12, fontweight="bold")

    # X축 라벨 설정 (날짜)
    labels = list(heatmap_data.columns)
    plt.xticks(ticks=np.arange(len(labels)) + 0.5, labels=labels, rotation=45, ha="right")

    # 레이아웃 조정
    plt.tight_layout()

    # 파일 저장
    output_path = "output/weekly_partner_nan_heatmap_final.png"
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close()

    print(f"✅ 최종 주간 히트맵 저장 완료: {output_path}")


def main():
    """메인 함수"""
    print("=== 로컬 JSON 파일로 주간 히트맵 생성 ===")

    # JSON 데이터 로드
    json_data = load_local_json_files()
    if not json_data:
        print("❌ 로드할 데이터가 없습니다.")
        return

    # 데이터 처리
    processed_data = process_data_for_heatmap(json_data)
    if not processed_data:
        print("❌ 처리된 데이터가 없습니다.")
        return

    # 주간 히트맵 생성
    generate_weekly_heatmap(processed_data)

    print("=== 주간 히트맵 생성 완료 ===")


if __name__ == "__main__":
    main()
