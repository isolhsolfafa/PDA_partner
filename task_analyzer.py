import logging
import re
from datetime import date, datetime, time, timedelta

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

# 공휴일 및 근무시간 설정 (PDA_partner.py와 동일)
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
        fmt = "%Y. %m. %d %p %I:%M:%S" if "AM" in s or "PM" in s else "%Y. %m. %d %H:%M:%S"
        return pd.to_datetime(s, format=fmt, errors="coerce")
    return pd.NaT


def calculate_working_hours_with_holidays(start_time, end_time):
    if pd.isna(start_time) or pd.isna(end_time):
        return 0
    total_hours = 0
    current_time = start_time
    while current_time < end_time:
        day_of_week = current_time.weekday()
        work_start = datetime.combine(current_time.date(), time(8, 0, 0) if day_of_week in [5, 6] else WORK_START)
        work_end = datetime.combine(current_time.date(), time(17, 0, 0) if day_of_week in [5, 6] else WORK_END)
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
                daily_hours -= (min(work_end, b_end_dt) - max(work_start, b_start_dt)).total_seconds() / 3600
        total_hours += min(daily_hours, 9 if day_of_week in [5, 6] else MAX_DAILY_HOURS)
        current_time = datetime.combine(current_time.date() + timedelta(days=1), WORK_START)
    return total_hours


# PDA_partner.py의 분류 기준 리스트

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
    tms_tasks = ["BURNER ASSY(TMS)", "WET TANK ASSY(TMS)", "COOLING UNIT(TMS)", "REACTOR ASSY(TMS)"]
    # DRAGON, DRAGON DUAL, SWS-I 모델에서는 tms_tasks도 기구로 분류
    if model_name in ["DRAGON", "DRAGON DUAL", "SWS-I"] and content in get_mechanical_tasks(model_name):
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
    # 완료: 시작/완료 시간이 모두 있는 경우
    df["진행율"] = df.apply(
        lambda row: 100.0 if pd.notna(row["시작 시간"]) and pd.notna(row["완료 시간"]) else 0.0, axis=1
    )
    df_valid = df.dropna(subset=["내용"])
    df_max = df_valid.groupby(["내용", "작업 분류"])["진행율"].max().reset_index()
    progress_summary = {}
    for category in ["기구", "전장", "TMS_반제품", "검사", "마무리", "기타"]:
        df_cat = df_max[df_max["작업 분류"] == category]
        total_tasks = len(df_cat)
        completed = df_cat[df_cat["진행율"] == 100.0]
        progress = (len(completed) / total_tasks * 100) if total_tasks > 0 else 0
        progress_summary[category] = round(progress, 1)
    return progress_summary


class TaskAnalyzer:
    """작업 데이터 분석 클래스"""

    def __init__(self, model_name):
        """
        Args:
            model_name (str): 모델명
        """
        self.model_name = model_name

    def analyze_tasks(self, df):
        """작업 데이터 분석

        Args:
            df (pd.DataFrame): WORKSHEET 시트 데이터

        Returns:
            pd.DataFrame: 분석된 작업 데이터
            pd.DataFrame: 작업명별 누적 소요시간 데이터
            dict: 작업 분류별(기구, 전장, 검사, 마무리 등) 누적 소요시간
        """
        try:
            columns = ["내용", "시작 시간", "완료 시간"]
            df = df[columns].copy()

            # 시간 데이터 처리
            df["시작 시간"] = df["시작 시간"].apply(parse_korean_datetime)
            df["완료 시간"] = df["완료 시간"].apply(parse_korean_datetime)

            # 워킹데이 소요시간 계산
            df["워킹데이 소요 시간"] = df.apply(
                lambda row: calculate_working_hours_with_holidays(row["시작 시간"], row["완료 시간"]), axis=1
            )

            # 작업 상태 추가
            df["작업 상태"] = df.apply(self._get_task_status, axis=1)

            # 작업명별 누적 소요시간 계산
            task_total_time = df.groupby("내용", as_index=False)["워킹데이 소요 시간"].sum()
            for idx, row in task_total_time.iterrows():
                logger.info(
                    f"[누적] 작업명: {row['내용']}, 누적 워킹데이 소요시간: {row['워킹데이 소요 시간']:.2f}시간"
                )

            # 작업 분류(카테고리)별 누적 소요시간 계산
            # 분류 기준: classify_task 함수 필요
            task_total_time["작업 분류"] = task_total_time["내용"].apply(lambda x: classify_task(x, self.model_name))
            category_sums = {}
            for cat in ["기구", "전장", "검사", "마무리", "TMS_반제품", "기타"]:
                category_sums[cat] = task_total_time[task_total_time["작업 분류"] == cat]["워킹데이 소요 시간"].sum()
                logger.info(f"[그룹] {cat} 전체 누적 워킹데이 소요시간: {category_sums[cat]:.2f}시간")
            # 진행률 계산
            progress_summary = calculate_progress_by_category(df, self.model_name)
            for cat, prog in progress_summary.items():
                logger.info(f"[진행률] {cat}: {prog}%")
            return df, task_total_time, category_sums, progress_summary

        except Exception as e:
            logger.error(f"❌ 작업 데이터 분석 중 오류 발생: {str(e)}")
            return None, None, None, None

    def _get_task_status(self, row):
        """작업 상태 판단

        Args:
            row (pd.Series): 작업 데이터 행

        Returns:
            str: 작업 상태 (미시작/진행중/완료)
        """
        if pd.isna(row["시작 시간"]):
            return "미시작"
        elif pd.isna(row["완료 시간"]):
            return "진행중"
        else:
            return "완료"
