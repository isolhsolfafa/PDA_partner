"""
작업 시간 계산 및 분석 로직
"""

import pandas as pd
from datetime import datetime, time, timedelta
from .settings import (
    logger, WORK_SCHEDULE, DEFAULT_TASKS,
    MODEL_TIME_FACTORS, BASE_WORK_HOURS, HOLIDAYS
)

class WorkTimeCalculator:
    """작업 시간 계산 클래스"""
    
    def __init__(self, model_name):
        self.model_name = model_name.upper()
        self.time_factor = MODEL_TIME_FACTORS.get(self.model_name, 1.0)
        
    def get_work_hours(self, category):
        """카테고리별 기준 작업 시간 계산"""
        base_hours = BASE_WORK_HOURS.get(category, 0)
        return base_hours * self.time_factor
        
    def is_holiday(self, date):
        """공휴일 여부 확인"""
        return date in HOLIDAYS
        
    def calculate_working_hours(self, start_time, end_time):
        """작업 시간 계산 (휴일 및 휴식 시간 고려)"""
        if pd.isna(start_time) or pd.isna(end_time):
            return 0
            
        total_hours = 0
        current_time = start_time
        
        while current_time < end_time:
            current_date = current_time.date()
            day_of_week = current_time.weekday()
            
            # 주말이나 공휴일이면 다음 날로 스킵
            if day_of_week >= 5 or self.is_holiday(current_date):
                current_time = datetime.combine(
                    current_date + timedelta(days=1),
                    time(*WORK_SCHEDULE['WORK_START'])
                )
                continue
            
            # 작업 시작/종료 시간 설정
            work_start = datetime.combine(
                current_date,
                time(*WORK_SCHEDULE['WORK_START'])
            )
            work_end = datetime.combine(
                current_date,
                time(*WORK_SCHEDULE['WORK_END'])
            )
            
            work_start = max(work_start, current_time)
            work_end = min(work_end, end_time)
            
            # 일일 작업 시간 계산
            daily_hours = (work_end - work_start).total_seconds() / 3600
            
            # 휴식 시간 제외
            breaks = [
                (WORK_SCHEDULE['LUNCH_START'], WORK_SCHEDULE['LUNCH_END']),
                (WORK_SCHEDULE['BREAK_1_START'], WORK_SCHEDULE['BREAK_1_END']),
                (WORK_SCHEDULE['BREAK_2_START'], WORK_SCHEDULE['BREAK_2_END']),
                (WORK_SCHEDULE['DINNER_START'], WORK_SCHEDULE['DINNER_END'])
            ]
            
            for b_start, b_end in breaks:
                b_start_dt = datetime.combine(current_date, time(*b_start))
                b_end_dt = datetime.combine(current_date, time(*b_end))
                if work_start < b_end_dt and b_start_dt < work_end:
                    daily_hours -= (
                        min(work_end, b_end_dt) - max(work_start, b_start_dt)
                    ).total_seconds() / 3600
            
            # 최대 작업 시간 제한
            total_hours += min(daily_hours, WORK_SCHEDULE['MAX_DAILY_HOURS'])
            
            # 다음 날로 이동
            current_time = datetime.combine(
                current_date + timedelta(days=1),
                time(*WORK_SCHEDULE['WORK_START'])
            )
        
        return round(total_hours, 2)

class TaskAnalyzer:
    """작업 분석 클래스"""
    
    def __init__(self, model_name):
        self.model_name = model_name.upper()
        self.calculator = WorkTimeCalculator(model_name)
        
    def classify_task(self, content):
        """작업 분류"""
        content = str(content).strip()
        
        # TMS 반제품 작업 확인
        tms_tasks = [
            "BURNER ASSY(TMS)", "WET TANK ASSY(TMS)",
            "COOLING UNIT(TMS)", "REACTOR ASSY(TMS)"
        ]
        if content in tms_tasks and self.model_name not in ["DRAGON", "DRAGON DUAL", "SWS-I"]:
            return "TMS"
            
        # 기구 작업 확인
        if content in DEFAULT_TASKS['MECHANICAL']:
            return "기구"
            
        # 전장 작업 확인
        if content in DEFAULT_TASKS['ELECTRICAL']:
            return "전장"
            
        # 검사 작업 확인
        if content in DEFAULT_TASKS['INSPECTION']:
            return "검사"
            
        # 마무리 작업 확인
        if content in DEFAULT_TASKS['FINISHING']:
            return "마무리"
            
        return "기타"
        
    def analyze_data(self, df):
        """데이터 분석"""
        df = df.copy()
        
        # 작업 시간 계산
        df['워킹데이 소요 시간'] = df.apply(
            lambda row: self.calculator.calculate_working_hours(
                row['시작 시간'], row['완료 시간']
            ),
            axis=1
        )
        
        # 작업 분류
        df['작업 분류'] = df['내용'].apply(self.classify_task)
        
        # 작업별 총 시간 계산
        task_total_time = df.groupby(['내용', '작업 분류'])['워킹데이 소요 시간'].sum().reset_index()
        
        return df, task_total_time
        
    def calculate_progress(self, df):
        """진행률 계산"""
        df = df.copy()
        
        # 완료된 작업 처리
        df['진행율'] = df.apply(
            lambda row: 100.0 if pd.isna(row['진행율']) and 
                       pd.notna(row['시작 시간']) and 
                       pd.notna(row['완료 시간'])
            else row['진행율'],
            axis=1
        )
        
        # 카테고리별 진행률 계산
        progress = {}
        for category in ['기구', '전장', 'TMS']:
            df_cat = df[df['작업 분류'] == category]
            if len(df_cat) > 0:
                completed = len(df_cat[df_cat['진행율'] == 100.0])
                progress[category] = round(completed / len(df_cat) * 100, 1)
            else:
                progress[category] = 0.0
                
        return progress

def analyze_occurrence_rates(df, task_total_time, model_name, tolerance=2,
                           mech_partner=None, elec_partner=None):
    """NaN & Overtime 발생률 분석"""
    calculator = WorkTimeCalculator(model_name)
    occurrence_stats = {}
    partner_stats = {}
    
    for category in ['기구', '전장', '검사', '마무리', 'TMS']:
        category_tasks = task_total_time[task_total_time['작업 분류'] == category]
        total_count = len(category_tasks)
        
        if total_count == 0:
            continue
            
        avg_time = calculator.get_work_hours(category)
        tolerance_time = avg_time * tolerance
        
        nan_count = category_tasks['워킹데이 소요 시간'].isna().sum()
        ot_count = len(category_tasks[
            category_tasks['워킹데이 소요 시간'] > tolerance_time
        ])
        
        occurrence_stats[category] = {
            "total_count": total_count,
            "nan_count": nan_count,
            "ot_count": ot_count,
            "nan_ratio": (nan_count / total_count * 100) if total_count > 0 else 0,
            "ot_ratio": (ot_count / total_count * 100) if total_count > 0 else 0
        }
        
        if category == '기구' and mech_partner:
            partner_stats[mech_partner] = occurrence_stats[category]
        elif category == '전장' and elec_partner:
            partner_stats[elec_partner] = occurrence_stats[category]
            
    return occurrence_stats, partner_stats 