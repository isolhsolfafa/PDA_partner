"""
유틸리티 함수
"""

import os
import json
import pytz
from datetime import datetime, timedelta
from .settings import logger

def format_time(dt, tz='Asia/Seoul'):
    """시간 포맷팅"""
    if not dt:
        return None
        
    try:
        seoul_tz = pytz.timezone(tz)
        if dt.tzinfo is None:
            dt = pytz.utc.localize(dt)
        dt = dt.astimezone(seoul_tz)
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        logger.error(f"❌ 시간 포맷팅 실패: {e}")
        return None

def calculate_duration(start_time, end_time):
    """시간 간격 계산"""
    if not start_time or not end_time:
        return None
        
    try:
        duration = end_time - start_time
        hours = duration.total_seconds() / 3600
        return round(hours, 2)
    except Exception as e:
        logger.error(f"❌ 시간 간격 계산 실패: {e}")
        return None

def format_duration(hours):
    """시간 간격 포맷팅"""
    if not hours:
        return None
        
    try:
        days = int(hours // 24)
        remaining_hours = int(hours % 24)
        minutes = int((hours * 60) % 60)
        
        parts = []
        if days > 0:
            parts.append(f"{days}일")
        if remaining_hours > 0:
            parts.append(f"{remaining_hours}시간")
        if minutes > 0:
            parts.append(f"{minutes}분")
            
        return " ".join(parts) if parts else "0분"
    except Exception as e:
        logger.error(f"❌ 시간 간격 포맷팅 실패: {e}")
        return None

def save_json(data, file_path):
    """JSON 파일 저장"""
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logger.info(f"✅ JSON 파일 저장 완료: {file_path}")
        return True
    except Exception as e:
        logger.error(f"❌ JSON 파일 저장 실패: {e}")
        return False

def load_json(file_path):
    """JSON 파일 로드"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        logger.info(f"✅ JSON 파일 로드 완료: {file_path}")
        return data
    except Exception as e:
        logger.error(f"❌ JSON 파일 로드 실패: {e}")
        return None

def create_report_content(model_name, progress_data, occurrence_stats, partner_stats=None):
    """리포트 내용 생성"""
    try:
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        content = f"""
        <h2>📊 {model_name} 작업 현황 리포트</h2>
        <p>생성 시간: {current_time}</p>
        
        <h3>🔄 진행률</h3>
        <ul>
        """
        
        for category, progress in progress_data.items():
            content += f"<li>{category}: {progress:.1f}%</li>"
            
        content += """
        </ul>
        
        <h3>📈 NaN & Overtime 발생률</h3>
        <ul>
        """
        
        for category, stats in occurrence_stats.items():
            content += f"""
            <li>{category}:
                <ul>
                    <li>총 작업 수: {stats['total_count']}</li>
                    <li>NaN 발생: {stats['nan_count']}건 ({stats['nan_ratio']:.1f}%)</li>
                    <li>Overtime 발생: {stats['ot_count']}건 ({stats['ot_ratio']:.1f}%)</li>
                </ul>
            </li>
            """
            
        if partner_stats:
            content += """
            <h3>🤝 파트너사 현황</h3>
            <ul>
            """
            
            for partner, stats in partner_stats.items():
                content += f"""
                <li>{partner}:
                    <ul>
                        <li>NaN 발생률: {stats['nan_ratio']:.1f}%</li>
                        <li>Overtime 발생률: {stats['ot_ratio']:.1f}%</li>
                    </ul>
                </li>
                """
                
            content += "</ul>"
            
        return content
        
    except Exception as e:
        logger.error(f"❌ 리포트 내용 생성 실패: {e}")
        return None

def get_file_paths(base_dir, model_name, current_time=None):
    """파일 경로 생성"""
    try:
        if current_time is None:
            current_time = datetime.now()
            
        date_str = current_time.strftime('%Y%m%d_%H%M%S')
        
        paths = {
            'heatmap': os.path.join(base_dir, 'output', f'{model_name}_heatmap_{date_str}.png'),
            'progress': os.path.join(base_dir, 'output', f'{model_name}_progress_{date_str}.png'),
            'timeline': os.path.join(base_dir, 'output', f'{model_name}_timeline_{date_str}.png'),
            'report': os.path.join(base_dir, 'output', f'{model_name}_report_{date_str}.html'),
            'data': os.path.join(base_dir, 'output', f'{model_name}_data_{date_str}.json')
        }
        
        # 디렉토리 생성
        for path in paths.values():
            os.makedirs(os.path.dirname(path), exist_ok=True)
            
        return paths
        
    except Exception as e:
        logger.error(f"❌ 파일 경로 생성 실패: {e}")
        return None 