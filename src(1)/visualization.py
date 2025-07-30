"""
데이터 시각화 기능
"""

import os
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from .settings import logger
from matplotlib.patches import Patch
import matplotlib.dates as mdates
import matplotlib.font_manager as fm

class ReportGenerator:
    """HTML 리포트 생성 클래스"""
    
    def __init__(self, save_path=None):
        self.save_path = save_path or 'partner.html'
        
    def generate_html_report(self, data, stats, model_name):
        """HTML 리포트 생성"""
        try:
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            week_number = datetime.now().isocalendar()[1]
            year = datetime.now().year
            
            # HTML 템플릿
            html_content = f"""
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
                .progress-bar {{
                    width: 100%;
                    background-color: #e0e0e0;
                    height: 12px;
                    border-radius: 3px;
                }}
                .progress-fill {{
                    background-color: orange;
                    height: 100%;
                    border-radius: 3px;
                }}
                .warning {{ color: red; font-weight: bold; }}
              </style>
            </head>
            <body>
            <div style="text-align: center; margin-bottom: 20px;">
                <img src="https://rainbow-haupia-cd8290.netlify.app/GST_banner.jpg" alt="Build up GST Banner" style="max-width: 100%; height: auto;">
            </div>
            <h1>PDA Dashboard - {year}년 {week_number}주차</h1>
            <h3>📌 [알림] PDA Overtime 및 NaN 체크 결과</h3>
            <p>📅 실행 시간: {current_time} (KST)</p>
            <p>📊 대시보드에서 상세 내용 확인하세요! (<a href="https://gstfactory.netlify.app">대시보드 바로가기</a>)</p>
            
            <h4>요약 테이블</h4>
            <table id="summaryTable">
            <tr>
                <th>Order</th>
                <th>모델명</th>
                <th>기구협력사</th>
                <th>전장협력사</th>
                <th>총 작업 수</th>
                <th>기구 NaN</th>
                <th>기구 OT</th>
                <th>기구 진행률</th>
                <th>전장 NaN</th>
                <th>전장 OT</th>
                <th>전장 진행률</th>
                <th>TMS NaN</th>
                <th>TMS OT</th>
                <th>TMS 진행률</th>
            </tr>
            """
            
            # 데이터 행 추가
            for order, order_data in data.items():
                html_content += f"""
                <tr>
                    <td><a href="{order_data['spreadsheet_url']}">{order}</a></td>
                    <td>{model_name}</td>
                    <td>{order_data.get('기구협력사', '-')}</td>
                    <td>{order_data.get('전장협력사', '-')}</td>
                    <td>{order_data.get('총작업수', 0)}</td>
                """
                
                # 각 분야별 NaN, OT, 진행률 추가
                for category in ['기구', '전장', 'TMS']:
                    nan_count = order_data.get(f'{category}_nan', 0)
                    ot_count = order_data.get(f'{category}_ot', 0)
                    progress = order_data.get(f'{category}_progress', 0)
                    
                    nan_class = 'warning' if nan_count > 0 else ''
                    ot_class = 'warning' if ot_count > 0 else ''
                    
                    html_content += f"""
                    <td class="{nan_class}">{nan_count}</td>
                    <td class="{ot_class}">{ot_count}</td>
                    <td>
                    """
                    
                    if progress == 100:
                        html_content += '<span title="작업 완료" style="font-size: 16px;">✅</span>'
                    else:
                        html_content += f"""
                        <div class="progress-bar">
                            <div class="progress-fill" style="width: {progress}%;" title="진행률: {progress}%"></div>
                        </div>
                        <span style="font-size: 12px;">{progress}%</span>
                        """
                    
                    html_content += "</td>"
                
                html_content += "</tr>"
            
            html_content += """
            </table>
            <script>
            // 필터링 기능 추가 예정
            </script>
            </body>
            </html>
            """
            
            # 파일 저장
            with open(self.save_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            logger.info(f"✅ HTML 리포트 생성 완료: {self.save_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ HTML 리포트 생성 실패: {e}")
            return False

class HeatmapGenerator:
    """히트맵 생성 클래스"""
    
    def __init__(self, save_path=None):
        self.save_path = save_path or 'heatmap.png'
        
    def create_heatmap(self, df, occurrence_stats, partner_stats=None):
        """NaN & Overtime 발생률 히트맵 생성"""
        try:
            # 데이터 준비
            categories = list(occurrence_stats.keys())
            metrics = ['nan_ratio', 'ot_ratio']
            data = np.zeros((len(categories), len(metrics)))
            
            for i, category in enumerate(categories):
                stats = occurrence_stats[category]
                data[i, 0] = stats['nan_ratio']
                data[i, 1] = stats['ot_ratio']
                
            # 히트맵 생성
            plt.figure(figsize=(10, 6))
            sns.heatmap(
                data,
                annot=True,
                fmt='.1f',
                cmap='YlOrRd',
                xticklabels=['NaN 발생률 (%)', 'Overtime 발생률 (%)'],
                yticklabels=categories,
                cbar_kws={'label': '발생률 (%)'}
            )
            
            plt.title('작업별 NaN & Overtime 발생률')
            plt.tight_layout()
            
            # 파트너사 정보 추가
            if partner_stats:
                plt.figtext(
                    0.02, 0.02,
                    f"파트너사 현황: {', '.join(f'{k}: NaN {v['nan_ratio']:.1f}%, OT {v['ot_ratio']:.1f}%' for k, v in partner_stats.items())}",
                    fontsize=8
                )
            
            # 저장
            plt.savefig(self.save_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"✅ 히트맵 생성 완료: {self.save_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 히트맵 생성 실패: {e}")
            return False
            
    def create_progress_chart(self, progress_data, save_path=None):
        """진행률 차트 생성"""
        try:
            save_path = save_path or 'progress.png'
            
            categories = list(progress_data.keys())
            values = list(progress_data.values())
            
            plt.figure(figsize=(10, 6))
            bars = plt.bar(categories, values)
            
            # 바 위에 값 표시
            for bar in bars:
                height = bar.get_height()
                plt.text(
                    bar.get_x() + bar.get_width()/2.,
                    height,
                    f'{height:.1f}%',
                    ha='center',
                    va='bottom'
                )
            
            plt.title('작업 진행률')
            plt.ylabel('진행률 (%)')
            plt.ylim(0, 100)
            
            # 현재 시간 표시
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            plt.figtext(
                0.02, 0.02,
                f'생성 시간: {current_time}',
                fontsize=8
            )
            
            plt.tight_layout()
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"✅ 진행률 차트 생성 완료: {save_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 진행률 차트 생성 실패: {e}")
            return False
            
    def create_timeline_chart(self, df):
        """타임라인 차트 생성"""
        try:
            # 날짜 형식 변환
            df['시작 시간'] = pd.to_datetime(df['시작 시간'])
            df['종료 시간'] = pd.to_datetime(df['종료 시간'])
            
            # 작업 시간 계산
            df['작업 시간'] = (df['종료 시간'] - df['시작 시간']).dt.total_seconds() / 3600  # 시간 단위
            
            # 작업 분류별 색상 설정
            colors = {
                '기구': '#FF9999',
                '전장': '#66B2FF',
                'TMS': '#99FF99',
                '검사': '#FFCC99',
                '마무리': '#FF99CC'
            }
            
            # 그래프 생성
            fig, ax = plt.subplots(figsize=(15, 8))
            
            # 각 작업별 막대 그리기
            for i, (idx, row) in enumerate(df.iterrows()):
                ax.barh(i, row['작업 시간'], left=row['시작 시간'],
                       color=colors.get(row['작업 분류'], '#CCCCCC'),
                       height=0.3)
                
            # y축 설정
            ax.set_yticks(range(len(df)))
            ax.set_yticklabels([f"{row['작업 내용']} ({row['작업 분류']})" for _, row in df.iterrows()])
            
            # x축 설정
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H:%M'))
            plt.xticks(rotation=45)
            
            # 제목 설정
            plt.title('작업 타임라인')
            
            # 레이아웃 조정
            plt.tight_layout()
            
            # 저장
            plt.savefig(self.save_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"✅ 타임라인 차트 생성 완료: {os.path.basename(self.save_path)}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 타임라인 차트 생성 실패: {str(e)}")
            return False

class ChartGenerator:
    """차트 생성 클래스"""
    
    def __init__(self, save_dir="output"):
        self.save_dir = save_dir
        os.makedirs(save_dir, exist_ok=True)
        
        # 한글 폰트 설정
        font_paths = [
            "/Users/kdkyu311/Library/Fonts/NanumGothic.ttf",
            "/System/Library/AssetsV2/com_apple_MobileAsset_Font7/bad9b4bf17cf1669dde54184ba4431c22dcad27b.asset/AssetData/NanumGothic.ttc",
            "/Library/Fonts/NanumGothic.ttf",
            "/System/Library/Fonts/Supplemental/NanumGothic.ttf",
            "/System/Library/Fonts/AppleSDGothicNeo.ttc"  # macOS 기본 한글 폰트 추가
        ]
        
        font_path = next((path for path in font_paths if os.path.exists(path)), None)
        if font_path:
            logger.info(f"✅ 한글 폰트 적용 완료: {font_path}")
            plt.rcParams['font.family'] = 'AppleGothic'  # macOS 기본 한글 폰트 사용
        else:
            logger.error("🚨 한글 폰트를 찾을 수 없습니다.")
            
        plt.rcParams['axes.unicode_minus'] = False
        
    def create_heatmap(self, df, model_name):
        """히트맵 생성"""
        try:
            # 데이터 전처리
            df['시작 시간'] = pd.to_datetime(df['시작 시간'])
            df['요일'] = df['시작 시간'].dt.dayofweek
            df['시간'] = df['시작 시간'].dt.hour
            
            # 피벗 테이블 생성
            pivot_table = pd.crosstab(df['요일'], df['시간'])
            
            # 요일 레이블 설정
            days = ['월', '화', '수', '목', '금', '토', '일']
            pivot_table.index = days
            
            # 플롯 생성
            plt.figure(figsize=(12, 6))
            plt.imshow(pivot_table, cmap='YlOrRd', aspect='auto')
            
            # 축 레이블 설정
            plt.xticks(range(24), range(24))
            plt.yticks(range(7), days)
            
            plt.colorbar(label='작업 수')
            plt.title(f'{model_name} 작업 시간대별 히트맵')
            plt.xlabel('시간')
            plt.ylabel('요일')
            
            # 저장
            save_path = os.path.join(self.save_dir, 'heatmap.png')
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"✅ 히트맵 생성 완료: heatmap.png")
            return save_path
            
        except Exception as e:
            logger.error(f"❌ 히트맵 생성 실패: {e}")
            return None
            
    def create_progress_chart(self, progress_data, model_name):
        """진행률 차트 생성"""
        try:
            categories = list(progress_data.keys())
            values = list(progress_data.values())
            
            plt.figure(figsize=(10, 6))
            bars = plt.bar(categories, values)
            
            # 바 위에 값 표시
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:.1f}%', ha='center', va='bottom')
                
            plt.title(f'{model_name} 작업 진행률')
            plt.ylabel('진행률 (%)')
            plt.ylim(0, 100)
            
            # 저장
            save_path = os.path.join(self.save_dir, 'progress.png')
            plt.tight_layout()
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"✅ 진행률 차트 생성 완료: progress.png")
            return save_path
            
        except Exception as e:
            logger.error(f"❌ 진행률 차트 생성 실패: {e}")
            return None
            
    def create_timeline_chart(self, df):
        """타임라인 차트 생성"""
        try:
            # 날짜 형식 변환
            df['시작 시간'] = pd.to_datetime(df['시작 시간'])
            df['종료 시간'] = pd.to_datetime(df['종료 시간'])
            
            # 작업 시간 계산
            df['작업 시간'] = (df['종료 시간'] - df['시작 시간']).dt.total_seconds() / 3600  # 시간 단위
            
            # 작업 분류별 색상 설정
            colors = {
                '기구': '#FF9999',
                '전장': '#66B2FF',
                'TMS': '#99FF99',
                '검사': '#FFCC99',
                '마무리': '#FF99CC'
            }
            
            # 그래프 생성
            fig, ax = plt.subplots(figsize=(15, 8))
            
            # 각 작업별 막대 그리기
            for i, (idx, row) in enumerate(df.iterrows()):
                ax.barh(i, row['작업 시간'], left=row['시작 시간'],
                       color=colors.get(row['작업 분류'], '#CCCCCC'),
                       height=0.3)
                
            # y축 설정
            ax.set_yticks(range(len(df)))
            ax.set_yticklabels([f"{row['작업 내용']} ({row['작업 분류']})" for _, row in df.iterrows()])
            
            # x축 설정
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H:%M'))
            plt.xticks(rotation=45)
            
            # 제목 설정
            plt.title('작업 타임라인')
            
            # 레이아웃 조정
            plt.tight_layout()
            
            # 저장
            plt.savefig(self.save_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"✅ 타임라인 차트 생성 완료: {os.path.basename(self.save_path)}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 타임라인 차트 생성 실패: {str(e)}")
            return False 