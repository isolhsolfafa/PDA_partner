"""
PDA (Production Data Analysis) 시스템 메인 실행 파일
"""

import argparse
import json
import os
from datetime import datetime

import numpy as np
import pandas as pd
from src.calculate import TaskAnalyzer, WorkTimeCalculator, analyze_occurrence_rates
from src.notification import EmailNotifier, KakaoNotifier
from src.settings import (
    GITHUB_REPO,
    GITHUB_USERNAME,
    INFO_RANGE,
    PROCESS_LIMIT,
    WORKSHEET_RANGE,
    get_google_services,
    logger,
)
from src.sheets import SheetsProcessor
from src.upload import DriveUploader, GitHubUploader
from src.utils import get_file_paths, save_json
from src.visualization import HeatmapGenerator, ReportGenerator

from task_analyzer import TaskAnalyzer


def main(args):
    """메인 실행 함수"""
    try:
        # 테스트 모드 확인
        if args.test_mode:
            logger.info("🔧 테스트 모드로 실행됩니다.")
            logger.info("⚠️ GitHub, Drive 업로드 및 알림 기능이 비활성화됩니다.")

        # Google 서비스 초기화
        sheets_service, drive_service = get_google_services()
        sheets_processor = SheetsProcessor(sheets_service)

        # 하이퍼링크 목록 가져오기
        target_sheet_name = os.getenv("TARGET_SHEET_NAME", "출하예정리스트(TEST)")
        hyperlinks = sheets_processor.get_hyperlinks(args.spreadsheet_id, f"'{target_sheet_name}'!A3:A")

        if not hyperlinks:
            logger.error("❌ 하이퍼링크를 찾을 수 없습니다.")
            return False

        logger.info(f"📋 처리할 스프레드시트 수: {min(len(hyperlinks), PROCESS_LIMIT)}")

        # 데이터 처리 결과 저장
        processed_data = []

        # 각 하이퍼링크 처리
        for i, hyperlink in enumerate(hyperlinks[:PROCESS_LIMIT]):
            # 스프레드시트 ID 추출
            spreadsheet_id = hyperlink["spreadsheet_id"]

            # 정보판 시트에서 모델명과 협력사 정보 가져오기
            info_data = sheets_processor.read_info_sheet(spreadsheet_id)
            if info_data is None:
                continue

            # WORKSHEET 시트에서 작업 데이터 가져오기
            worksheet_df = sheets_processor.read_sheet(spreadsheet_id)
            if worksheet_df is None:
                logger.error(f"❌ WORKSHEET 데이터를 읽을 수 없습니다: {spreadsheet_id}")
                continue

            # 데이터 분석
            analyzer = TaskAnalyzer(info_data.get("모델명", "Unknown"))
            analyzed_df, task_total_time, category_sums, progress_summary = analyzer.analyze_tasks(worksheet_df)
            if analyzed_df is None:
                logger.error(f"❌ 데이터 분석에 실패했습니다: {spreadsheet_id}")
                continue

            processed_data.append(
                {
                    "spreadsheet_id": spreadsheet_id,
                    "model_name": info_data.get("모델명", "Unknown"),
                    "mechanical_partner": info_data.get("기구협력사", "Unknown"),
                    "electrical_partner": info_data.get("전장협력사", "Unknown"),
                    "tasks": analyzed_df,
                    "task_total_time": task_total_time,
                    "category_sums": category_sums,
                    "progress_summary": progress_summary,
                }
            )

        if not processed_data:
            logger.error("❌ 처리된 데이터가 없습니다.")
            return False

        # 시각화
        chart_generator = HeatmapGenerator()

        # 히트맵 생성
        heatmap_generator = HeatmapGenerator()
        for data in processed_data:
            heatmap_generator.create_heatmap(data["tasks"], data["model_name"])

        # 진행률 차트 생성
        chart_generator.create_progress_chart(processed_data)

        # 타임라인 차트 생성
        for spreadsheet_id, data in processed_data:
            chart_generator.create_timeline_chart(data["tasks"])

        # HTML 리포트 생성
        report_generator = ReportGenerator()
        report_generator.generate_html_report(processed_data, "PDA_Report")

        # JSON 파일로 저장
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_path = os.path.join("output", f"PDA_Report_data_{timestamp}.json")

        # int64를 int로 변환
        for spreadsheet_id, data in processed_data:
            tasks_dict = data["tasks"].to_dict("records")
            for task in tasks_dict:
                for key, value in task.items():
                    if isinstance(value, np.int64):
                        task[key] = int(value)
                    elif isinstance(value, pd.Timestamp):
                        task[key] = value.strftime("%Y-%m-%d %H:%M:%S")
            data["tasks"] = tasks_dict

        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(processed_data, f, ensure_ascii=False, indent=2)
            logger.info(f"✅ JSON 파일 저장 완료: {os.path.basename(json_path)}")
        except Exception as e:
            logger.error(f"❌ JSON 파일 저장 실패: {str(e)}")

        # 테스트 모드 확인
        if args.test_mode:
            logger.info("🔧 테스트 모드: 파일 생성만 완료되었습니다.")
            logger.info("📁 생성된 파일 목록:")
            logger.info(f"  - data: {json_path}")
            return True

        # GitHub에 업로드
        github_uploader = GitHubUploader()
        github_uploader.upload_files()

        # Google Drive에 업로드
        drive_uploader = DriveUploader(drive_service)
        drive_uploader.upload_files()

        # 알림 전송
        notifier = EmailNotifier()
        notifier.send_email(
            args.email_to.split(","),
            "PDA_Report 작업 현황 리포트",
            open("PDA_Report.html", "r", encoding="utf-8").read(),
            [f"PDA_Report_heatmap.png", f"PDA_Report_progress.png", f"PDA_Report_timeline.png"],
        )

        logger.info("✅ 모든 작업이 완료되었습니다.")
        return True

    except Exception as e:
        logger.error(f"❌ 작업 중 오류가 발생했습니다: {str(e)}")
        return False


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDA (Production Data Analysis) 시스템")

    parser.add_argument("--spreadsheet-id", required=True, help="Google Spreadsheet ID")
    parser.add_argument("--range-name", required=False, help="시트 범위 (기본값: 환경 변수 WORKSHEET_RANGE 사용)")
    parser.add_argument(
        "--model-name", required=False, help="모델 이름 (기본값: 스프레드시트의 '정보판'에서 자동으로 가져옴)"
    )
    parser.add_argument(
        "--mech-partner", required=False, help="기구 협력사 (기본값: 스프레드시트의 '정보판'에서 자동으로 가져옴)"
    )
    parser.add_argument(
        "--elec-partner", required=False, help="전장 협력사 (기본값: 스프레드시트의 '정보판'에서 자동으로 가져옴)"
    )
    parser.add_argument("--tolerance", type=float, default=2.0, help="Overtime 판단 기준 (기준 시간 대비 배수)")
    parser.add_argument("--upload-github", action="store_true", help="GitHub 업로드 여부")
    parser.add_argument("--upload-drive", action="store_true", help="Google Drive 업로드 여부")
    parser.add_argument("--email", action="store_true", help="이메일 알림 여부")
    parser.add_argument("--email-to", help="이메일 수신자 (쉼표로 구분)")
    parser.add_argument("--kakao", action="store_true", help="카카오톡 알림 여부")
    parser.add_argument("--test-mode", action="store_true", help="테스트 모드 실행 (업로드 및 알림 기능 비활성화)")

    args = parser.parse_args()
    success = main(args)
    exit(0 if success else 1)
