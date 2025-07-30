"""
Google Sheets 데이터 처리
"""

import re
import pandas as pd
from datetime import datetime
from .settings import logger, WORKSHEET_RANGE
from backoff import on_exception, expo
from googleapiclient.errors import HttpError

@on_exception(expo, HttpError, max_tries=8, max_time=120, giveup=lambda e: e.response.status_code not in [429, 503])
def api_call_with_backoff(func, *args, **kwargs):
    """API 호출에 대한 지수 백오프 재시도 데코레이터"""
    try:
        return func(*args, **kwargs)
    except HttpError as e:
        logger.error(f"⚠️ [Retrying] API 호출 실패: {e}")
        raise

class SheetsProcessor:
    """스프레드시트 처리 클래스"""
    
    def __init__(self, sheets_service):
        self.service = sheets_service
        
    def extract_spreadsheet_id(self, hyperlink):
        """하이퍼링크에서 스프레드시트 ID 추출"""
        try:
            # =HYPERLINK("https://docs.google.com/spreadsheets/d/XXXXX...", "표시될 텍스트")
            match = re.search(r'spreadsheets/d/([a-zA-Z0-9-_]+)', hyperlink)
            if match:
                return match.group(1)
            return None
        except Exception as e:
            logger.error(f"❌ 스프레드시트 ID 추출 실패: {e}")
            return None

    def get_hyperlinks(self, spreadsheet_id, range_name):
        """하이퍼링크 목록 가져오기"""
        try:
            result = api_call_with_backoff(
                self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueRenderOption="FORMULA"
                ).execute
            )
            
            values = result.get('values', [])
            if not values:
                logger.error("❌ 하이퍼링크를 찾을 수 없습니다.")
                return []
                
            hyperlinks = []
            for row in values:
                if not row:  # 빈 행 건너뛰기
                    continue
                    
                cell = row[0]
                if not cell.startswith('=HYPERLINK('):
                    continue
                    
                # 하이퍼링크 URL 추출
                url_start = cell.find('\"') + 1
                url_end = cell.find('\"', url_start)
                if url_start == -1 or url_end == -1:
                    continue
                    
                url = cell[url_start:url_end]
                
                # 스프레드시트 ID 추출
                if '/spreadsheets/d/' in url:
                    spreadsheet_id = url.split('/spreadsheets/d/')[1].split('/')[0]
                    hyperlinks.append({
                        'url': url,
                        'spreadsheet_id': spreadsheet_id
                    })
                    
            if hyperlinks:
                logger.info(f"✅ 하이퍼링크 {len(hyperlinks)}개를 찾았습니다.")
            return hyperlinks
            
        except Exception as e:
            logger.error(f"❌ 하이퍼링크 가져오기 실패: {str(e)}")
            return []
            
    def get_spreadsheet_title(self, spreadsheet_id):
        """스프레드시트 제목 가져오기"""
        try:
            info = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields='properties.title'
            ).execute()
            return info['properties']['title']
        except Exception as e:
            logger.error(f"❌ 스프레드시트 제목 가져오기 실패: {spreadsheet_id} -> {e}")
            return f"Unknown_{spreadsheet_id}"
        
    def get_sheet_names(self, spreadsheet_id):
        """스프레드시트의 시트 목록 가져오기"""
        try:
            result = api_call_with_backoff(
                self.service.spreadsheets().get(
                    spreadsheetId=spreadsheet_id,
                    fields="sheets.properties.title"
                ).execute
            )
            
            sheets = result.get('sheets', [])
            sheet_names = [sheet['properties']['title'] for sheet in sheets]
            logger.info(f"📋 시트 목록: {sheet_names}")
            return sheet_names
            
        except Exception as e:
            logger.error(f"❌ 시트 목록 가져오기 실패: {e}")
            return None
            
    def get_info_data(self, spreadsheet_id):
        """스프레드시트의 '정보판' 시트에서 모델명과 협력사 정보 가져오기"""
        try:
            # 시트 목록 확인
            sheet_names = self.get_sheet_names(spreadsheet_id)
            if not sheet_names:
                return None, None, None
                
            if '정보판' not in sheet_names:
                logger.error(f"❌ '정보판' 시트를 찾을 수 없습니다: {spreadsheet_id}")
                return None, None, None
                
            # 한 번의 API 호출로 여러 범위의 데이터를 가져옴
            ranges = [("'정보판'!D4", "model_name"), ("'정보판'!B5", "mech_partner"), ("'정보판'!D5", "elec_partner")]
            batch_request = self.service.spreadsheets().values().batchGet(
                spreadsheetId=spreadsheet_id,
                ranges=[rng for rng, _ in ranges],
                valueRenderOption="FORMATTED_VALUE"
            )
            
            result = api_call_with_backoff(batch_request.execute)
            results = {}
            for (rng, key), response in zip(ranges, result.get("valueRanges", [])):
                values = response.get("values", [[]])
                results[key] = values[0][0].strip() if values else "미정"
                
            # 디버깅 정보 출력
            logger.info(f"📌 [디버깅] 모델명: {results['model_name']}, 기구협력사: {results['mech_partner']}, 전장협력사: {results['elec_partner']}")
            
            return results['model_name'] or "NoValue", results['mech_partner'], results['elec_partner']
            
        except Exception as e:
            logger.error(f"❌ 정보판 데이터 가져오기 실패: {e}")
            return None, None, None
            
    def get_model_name(self, spreadsheet_id):
        """스프레드시트의 '정보판' 시트에서 모델명 가져오기"""
        model_name, _, _ = self.get_info_data(spreadsheet_id)
        return model_name
            
    def read_sheet(self, spreadsheet_id, range_name=None):
        """스프레드시트에서 데이터 읽기"""
        try:
            # 시트 목록 확인
            sheet_names = self.get_sheet_names(spreadsheet_id)
            if not sheet_names:
                return None
                
            if 'WORKSHEET' not in sheet_names:
                logger.error(f"❌ 'WORKSHEET' 시트를 찾을 수 없습니다: {spreadsheet_id}")
                return None
                
            # 데이터 가져오기
            result = api_call_with_backoff(
                self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range="'WORKSHEET'!A1:AZ200",
                    valueRenderOption="FORMATTED_VALUE"
                ).execute
            )
            
            values = result.get('values', [])
            if not values:
                logger.error("❌ 데이터를 찾을 수 없습니다.")
                return None
                
            # 헤더를 7행으로 설정
            header_row = 6  # 0-based index for 7th row
            if len(values) <= header_row:
                logger.error("❌ 7행에 데이터가 없습니다.")
                logger.info(f"[DEBUG] WORKSHEET 데이터: {values[:5]}")
                return None
                
            # 헤더와 데이터 분리
            headers = values[header_row]
            logger.info(f"[DEBUG] 헤더: {headers}")
            data = []
            
            # 데이터 행 처리
            for row in values[header_row + 1:]:
                if not row:
                    continue
                # 부족한 셀은 빈 문자열로 채워서 길이 맞추기
                row_extended = row + [""] * (len(headers) - len(row))
                data.append(row_extended[:len(headers)])
                
            # DataFrame 생성
            df = pd.DataFrame(data, columns=headers) if data else pd.DataFrame(columns=headers)
            logger.info(f"✅ 데이터 읽기 완료: {len(df)}행")
            return df
            
        except Exception as e:
            logger.error(f"❌ 데이터 읽기 실패: {str(e)}")
            return None
            
    def update_sheet(self, spreadsheet_id, range_name, values):
        """스프레드시트 데이터 업데이트"""
        try:
            body = {
                'values': values
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            logger.info(f"✅ 스프레드시트 업데이트 완료: {result.get('updatedCells')} 셀")
            return True
            
        except Exception as e:
            logger.error(f"❌ 스프레드시트 업데이트 실패: {e}")
            return False
            
    def append_sheet(self, spreadsheet_id, range_name, values):
        """스프레드시트에 데이터 추가"""
        try:
            body = {
                'values': values
            }
            
            result = self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
            logger.info(f"✅ 스프레드시트 데이터 추가 완료: {result.get('updates').get('updatedRows')} 행")
            return True
            
        except Exception as e:
            logger.error(f"❌ 스프레드시트 데이터 추가 실패: {e}")
            return False
            
    def create_summary_sheet(self, spreadsheet_id, summary_data, sheet_name='Summary'):
        """요약 시트 생성"""
        try:
            # 새 시트 생성
            body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                }]
            }
            
            try:
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=body
                ).execute()
            except Exception:
                logger.info("이미 Summary 시트가 존재합니다.")
                
            # 데이터 준비
            headers = ['카테고리', '완료 작업 수', '총 작업 수', '진행률 (%)',
                      'NaN 발생률 (%)', 'Overtime 발생률 (%)']
            rows = []
            
            for category, stats in summary_data.items():
                rows.append([
                    category,
                    stats.get('completed_tasks', 0),
                    stats.get('total_tasks', 0),
                    stats.get('progress', 0),
                    stats.get('nan_ratio', 0),
                    stats.get('ot_ratio', 0)
                ])
                
            # 현재 시간 추가
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            rows.append(['생성 시간:', current_time])
            
            # 데이터 업데이트
            values = [headers] + rows
            range_name = f'{sheet_name}!A1'
            
            self.update_sheet(spreadsheet_id, range_name, values)
            
            logger.info(f"✅ 요약 시트 생성 완료: {sheet_name}")
            return True
            
        except Exception as e:
            logger.error(f"❌ 요약 시트 생성 실패: {e}")
            return False
            
    def read_info_sheet(self, spreadsheet_id):
        """정보판 시트에서 모델명과 협력사 정보 가져오기 (한 행에 여러 정보가 있을 때도 파싱)"""
        try:
            # 시트 목록 확인
            sheet_names = self.get_sheet_names(spreadsheet_id)
            if not sheet_names:
                return None
                
            if '정보판' not in sheet_names:
                logger.error(f"❌ '정보판' 시트를 찾을 수 없습니다: {spreadsheet_id}")
                return None
                
            # 데이터 가져오기
            result = api_call_with_backoff(
                self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range="'정보판'!A1:Z100",
                    valueRenderOption="FORMATTED_VALUE"
                ).execute
            )
            
            values = result.get('values', [])
            print(f"[DEBUG] 정보판 원본 데이터: {values}")
            logger.info(f"[DEBUG] 정보판 원본 데이터: {values}")
            if not values:
                logger.error(f"❌ '정보판' 시트에서 데이터를 찾을 수 없습니다: {spreadsheet_id}")
                return None
                
            # 데이터 파싱 (한 행에 여러 정보가 있을 때도 처리)
            info_data = {}
            for row in values:
                for i, cell in enumerate(row):
                    key = str(cell).strip().upper()
                    # 모델명
                    if key == 'MODEL' and i+1 < len(row):
                        info_data['모델명'] = row[i+1].strip()
                    # 기구협력사
                    elif key in ['기구외주', '기구협력사'] and i+1 < len(row):
                        info_data['기구협력사'] = row[i+1].strip()
                    # 전장협력사
                    elif key in ['전장외주', '전장협력사'] and i+1 < len(row):
                        info_data['전장협력사'] = row[i+1].strip()
            # 필수 정보 확인
            if '모델명' not in info_data:
                logger.error(f"❌ '정보판' 시트에서 모델명을 찾을 수 없습니다: {spreadsheet_id}")
                return None
            if '기구협력사' not in info_data:
                info_data['기구협력사'] = ""
            if '전장협력사' not in info_data:
                info_data['전장협력사'] = ""
            logger.info(f"✅ 정보판 데이터 읽기 완료: {info_data}")
            return info_data
        except Exception as e:
            logger.error(f"❌ '정보판' 시트 읽기 실패: {str(e)}")
            return None 