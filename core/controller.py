# -*- coding: utf-8 -*-
"""
메인 컨트롤러 - 엑셀 취합 프로세스 전체 제어
"""

import os
import pandas as pd
from typing import List, Dict, Any, Optional, Callable
from pathlib import Path
from datetime import datetime
import threading
import time

from .processor import ExcelProcessor, FileStructure
from .normalizer import DataNormalizer
import config

class MergeController:
    """엑셀 병합 컨트롤러"""
    
    def __init__(self):
        """컨트롤러 초기화"""
        self.processor = ExcelProcessor()
        self.normalizer = DataNormalizer()
        
        # 상태 관리
        self.is_processing = False
        self.is_cancelled = False
        self.current_progress = 0.0
        self.current_status = "준비 중"
        
        # UI 참조
        self.ui = None
        
        # 처리 결과
        self.merge_result = None
        self.change_logs = []
        self.validation_passed = False  # 검수 통과 여부
        
    def set_ui(self, ui):
        """UI 참조 설정"""
        self.ui = ui
        
    def cleanup(self):
        """정리 작업"""
        self.is_cancelled = True
        
    def start_merge(self, reference_file: str, merge_files: List[str], output_file: str = None) -> bool:
        """
        병합 프로세스 시작 (비동기)
        
        Args:
            reference_file: 원본 파일 경로
            merge_files: 취합할 파일 목록
            output_file: 출력 파일 경로 (None이면 자동 생성)
            
        Returns:
            bool: 시작 성공 여부
        """
        if self.is_processing:
            self._show_error("이미 처리 중입니다.")
            return False
        
        # 입력 파일 검증
        if not reference_file or not merge_files:
            self._show_error("원본 파일과 취합 파일을 모두 선택해주세요.")
            return False
        
        # 출력 파일명 처리
        if not output_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"통합_데이터_{timestamp}.xlsx"
        
        print(f"🚀 취합 프로세스 시작")
        print(f"📁 원본 파일: {Path(reference_file).name}")
        print(f"📊 취합 파일: {len(merge_files)}개")
        for i, file in enumerate(merge_files, 1):
            print(f"   {i}. {Path(file).name}")
        print(f"💾 출력 파일: {output_file}")
        
        # 파일 존재 확인
        if not os.path.exists(reference_file):
            self._show_error(f"원본 파일이 존재하지 않습니다: {reference_file}")
            return False
        
        for merge_file in merge_files:
            if not os.path.exists(merge_file):
                self._show_error(f"취합 파일이 존재하지 않습니다: {merge_file}")
                return False
        
        # UI 상태 업데이트
        if self.ui:
            self.ui.set_merge_state(True)
            self.ui.update_progress(0, "취합 시작...")
        
        # 백그라운드 스레드에서 실행
        thread = threading.Thread(
            target=self._merge_process,
            args=(reference_file, merge_files, output_file)
        )
        thread.daemon = True
        thread.start()
        
        print(f"✅ 백그라운드 스레드 시작됨")
        return True
    
    def cancel_merge(self):
        """병합 프로세스 취소"""
        self.is_cancelled = True
        self._update_status("취소 중...")
        print("⏹️ 사용자가 취합을 취소했습니다.")
    
    def get_progress(self) -> float:
        """현재 진행률 반환"""
        return self.current_progress
    
    def get_status(self) -> str:
        """현재 상태 반환"""
        return self.current_status
    
    def is_validation_passed(self) -> bool:
        """검수 통과 여부 반환"""
        return self.validation_passed
    
    def _merge_process(self, reference_file: str, merge_files: List[str], output_file: str):
        """
        실제 병합 처리 로직
        
        Args:
            reference_file: 원본 파일 경로
            merge_files: 취합할 파일 목록
            output_file: 출력 파일 경로
        """
        print(f"🔄 _merge_process 시작됨")
        print(f"📁 reference_file: {reference_file}")
        print(f"📊 merge_files: {merge_files}")
        print(f"💾 output_file: {output_file}")
        
        try:
            self.is_processing = True
            self.is_cancelled = False
            self.current_progress = 0.0
            self.validation_passed = False
            
            print(f"🏗️ 처리 상태 초기화 완료")
            
            total_steps = 6 + len(merge_files)  # 분석, 검수, 서식추출, 병합처리, 데이터검증, 저장 + 각 파일 처리
            current_step = 0
            
            print(f"📋 총 단계: {total_steps}개")
            
            # 1. 원본 파일 분석
            self._update_progress(current_step / total_steps, "원본 파일 분석 중...")
            
            reference_structure = self.processor.read_file_structure(reference_file)
            if reference_structure.error_message:
                raise Exception(f"원본 파일 분석 실패: {reference_structure.error_message}")
            
            if not reference_structure.sheets:
                raise Exception("원본 파일에 유효한 시트가 없습니다.")
            
            # 원본 파일 경로 설정 (중요!)
            reference_structure.file_path = reference_file
            
            current_step += 1
            
            if self.is_cancelled:
                self._handle_cancellation()
                return
                
            # 2. 사전 검수 - 파일 호환성 확인
            self._update_progress(current_step / total_steps, "파일 호환성 검수 중...")
            
            validation_result = self._validate_files_compatibility(reference_file, merge_files)
            if not validation_result['success']:
                raise Exception(f"파일 호환성 검수 실패: {validation_result['message']}")
            
            current_step += 1
            
            if self.is_cancelled:
                self._handle_cancellation()
                return
            
            # 3. 서식 기준 추출
            self._update_progress(current_step / total_steps, "서식 기준 추출 중...")
            
            format_standards = self._extract_format_standards(reference_structure)
            
            current_step += 1
            
            if self.is_cancelled:
                self._handle_cancellation()
                return
            
            # 4. 각 취합 파일 처리
            merged_data = {}  # {sheet_name: [dataframes]}
            processing_errors = []
            
            for i, merge_file in enumerate(merge_files):
                self._update_progress(
                    current_step / total_steps, 
                    f"파일 처리 중... ({i+1}/{len(merge_files)}) {Path(merge_file).name}"
                )
                
                try:
                    file_data = self._process_merge_file(
                        merge_file, reference_structure, format_standards
                    )
                    
                    # 데이터 병합
                    for sheet_name, df in file_data.items():
                        if sheet_name not in merged_data:
                            merged_data[sheet_name] = []
                        merged_data[sheet_name].append(df)
                        
                except Exception as e:
                    error_msg = f"파일 처리 오류 ({Path(merge_file).name}): {str(e)}"
                    processing_errors.append(error_msg)
                    self._log_error(error_msg)
                
                current_step += 1
                
                if self.is_cancelled:
                    self._handle_cancellation()
                    return
            
            # 5. 데이터 통합
            self._update_progress(current_step / total_steps, "데이터 통합 중...")
            
            combined_data = self._combine_sheet_data(merged_data, reference_structure)
            
            current_step += 1
            
            if self.is_cancelled:
                self._handle_cancellation()
                return
                
            # 6. 데이터 품질 검증
            self._update_progress(current_step / total_steps, "데이터 품질 검증 중...")
            
            validation_result = self._validate_merged_data(combined_data, reference_structure)
            if not validation_result['success']:
                self._log_error(f"데이터 품질 검증 경고: {validation_result['message']}")
            else:
                self.validation_passed = True
                print("✅ 데이터 품질 검증 통과")
            
            current_step += 1
            
            if self.is_cancelled:
                self._handle_cancellation()
                return
            
            # 7. 결과 파일 저장
            self._update_progress(current_step / total_steps, "결과 파일 저장 중...")
            
            self._save_merged_file(combined_data, output_file, reference_structure)
            
            # 완료
            self._update_progress(1.0, "처리 완료!")
            
            # 최종 상태 메시지
            success_files = len(merge_files) - len(processing_errors)
            status_msg = f"성공적으로 완료! ({success_files}/{len(merge_files)}개 파일 처리됨)"
            
            if processing_errors:
                status_msg += f" - {len(processing_errors)}개 파일에서 오류 발생"
            
            self._update_status(status_msg)
            
            # 결과 보고서 생성
            self._create_completion_report(success_files, combined_data, processing_errors)
            
            # 검수 통과 여부에 따른 메시지
            if self.validation_passed:
                self._show_success(f"✅ 취합이 성공적으로 완료되었습니다!\n" + 
                                 f"📊 검수 통과: 데이터 품질이 검증되었습니다.\n" +
                                 f"📁 결과 파일: {output_file}")
            else:
                self._show_success(f"⚠️ 취합이 완료되었으나 일부 문제가 있습니다.\n" + 
                                 f"📋 처리 로그를 확인해주세요.\n" +
                                 f"📁 결과 파일: {output_file}")
            
        except Exception as e:
            self._handle_error(str(e))
        finally:
            self.is_processing = False
            if self.ui:
                self.ui.set_merge_state(False)
    
    def _validate_files_compatibility(self, reference_file: str, merge_files: List[str]) -> Dict[str, Any]:
        """
        파일 호환성 사전 검수
        
        Args:
            reference_file: 원본 파일
            merge_files: 취합할 파일들
            
        Returns:
            Dict: 검수 결과
        """
        try:
            print("🔍 파일 호환성 검수 시작...")
            
            # 원본 파일 구조 분석
            ref_structure = self.processor.read_file_structure(reference_file)
            if ref_structure.error_message:
                return {
                    'success': False, 
                    'message': f"원본 파일 분석 실패: {ref_structure.error_message}"
                }
            
            ref_sheet_names = [sheet.name for sheet in ref_structure.sheets]
            compatibility_issues = []
            
            # 각 취합 파일 검수
            for i, merge_file in enumerate(merge_files, 1):
                try:
                    file_structure = self.processor.read_file_structure(merge_file)
                    if file_structure.error_message:
                        compatibility_issues.append(f"파일 {i} ({Path(merge_file).name}): 구조 분석 실패")
                        continue
                    
                    # 시트 존재 여부 확인
                    merge_sheet_names = [sheet.name for sheet in file_structure.sheets]
                    sheet_mappings = self.processor.match_sheets(ref_sheet_names, merge_sheet_names)
                    
                    if not sheet_mappings:
                        compatibility_issues.append(f"파일 {i} ({Path(merge_file).name}): 매칭되는 시트가 없음")
                    
                except Exception as e:
                    compatibility_issues.append(f"파일 {i} ({Path(merge_file).name}): 검수 중 오류 - {str(e)}")
            
            # 결과 판정
            if compatibility_issues:
                issues_text = "\n".join([f"  - {issue}" for issue in compatibility_issues[:5]])  # 최대 5개만 표시
                if len(compatibility_issues) > 5:
                    issues_text += f"\n  - ... 및 {len(compatibility_issues) - 5}개 추가 문제"
                
                return {
                    'success': False,
                    'message': f"호환성 문제가 발견되었습니다:\n{issues_text}",
                    'issues': compatibility_issues
                }
            
            print(f"✅ 모든 파일이 호환성 검수를 통과했습니다.")
            return {'success': True, 'message': '호환성 검수 통과'}
            
        except Exception as e:
            return {
                'success': False,
                'message': f"검수 중 예상치 못한 오류: {str(e)}"
            }
    
    def _validate_merged_data(self, combined_data: Dict[str, pd.DataFrame], 
                             reference_structure: FileStructure) -> Dict[str, Any]:
        """
        병합된 데이터의 품질 검증
        
        Args:
            combined_data: 병합된 데이터
            reference_structure: 원본 파일 구조
            
        Returns:
            Dict: 검증 결과
        """
        try:
            print("🔍 데이터 품질 검증 시작...")
            
            validation_issues = []
            
            for sheet in reference_structure.sheets:
                sheet_name = sheet.name
                
                if sheet_name not in combined_data:
                    validation_issues.append(f"시트 '{sheet_name}'이 결과에 포함되지 않음")
                    continue
                
                df = combined_data[sheet_name]
                
                # 데이터 존재 여부 확인
                if df.empty:
                    validation_issues.append(f"시트 '{sheet_name}'에 데이터가 없음")
                    continue
                
                # 필수 열 존재 여부 확인
                missing_columns = set(sheet.columns) - set(df.columns)
                if missing_columns:
                    validation_issues.append(f"시트 '{sheet_name}'에서 누락된 열: {', '.join(missing_columns)}")
                
                # 데이터 행 수 확인
                if len(df) == 0:
                    validation_issues.append(f"시트 '{sheet_name}'에 데이터 행이 없음")
                
                # 중복 데이터 확인 (간단한 체크)
                if len(df) != len(df.drop_duplicates()):
                    duplicate_count = len(df) - len(df.drop_duplicates())
                    validation_issues.append(f"시트 '{sheet_name}'에서 {duplicate_count}개의 중복 행 발견")
            
            # 결과 판정
            if validation_issues:
                issues_text = "\n".join([f"  - {issue}" for issue in validation_issues[:5]])
                if len(validation_issues) > 5:
                    issues_text += f"\n  - ... 및 {len(validation_issues) - 5}개 추가 문제"
                
                return {
                    'success': False,
                    'message': f"데이터 품질 문제가 발견되었습니다:\n{issues_text}",
                    'issues': validation_issues
                }
            
            print(f"✅ 데이터 품질 검증을 통과했습니다.")
            return {'success': True, 'message': '데이터 품질 검증 통과'}
            
        except Exception as e:
            return {
                'success': False,
                'message': f"검증 중 예상치 못한 오류: {str(e)}"
            }
    
    def _extract_format_standards(self, reference_structure: FileStructure) -> Dict[str, Dict[str, Dict[str, Any]]]:
        """
        원본 파일로부터 서식 기준 추출 (개선된 버전)
        
        Args:
            reference_structure: 원본 파일 구조
            
        Returns:
            Dict: 시트별 서식 기준 정보
        """
        format_standards = {}
        
        for sheet in reference_structure.sheets:
            sheet_name = sheet.name
            
            # 실제 데이터를 읽어서 서식 기준 추출
            try:
                df = self.processor.read_sheet_data(reference_structure.file_path, sheet_name)
                
                # 정규화기를 사용해서 실제 서식 기준 추출
                detailed_standards = self.normalizer.extract_format_standards(df, sheet.data_types)
                
                format_standards[sheet_name] = detailed_standards
                
                print(f"📋 {sheet_name} 서식 기준:")
                for col_name, standard in detailed_standards.items():
                    print(f"  {col_name}: {standard}")
                
            except Exception as e:
                print(f"⚠️ {sheet_name} 서식 추출 실패: {e}")
                # 기본 서식으로 대체
                format_standards[sheet_name] = {
                    col: {'type': sheet.data_types.get(col, 'text')} 
                    for col in sheet.columns
                }
        
        return format_standards
    
    def _process_merge_file(self, merge_file: str, reference_structure: FileStructure, 
                           format_standards: Dict[str, Dict[str, Dict[str, Any]]]) -> Dict[str, pd.DataFrame]:
        """
        개별 취합 파일 처리 (로그 추가)
        
        Args:
            merge_file: 취합할 파일 경로
            reference_structure: 원본 파일 구조
            format_standards: 서식 기준
            
        Returns:
            Dict[str, pd.DataFrame]: 시트별 처리된 데이터
        """
        result_data = {}
        file_name = Path(merge_file).name
        
        try:
            # 파일 구조 분석
            file_structure = self.processor.read_file_structure(merge_file)
            
            if file_structure.error_message:
                raise Exception(f"파일 분석 실패: {file_structure.error_message}")
            
            # 시트 매칭
            reference_sheet_names = [sheet.name for sheet in reference_structure.sheets]
            target_sheet_names = [sheet.name for sheet in file_structure.sheets]
            
            sheet_mappings = self.processor.match_sheets(reference_sheet_names, target_sheet_names)
            
            # 시트 매칭 로그 기록
            for ref_sheet, target_sheet in sheet_mappings.items():
                if ref_sheet != target_sheet:
                    # 시트명이 다를 때만 로그 기록
                    self.normalizer._log_change_dict(
                        file_name, ref_sheet, -1, 'SHEET_NAME',
                        'sheet_mapping', target_sheet, ref_sheet
                    )
            
            # 각 시트 처리
            for ref_sheet_name, target_sheet_name in sheet_mappings.items():
                # 참조 시트 정보
                ref_sheet = next((s for s in reference_structure.sheets if s.name == ref_sheet_name), None)
                if not ref_sheet:
                    continue
                
                # 데이터 읽기
                df = self.processor.read_sheet_data(merge_file, target_sheet_name, chunk_size=None)
                
                if df.empty:
                    continue
                
                # 열 매칭 및 정렬
                column_mappings = self.processor.match_columns(ref_sheet.columns, df.columns.tolist())
                
                # 열 매칭 로그 기록
                for ref_col, target_col in column_mappings.items():
                    if ref_col != target_col:
                        # 열명이 다를 때만 로그 기록
                        self.normalizer._log_change_dict(
                            file_name, ref_sheet_name, -1, ref_col,
                            'column_mapping', target_col, ref_col
                        )
                
                aligned_df = self._align_columns(df, ref_sheet.columns, column_mappings)
                
                # 데이터 정규화
                normalized_df, change_log = self.normalizer.normalize_data(
                    aligned_df, format_standards[ref_sheet_name], file_name, ref_sheet_name
                )
                
                # 정규화 결과 확인 로그
                print(f"🔧 정규화 완료: {file_name} -> {len(change_log)}개 변경사항")
                
                # 파일명 정보를 각 행의 비고에 추가하기 위해 메타데이터 저장
                normalized_df['_source_file'] = file_name  # 임시 컬럼
                
                result_data[ref_sheet_name] = normalized_df
        
        except Exception as e:
            raise Exception(f"파일 처리 오류 ({Path(merge_file).name}): {str(e)}")
        
        return result_data
    
    def _align_columns(self, df: pd.DataFrame, reference_columns: List[str], 
                      column_mappings: Dict[str, str]) -> pd.DataFrame:
        """
        열 순서를 원본과 일치하도록 정렬
        
        Args:
            df: 입력 데이터프레임
            reference_columns: 원본 열 순서
            column_mappings: 열 매핑 정보
            
        Returns:
            pd.DataFrame: 정렬된 데이터프레임
        """
        aligned_df = pd.DataFrame()
        
        for ref_col in reference_columns:
            if ref_col in column_mappings:
                # 매핑된 열이 있는 경우
                source_col = column_mappings[ref_col]
                if source_col in df.columns:
                    aligned_df[ref_col] = df[source_col]
                else:
                    aligned_df[ref_col] = None
            else:
                # 매핑된 열이 없는 경우 빈 열 추가
                aligned_df[ref_col] = None
        
        return aligned_df
    
    def _combine_sheet_data(self, merged_data: Dict[str, List[pd.DataFrame]], 
                           reference_structure: FileStructure) -> Dict[str, pd.DataFrame]:
        """
        시트별 데이터 결합 (비고 컬럼 추가)
        
        Args:
            merged_data: 시트별 데이터프레임 목록
            reference_structure: 원본 파일 구조
            
        Returns:
            Dict[str, pd.DataFrame]: 결합된 시트별 데이터
        """
        combined_data = {}
        
        for sheet in reference_structure.sheets:
            sheet_name = sheet.name
            
            if sheet_name not in merged_data:
                # 원본 데이터만 있는 경우
                original_df = self.processor.read_sheet_data(
                    reference_structure.file_path, sheet_name
                )
                # 비고 컬럼 추가
                original_df['비고'] = ''
                combined_data[sheet_name] = original_df
            else:
                # 원본 + 취합 데이터
                dataframes = merged_data[sheet_name]
                
                # 원본 데이터 추가
                original_df = self.processor.read_sheet_data(
                    reference_structure.file_path, sheet_name
                )
                # 원본 데이터에 비고 컬럼 추가
                original_df['비고'] = ''
                dataframes.insert(0, original_df)
                
                # 모든 데이터프레임에 비고 컬럼이 있는지 확인하고 추가
                for i, df in enumerate(dataframes):
                    if '비고' not in df.columns:
                        df['비고'] = ''
                
                # 모든 데이터프레임 결합
                combined_df = pd.concat(dataframes, ignore_index=True)
                
                # 변경 로그를 비고에 반영
                combined_df = self._add_change_notes_to_dataframe(combined_df, sheet_name)
                
                combined_data[sheet_name] = combined_df
        
        return combined_data
    
    def _add_change_notes_to_dataframe(self, df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
        """
        데이터프레임에 변경 사항을 비고로 추가 (개선된 버전)
        
        Args:
            df: 대상 데이터프레임
            sheet_name: 시트명
            
        Returns:
            pd.DataFrame: 비고가 추가된 데이터프레임
        """
        # 임시 컬럼에서 파일 출처 정보 추출
        if '_source_file' in df.columns:
            for idx, row in df.iterrows():
                source_file = row['_source_file']
                if pd.notna(source_file) and str(source_file).strip():
                    # 원본 파일이 아닌 경우에만 출처 정보 추가
                    if not source_file.startswith('원본'):
                        existing_note = str(df.iloc[idx]['비고']) if pd.notna(df.iloc[idx]['비고']) else ''
                        source_note = f"출처:{source_file}"
                        
                        if existing_note:
                            new_note = f"{existing_note}; {source_note}"
                        else:
                            new_note = source_note
                        
                        df.iloc[idx, df.columns.get_loc('비고')] = new_note
            
            # 임시 컬럼 제거
            df = df.drop(columns=['_source_file'])
        
        # 정규화기에서 수집한 변경 로그 가져오기
        if hasattr(self.normalizer, 'change_logs'):
            change_logs = self.normalizer.change_logs
            
            # 파일별로 로그 그룹화
            file_logs = {}
            for log in change_logs:
                if log.get('sheet_name') == sheet_name:
                    file_name = log.get('file_name', '')
                    if file_name not in file_logs:
                        file_logs[file_name] = []
                    file_logs[file_name].append(log)
            
            # 파일별로 변경 사항 처리
            for file_name, logs in file_logs.items():
                # 시트 및 열 매핑 정보 수집
                sheet_mappings = [log for log in logs if log.get('change_type') == 'sheet_mapping']
                column_mappings = [log for log in logs if log.get('change_type') == 'column_mapping']
                
                # 해당 파일의 모든 행에 매핑 정보 추가
                if sheet_mappings or column_mappings:
                    for idx, row in df.iterrows():
                        if '_source_file' in df.columns and df.iloc[idx]['_source_file'] == file_name:
                            continue  # 이미 처리됨
                        
                        # 해당 파일에서 온 데이터인지 확인 (간접적으로)
                        existing_note = str(df.iloc[idx]['비고']) if pd.notna(df.iloc[idx]['비고']) else ''
                        if file_name in existing_note or (not existing_note and idx >= 3):  # 원본 다음부터
                            
                            mapping_notes = []
                            
                            # 시트 매핑 정보
                            for log in sheet_mappings:
                                original = log.get('original_value', '')
                                mapped = log.get('new_value', '')
                                if original != mapped:
                                    mapping_notes.append(f"시트매칭({original}→{mapped})")
                            
                            # 열 매핑 정보
                            for log in column_mappings:
                                original = log.get('original_value', '')
                                mapped = log.get('new_value', '')
                                column = log.get('column_name', '')
                                if original != mapped:
                                    mapping_notes.append(f"열매칭({column}:{original}→{mapped})")
                            
                            if mapping_notes:
                                mapping_text = "; ".join(mapping_notes)
                                if existing_note:
                                    new_note = f"{existing_note}; {mapping_text}"
                                else:
                                    new_note = mapping_text
                                
                                df.iloc[idx, df.columns.get_loc('비고')] = new_note
                
                # 행별 변경 사항 (날짜 형식 등)
                row_changes = [log for log in logs if log.get('row_index', -1) >= 0]
                for log in row_changes:
                    row_idx = log.get('row_index', -1)
                    if 0 <= row_idx < len(df):
                        change_type = log.get('change_type', '')
                        column_name = log.get('column_name', '')
                        original_value = log.get('original_value', '')
                        new_value = log.get('new_value', '')
                        
                        # 변경 사항 텍스트 생성
                        if change_type == 'date_format':
                            note = f"날짜형식변경({column_name}:{original_value}→{new_value})"
                        elif change_type == 'number_format':
                            note = f"숫자형식변경({column_name}:{original_value}→{new_value})"
                        elif change_type == 'text_format':
                            note = f"텍스트정리({column_name}:{original_value}→{new_value})"
                        else:
                            note = f"{change_type}({column_name}:{original_value}→{new_value})"
                        
                        # 기존 비고에 추가
                        existing_note = str(df.iloc[row_idx]['비고']) if pd.notna(df.iloc[row_idx]['비고']) else ''
                        if existing_note:
                            new_note = f"{existing_note}; {note}"
                        else:
                            new_note = note
                        
                        df.iloc[row_idx, df.columns.get_loc('비고')] = new_note
        
        return df
    
    def _save_merged_file(self, combined_data: Dict[str, pd.DataFrame], 
                         output_file: str, reference_structure: FileStructure):
        """
        결합된 데이터를 엑셀 파일로 저장
        
        Args:
            combined_data: 결합된 데이터
            output_file: 출력 파일 경로
            reference_structure: 원본 파일 구조
        """
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in combined_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 열 너비 자동 조정
                    worksheet = writer.sheets[sheet_name]
                    self._adjust_column_widths(worksheet, df)
            
            print(f"✅ 결과 파일 저장 완료: {output_file}")
            
        except Exception as e:
            raise Exception(f"파일 저장 실패: {str(e)}")
    
    def _adjust_column_widths(self, worksheet, df: pd.DataFrame):
        """열 너비 자동 조정"""
        try:
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        except Exception:
            pass  # 열 너비 조정 실패는 무시
    
    def _update_progress(self, progress: float, message: str):
        """진행률 업데이트"""
        self.current_progress = progress * 100
        self.current_status = message
        
        if self.ui:
            self.ui.update_progress(self.current_progress, message)
        
        print(f"📊 {self.current_progress:.1f}% - {message}")
    
    def _update_status(self, status: str):
        """상태 업데이트"""
        self.current_status = status
        print(f"ℹ️ {status}")
    
    def _handle_error(self, error_message: str):
        """오류 처리"""
        self.current_status = f"오류: {error_message}"
        print(f"❌ 오류 발생: {error_message}")
        
        if self.ui:
            self.ui.show_error(error_message)
    
    def _handle_cancellation(self):
        """취소 처리"""
        self.current_status = "취소됨"
        print("⏹️ 작업이 취소되었습니다.")
        
        if self.ui:
            self.ui.update_progress(0, "취소됨")
    
    def _show_error(self, message: str):
        """오류 메시지 표시"""
        if self.ui:
            self.ui.show_error(message)
        else:
            print(f"❌ 오류: {message}")
    
    def _show_success(self, message: str):
        """성공 메시지 표시"""
        if self.ui:
            self.ui.show_success(message)
        else:
            print(f"✅ 성공: {message}")
    
    def _log_error(self, error_message: str):
        """오류 로그 기록"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {error_message}"
        
        self.change_logs.append(log_entry)
        print(f"⚠️ {log_entry}")
        
        # 파일에도 로그 저장
        try:
            with open("merge_errors.log", "a", encoding="utf-8") as f:
                f.write(f"{log_entry}\n")
        except:
            pass
    
    def _create_completion_report(self, processed_files_count: int, 
                                 combined_data: Dict[str, pd.DataFrame],
                                 processing_errors: List[str]):
        """완료 보고서 생성"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            report = []
            report.append(f"📊 엑셀 취합 완료 보고서")
            report.append(f"📅 완료 시간: {timestamp}")
            report.append(f"📁 처리된 파일 수: {processed_files_count}개")
            report.append(f"📋 생성된 시트 수: {len(combined_data)}개")
            report.append("")
            
            # 시트별 상세 정보
            for sheet_name, df in combined_data.items():
                report.append(f"• {sheet_name}: {len(df)}행 × {len(df.columns)}열")
            
            if self.change_logs:
                report.append("")
                report.append("⚠️ 처리 중 발생한 경고/오류:")
                for log in self.change_logs[-10:]:  # 최근 10개만
                    report.append(f"  - {log}")
            
            if processing_errors:
                report.append("")
                report.append("⚠️ 처리 중 발생한 오류:")
                for error in processing_errors[:5]:  # 최대 5개만
                    report.append(f"  - {error}")
            
            report_text = "\n".join(report)
            print("\n" + "="*50)
            print(report_text)
            print("="*50)
            
        except Exception as e:
            print(f"⚠️ 보고서 생성 실패: {e}") 