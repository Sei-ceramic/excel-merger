#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
데이터 정규화 모듈 - 서식 통일 및 데이터 정제
"""

import pandas as pd
from typing import Dict, List, Any, Tuple, Optional
import re
from datetime import datetime, timedelta
import config

class DataNormalizer:
    """데이터 정규화 클래스"""
    
    def __init__(self):
        """정규화기 초기화"""
        self.change_logs = []  # 딕셔너리 형태로 변경 로그 저장
        
    def normalize_dataframe(self, df: pd.DataFrame, data_types: Dict[str, str]) -> pd.DataFrame:
        """
        데이터프레임 정규화 (간단한 버전)
        
        Args:
            df: 정규화할 DataFrame
            data_types: 열별 데이터 타입 정보
            
        Returns:
            pd.DataFrame: 정규화된 DataFrame
        """
        try:
            normalized_df = df.copy()
            
            for column in normalized_df.columns:
                if column in data_types:
                    data_type = data_types[column]
                    
                    if data_type == 'number':
                        normalized_df[column] = self._simple_normalize_numbers(normalized_df[column])
                    elif data_type == 'date':
                        normalized_df[column] = self._simple_normalize_dates(normalized_df[column])
                    else:  # text
                        normalized_df[column] = self._simple_normalize_text(normalized_df[column])
            
            return normalized_df
            
        except Exception as e:
            print(f"⚠️ 데이터 정규화 중 오류: {e}")
            return df  # 오류 시 원본 반환
    
    def _simple_normalize_numbers(self, series: pd.Series) -> pd.Series:
        """간단한 숫자 정규화"""
        result = series.copy()
        
        for idx, value in series.items():
            if pd.isna(value):
                continue
            try:
                if isinstance(value, str):
                    # 쉼표 제거 후 숫자 변환
                    clean_value = value.replace(',', '').strip()
                    if clean_value:
                        result.iloc[idx] = float(clean_value)
                elif isinstance(value, (int, float)):
                    result.iloc[idx] = float(value)
            except:
                continue  # 변환 실패 시 원본 유지
        
        return result
    
    def _simple_normalize_dates(self, series: pd.Series) -> pd.Series:
        """간단한 날짜 정규화"""
        result = series.copy()
        
        for idx, value in series.items():
            if pd.isna(value):
                continue
            try:
                if isinstance(value, datetime):
                    result.iloc[idx] = value.strftime('%Y-%m-%d')
                elif isinstance(value, str):
                    # 간단한 날짜 변환 시도
                    date_formats = ['%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%Y.%m.%d']
                    for fmt in date_formats:
                        try:
                            parsed_date = datetime.strptime(value.strip(), fmt)
                            result.iloc[idx] = parsed_date.strftime('%Y-%m-%d')
                            break
                        except ValueError:
                            continue
            except:
                continue  # 변환 실패 시 원본 유지
        
        return result
    
    def _simple_normalize_text(self, series: pd.Series) -> pd.Series:
        """간단한 텍스트 정규화"""
        result = series.copy()
        
        for idx, value in series.items():
            if pd.isna(value):
                continue
            try:
                # 문자열로 변환 후 앞뒤 공백 제거
                result.iloc[idx] = str(value).strip()
            except:
                continue
        
        return result

    def normalize_data(self, df: pd.DataFrame, format_standards: Dict[str, Dict[str, Any]], 
                      file_name: str = "", sheet_name: str = "") -> Tuple[pd.DataFrame, List[Dict]]:
        """
        데이터 정규화 메인 함수
        
        Args:
            df: 정규화할 DataFrame
            format_standards: 서식 기준 정보
            file_name: 파일명 (로그용)
            sheet_name: 시트명 (로그용)
            
        Returns:
            Tuple[pd.DataFrame, List[Dict]]: 정규화된 DataFrame과 변경 로그
        """
        self.change_logs = []
        normalized_df = df.copy()
        
        for column in normalized_df.columns:
            if column in format_standards:
                standard = format_standards[column]
                
                # 열별 정규화 수행
                normalized_df[column] = self._normalize_column(
                    normalized_df[column], 
                    standard,
                    column,
                    file_name,
                    sheet_name
                )
        
        # 추가적인 데이터 정제
        normalized_df = self._clean_data(normalized_df, file_name, sheet_name)
        
        return normalized_df, self.change_logs
    
    def _normalize_column(self, series: pd.Series, standard: Dict[str, Any], 
                         column_name: str, file_name: str, sheet_name: str) -> pd.Series:
        """
        개별 열 정규화
        
        Args:
            series: 정규화할 시리즈
            standard: 서식 기준
            column_name: 열명
            file_name: 파일명 (로그용)
            sheet_name: 시트명 (로그용)
            
        Returns:
            pd.Series: 정규화된 시리즈
        """
        data_type = standard.get('type', 'text')
        
        if data_type == 'number':
            return self._normalize_numbers(series, standard, column_name, file_name, sheet_name)
        elif data_type == 'date':
            return self._normalize_dates(series, standard, column_name, file_name, sheet_name)
        elif data_type == 'text':
            return self._normalize_text(series, standard, column_name, file_name, sheet_name)
        else:
            return series
    
    def _normalize_numbers(self, series: pd.Series, standard: Dict[str, Any], 
                          column_name: str, file_name: str, sheet_name: str) -> pd.Series:
        """
        숫자 데이터 정규화
        
        Args:
            series: 숫자 시리즈
            standard: 숫자 서식 기준
            column_name: 열명
            file_name: 파일명
            sheet_name: 시트명
            
        Returns:
            pd.Series: 정규화된 숫자 시리즈
        """
        result = series.copy()
        
        # 기본 설정값
        decimal_places = standard.get('decimal_places', 2)
        use_thousands_separator = standard.get('thousands_separator', True)
        
        for idx, value in series.items():
            if pd.isna(value):
                continue
                
            original_value = value
            converted_value = None
            
            try:
                # 문자열인 경우 숫자로 변환
                if isinstance(value, str):
                    # 쉼표 제거 후 숫자 변환
                    clean_value = value.replace(',', '')
                    converted_value = float(clean_value)
                elif isinstance(value, (int, float)):
                    converted_value = float(value)
                
                if converted_value is not None:
                    # 소수점 자릿수 조정
                    if decimal_places >= 0:
                        converted_value = round(converted_value, decimal_places)
                    
                    result.iloc[idx] = converted_value
                    
                    # 변경 로그 기록
                    if str(original_value) != str(converted_value):
                        self._log_change_dict(
                            file_name, sheet_name, idx, column_name,
                            'number_format', original_value, converted_value
                        )
                        
            except (ValueError, TypeError):
                # 변환 실패 시 원본 유지
                continue
        
        return result
    
    def _normalize_dates(self, series: pd.Series, standard: Dict[str, Any], 
                        column_name: str, file_name: str, sheet_name: str) -> pd.Series:
        """
        날짜 데이터 정규화 (완전 통일 버전)
        
        Args:
            series: 날짜 시리즈
            standard: 날짜 서식 기준
            column_name: 열명
            file_name: 파일명
            sheet_name: 시트명
            
        Returns:
            pd.Series: 정규화된 날짜 시리즈
        """
        result = series.copy()
        
        # 기본 날짜 형식 (원본 파일 기준) - 무조건 이 형식으로 통일
        target_format = standard.get('format', '%Y-%m-%d')
        
        # 확장된 날짜 형식 목록
        date_formats = [
            # 기본 형식들
            '%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y',
            '%Y.%m.%d', '%m.%d.%Y', '%d.%m.%Y',
            # 시간이 포함된 형식들
            '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
            '%Y.%m.%d %H:%M:%S', '%m.%d.%Y %H:%M:%S',
            # 한국어 형식들
            '%Y년 %m월 %d일', '%Y년%m월%d일',
            # 축약된 연도 형식들
            '%y-%m-%d', '%y/%m/%d', '%m/%d/%y', '%d/%m/%y',
            '%y.%m.%d', '%m.%d.%y', '%d.%m.%y',
            # 점으로 구분된 형식들 (독일식 등)
            '%d.%m.%Y', '%d.%m.%y',
            # 특수 형식들
            '%Y%m%d', '%y%m%d',  # 연속된 숫자
            '%B %d, %Y', '%b %d, %Y',  # 영어 월명
        ]
        
        for idx, value in series.items():
            if pd.isna(value):
                continue
                
            original_value = value
            converted_value = None
            
            try:
                # 이미 datetime 객체인 경우
                if isinstance(value, datetime):
                    converted_value = value.strftime(target_format)
                # 숫자인 경우 (엑셀 시리얼 날짜)
                elif isinstance(value, (int, float)):
                    if 1 <= value <= 2958465:  # 엑셀 날짜 범위
                        try:
                            # 엑셀 기준일 (1900-01-01)에서 계산
                            excel_date = datetime(1900, 1, 1) + timedelta(days=value - 2)
                            converted_value = excel_date.strftime(target_format)
                        except:
                            continue
                # 문자열인 경우
                elif isinstance(value, str):
                    value_str = str(value).strip()
                    
                    # 빈 문자열 스킵
                    if not value_str:
                        continue
                    
                    # 이미 목표 형식인지 확인
                    if self._is_target_format(value_str, target_format):
                        converted_value = value_str  # 이미 올바른 형식
                    else:
                        # 다양한 날짜 형식 시도
                        for fmt in date_formats:
                            try:
                                parsed_date = datetime.strptime(value_str, fmt)
                                converted_value = parsed_date.strftime(target_format)
                                break
                            except ValueError:
                                continue
                        
                        # 정규표현식을 사용한 추가 파싱 시도
                        if converted_value is None:
                            converted_value = self._parse_date_with_regex(value_str, target_format)
                
                # 결과 설정 및 로그 기록
                if converted_value is not None:
                    result.iloc[idx] = converted_value
                    
                    # 변경이 있었을 때만 로그 기록
                    if str(original_value) != str(converted_value):
                        self._log_change_dict(
                            file_name, sheet_name, idx, column_name,
                            'date_format', original_value, converted_value
                        )
                        
            except (ValueError, TypeError, OverflowError):
                # 변환 실패 시 원본 유지
                continue
        
        return result
    
    def _is_target_format(self, date_str: str, target_format: str) -> bool:
        """
        문자열이 이미 목표 형식인지 확인
        
        Args:
            date_str: 날짜 문자열
            target_format: 목표 형식
            
        Returns:
            bool: 목표 형식 여부
        """
        try:
            parsed = datetime.strptime(date_str, target_format)
            reformatted = parsed.strftime(target_format)
            return reformatted == date_str
        except ValueError:
            return False
    
    def _parse_date_with_regex(self, date_str: str, target_format: str) -> Optional[str]:
        """
        정규표현식을 사용한 날짜 파싱
        
        Args:
            date_str: 날짜 문자열
            target_format: 목표 형식
            
        Returns:
            str: 파싱된 날짜 문자열 또는 None
        """
        import re
        
        # 숫자만 추출하여 날짜 파싱 시도
        numbers = re.findall(r'\d+', date_str)
        
        if len(numbers) >= 3:
            try:
                # 연도, 월, 일 추정
                year, month, day = None, None, None
                
                # 첫 번째 숫자가 4자리면 연도로 추정
                if len(numbers[0]) == 4:
                    year, month, day = int(numbers[0]), int(numbers[1]), int(numbers[2])
                # 마지막 숫자가 4자리면 연도로 추정
                elif len(numbers[-1]) == 4:
                    month, day, year = int(numbers[0]), int(numbers[1]), int(numbers[2])
                # 중간 숫자가 4자리면 연도로 추정
                elif len(numbers) >= 3 and len(numbers[1]) == 4:
                    day, year, month = int(numbers[0]), int(numbers[1]), int(numbers[2])
                else:
                    # 2자리 연도 처리
                    year_candidate = int(numbers[-1]) if len(numbers[-1]) == 2 else int(numbers[0])
                    if year_candidate < 50:
                        year = 2000 + year_candidate
                    else:
                        year = 1900 + year_candidate
                    
                    if len(numbers[0]) == 2 and year_candidate == int(numbers[0]):
                        month, day = int(numbers[1]), int(numbers[2])
                    else:
                        month, day = int(numbers[0]), int(numbers[1])
                
                # 유효성 검사
                if 1 <= month <= 12 and 1 <= day <= 31 and 1900 <= year <= 2100:
                    parsed_date = datetime(year, month, day)
                    return parsed_date.strftime(target_format)
                    
            except (ValueError, IndexError):
                pass
        
        return None
    
    def _normalize_text(self, series: pd.Series, standard: Dict[str, Any], 
                       column_name: str, file_name: str, sheet_name: str) -> pd.Series:
        """
        텍스트 데이터 정규화
        
        Args:
            series: 텍스트 시리즈
            standard: 텍스트 서식 기준
            column_name: 열명
            file_name: 파일명
            sheet_name: 시트명
            
        Returns:
            pd.Series: 정규화된 텍스트 시리즈
        """
        result = series.copy()
        
        # 기본 설정값
        trim_whitespace = standard.get('trim_whitespace', True)
        normalize_spaces = standard.get('normalize_spaces', True)
        case_normalization = standard.get('case', None)  # 'upper', 'lower', 'title', None
        
        for idx, value in series.items():
            if pd.isna(value):
                continue
                
            original_value = value
            converted_value = str(value)
            
            # 앞뒤 공백 제거
            if trim_whitespace:
                converted_value = converted_value.strip()
            
            # 중간 공백 정리 (연속된 공백을 하나로)
            if normalize_spaces:
                converted_value = re.sub(r'\s+', ' ', converted_value)
            
            # 대소문자 변환
            if case_normalization == 'upper':
                converted_value = converted_value.upper()
            elif case_normalization == 'lower':
                converted_value = converted_value.lower()
            elif case_normalization == 'title':
                converted_value = converted_value.title()
            
            result.iloc[idx] = converted_value
            
            # 변경 로그 기록
            if str(original_value) != str(converted_value):
                self._log_change_dict(
                    file_name, sheet_name, idx, column_name,
                    'text_format', original_value, converted_value
                )
        
        return result
    
    def _clean_data(self, df: pd.DataFrame, file_name: str, sheet_name: str) -> pd.DataFrame:
        """
        추가적인 데이터 정제
        
        Args:
            df: 정제할 DataFrame
            file_name: 파일명
            sheet_name: 시트명
            
        Returns:
            pd.DataFrame: 정제된 DataFrame
        """
        cleaned_df = df.copy()
        
        # 완전히 빈 행 제거
        empty_rows_before = len(cleaned_df)
        cleaned_df = cleaned_df.dropna(how='all')
        empty_rows_after = len(cleaned_df)
        
        if empty_rows_before != empty_rows_after:
            self._log_change_dict(
                file_name, sheet_name, -1, 'ALL_COLUMNS',
                'remove_empty_rows', 
                f"{empty_rows_before}행", 
                f"{empty_rows_after}행"
            )
        
        # 인덱스 재설정
        cleaned_df = cleaned_df.reset_index(drop=True)
        
        return cleaned_df
    
    def extract_format_standards(self, df: pd.DataFrame, data_types: Dict[str, str]) -> Dict[str, Dict[str, Any]]:
        """
        DataFrame에서 서식 기준 추출
        
        Args:
            df: 기준이 될 DataFrame
            data_types: 열별 데이터 타입 정보
            
        Returns:
            Dict[str, Dict[str, Any]]: 열별 서식 기준
        """
        standards = {}
        
        for column in df.columns:
            if column not in data_types:
                continue
                
            data_type = data_types[column]
            standard = {'type': data_type}
            
            # 해당 열의 샘플 데이터 추출 (null이 아닌 처음 10개)
            sample_data = df[column].dropna().head(10).tolist()
            
            if data_type == 'number':
                standard.update(self._analyze_number_format(sample_data))
            elif data_type == 'date':
                standard.update(self._analyze_date_format(sample_data))
            elif data_type == 'text':
                standard.update(self._analyze_text_format(sample_data))
            
            standards[column] = standard
        
        return standards
    
    def _analyze_number_format(self, sample_data: List[Any]) -> Dict[str, Any]:
        """
        숫자 형식 분석
        
        Args:
            sample_data: 샘플 데이터
            
        Returns:
            Dict[str, Any]: 숫자 서식 정보
        """
        decimal_places = 0
        has_thousands_separator = False
        
        for value in sample_data:
            if isinstance(value, str):
                # 천단위 구분자 확인
                if ',' in value:
                    has_thousands_separator = True
                
                # 소수점 자릿수 확인
                if '.' in value:
                    decimal_part = value.split('.')[-1]
                    decimal_places = max(decimal_places, len(decimal_part))
                    
            elif isinstance(value, float):
                # float의 소수점 자릿수 확인
                decimal_str = str(value)
                if '.' in decimal_str:
                    decimal_part = decimal_str.split('.')[-1]
                    decimal_places = max(decimal_places, len(decimal_part))
        
        return {
            'decimal_places': decimal_places,
            'thousands_separator': has_thousands_separator
        }
    
    def _analyze_date_format(self, sample_data: List[Any]) -> Dict[str, Any]:
        """
        날짜 형식 분석
        
        Args:
            sample_data: 샘플 데이터
            
        Returns:
            Dict[str, Any]: 날짜 서식 정보
        """
        # 가장 일반적인 형식들의 우선순위
        common_formats = [
            '%Y-%m-%d',
            '%Y/%m/%d', 
            '%m/%d/%Y',
            '%d/%m/%Y',
            '%Y.%m.%d'
        ]
        
        format_counts = {fmt: 0 for fmt in common_formats}
        
        for value in sample_data:
            if isinstance(value, str):
                for fmt in common_formats:
                    try:
                        datetime.strptime(value.strip(), fmt)
                        format_counts[fmt] += 1
                        break
                    except ValueError:
                        continue
            elif isinstance(value, datetime):
                # 기본적으로 ISO 형식 선호
                format_counts['%Y-%m-%d'] += 1
        
        # 가장 많이 매칭된 형식 선택
        best_format = max(format_counts, key=format_counts.get)
        
        return {
            'format': best_format
        }
    
    def _analyze_text_format(self, sample_data: List[Any]) -> Dict[str, Any]:
        """
        텍스트 형식 분석
        
        Args:
            sample_data: 샘플 데이터
            
        Returns:
            Dict[str, Any]: 텍스트 서식 정보
        """
        has_leading_trailing_spaces = False
        has_multiple_spaces = False
        case_pattern = None
        
        for value in sample_data:
            if isinstance(value, str):
                # 앞뒤 공백 확인
                if value != value.strip():
                    has_leading_trailing_spaces = True
                
                # 연속 공백 확인
                if re.search(r'\s{2,}', value):
                    has_multiple_spaces = True
                
                # 대소문자 패턴 분석
                if value.isupper() and case_pattern != 'mixed':
                    case_pattern = 'upper' if case_pattern is None else 'mixed'
                elif value.islower() and case_pattern != 'mixed':
                    case_pattern = 'lower' if case_pattern is None else 'mixed'
                elif value.istitle() and case_pattern != 'mixed':
                    case_pattern = 'title' if case_pattern is None else 'mixed'
                else:
                    case_pattern = 'mixed'
        
        return {
            'trim_whitespace': has_leading_trailing_spaces,
            'normalize_spaces': has_multiple_spaces,
            'case': case_pattern if case_pattern != 'mixed' else None
        }
    
    def _log_change_dict(self, file_name: str, sheet_name: str, row_index: int, 
                        column_name: str, change_type: str, original_value: Any, new_value: Any):
        """
        변경 사항을 딕셔너리 형태로 로그에 기록
        
        Args:
            file_name: 파일명
            sheet_name: 시트명
            row_index: 행 인덱스
            column_name: 열명
            change_type: 변경 유형
            original_value: 원본 값
            new_value: 변경된 값
        """
        log_entry = {
            'file_name': file_name,
            'sheet_name': sheet_name,
            'row_index': row_index,
            'column_name': column_name,
            'change_type': change_type,
            'original_value': original_value,
            'new_value': new_value,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        self.change_logs.append(log_entry)
        
        # 기존 텍스트 로그도 유지 (호환성)
        self._log_change(file_name, sheet_name, row_index, column_name, change_type, original_value, new_value)
    
    def _log_change(self, file_name: str, sheet_name: str, row_index: int, 
                   column_name: str, change_type: str, original_value: Any, new_value: Any):
        """
        기존 텍스트 로그 기록 (호환성 유지)
        
        Args:
            file_name: 파일명
            sheet_name: 시트명
            row_index: 행 인덱스
            column_name: 열명
            change_type: 변경 유형
            original_value: 원본 값
            new_value: 변경된 값
        """
        # 기존 change_log 리스트가 있다면 사용, 없다면 무시
        if hasattr(self, 'change_log'):
            self.change_log.append({
                'file_name': file_name,
                'sheet_name': sheet_name,
                'row_index': row_index,
                'column_name': column_name,
                'change_type': change_type,
                'original_value': original_value,
                'new_value': new_value,
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
    
    def get_change_summary(self) -> Dict[str, int]:
        """
        변경사항 요약 통계
        
        Returns:
            Dict[str, int]: 변경 유형별 개수
        """
        summary = {}
        for change in self.change_logs:
            change_type = change['change_type']
            summary[change_type] = summary.get(change_type, 0) + 1
        
        return summary
    
    def create_change_report(self) -> str:
        """
        변경사항 보고서 생성
        
        Returns:
            str: 변경사항 보고서 텍스트
        """
        if not self.change_logs:
            return "변경사항이 없습니다."
        
        summary = self.get_change_summary()
        
        report = "=== 데이터 정규화 보고서 ===\n\n"
        report += f"총 변경사항: {len(self.change_logs)}건\n\n"
        
        report += "변경 유형별 통계:\n"
        for change_type, count in summary.items():
            report += f"- {change_type}: {count}건\n"
        
        report += "\n상세 변경사항 (최대 100건):\n"
        for i, change in enumerate(self.change_logs[:100]):
            report += f"{i+1}. [{change['file_name']}] {change['sheet_name']} 시트, "
            report += f"행 {change['row_index']}, {change['column_name']} 열: "
            report += f"{change['original_value']} → {change['new_value']}\n"
        
        if len(self.change_logs) > 100:
            report += f"\n... 외 {len(self.change_logs) - 100}건 더\n"
        
        return report 