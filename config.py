# -*- coding: utf-8 -*-
"""
수창_엑셀 취합프로그램 - 설정 파일
"""

# 성능 설정
MAX_MEMORY_MB = 400  # 최대 메모리 사용량 (MB)
CHUNK_SIZE = 1000    # 데이터 청크 크기
MAX_WORKERS = 4      # 최대 스레드 수

# UI 설정
WINDOW_WIDTH = 800
WINDOW_HEIGHT = 600
WINDOW_TITLE = "수창 엑셀 취합프로그램"
THEME_COLOR = "#2E86AB"
BACKGROUND_COLOR = "#F8F9FA"

# 파일 처리 설정
SUPPORTED_FORMATS = ['.xlsx', '.xls', '.xlsm', '.csv']
MAX_FILE_SIZE_MB = 50
MAX_FILES_COUNT = 100
SIMILARITY_THRESHOLD = 0.6      # 시트명 유사도 임계값
COLUMN_SIMILARITY_THRESHOLD = 0.5  # 열제목 유사도 임계값

# 처리 타임아웃 설정
PROCESSING_TIMEOUT_MINUTES = 30

# 로그 설정
LOG_LEVEL = "INFO"
LOG_FILE = "excel_merger.log"

# 기본 출력 파일명 패턴
DEFAULT_OUTPUT_PATTERN = "통합_데이터_{datetime}.xlsx"

# 에러 메시지
ERROR_MESSAGES = {
    'FILE_NOT_FOUND': '파일을 찾을 수 없습니다.',
    'INVALID_FILE_FORMAT': '지원하지 않는 파일 형식입니다.',
    'FILE_TOO_LARGE': f'파일 크기가 {MAX_FILE_SIZE_MB}MB를 초과합니다.',
    'TOO_MANY_FILES': f'최대 {MAX_FILES_COUNT}개 파일까지 처리 가능합니다.',
    'MEMORY_LIMIT_EXCEEDED': '메모리 사용량이 한계를 초과했습니다.',
    'PROCESSING_TIMEOUT': '처리 시간이 초과되었습니다.',
    'UNKNOWN_ERROR': '알 수 없는 오류가 발생했습니다.'
}

# 성공 메시지
SUCCESS_MESSAGES = {
    'MERGE_COMPLETED': '파일 통합이 완료되었습니다.',
    'FILE_SAVED': '파일이 성공적으로 저장되었습니다.'
} 