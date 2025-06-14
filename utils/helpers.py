# -*- coding: utf-8 -*-
"""
공통 유틸리티 함수들
"""

import os
import time
import psutil
from pathlib import Path
from typing import List, Optional
from datetime import datetime
import config

def format_file_size(size_bytes: int) -> str:
    """
    파일 크기를 사람이 읽기 쉬운 형태로 변환
    
    Args:
        size_bytes: 바이트 단위 파일 크기
        
    Returns:
        str: 형식화된 파일 크기 문자열
    """
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"

def get_memory_usage() -> float:
    """
    현재 프로세스의 메모리 사용량 반환
    
    Returns:
        float: 메모리 사용량 (MB)
    """
    try:
        process = psutil.Process()
        memory_info = process.memory_info()
        return memory_info.rss / (1024 * 1024)  # MB 단위
    except:
        return 0.0

def is_memory_limit_exceeded(threshold_mb: float = None) -> bool:
    """
    메모리 사용량이 임계값을 초과했는지 확인
    
    Args:
        threshold_mb: 임계값 (MB), None이면 설정값 사용
        
    Returns:
        bool: 임계값 초과 여부
    """
    if threshold_mb is None:
        threshold_mb = config.MAX_MEMORY_MB
    
    current_usage = get_memory_usage()
    return current_usage > threshold_mb

def create_safe_filename(filename: str) -> str:
    """
    안전한 파일명 생성 (특수문자 제거)
    
    Args:
        filename: 원본 파일명
        
    Returns:
        str: 안전한 파일명
    """
    # 금지된 문자들 제거
    forbidden_chars = '<>:"/\\|?*'
    safe_name = filename
    
    for char in forbidden_chars:
        safe_name = safe_name.replace(char, '_')
    
    # 연속된 언더스코어 제거
    while '__' in safe_name:
        safe_name = safe_name.replace('__', '_')
    
    # 앞뒤 공백 및 점 제거
    safe_name = safe_name.strip(' .')
    
    # 빈 문자열 방지
    if not safe_name:
        safe_name = "unnamed_file"
    
    return safe_name

def generate_output_filename(base_name: str = None) -> str:
    """
    출력 파일명 생성
    
    Args:
        base_name: 기본 파일명
        
    Returns:
        str: 생성된 파일명
    """
    if base_name is None:
        base_name = "통합_데이터"
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f"{base_name}_{timestamp}.xlsx"

def validate_file_paths(file_paths: List[str]) -> List[str]:
    """
    파일 경로 목록 유효성 검사
    
    Args:
        file_paths: 검사할 파일 경로 목록
        
    Returns:
        List[str]: 유효한 파일 경로 목록
    """
    valid_paths = []
    
    for path in file_paths:
        if os.path.exists(path) and os.path.isfile(path):
            # 파일 크기 확인
            size_mb = os.path.getsize(path) / (1024 * 1024)
            if size_mb <= config.MAX_FILE_SIZE_MB:
                # 확장자 확인
                ext = Path(path).suffix.lower()
                if ext in config.SUPPORTED_FORMATS:
                    valid_paths.append(path)
    
    return valid_paths

def estimate_processing_time(file_count: int, total_size_mb: float) -> float:
    """
    처리 시간 추정
    
    Args:
        file_count: 파일 개수
        total_size_mb: 총 파일 크기 (MB)
        
    Returns:
        float: 예상 처리 시간 (분)
    """
    # 간단한 추정 공식 (실제로는 더 정교한 계산 필요)
    base_time = file_count * 0.5  # 파일당 30초
    size_factor = total_size_mb * 0.1  # MB당 6초
    
    return max(base_time + size_factor, 1.0)  # 최소 1분

def cleanup_temp_files(temp_dir: Optional[str] = None):
    """
    임시 파일들 정리
    
    Args:
        temp_dir: 임시 디렉토리 경로 (None이면 기본 임시 디렉토리)
    """
    import tempfile
    import glob
    
    if temp_dir is None:
        temp_dir = tempfile.gettempdir()
    
    # 프로그램 관련 임시 파일 패턴
    patterns = [
        "excel_merger_*.tmp",
        "~$*.xlsx",
        "~$*.xls"
    ]
    
    for pattern in patterns:
        for file_path in glob.glob(os.path.join(temp_dir, pattern)):
            try:
                os.remove(file_path)
            except:
                pass  # 삭제 실패는 무시

def log_system_info():
    """시스템 정보 로깅"""
    try:
        import platform
        
        info = {
            "OS": platform.system(),
            "OS Version": platform.version(),
            "Python Version": platform.python_version(),
            "CPU Count": psutil.cpu_count(),
            "Total Memory": f"{psutil.virtual_memory().total / (1024**3):.1f} GB",
            "Available Memory": f"{psutil.virtual_memory().available / (1024**3):.1f} GB"
        }
        
        print("=== 시스템 정보 ===")
        for key, value in info.items():
            print(f"{key}: {value}")
        print("=" * 20)
        
    except Exception as e:
        print(f"시스템 정보 수집 실패: {e}")

def create_backup(file_path: str) -> Optional[str]:
    """
    파일 백업 생성
    
    Args:
        file_path: 백업할 파일 경로
        
    Returns:
        Optional[str]: 백업 파일 경로 (실패 시 None)
    """
    try:
        if not os.path.exists(file_path):
            return None
        
        path_obj = Path(file_path)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = path_obj.parent / f"{path_obj.stem}_backup_{timestamp}{path_obj.suffix}"
        
        import shutil
        shutil.copy2(file_path, backup_path)
        
        return str(backup_path)
        
    except Exception as e:
        print(f"백업 생성 실패: {e}")
        return None

class PerformanceTimer:
    """성능 측정용 타이머"""
    
    def __init__(self):
        self.start_time = None
        self.end_time = None
        
    def start(self):
        """타이머 시작"""
        self.start_time = time.time()
        
    def stop(self):
        """타이머 종료"""
        self.end_time = time.time()
        
    def elapsed(self) -> float:
        """경과 시간 반환 (초)"""
        if self.start_time is None:
            return 0.0
        
        end = self.end_time if self.end_time else time.time()
        return end - self.start_time
    
    def elapsed_str(self) -> str:
        """경과 시간을 문자열로 반환"""
        elapsed = self.elapsed()
        
        if elapsed < 60:
            return f"{elapsed:.1f}초"
        elif elapsed < 3600:
            minutes = elapsed // 60
            seconds = elapsed % 60
            return f"{int(minutes)}분 {seconds:.1f}초"
        else:
            hours = elapsed // 3600
            minutes = (elapsed % 3600) // 60
            return f"{int(hours)}시간 {int(minutes)}분" 