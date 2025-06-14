#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
수창_엑셀 취합프로그램 - 메인 실행 파일
"""

import sys
import os
from pathlib import Path
import traceback

# 프로젝트 루트 디렉토리를 Python path에 추가
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def main():
    """애플리케이션 메인 함수"""
    try:
        from ui.main_window import MainWindow
        from core.controller import MergeController
        
        # 컨트롤러 생성
        controller = MergeController()
        
        # 메인 윈도우 생성
        app = MainWindow()
        
        # 콜백 연결
        app.set_callbacks(
            on_start_merge=controller.start_merge,
            on_cancel_merge=controller.cancel_merge,
            on_close=controller.cleanup
        )
        
        # 컨트롤러에 UI 참조 설정
        controller.set_ui(app)
        
        print("✅ 수창 엑셀 취합프로그램이 시작되었습니다.")
        print("📁 지원 형식: .xlsx, .xls, .xlsm, .csv")
        
        # 앱 실행
        app.run()
        
    except ImportError as e:
        print(f"❌ 모듈 import 오류: {e}")
        print("필요한 라이브러리가 설치되어 있는지 확인하세요:")
        print("pip install -r requirements.txt")
        input("아무 키나 누르세요...")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 애플리케이션 실행 중 오류 발생: {e}")
        print("상세 오류 정보:")
        traceback.print_exc()
        input("아무 키나 누르세요...")
        sys.exit(1)

if __name__ == "__main__":
    main() 