#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXE 파일 빌드 스크립트
"""

import os
import sys
import shutil
from pathlib import Path
import subprocess

def build_exe():
    """PyInstaller를 사용해서 EXE 파일 빌드"""
    
    print("🔨 EXE 파일 빌드 시작...")
    print("=" * 50)
    
    # 현재 디렉토리 확인
    current_dir = Path.cwd()
    print(f"📁 작업 디렉토리: {current_dir}")
    
    # 빌드 디렉토리 정리
    build_dirs = ['build', 'dist', '__pycache__']
    for dir_name in build_dirs:
        if os.path.exists(dir_name):
            print(f"🗑️ 기존 {dir_name} 디렉토리 삭제...")
            shutil.rmtree(dir_name)
    
    # .spec 파일 삭제
    spec_files = list(current_dir.glob("*.spec"))
    for spec_file in spec_files:
        print(f"🗑️ 기존 spec 파일 삭제: {spec_file}")
        spec_file.unlink()
    
    # PyInstaller 명령어 구성
    cmd = [
        "pyinstaller",
        "--onefile",                    # 단일 파일로 생성
        "--windowed",                   # 콘솔창 숨김
        "--name=엑셀취합프로그램",          # 실행파일 이름
        "--icon=icon.ico",              # 아이콘 (있는 경우)
        "--add-data=ui;ui",             # UI 폴더 포함
        "--add-data=core;core",         # core 폴더 포함
        "--add-data=utils;utils",       # utils 폴더 포함
        "--hidden-import=pandas",       # pandas 명시적 포함
        "--hidden-import=openpyxl",     # openpyxl 명시적 포함
        "--hidden-import=xlrd",         # xlrd 명시적 포함
        "--hidden-import=PyQt5",        # PyQt5 명시적 포함
        "--hidden-import=xlsxwriter",   # xlsxwriter 명시적 포함
        "--exclude-module=tkinter",     # tkinter 제외
        "--exclude-module=matplotlib",  # matplotlib 제외
        "--exclude-module=numpy.f2py",  # 불필요한 numpy 모듈 제외
        "--optimize=2",                 # 최적화 레벨 2
        "main.py"                       # 메인 파일
    ]
    
    # 아이콘 파일이 없으면 아이콘 옵션 제거
    if not os.path.exists("icon.ico"):
        cmd = [c for c in cmd if not c.startswith("--icon")]
        print("⚠️ 아이콘 파일(icon.ico)이 없어서 기본 아이콘을 사용합니다.")
    
    print("🔧 PyInstaller 실행...")
    print(f"명령어: {' '.join(cmd)}")
    print()
    
    try:
        # PyInstaller 실행
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ PyInstaller 실행 완료!")
        
        # 결과 확인
        exe_path = Path("dist") / "엑셀취합프로그램.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"✅ EXE 파일 생성 성공!")
            print(f"📁 경로: {exe_path}")
            print(f"📊 크기: {size_mb:.1f} MB")
            
            # 테스트용 배치 파일 생성
            create_test_batch()
            
            print()
            print("🎉 빌드 완료!")
            print(f"👉 실행 파일: {exe_path}")
            print("👉 테스트용 배치 파일: dist/test_run.bat")
            
        else:
            print("❌ EXE 파일 생성 실패!")
            
    except subprocess.CalledProcessError as e:
        print(f"❌ PyInstaller 실행 실패: {e}")
        print("오류 출력:")
        print(e.stderr)
        
    except FileNotFoundError:
        print("❌ PyInstaller가 설치되지 않았습니다.")
        print("다음 명령어로 설치하세요: pip install pyinstaller")

def create_test_batch():
    """테스트용 배치 파일 생성"""
    
    batch_content = '''@echo off
chcp 65001 > nul
echo 🚀 엑셀 취합 프로그램 테스트 실행
echo.
echo 프로그램을 실행합니다...
echo.

"엑셀취합프로그램.exe"

echo.
echo 프로그램이 종료되었습니다.
pause
'''
    
    batch_path = Path("dist") / "test_run.bat"
    with open(batch_path, 'w', encoding='utf-8') as f:
        f.write(batch_content)
    
    print(f"📝 테스트 배치 파일 생성: {batch_path}")

def create_release_files():
    """릴리즈용 파일들 생성"""
    
    print("📦 릴리즈 파일 생성...")
    
    # 릴리즈 폴더 생성
    release_dir = Path("release")
    if release_dir.exists():
        shutil.rmtree(release_dir)
    release_dir.mkdir()
    
    # EXE 파일 복사
    exe_source = Path("dist") / "엑셀취합프로그램.exe"
    exe_dest = release_dir / "엑셀취합프로그램.exe"
    
    if exe_source.exists():
        shutil.copy2(exe_source, exe_dest)
        print(f"✅ EXE 파일 복사: {exe_dest}")
    
    # 사용법 파일 생성
    usage_content = '''# 엑셀 취합 프로그램 사용법

## 📋 시작하기

1. **엑셀취합프로그램.exe** 파일을 더블클릭하여 실행합니다.
2. **원본 파일 선택**: 기준이 되는 엑셀 파일을 선택합니다.
3. **취합 파일 추가**: 병합할 엑셀 파일들을 추가합니다.
4. **출력 파일명 설정**: 결과 파일의 이름을 입력합니다.
5. **취합 시작**: 버튼을 클릭하여 자동 처리를 시작합니다.

## ⚠️ 주의사항

- 엑셀 파일이 다른 프로그램에서 열려있으면 안됩니다.
- 큰 파일의 경우 처리 시간이 오래 걸릴 수 있습니다.
- 처리 중에는 컴퓨터를 끄지 마세요.

## 🆘 문제 해결

프로그램이 실행되지 않을 때:
1. Windows Defender 예외 설정 확인
2. 바이러스 백신 프로그램 예외 설정
3. 관리자 권한으로 실행

자세한 내용은 GitHub 페이지를 참조하세요.
'''
    
    usage_path = release_dir / "사용법.md"
    with open(usage_path, 'w', encoding='utf-8') as f:
        f.write(usage_content)
    
    print(f"📝 사용법 파일 생성: {usage_path}")
    
    # 샘플 파일 폴더 생성
    sample_dir = release_dir / "샘플파일"
    sample_dir.mkdir()
    
    print(f"📁 샘플 파일 폴더 생성: {sample_dir}")
    print("✅ 릴리즈 파일 생성 완료!")

if __name__ == "__main__":
    print("🔨 엑셀 취합 프로그램 빌드")
    print("=" * 40)
    print()
    
    # 필요한 패키지 확인
    try:
        import PyInstaller
        print("✅ PyInstaller 확인됨")
    except ImportError:
        print("❌ PyInstaller가 설치되어 있지 않습니다.")
        print("설치 명령어: pip install pyinstaller")
        sys.exit(1)
    
    # 메인 파일 확인
    if not os.path.exists("main.py"):
        print("❌ main.py 파일이 없습니다.")
        sys.exit(1)
    
    print("✅ main.py 확인됨")
    print()
    
    # 빌드 실행
    build_exe()
    
    # 릴리즈 파일 생성
    create_release_files()
    
    print()
    print("🎉 모든 작업 완료!")
    print("📁 dist/ 폴더에서 EXE 파일을 확인하세요.")
    print("📁 release/ 폴더에서 배포용 파일들을 확인하세요.") 