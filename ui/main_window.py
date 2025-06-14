# -*- coding: utf-8 -*-
"""
메인 윈도우 - 애플리케이션의 주 인터페이스
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import os
from pathlib import Path
from typing import List, Optional, Callable
from datetime import datetime
import config

class ModernButton(tk.Button):
    """현대적인 스타일의 버튼"""
    def __init__(self, parent, text, command=None, style="primary", **kwargs):
        # 기본 스타일 설정
        styles = {
            "primary": {
                "bg": "#2563eb",
                "fg": "white",
                "activebackground": "#1d4ed8",
                "activeforeground": "white",
                "relief": "flat",
                "borderwidth": 0,
                "font": ("Malgun Gothic", 10, "bold"),
                "cursor": "hand2"
            },
            "secondary": {
                "bg": "#6b7280",
                "fg": "white",
                "activebackground": "#374151",
                "activeforeground": "white",
                "relief": "flat",
                "borderwidth": 0,
                "font": ("Malgun Gothic", 10),
                "cursor": "hand2"
            },
            "success": {
                "bg": "#059669",
                "fg": "white",
                "activebackground": "#047857",
                "activeforeground": "white",
                "relief": "flat",
                "borderwidth": 0,
                "font": ("Malgun Gothic", 10, "bold"),
                "cursor": "hand2"
            },
            "danger": {
                "bg": "#dc2626",
                "fg": "white",
                "activebackground": "#b91c1c",
                "activeforeground": "white",
                "relief": "flat",
                "borderwidth": 0,
                "font": ("Malgun Gothic", 10),
                "cursor": "hand2"
            }
        }
        
        button_style = styles.get(style, styles["primary"])
        button_style.update(kwargs)
        
        super().__init__(parent, text=text, command=command, **button_style)
        
        # 호버 효과 속성 설정
        self.default_bg = button_style["bg"]
        self.hover_bg = button_style["activebackground"]
        
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
    
    def _on_enter(self, event):
        self.config(bg=self.hover_bg)
    
    def _on_leave(self, event):
        self.config(bg=self.default_bg)

class FileDropFrame(tk.Frame):
    """드래그 앤 드롭 지원 프레임"""
    def __init__(self, parent, label_text, on_files_dropped=None, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.on_files_dropped = on_files_dropped
        self.files = []
        
        # 스타일 설정
        self.config(
            bg="#f8fafc",
            relief="solid",
            borderwidth=2,
            highlightbackground="#e2e8f0",
            highlightcolor="#3b82f6",
            highlightthickness=1
        )
        
        # 라벨
        self.label = tk.Label(
            self,
            text=label_text,
            font=("Malgun Gothic", 12, "bold"),
            bg="#f8fafc",
            fg="#1f2937"
        )
        self.label.pack(pady=10)
        
        # 드래그 앤 드롭 안내
        self.drop_label = tk.Label(
            self,
            text="여기에 파일을 드래그하거나 클릭하여 선택하세요",
            font=("Malgun Gothic", 10),
            bg="#f8fafc",
            fg="#6b7280"
        )
        self.drop_label.pack(pady=5)
        
        # 파일 목록
        self.file_listbox = tk.Listbox(
            self,
            height=5,
            font=("Malgun Gothic", 9),
            bg="white",
            fg="#374151",
            selectbackground="#3b82f6",
            selectforeground="white",
            relief="flat",
            borderwidth=0
        )
        self.file_listbox.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 버튼 프레임
        button_frame = tk.Frame(self, bg="#f8fafc")
        button_frame.pack(fill="x", padx=10, pady=5)
        
        # 파일 선택 버튼
        self.select_btn = ModernButton(
            button_frame,
            "📁 파일 선택",
            command=self._select_files,
            style="secondary"
        )
        self.select_btn.pack(side="left", padx=5)
        
        # 초기화 버튼
        self.clear_btn = ModernButton(
            button_frame,
            "🗑️ 초기화",
            command=self._clear_files,
            style="danger"
        )
        self.clear_btn.pack(side="right", padx=5)
        
        # 드래그 앤 드롭 설정
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_drop)
        
        # 클릭 이벤트
        self.bind("<Button-1>", lambda e: self._select_files())
        self.drop_label.bind("<Button-1>", lambda e: self._select_files())
        
        # 호버 효과
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
    
    def _on_enter(self, event):
        self.config(highlightbackground="#3b82f6")
    
    def _on_leave(self, event):
        self.config(highlightbackground="#e2e8f0")
    
    def _on_drop(self, event):
        """드래그 앤 드롭 처리 (개선된 경로 처리)"""
        try:
            # 드롭된 데이터 파싱 (여러 형태 지원)
            drop_data = event.data
            
            # 다양한 형태의 경로 데이터 처리
            if isinstance(drop_data, (list, tuple)):
                file_paths = [str(path) for path in drop_data]
            else:
                # 문자열인 경우 공백으로 분할하되, 중괄호로 감싸진 경로 처리
                import re
                # {경로} 형태 또는 일반 경로 추출
                file_paths = re.findall(r'\{([^}]+)\}|([^\s]+)', str(drop_data))
                file_paths = [path[0] if path[0] else path[1] for path in file_paths]
            
            valid_files = []
            invalid_files = []
            
            for file_path in file_paths:
                # 경로 정리
                file_path = file_path.strip().strip('{}').strip('"').strip("'")
                
                if not file_path:
                    continue
                
                path_obj = Path(file_path)
                
                # 파일 존재 확인
                if not path_obj.exists():
                    invalid_files.append(f"{file_path} (파일 없음)")
                    continue
                
                # 디렉토리인 경우 스킵
                if path_obj.is_dir():
                    invalid_files.append(f"{file_path} (폴더는 지원하지 않음)")
                    continue
                
                # 파일 확장자 확인
                file_ext = path_obj.suffix.lower()
                if file_ext not in config.SUPPORTED_FORMATS:
                    invalid_files.append(f"{path_obj.name} (지원하지 않는 형식: {file_ext})")
                    continue
                
                # 파일 크기 확인
                try:
                    file_size_mb = path_obj.stat().st_size / (1024 * 1024)
                    if file_size_mb > config.MAX_FILE_SIZE_MB:
                        invalid_files.append(f"{path_obj.name} (파일 크기 초과: {file_size_mb:.1f}MB)")
                        continue
                except:
                    pass
                
                valid_files.append(str(path_obj.resolve()))
            
            # 결과 처리
            if valid_files:
                self.files.extend(valid_files)
                self._update_file_list()
                if self.on_files_dropped:
                    self.on_files_dropped(valid_files)
                
                # 성공 메시지
                if len(valid_files) == 1:
                    print(f"✅ 파일 추가됨: {Path(valid_files[0]).name}")
                else:
                    print(f"✅ {len(valid_files)}개 파일 추가됨")
            
            # 오류 파일이 있으면 경고 표시
            if invalid_files:
                error_msg = f"다음 파일들을 처리할 수 없습니다:\n\n"
                error_msg += "\n".join(invalid_files[:10])  # 최대 10개만 표시
                if len(invalid_files) > 10:
                    error_msg += f"\n... 외 {len(invalid_files) - 10}개"
                error_msg += f"\n\n지원 형식: {', '.join(config.SUPPORTED_FORMATS)}"
                error_msg += f"\n최대 파일 크기: {config.MAX_FILE_SIZE_MB}MB"
                messagebox.showwarning("파일 처리 경고", error_msg)
            
            # 모든 파일이 무효한 경우
            if not valid_files and not invalid_files:
                messagebox.showwarning("경고", "유효한 파일을 찾을 수 없습니다.")
        
        except Exception as e:
            print(f"❌ 드래그앤드롭 처리 오류: {e}")
            messagebox.showerror("오류", f"파일 드롭 처리 중 오류가 발생했습니다:\n{e}")
    
    def _select_files(self):
        """파일 선택 대화상자"""
        file_types = [
            ("모든 지원 파일", " ".join(f"*{ext}" for ext in config.SUPPORTED_FORMATS)),
            ("Excel 파일", "*.xlsx *.xls *.xlsm"),
            ("CSV 파일", "*.csv"),
            ("모든 파일", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="파일 선택",
            filetypes=file_types
        )
        
        if files:
            self.files.extend(files)
            self._update_file_list()
            if self.on_files_dropped:
                self.on_files_dropped(list(files))
    
    def _clear_files(self):
        """파일 목록 초기화"""
        self.files.clear()
        self._update_file_list()
    
    def _update_file_list(self):
        """파일 목록 업데이트"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.files:
            self.file_listbox.insert(tk.END, Path(file_path).name)
    
    def get_files(self):
        """선택된 파일 목록 반환"""
        return self.files.copy()

class MainWindow:
    """메인 윈도우 클래스"""
    
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.setup_window()
        self.setup_ui()
        
        # 콜백 함수들
        self.on_start_merge = None
        self.on_cancel_merge = None
        self.on_close = None
        
        # 상태 변수
        self.is_merging = False
        self.current_progress = 0
        self.current_status = "준비됨"
    
    def setup_window(self):
        """윈도우 기본 설정"""
        self.root.title("수창 엑셀 취합프로그램")
        self.root.geometry("800x700")
        self.root.minsize(600, 500)
        
        # 아이콘 설정 (선택적)
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # 배경색
        self.root.configure(bg="#ffffff")
        
        # 종료 이벤트
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def setup_ui(self):
        """UI 구성"""
        # 메인 컨테이너
        main_container = tk.Frame(self.root, bg="#ffffff")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 헤더
        self._setup_header(main_container)
        
        # 파일 선택 영역
        self._setup_file_selection(main_container)
        
        # 출력 경로 설정 영역
        self._setup_output_section(main_container)
        
        # 진행률 및 상태 영역
        self._setup_progress_section(main_container)
        
        # 버튼 영역
        self._setup_buttons(main_container)
        
        # 하단 정보
        self._setup_footer(main_container)
    
    def _setup_header(self, parent):
        """헤더 영역 설정"""
        header_frame = tk.Frame(parent, bg="#ffffff")
        header_frame.pack(fill="x", pady=(0, 20))
        
        # 프로그램 제목
        title_label = tk.Label(
            header_frame,
            text="수창 엑셀 취합프로그램",
            font=("Malgun Gothic", 24, "bold"),
            bg="#ffffff",
            fg="#1f2937"
        )
        title_label.pack(anchor="w")
        
        # 지원 형식 표시
        formats_text = f"지원 형식: {', '.join(config.SUPPORTED_FORMATS)}"
        formats_label = tk.Label(
            header_frame,
            text=formats_text,
            font=("Malgun Gothic", 10),
            bg="#ffffff",
            fg="#6b7280"
        )
        formats_label.pack(anchor="w", pady=(5, 0))
        
        # 구분선
        separator = tk.Frame(header_frame, height=2, bg="#e5e7eb")
        separator.pack(fill="x", pady=10)
    
    def _setup_file_selection(self, parent):
        """파일 선택 영역 설정"""
        file_frame = tk.Frame(parent, bg="#ffffff")
        file_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        # 좌우 분할
        left_frame = tk.Frame(file_frame, bg="#ffffff")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_frame = tk.Frame(file_frame, bg="#ffffff")
        right_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # 원본 파일 영역
        self.reference_frame = FileDropFrame(
            left_frame,
            "📄 원본 파일",
            on_files_dropped=self._on_reference_files_dropped,
            bg="#f8fafc",
            width=300,
            height=200
        )
        self.reference_frame.pack(fill="both", expand=True)
        
        # 취합 파일 영역
        self.target_frame = FileDropFrame(
            right_frame,
            "📊 취합 파일들",
            on_files_dropped=self._on_target_files_dropped,
            bg="#f8fafc",
            width=300,
            height=200
        )
        self.target_frame.pack(fill="both", expand=True)
    
    def _setup_output_section(self, parent):
        """출력 경로 설정 영역"""
        output_frame = tk.Frame(parent, bg="#ffffff")
        output_frame.pack(fill="x", pady=(0, 20))
        
        # 라벨
        output_label = tk.Label(
            output_frame,
            text="💾 출력 파일 설정",
            font=("Malgun Gothic", 12, "bold"),
            bg="#ffffff",
            fg="#1f2937"
        )
        output_label.pack(anchor="w", pady=(0, 5))
        
        # 경로 선택 프레임
        path_frame = tk.Frame(output_frame, bg="#ffffff")
        path_frame.pack(fill="x", pady=(0, 5))
        
        # 출력 경로 표시
        self.output_path_var = tk.StringVar()
        self.output_path_var.set(os.getcwd())  # 기본값: 현재 디렉토리
        
        self.output_path_entry = tk.Entry(
            path_frame,
            textvariable=self.output_path_var,
            font=("Malgun Gothic", 10),
            bg="#f8fafc",
            fg="#374151",
            relief="solid",
            borderwidth=1
        )
        self.output_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # 경로 선택 버튼
        self.browse_btn = ModernButton(
            path_frame,
            "📁 경로 선택",
            command=self._browse_output_path,
            style="secondary"
        )
        self.browse_btn.pack(side="right")
        
        # 파일명 설정 프레임
        filename_frame = tk.Frame(output_frame, bg="#ffffff")
        filename_frame.pack(fill="x")
        
        # 파일명 라벨
        filename_label = tk.Label(
            filename_frame,
            text="파일명:",
            font=("Malgun Gothic", 10),
            bg="#ffffff",
            fg="#6b7280"
        )
        filename_label.pack(side="left", padx=(0, 5))
        
        # 파일명 입력
        self.output_filename_var = tk.StringVar()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_filename_var.set(f"통합_데이터_{timestamp}.xlsx")
        
        self.output_filename_entry = tk.Entry(
            filename_frame,
            textvariable=self.output_filename_var,
            font=("Malgun Gothic", 10),
            bg="#f8fafc",
            fg="#374151",
            relief="solid",
            borderwidth=1,
            width=30
        )
        self.output_filename_entry.pack(side="left", padx=(0, 10))
        
        # 자동 이름 생성 버튼
        self.auto_name_btn = ModernButton(
            filename_frame,
            "🔄 자동 이름",
            command=self._generate_auto_name,
            style="secondary"
        )
        self.auto_name_btn.pack(side="right")
    
    def _browse_output_path(self):
        """출력 경로 선택"""
        from tkinter import filedialog
        
        directory = filedialog.askdirectory(
            title="출력 폴더 선택",
            initialdir=self.output_path_var.get()
        )
        
        if directory:
            self.output_path_var.set(directory)
            print(f"📁 출력 경로 설정: {directory}")
    
    def _generate_auto_name(self):
        """자동 파일명 생성"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        auto_name = f"통합_데이터_{timestamp}.xlsx"
        self.output_filename_var.set(auto_name)
        print(f"📝 파일명 자동 생성: {auto_name}")
    
    def get_output_file_path(self):
        """완전한 출력 파일 경로 반환"""
        output_dir = self.output_path_var.get()
        filename = self.output_filename_var.get()
        
        # .xlsx 확장자가 없으면 추가
        if not filename.lower().endswith('.xlsx'):
            filename += '.xlsx'
        
        return os.path.join(output_dir, filename)

    def _setup_progress_section(self, parent):
        """진행률 및 상태 영역 설정"""
        progress_frame = tk.Frame(parent, bg="#ffffff")
        progress_frame.pack(fill="x", pady=(0, 20))
        
        # 상태 라벨
        self.status_label = tk.Label(
            progress_frame,
            text="준비됨",
            font=("Malgun Gothic", 12),
            bg="#ffffff",
            fg="#374151"
        )
        self.status_label.pack(anchor="w", pady=(0, 5))
        
        # 진행률 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill="x", pady=(0, 5))
        
        # 진행률 텍스트
        self.progress_label = tk.Label(
            progress_frame,
            text="0%",
            font=("Malgun Gothic", 10),
            bg="#ffffff",
            fg="#6b7280"
        )
        self.progress_label.pack(anchor="w")
        
        # 스타일 설정
        style = ttk.Style()
        style.configure(
            "Custom.Horizontal.TProgressbar",
            background="#3b82f6",
            troughcolor="#e5e7eb",
            borderwidth=1,
            lightcolor="#3b82f6",
            darkcolor="#3b82f6"
        )
    
    def _setup_buttons(self, parent):
        """버튼 영역 설정"""
        button_frame = tk.Frame(parent, bg="#ffffff")
        button_frame.pack(fill="x", pady=(0, 20))
        
        # 시작 버튼
        self.start_button = ModernButton(
            button_frame,
            "🚀 취합 시작",
            command=self._on_start_clicked,
            style="success"
        )
        self.start_button.pack(side="left", padx=(0, 10))
        
        # 중지 버튼
        self.cancel_button = ModernButton(
            button_frame,
            "⏹️ 중지",
            command=self._on_cancel_clicked,
            style="danger"
        )
        self.cancel_button.pack(side="left", padx=10)
        self.cancel_button.config(state="disabled")
        
        # 종료 버튼
        self.exit_button = ModernButton(
            button_frame,
            "❌ 종료",
            command=self._on_close,
            style="secondary"
        )
        self.exit_button.pack(side="right")
    
    def _setup_footer(self, parent):
        """하단 정보 영역 설정"""
        footer_frame = tk.Frame(parent, bg="#ffffff")
        footer_frame.pack(fill="x")
        
        # 구분선
        separator = tk.Frame(footer_frame, height=1, bg="#e5e7eb")
        separator.pack(fill="x", pady=(0, 10))
        
        # 정보 텍스트
        info_text = (
            "💡 사용법: 좌측에 원본 파일 1개, 우측에 취합할 파일들을 추가한 후 '취합 시작'을 클릭하세요.\n"
            "📝 자동으로 시트명과 열제목을 매칭하여 데이터를 통합합니다."
        )
        info_label = tk.Label(
            footer_frame,
            text=info_text,
            font=("Malgun Gothic", 9),
            bg="#ffffff",
            fg="#6b7280",
            justify="left"
        )
        info_label.pack(anchor="w")
    
    def _on_reference_files_dropped(self, files):
        """원본 파일 드롭 처리"""
        if len(files) > 1:
            messagebox.showwarning("경고", "원본 파일은 1개만 선택할 수 있습니다.")
            self.reference_frame.files = [files[0]]
            self.reference_frame._update_file_list()
    
    def _on_target_files_dropped(self, files):
        """취합 파일 드롭 처리"""
        pass  # 여러 파일 허용
    
    def _on_start_clicked(self):
        """취합 시작 버튼 클릭"""
        reference_files = self.reference_frame.get_files()
        target_files = self.target_frame.get_files()
        
        if not reference_files:
            messagebox.showerror("오류", "원본 파일을 선택해주세요.")
            return
        
        if not target_files:
            messagebox.showerror("오류", "취합할 파일들을 선택해주세요.")
            return
        
        if len(reference_files) > 1:
            messagebox.showerror("오류", "원본 파일은 1개만 선택할 수 있습니다.")
            return
        
        # 출력 파일 경로 확인
        output_file_path = self.get_output_file_path()
        
        # 출력 디렉토리 존재 확인
        output_dir = os.path.dirname(output_file_path)
        if not os.path.exists(output_dir):
            if messagebox.askyesno("확인", f"출력 폴더가 존재하지 않습니다.\n폴더를 생성하시겠습니까?\n\n{output_dir}"):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                    print(f"📁 출력 폴더 생성: {output_dir}")
                except Exception as e:
                    messagebox.showerror("오류", f"폴더 생성 실패:\n{e}")
                    return
            else:
                return
        
        # 파일 덮어쓰기 확인
        if os.path.exists(output_file_path):
            if not messagebox.askyesno("확인", f"파일이 이미 존재합니다.\n덮어쓰시겠습니까?\n\n{os.path.basename(output_file_path)}"):
                return
        
        print(f"🚀 취합 시작!")
        print(f"📁 원본 파일: {Path(reference_files[0]).name}")
        print(f"📊 취합 파일: {len(target_files)}개")
        print(f"💾 출력 파일: {output_file_path}")
        
        # 콜백 호출 (출력 경로 포함)
        if self.on_start_merge:
            self.on_start_merge(reference_files[0], target_files, output_file_path)
    
    def _on_cancel_clicked(self):
        """중지 버튼 클릭"""
        if self.on_cancel_merge:
            self.on_cancel_merge()
    
    def _on_close(self):
        """종료 처리"""
        try:
            if self.is_merging:
                if messagebox.askyesno("확인", "작업이 진행 중입니다. 정말 종료하시겠습니까?"):
                    if self.on_close:
                        self.on_close()
                    self.root.quit()
                    self.root.destroy()
            else:
                if self.on_close:
                    self.on_close()
                self.root.quit()
                self.root.destroy()
        except Exception as e:
            print(f"종료 중 오류: {e}")
            self.root.quit()
            self.root.destroy()
    
    def set_callbacks(self, on_start_merge=None, on_cancel_merge=None, on_close=None):
        """콜백 함수 설정"""
        self.on_start_merge = on_start_merge
        self.on_cancel_merge = on_cancel_merge
        self.on_close = on_close
    
    def update_progress(self, progress: float, status: str = None):
        """진행률 업데이트"""
        self.current_progress = max(0, min(100, progress))
        self.progress_var.set(self.current_progress)
        self.progress_label.config(text=f"{self.current_progress:.1f}%")
        
        if status:
            self.current_status = status
            self.status_label.config(text=status)
        
        self.root.update_idletasks()
    
    def set_merge_state(self, is_merging: bool):
        """병합 상태 설정"""
        self.is_merging = is_merging
        
        if is_merging:
            self.start_button.config(state="disabled")
            self.cancel_button.config(state="normal")
            self.reference_frame.select_btn.config(state="disabled")
            self.reference_frame.clear_btn.config(state="disabled")
            self.target_frame.select_btn.config(state="disabled")
            self.target_frame.clear_btn.config(state="disabled")
        else:
            self.start_button.config(state="normal")
            self.cancel_button.config(state="disabled")
            self.reference_frame.select_btn.config(state="normal")
            self.reference_frame.clear_btn.config(state="normal")
            self.target_frame.select_btn.config(state="normal")
            self.target_frame.clear_btn.config(state="normal")
    
    def show_error(self, message: str):
        """오류 메시지 표시"""
        messagebox.showerror("오류", message)
    
    def show_success(self, message: str):
        """성공 메시지 표시"""
        messagebox.showinfo("완료", message)
    
    def show_warning(self, message: str):
        """경고 메시지 표시"""
        messagebox.showwarning("경고", message)
    
    def run(self):
        """앱 실행"""
        self.root.mainloop()