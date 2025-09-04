import customtkinter as ctk
import threading
import json
import os
import traceback
from tkinter import filedialog, messagebox
from hwp_assistant import HWPAssistant  # 기존 클래스


class ErrorHandler:
    """통합 에러 처리 클래스"""
    
    @staticmethod
    def handle_error(func, error_callback=None):
        """데코레이터 패턴으로 에러 처리"""
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                error_msg = f"오류 발생: {str(e)}"
                print(f"[ERROR] {error_msg}\n{traceback.format_exc()}")
                if error_callback:
                    error_callback(error_msg)
                else:
                    messagebox.showerror("오류", error_msg)
                return None
        return wrapper

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 기본 설정
        self.title("HWP AI 어시스턴트 v3.0 통합 GUI")
        self.geometry("800x700")
        self.assistant = HWPAssistant()
        self.current_file = ""
        
        # GUI 초기화
        self._setup_gui()
        self._load_styles()
        

    def _setup_gui(self):
        """메인 GUI 레이아웃 구성 - 버튼 변수 할당 수정"""
        
        # === 제목 ===
        title_label = ctk.CTkLabel(self, text="🤖 HWP AI 어시스턴트", 
                                font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=20)
        
        # === 파일 관리 섹션 ===
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(file_frame, text="📄 파일 관리", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)
        
        file_buttons = ctk.CTkFrame(file_frame)
        file_buttons.pack(fill="x", padx=10, pady=5)
        
        self.open_button = ctk.CTkButton(file_buttons, text="파일 열기", command=self._open_file)
        self.open_button.pack(side="left", padx=5)
        
        self.close_button = ctk.CTkButton(file_buttons, text="파일 닫기", command=self._close_file)
        self.close_button.pack(side="left", padx=5)
        
        self.file_status = ctk.CTkLabel(file_buttons, text="파일이 열리지 않음")
        self.file_status.pack(side="right", padx=10)
        
        # === 텍스트 수정 섹션 ===
        text_frame = ctk.CTkFrame(self)
        text_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(text_frame, text="✏️ AI 텍스트 수정", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)
        
        self.request_entry = ctk.CTkEntry(text_frame, 
                                        placeholder_text="수정 요청을 입력하세요 (예: 더 친근하게 바꿔줘)")
        self.request_entry.pack(fill="x", padx=10, pady=5)
        
        self.context_entry = ctk.CTkEntry(text_frame, 
                                        placeholder_text="스타일 파일 (선택사항, 예: @style.md)")
        self.context_entry.pack(fill="x", padx=10, pady=5)
        
        # ✨ 핵심 수정 ✨
        self.modify_button = ctk.CTkButton(text_frame, text="선택된 텍스트 수정", 
                                        command=self._modify_selected_text)
        self.modify_button.pack(pady=5)
        
        # === 표 생성 섹션 ===
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(table_frame, text="📊 표 생성", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)
        
        # ✨ 핵심 수정 ✨
        self.table_button = ctk.CTkButton(table_frame, text="선택된 텍스트를 표로 변환", 
                                        command=self._create_table)
        self.table_button.pack(pady=5)
        

        # === 스타일` 및 서식 섹션 ===
        style_frame = ctk.CTkFrame(self)
        style_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(style_frame, text="🎨 스타일 및 서식", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)

        style_controls = ctk.CTkFrame(style_frame)
        style_controls.pack(fill="x", padx=10, pady=5)

        self.style_combo = ctk.CTkComboBox(
            style_controls, 
            values=["스타일 없음"],
            width=200
        )
        self.style_combo.pack(side="left", padx=5)

        # 새로고침 버튼 추가
        ctk.CTkButton(style_controls, text="🔄", width=30, 
                    command=self._load_styles).pack(side="left", padx=5)

        self.apply_style_button = ctk.CTkButton(style_controls, text="선택 영역에 적용", 
                                            command=self._apply_style)
        self.apply_style_button.pack(side="left", padx=5)

        ctk.CTkButton(style_controls, text="🤖 AI 스마트 적용", 
              command=self._open_smart_style_window).pack(side="left", padx=5)

        # === 템플릿 관리 섹션 ===
        template_frame = ctk.CTkFrame(self)
        template_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(template_frame, text="🏗️ 템플릿 관리", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)
        
        template_buttons = ctk.CTkFrame(template_frame)
        template_buttons.pack(fill="x", padx=10, pady=5)
        
        self.template_create_button = ctk.CTkButton(template_buttons, text="템플릿 생성", 
                                                command=self._open_template_creation)
        self.template_create_button.pack(side="left", padx=5)
        
        self.template_use_button = ctk.CTkButton(template_buttons, text="템플릿 사용", 
                                                command=self._open_template_usage)
        self.template_use_button.pack(side="left", padx=5)
        
        # === 로그 출력 섹션 ===
        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        ctk.CTkLabel(log_frame, text="📋 작업 로그", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=5)
        
        self.log_textbox = ctk.CTkTextbox(log_frame, height=150)
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=5)

    def log(self, message):
        """로그 메시지 추가"""
        self.log_textbox.insert("end", f"{message}\n")
        self.log_textbox.see("end")
        self.update_idletasks()
        
    def _run_in_thread(self, target_func):
        """GUI 블로킹 방지를 위한 스레드 실행"""
        thread = threading.Thread(target=target_func)
        thread.daemon = True
        thread.start()

    @ErrorHandler.handle_error
    def _open_file(self):
        """파일 열기 - 메인 스레드 직접 실행"""
        file_path = filedialog.askopenfilename(
            title="HWP 파일 선택",
            filetypes=[("한글 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")]
        )
        
        if file_path:
            try:
                self._show_progress("📂 파일을 열고 있습니다...")
                
                if self.assistant.open_file(file_path):
                    self.current_file = file_path
                    filename = os.path.basename(file_path)
                    self.file_status.configure(text=f"열림: {filename}")
                    self.log(f"✅ 파일 열기 성공: {filename}")
                else:
                    self.log("❌ 파일 열기 실패")
            except Exception as e:
                self.log(f"❌ 파일 열기 오류: {e}")

    @ErrorHandler.handle_error    
    def _close_file(self):
        """파일 닫기 - 메인 스레드 직접 실행"""
        try:
            if self.assistant.is_opened:
                self.assistant.close_file()
                self.current_file = ""
                self.file_status.configure(text="파일이 열리지 않음")
                self.log("📁 파일이 닫혔습니다")
            else:
                self.log("⚠️ 열린 파일이 없습니다")
        except Exception as e:
            self.log(f"❌ 파일 닫기 오류: {e}")

    def _show_progress(self, message):
        """진행 상황 표시"""
        self.log(message)
        self.update()  # GUI 즉시 업데이트

    def _modify_selected_text(self):
        """✨ 완전히 메인 스레드에서 실행되는 텍스트 수정"""
        if not self.assistant.is_opened:
            self.log("⚠️ 먼저 파일을 열어주세요")
            return
            
        request = self.request_entry.get().strip()
        if not request:
            self.log("⚠️ 수정 요청을 입력해주세요")
            return

        # 버튼 비활성화 (중복 실행 방지)
        self.modify_button.configure(state="disabled", text="처리 중...")
        
        try:
            # 1단계: 선택된 텍스트 가져오기 (메인 스레드)
            self._show_progress("📌 선택된 텍스트를 가져오는 중...")
            selected_text = self.assistant.get_selected_text()
            
            if not selected_text:
                self.log("⚠️ 텍스트를 선택해주세요")
                return
                
            self._show_progress(f"📝 선택된 텍스트: '{selected_text[:50]}...'")
            
            # 2단계: AI 처리 (메인 스레드에서 블로킹 실행)
            context = self.context_entry.get().strip()
            full_request = f"{request} {context}".strip()
            
            self._show_progress("🤖 AI가 텍스트를 분석하고 있습니다... (잠시만 기다려주세요)")
            
            # ✨ 핵심: 메인 스레드에서 블로킹 실행 (스레드 사용 안 함)
            modified_text = self.assistant.call_gemini(full_request, selected_text)
            
            if not modified_text:
                self.log("❌ AI 수정 실패")
                return
                
            self._show_progress("✅ AI 수정 완료")
            
            # 3단계: 결과 확인 및 적용
            self._show_modification_result(modified_text, selected_text)
            
        except Exception as e:
            self.log(f"❌ 처리 오류: {e}")
        finally:
            # 버튼 재활성화
            self.modify_button.configure(state="normal", text="선택된 텍스트 수정")

    def _show_modification_result(self, modified_text, original_text):
        """수정 결과 확인 창"""
        result_window = ctk.CTkToplevel(self)
        result_window.title("수정 결과 확인")
        result_window.geometry("700x600")
        result_window.grab_set()
        
        # 원본 텍스트
        ctk.CTkLabel(result_window, text="📌 원본 텍스트:", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        original_box = ctk.CTkTextbox(result_window, height=120)
        original_box.pack(fill="x", padx=20, pady=5)
        original_box.insert("0.0", original_text)
        original_box.configure(state="disabled")
        
        # 수정 결과
        ctk.CTkLabel(result_window, text="✨ AI 수정 결과 (편집 가능):", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        result_box = ctk.CTkTextbox(result_window, height=200)
        result_box.pack(fill="x", padx=20, pady=5)
        result_box.insert("0.0", modified_text)
        
        # 버튼
        button_frame = ctk.CTkFrame(result_window)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        def apply_changes():
            """✨ 메인 스레드에서 직접 텍스트 교체"""
            try:
                final_text = result_box.get("0.0", "end-1c")
                
                self.log("🔄 텍스트를 교체하고 있습니다...")
                
                # COM 작업을 메인 스레드에서 직접 실행
                if self.assistant.replace_selected_text(final_text):
                    self.log("✅ 텍스트 교체 성공!")
                    result_window.destroy()
                else:
                    self.log("❌ 텍스트 교체 실패")
            except Exception as e:
                self.log(f"❌ 교체 오류: {e}")
                
        def cancel_changes():
            self.log("❌ 텍스트 교체 취소")
            result_window.destroy()
        
        ctk.CTkButton(button_frame, text="✅ 적용", 
                     command=apply_changes, width=120).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="❌ 취소", 
                     command=cancel_changes, width=120).pack(side="right", padx=10)

    def _create_table(self):
        """표 생성 - 완전히 메인 스레드에서 실행"""
        if not self.assistant.is_opened:
            self.log("⚠️ 먼저 파일을 열어주세요")
            return
            
        # 버튼 비활성화
        self.table_button.configure(state="disabled", text="표 생성 중...")
        
        try:
            # 1단계: 선택된 텍스트 가져오기
            self._show_progress("📌 선택된 텍스트를 가져오는 중...")
            selected_text = self.assistant.get_selected_text()
            
            if not selected_text:
                self.log("⚠️ 표로 만들 텍스트를 선택해주세요")
                return
                
            self._show_progress(f"📝 선택된 텍스트: '{selected_text[:50]}...'")
            
            # 2단계: AI 처리 (메인 스레드에서 블로킹)
            self._show_progress("🤖 AI가 표를 생성하고 있습니다...")
            modified_text = self.assistant.call_gemini("이 내용을 표로 만들어줘", selected_text)
            
            if not (modified_text and modified_text.strip().startswith('|')):
                self.log("❌ 표 형식 생성 실패")
                return
                
            # 3단계: 표 삽입 (메인 스레드)
            self._show_progress("📊 문서에 표를 삽입하고 있습니다...")
            self.assistant.move_caret_right()  # 커서 이동
            
            if self.assistant.insert_table(modified_text):
                self.log("✅ 표 삽입 성공!")
            else:
                self.log("❌ 표 삽입 실패")
                
        except Exception as e:
            self.log(f"❌ 표 생성 오류: {e}")
        finally:
            # 버튼 재활성화
            self.table_button.configure(state="normal", text="선택된 텍스트를 표로 변환")
            
    def _open_template_creation(self):
        """템플릿 생성 윈도우 열기"""
        if not self.assistant.is_opened:
            self.log("⚠️ 먼저 템플릿으로 만들 파일을 열어주세요")
            return
            
        TemplateCreationWindow(self, self.assistant)

    def _open_template_usage(self):
        """템플릿 사용 윈도우 열기"""
        TemplateUsageWindow(self, self.assistant)

    def _load_styles(self):
        """사용 가능한 스타일 목록을 콤보박스에 로드합니다."""
        try:
            styles = self.assistant.get_style_list()
            if styles:
                self.style_combo.configure(values=styles)
                self.style_combo.set(styles[0])
                self.log(f"✅ {len(styles)}개의 스타일 로드 완료")
            else:
                self.style_combo.configure(values=["스타일 없음"])
                self.style_combo.set("스타일 없음")
        except Exception as e:
            self.log(f"❌ 스타일 로딩 오류: {e}")

    def _apply_style(self):
        """선택된 스타일을 적용합니다."""
        if not self.assistant.is_opened:
            self.log("⚠️ 먼저 파일을 열어주세요.")
            return

        style_name = self.style_combo.get()
        if style_name == "스타일 없음":
            self.log("⚠️ 적용할 스타일을 선택하세요.")
            return
            
        self.apply_style_button.configure(state="disabled", text="적용 중...")
        
        try:
            self.log(f"🎨 스타일 '{style_name}'을(를) 적용합니다...")
            
            # JSON 파일 읽기
            style_path = os.path.join(os.getcwd(), "styles", f"{style_name}.json")
            with open(style_path, 'r', encoding='utf-8') as f:
                style_data = json.load(f)
                
            # 백엔드 메서드 호출 (메인 스레드에서 직접 실행)
            if self.assistant.apply_style_to_selection(style_data):
                self.log("✅ 스타일 적용 성공!")
            else:
                self.log("❌ 스타일 적용 실패")

        except Exception as e:
            self.log(f"❌ 스타일 적용 오류: {e}")
        finally:
            self.apply_style_button.configure(state="normal", text="선택된 영역에 적용")

    def _open_smart_style_window(self):
        """스마트 스타일 적용 윈도우 열기"""
        if not self.assistant.is_opened:
            self.log("⚠️ 먼저 문서를 열어주세요.")
            return
        SmartStyleWindow(self, self.assistant)


class TemplateCreationWindow(ctk.CTkToplevel):
    def __init__(self, parent, assistant):
        super().__init__(parent)
        
        self.parent = parent
        self.assistant = assistant
        self.template_fields = []
        
        self.title("템플릿 생성")
        self.geometry("700x600")
        self.grab_set()
        
        self._setup_gui()
        
        # ✨ 핵심 수정: 메인 스레드에서 직접 분석 실행
        self.after(100, self._analyze_document_main_thread)
        
    def _setup_gui(self):
        """템플릿 생성 GUI 구성"""
        ctk.CTkLabel(self, text="🏗️ 템플릿 생성", 
                    font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        
        # 템플릿 이름 입력
        name_frame = ctk.CTkFrame(self)
        name_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(name_frame, text="템플릿 이름:").pack(side="left", padx=10)
        self.name_entry = ctk.CTkEntry(name_frame, placeholder_text="예: 내부공문")
        self.name_entry.pack(side="right", expand=True, fill="x", padx=10)
        
        # 진행 상황 표시
        self.progress_label = ctk.CTkLabel(self, text="🔄 문서를 분석하고 있습니다...")
        self.progress_label.pack(pady=10)
        
        # 필드 선택 영역
        self.fields_frame = ctk.CTkScrollableFrame(self, label_text="🔧 템플릿 필드 선택")
        self.fields_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 버튼 영역
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        self.create_button = ctk.CTkButton(button_frame, text="템플릿 생성", 
                                         command=self._create_template_main_thread, 
                                         state="disabled")
        self.create_button.pack(side="left", padx=10)
        
        ctk.CTkButton(button_frame, text="취소", 
                     command=self.destroy).pack(side="right", padx=10)

    def _show_progress(self, message):
        """진행 상황 업데이트"""
        self.progress_label.configure(text=message)
        self.update()

    def _analyze_document_main_thread(self):
        """✨ 강화된 디버깅과 함께 문서 분석 실행"""
        try:
            self._show_progress("📄 문서 구조를 분석하고 있습니다...")
            
            # 1단계: 문서 분석
            structure = self.assistant.analyze_document_for_template()
            if not structure:
                self._show_error("문서 분석 실패")
                return
                
            self._show_progress("🤖 AI가 템플릿 필드를 분석하고 있습니다...")
            
            # 2단계: Gemini 분석
            analysis_request = "이 문서를 분석하여 템플릿으로 만들 변수들을 제안해줘."
            template_plan_str = self.assistant.call_gemini(
                analysis_request, 
                json.dumps(structure, ensure_ascii=False, indent=2), 
                mode="template_analysis"
            )
            
            # ✨ 핵심 수정: 응답 디버깅 및 강화된 처리
            if not template_plan_str:
                self._show_error("AI가 빈 응답을 반환했습니다")
                return
                
            # 디버깅을 위해 원본 응답 로그
            self.parent.log(f"📋 Gemini 원본 응답: {template_plan_str[:200]}...")
            
            # 3단계: 강화된 JSON 추출
            try:
                clean_json = self._robust_extract_json(template_plan_str)
                if not clean_json:
                    self._show_error("JSON 추출 실패: 유효한 JSON을 찾을 수 없습니다")
                    return
                    
                # 추출된 JSON 로그
                self.parent.log(f"🔧 추출된 JSON: {clean_json[:200]}...")
                
                template_plan = json.loads(clean_json)
                fields = template_plan.get("template_fields", [])
                
                if fields:
                    self._show_progress("✅ 분석 완료! 템플릿 필드를 선택하세요.")
                    self._display_fields(fields)
                    self.create_button.configure(state="normal")
                else:
                    self._show_error("AI가 템플릿 필드를 생성하지 못했습니다")
                    
            except json.JSONDecodeError as e:
                error_msg = f"JSON 파싱 오류: {str(e)}"
                self.parent.log(f"❌ 파싱 실패한 텍스트: '{clean_json}'")
                self._show_error(error_msg)
            except Exception as e:
                error_msg = f"결과 처리 오류: {str(e)}"
                self._show_error(error_msg)
                    
        except Exception as e:
            error_msg = str(e)
            self._show_error(f"분석 오류: {error_msg}")

    def _robust_extract_json(self, text):
        """✨ 강화된 JSON 추출 함수"""
        if not text or not text.strip():
            return ""
        
        text = text.strip()
        
        # 방법 1: 마크다운 코드 블록 제거 (``````)
        import re
        json_match = re.search(r'``````', text, re.DOTALL)
        if json_match:
            extracted = json_match.group(1).strip()
            if extracted and (extracted.startswith('{') or extracted.startswith('[')):
                return extracted
        
        # 방법 2: 단순 코드 블록 제거 (``````)
        if text.startswith('``````'):
            extracted = text[3:-3].strip()
            # 첫 줄이 'json'인 경우 제거
            lines = extracted.split('\n')
            if lines and lines[0].strip().lower() in ['json', 'javascript']:
                extracted = '\n'.join(lines[1:]).strip()
            if extracted and (extracted.startswith('{') or extracted.startswith('[')):
                return extracted
        
        # 방법 3: JSON 객체/배열 직접 추출
        json_start = -1
        for i, char in enumerate(text):
            if char in ['{', '[']:
                json_start = i
                break
        
        if json_start >= 0:
            # 마지막 } 또는 ] 찾기
            json_end = -1
            bracket_count = 0
            start_char = text[json_start]
            end_char = '}' if start_char == '{' else ']'
            
            for i in range(json_start, len(text)):
                if text[i] == start_char:
                    bracket_count += 1
                elif text[i] == end_char:
                    bracket_count -= 1
                    if bracket_count == 0:
                        json_end = i
                        break
            
            if json_end > json_start:
                return text[json_start:json_end + 1]
        
        # 방법 4: 원본 그대로 반환 (JSON인 경우)
        if text.startswith('{') or text.startswith('['):
            return text
        
        return ""

    def _extract_json_from_markdown(self, text):
        """기존 함수를 강화된 버전으로 대체"""
        return self._robust_extract_json(text)

    def _create_template_main_thread(self):
        """✨ 메인 스레드에서 템플릿 생성 실행"""
        template_name = self.name_entry.get().strip()
        if not template_name:
            self._show_error("템플릿 이름을 입력하세요")
            return
            
        selected_fields = [
            item['field'] for item in self.template_fields 
            if item['selected'].get()
        ]
        
        if not selected_fields:
            self._show_error("하나 이상의 필드를 선택하세요")
            return

        # ✨ 문서가 열려있는지 확인
        if not self.assistant.is_opened:
            self._show_error("템플릿으로 만들 문서가 열려있지 않습니다.")
            return
            
        # ✨ HWP 객체 상태 확인
        if not self.assistant.hwp:
            self._show_error("HWP 객체가 초기화되지 않았습니다.")
            return
            
        # 버튼 비활성화
        self.create_button.configure(state="disabled", text="생성 중...")
        
        try:
            self._show_progress("🔄 누름틀을 생성하고 있습니다...")
            
            # 모든 COM 작업을 메인 스레드에서 실행
            success_count = 0
            total_fields = len(selected_fields)
            
            for i, field in enumerate(selected_fields):
                self._show_progress(f"🔄 누름틀 생성 중... ({i+1}/{total_fields})")
                
                if self.assistant.convert_text_to_field(
                    field['original_text'], 
                    field['field_name']
                ):
                    success_count += 1
                    
            if success_count > 0:
                self._show_progress("💾 템플릿을 저장하고 있습니다...")
                
                if self.assistant.create_template_from_current(template_name):
                    success_msg = f"템플릿 '{template_name}' 생성 완료!"
                    self._show_success(success_msg)
                    self.destroy()  # 성공 시 창 닫기
                else:
                    self._show_error("템플릿 저장 실패")
            else:
                self._show_error("필드 생성 실패")
                
        except Exception as e:
            error_msg = str(e)
            self._show_error(f"생성 오류: {error_msg}")
        finally:
            # 버튼 재활성화
            self.create_button.configure(state="normal", text="템플릿 생성")

    def _extract_json_from_markdown(self, text):
        """마크다운 코드 블록에서 JSON 추출"""
        text = text.strip()
        if text.startswith('``````'):
            return text[3:-3].strip()
        return text
        
    def _display_fields(self, fields):
        """동적으로 필드 체크박스 생성"""
        for field in fields:
            field_frame = ctk.CTkFrame(self.fields_frame)
            field_frame.pack(fill="x", padx=5, pady=5)
            
            var = ctk.BooleanVar(value=True)
            checkbox = ctk.CTkCheckBox(
                field_frame, 
                text=f"{field.get('field_name', 'unknown')}",
                variable=var
            )
            checkbox.pack(side="left", padx=10)
            
            desc_label = ctk.CTkLabel(
                field_frame, 
                text=f"원본: '{field.get('original_text', '')[:50]}...'"
            )
            desc_label.pack(side="right", padx=10)
            
            self.template_fields.append({
                'field': field,
                'selected': var
            })
            
    def _show_error(self, message):
        """에러 표시"""
        self._show_progress(f"❌ {message}")
        messagebox.showerror("오류", message)
        self.parent.log(f"❌ {message}")
        
    def _show_success(self, message):
        """성공 표시"""
        messagebox.showinfo("성공", message)
        self.parent.log(f"✅ {message}")


class TemplateUsageWindow(ctk.CTkToplevel):
    def __init__(self, parent, assistant):
        super().__init__(parent)
        
        self.parent = parent
        self.assistant = assistant
        self.field_entries = {}
        
        self.title("템플릿 사용")
        self.geometry("600x500")
        self.grab_set()
        
        self._setup_gui()
        self._load_templates()
        
    def _setup_gui(self):
        """템플릿 사용 GUI 구성"""
        ctk.CTkLabel(self, text="📄 템플릿 사용", 
                    font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        
        # 템플릿 선택
        select_frame = ctk.CTkFrame(self)
        select_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(select_frame, text="템플릿:").pack(side="left", padx=10)
        self.template_combo = ctk.CTkComboBox(
            select_frame, 
            values=["템플릿을 로딩 중..."],
            command=self._on_template_selected
        )
        self.template_combo.pack(side="right", expand=True, fill="x", padx=10)
        
        # 동적 필드 입력 영역
        self.fields_frame = ctk.CTkScrollableFrame(self, label_text="📝 필드 입력")
        self.fields_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 버튼
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkButton(button_frame, text="문서 생성", 
                     command=self._create_document).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="취소", 
                     command=self.destroy).pack(side="right", padx=10)
        
    def _load_templates(self):
        """템플릿 목록 로드"""
        try:
            template_dir = os.path.join(os.getcwd(), "templates")
            if os.path.exists(template_dir):
                templates = [f[:-4] for f in os.listdir(template_dir) if f.endswith('.hwp')]
                if templates:
                    self.template_combo.configure(values=templates)
                    self.template_combo.set(templates[0])
                    # ✨ 첫 번째 템플릿 필드 자동 로드
                    self._on_template_selected(templates[0])
                else:
                    self.template_combo.configure(values=["템플릿 없음"])
            else:
                self.template_combo.configure(values=["템플릿 폴더 없음"])
        except Exception as e:
            self.template_combo.configure(values=[f"오류: {e}"])
            
    def _on_template_selected(self, template_name):
        """✨ 템플릿 선택 시 실제 필드 로드"""
        # 기존 필드 제거
        for widget in self.fields_frame.winfo_children():
            widget.destroy()
        self.field_entries.clear()
        
        try:
            # 템플릿 파일에서 필드 목록 가져오기
            fields = self.assistant.get_field_list_from_file(template_name)
            
            if not fields:
                ctk.CTkLabel(self.fields_frame, text="템플릿에서 필드를 찾을 수 없습니다.").pack()
                return

            # 동적으로 필드 입력 위젯 생성
            for field_name in fields:
                field_frame = ctk.CTkFrame(self.fields_frame)
                field_frame.pack(fill="x", padx=5, pady=5)
                
                label = ctk.CTkLabel(field_frame, text=field_name, width=150)
                label.pack(side="left", padx=10)
                
                entry = ctk.CTkEntry(field_frame, placeholder_text=f"{field_name} 입력")
                entry.pack(side="right", expand=True, fill="x", padx=10)
                
                self.field_entries[field_name] = entry
        
        except Exception as e:
            ctk.CTkLabel(self.fields_frame, text=f"필드 로딩 오류: {e}").pack()

            
    def _create_document(self):
        """템플릿으로 문서 생성"""
        template_name = self.template_combo.get()
        if not template_name or template_name in ["템플릿 없음", "템플릿 폴더 없음"]:
            messagebox.showerror("오류", "유효한 템플릿을 선택하세요")
            return
            
        # 필드 값 수집
        field_values = {}
        for name, entry in self.field_entries.items():
            value = entry.get().strip()
            if value:
                field_values[name] = value
                
        if not field_values:
            messagebox.showerror("오류", "하나 이상의 필드에 값을 입력하세요")
            return
            
        def create_task():
            try:
                if self.assistant.create_document_from_template(template_name, field_values):
                    self.after(0, lambda: self._show_success(f"'{template_name}' 템플릿으로 문서 생성 완료!"))
                else:
                    self.after(0, lambda: self._show_error("문서 생성 실패"))
            except Exception as e:
                self.after(0, lambda: self._show_error(f"생성 오류: {e}"))
                
        threading.Thread(target=create_task, daemon=True).start()
        
    def _show_error(self, message):
        messagebox.showerror("오류", message)
        self.parent.log(f"❌ {message}")
        
    def _show_success(self, message):
        messagebox.showinfo("성공", message)
        self.parent.log(f"✅ {message}")
        self.destroy()

class SmartStyleWindow(ctk.CTkToplevel):
    def __init__(self, parent, assistant):
        super().__init__(parent)
        
        self.parent = parent
        self.assistant = assistant
        self.style_plan = []
        
        self.title("🤖 AI 기반 스마트 스타일 적용")
        self.geometry("800x600")
        self.grab_set()
        
        self._setup_gui()
        self._analyze_document()
        
    def _setup_gui(self):
        # 진행 상황 표시
        self.progress_label = ctk.CTkLabel(self, text="🔍 문서 구조를 분석하고 있습니다...")
        self.progress_label.pack(pady=20)
        
        # 분석 결과 표시 영역
        self.result_frame = ctk.CTkScrollableFrame(self, label_text="📋 스타일 적용 계획")
        self.result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 버튼 영역
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        self.apply_button = ctk.CTkButton(button_frame, text="스타일 적용 실행", 
                                        command=self._apply_smart_styles, 
                                        state="disabled")
        self.apply_button.pack(side="left", padx=10)
        
        ctk.CTkButton(button_frame, text="취소", 
                     command=self.destroy).pack(side="right", padx=10)
    
    def _analyze_document(self):
        """문서 분석 실행"""
        self.after(100, self._run_analysis)
    
    def _extract_json(self, text):
        """마크다운 코드 블록에서 JSON 추출"""
        if not text or not text.strip():
            return ""
        
        text = text.strip()
        print(text)
        # 방법 1: `````` 패턴에서 JSON 추출
        import re
        json_match = re.search(r'``````', text, re.DOTALL)
        if json_match:
            extracted = json_match.group(1).strip()
            if extracted and (extracted.startswith('{') or extracted.startswith('[')):
                return extracted
        
        # 방법 2: 단순 코드 블록 제거
        if text.startswith('``````'):
            extracted = text[3:-3].strip()
            # 첫 줄이 'json'인 경우 제거
            lines = extracted.split('\n')
            if lines and lines[0].strip().lower() in ['json', 'javascript']:
                extracted = '\n'.join(lines[1:]).strip()
            if extracted and (extracted.startswith('{') or extracted.startswith('[')):
                return extracted
        
        # 방법 3: JSON 객체 직접 추출
        json_start = -1
        for i, char in enumerate(text):
            if char in ['{', '[']:
                json_start = i
                break
        
        if json_start >= 0:
            bracket_count = 0
            start_char = text[json_start]
            end_char = '}' if start_char == '{' else ']'
            
            for i in range(json_start, len(text)):
                if text[i] == start_char:
                    bracket_count += 1
                elif text[i] == end_char:
                    bracket_count -= 1
                    if bracket_count == 0:
                        return text[json_start:i + 1]
        
        # 방법 4: 원본이 이미 JSON인 경우
        if text.startswith('{') or text.startswith('['):
            return text
        
        return ""

    def _robust_extract_json(self, text):
        """✨ 강화된 JSON 추출 함수"""
        if not text or not text.strip():
            return ""
        
        text = text.strip()
        
        # 방법 1: 마크다운 코드 블록 제거 (``````)
        import re
        json_match = re.search(r'``````', text, re.DOTALL)
        if json_match:
            extracted = json_match.group(1).strip()
            if extracted and (extracted.startswith('{') or extracted.startswith('[')):
                return extracted
        
        # 방법 2: 단순 코드 블록 제거 (``````)
        if text.startswith('``````'):
            extracted = text[3:-3].strip()
            # 첫 줄이 'json'인 경우 제거
            lines = extracted.split('\n')
            if lines and lines[0].strip().lower() in ['json', 'javascript']:
                extracted = '\n'.join(lines[1:]).strip()
            if extracted and (extracted.startswith('{') or extracted.startswith('[')):
                return extracted
        
        # 방법 3: JSON 객체/배열 직접 추출
        json_start = -1
        for i, char in enumerate(text):
            if char in ['{', '[']:
                json_start = i
                break
        
        if json_start >= 0:
            # 마지막 } 또는 ] 찾기
            json_end = -1
            bracket_count = 0
            start_char = text[json_start]
            end_char = '}' if start_char == '{' else ']'
            
            for i in range(json_start, len(text)):
                if text[i] == start_char:
                    bracket_count += 1
                elif text[i] == end_char:
                    bracket_count -= 1
                    if bracket_count == 0:
                        json_end = i
                        break
            
            if json_end > json_start:
                return text[json_start:json_end + 1]
        
        # 방법 4: 원본 그대로 반환 (JSON인 경우)
        if text.startswith('{') or text.startswith('['):
            return text
        
        return ""

    def _run_analysis(self):
        """문서 분석 실행 - 다양한 JSON 구조 처리"""
        try:
            self.progress_label.configure(text="🤖 AI가 문서 구조를 분석하고 있습니다...")
            self.update()
            
            result = self.assistant.analyze_document_structure()
            if not result:
                self._show_error("문서 분석에 실패했습니다.")
                return
            
            self.parent.log(f"🔍 AI 분석 결과: {result[:300]}...")
            
            clean_json = self._extract_json(result)
            if clean_json:
                try:
                    analysis_data = json.loads(clean_json)
                    self.parent.log("✅ JSON 형식 파싱 성공")
                except json.JSONDecodeError:
                    self.parent.log("⚠️ JSON 파싱 실패, 테이블 파싱 시도...")
                    analysis_data = None
            
            # 2단계: JSON 파싱 실패 시 테이블 파싱 시도
            if not analysis_data:
                if '|' in result:  # 테이블 형태인지 확인
                    analysis_data = self._parse_markdown_table_to_json(result)
                    if analysis_data:
                        self.parent.log("✅ 마크다운 테이블 파싱 성공")
                
            if not analysis_data:
                self._show_error("AI 응답을 JSON 또는 테이블 형태로 파싱할 수 없습니다.")
                return                
            
            # ✨ 핵심 수정: 두 가지 JSON 구조 모두 처리
            raw_plan = []
            if isinstance(analysis_data, dict) and "style_plan" in analysis_data:
                # 기대했던 형식: {"style_plan": [...]}
                raw_plan = analysis_data["style_plan"]
            elif isinstance(analysis_data, list):
                # 새로 받은 형식: [{...}, {...}]
                raw_plan = analysis_data
            
            if not raw_plan:
                self._show_error("스타일 적용 계획을 생성하지 못했습니다.")
                return
                
            # ✨ 내부 형식으로 변환
            self.style_plan = self._normalize_style_plan(raw_plan)
            
            if self.style_plan:
                self._display_style_plan()
                self.apply_button.configure(state="normal")
                self.progress_label.configure(text="✅ 분석 완료! 계획을 확인하고 적용하세요.")
                self.parent.log(f"🎯 총 {len(self.style_plan)}개 구간 분석 완료")
            else:
                self._show_error("유효한 스타일 계획을 변환하지 못했습니다.")
                
        except Exception as e:
            error_msg = f"분석 오류: {e}"
            self._show_error(error_msg)
            import traceback
            self.parent.log(f"🚨 상세 오류: {traceback.format_exc()}")


    def _parse_markdown_table_to_json(self, text):
        """마크다운 테이블을 JSON으로 변환하는 fallback 함수"""
        try:
            lines = text.strip().split('\n')
            style_plan = []
            line_number = 1
            
            for line in lines:
                if '|' in line and not line.startswith('|---'):
                    # 테이블 행 파싱
                    cells = [cell.strip() for cell in line.split('|')[1:-1]]
                    if len(cells) >= 2:
                        content = cells[0].replace('**', '').strip()
                        style_info = cells[1].strip()
                        
                        # 스타일 타입 결정
                        style_type = "본문"
                        if 'Heading 1' in style_info or '24pt' in style_info:
                            style_type = "대제목"
                        elif 'Heading 2' in style_info or '18pt' in style_info:
                            style_type = "소제목" 
                        elif 'Heading 3' in style_info or '14pt' in style_info:
                            style_type = "소제목"
                        elif 'Caption' in style_info:
                            style_type = "표제목"
                        
                        if content and not content.startswith('<'):  # HTML 태그 제외
                            style_plan.append({
                                "start_line": line_number,
                                "end_line": line_number,
                                "content_preview": content[:50],
                                "style_type": style_type,
                                "confidence": 0.8
                            })
                            line_number += 1
            
            return {"style_plan": style_plan}
        except Exception as e:
            print(f"❌ 테이블 파싱 실패: {e}")
            return None
    
    def _display_style_plan(self):
        """스타일 계획을 GUI에 표시 - 사용 가능한 모든 스타일 로드"""
        # 사용 가능한 모든 스타일 목록 가져오기
        available_styles = self.parent.assistant.get_style_list()
        if not available_styles:
            available_styles = ["스타일 없음"]
            
        for i, plan in enumerate(self.style_plan):
            plan_frame = ctk.CTkFrame(self.result_frame)
            plan_frame.pack(fill="x", padx=5, pady=5)
            
            info_text = f"📍 {plan['start_line']}~{plan['end_line']}행: {plan['style_type']}"
            ctk.CTkLabel(plan_frame, text=info_text, font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=5)
            
            preview_text = f"내용: {plan['content_preview'][:50]}..."
            ctk.CTkLabel(plan_frame, text=preview_text).pack(anchor="w", padx=10, pady=2)
            
            # ✨ 핵심 수정: 콤보박스에 모든 스타일 표시
            style_combo = ctk.CTkComboBox(plan_frame, values=available_styles)
            style_combo.set(self._get_default_style(plan['style_type']))
            style_combo.pack(anchor="w", padx=10, pady=5)
            
            plan['style_combo'] = style_combo

    def _get_default_style(self, style_type):
        """스타일 타입에 따른 기본 스타일 반환"""
        mapping = {
            "대제목": "보고서_대제목",
            "소제목": "보고서_소제목", 
            "본문": "보고서_본문",
            "표제목": "표_제목"
        }
        # 매핑된 스타일이 없으면 기본값으로 '보고서_본문' 사용
        return mapping.get(style_type, "보고서_본문")

    
    
    def _apply_smart_styles(self):
        """스마트 스타일 적용 실행"""
        try:
            # 스타일 매핑 수집
            style_mapping = {}
            for plan in self.style_plan:
                style_mapping[plan['style_type']] = plan['style_combo'].get()
            
            # 자동 적용 실행
            if self.assistant.apply_smart_styles(self.style_plan, style_mapping):
                messagebox.showinfo("성공", "스마트 스타일 적용이 완료되었습니다!")
                self.destroy()
            else:
                messagebox.showerror("실패", "스타일 적용 중 오류가 발생했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"스타일 적용 오류: {e}")

    def _show_error(self, message):
        """에러 메시지 표시"""
        from tkinter import messagebox
        self.progress_label.configure(text=f"❌ {message}")
        messagebox.showerror("오류", message)
        self.parent.log(f"❌ {message}")
        
    def _show_success(self, message):
        """성공 메시지 표시"""
        from tkinter import messagebox
        messagebox.showinfo("성공", message)
        self.parent.log(f"✅ {message}")

    def _parse_markdown_table_to_json(self, text):
        """마크다운 테이블을 JSON으로 변환하는 fallback 함수"""
        try:
            lines = text.strip().split('\n')
            style_plan = []
            line_number = 1
            
            for line in lines:
                if '|' in line and not line.startswith('|---'):
                    # 테이블 행 파싱
                    cells = [cell.strip() for cell in line.split('|')[1:-1]]
                    if len(cells) >= 2:
                        content = cells[0].replace('**', '').strip()
                        style_info = cells[1].strip()
                        
                        # 스타일 타입 결정
                        style_type = "본문"
                        if 'Heading 1' in style_info or '24pt' in style_info:
                            style_type = "대제목"
                        elif 'Heading 2' in style_info or '18pt' in style_info:
                            style_type = "소제목" 
                        elif 'Heading 3' in style_info or '14pt' in style_info:
                            style_type = "소제목"
                        elif 'Caption' in style_info:
                            style_type = "표제목"
                        
                        if content and not content.startswith('<'):  # HTML 태그 제외
                            style_plan.append({
                                "start_line": line_number,
                                "end_line": line_number,
                                "content_preview": content[:50],
                                "style_type": style_type,
                                "confidence": 0.8
                            })
                            line_number += 1
            
            return {"style_plan": style_plan}
        except Exception as e:
            print(f"❌ 테이블 파싱 실패: {e}")
            return None

    def _normalize_style_plan(self, raw_plan):
        """다양한 형식의 AI 응답을 내부 표준 형식으로 변환"""
        normalized_plan = []
        
        for i, item in enumerate(raw_plan):
            # start_line, end_line이 없는 경우 임의로 설정
            start_line = item.get("start_line", i + 1)
            end_line = item.get("end_line", i + 1)
            
            # content_preview 또는 content 키 사용
            content_preview = item.get("content_preview", item.get("content", ""))
            
            # style_type 또는 suggested_style 키 사용 및 변환
            style_type_raw = item.get("style_type", item.get("suggested_style", "본문"))
            style_type = style_type_raw.replace("보고서_", "") # "보고서_" 접두사 제거
            
            # confidence가 없는 경우 기본값 사용
            confidence = item.get("confidence", 0.9)
            
            normalized_plan.append({
                "start_line": start_line,
                "end_line": end_line,
                "content_preview": content_preview,
                "style_type": style_type,
                "confidence": confidence
            })
            
        return normalized_plan



if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
