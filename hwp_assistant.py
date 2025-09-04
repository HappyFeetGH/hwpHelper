import win32com.client as win32
import subprocess
import json
import sys
import os
import re
import win32clipboard as cb, win32con
import pythoncom

class HWPAssistant:
    def __init__(self):
        try:
            pythoncom.CoInitialize()
        except:
            pass

        self.hwp = None
        self.is_opened = False
        self.current_file = ""
        self.document_context = ""

    def open_file(self, file_path):
        if self.is_opened:
            print("⚠️  이미 파일이 열려있습니다. 'close' 명령으로 먼저 닫아주세요.")
            return False
        try:
            if self.hwp is None:
                pythoncom.CoInitialize()
                self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
                self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
                self.hwp.XHwpWindows.Item(0).Visible = True

            self.hwp.Open(file_path)
            self.is_opened = True
            self.current_file = os.path.abspath(file_path)
            
            full_text = self.hwp.GetTextFile("TEXT", "")
            self.document_context = f"""
### 현재 문서 컨텍스트
- **파일명**: {os.path.basename(file_path)}
- **문서 유형 추정**: {self._detect_document_type(full_text)}
- **내용 미리보기 (상위 1000자)**:
{full_text[:1000]}...
"""
            print(f"✅ 파일이 열렸습니다: {file_path}")
            print("🖥️  HWP 창이 화면에 표시되었습니다. 이제 텍스트를 선택하고 명령을 내리세요.")
            return True
        except Exception as e:
            print(f"❌ 파일 열기 실패: {e}")
            if self.hwp: self.hwp.Quit()
            return False

    def _detect_document_type(self, text):
        if "논문" in text: return "학술논문"
        if "보고서" in text: return "업무보고서"
        if "공문" in text: return "공문서"
        return "일반문서"

    def get_selected_text(self):
        if not self.is_opened: return ""
        try:
            self.hwp.InitScan(0x01, 0x00ff); texts = []
            while True:
                status, txt = self.hwp.GetText()
                if status == 0: break
                texts.append(txt)
            self.hwp.ReleaseScan(); return "".join(texts).strip()
        except Exception: return ""

    def replace_selected_text(self, new_text):
        if not self.is_opened: return False
        try:
            pset = self.hwp.HParameterSet.HInsertText
            pset.Text = new_text
            self.hwp.HAction.Execute("InsertText", pset.HSet)
            return True
        except Exception as e:
            print(f"❌ 텍스트 교체 실패: {e}", file=sys.stderr); return False

    def _find_context_file(self, filename):
        """컨텍스트 파일을 여러 경로에서 찾기"""
        # 1. 현재 작업 디렉토리
        if os.path.exists(filename):
            return filename
        
        # 2. 스크립트가 있는 디렉토리
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_dir, filename)
        if os.path.exists(script_path):
            return script_path
        
        # 3. 열린 HWP 파일과 같은 디렉토리
        if self.current_file:
            hwp_dir = os.path.dirname(self.current_file)
            hwp_path = os.path.join(hwp_dir, filename)
            if os.path.exists(hwp_path):
                return hwp_path
        
        # 4. 일반적인 컨텍스트 파일 경로들
        common_paths = [
            os.path.join(os.getcwd(), "context", filename),
            os.path.join(os.getcwd(), "instructions", filename),
            os.path.join(script_dir, "context", filename),
            os.path.join(script_dir, "instructions", filename)
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                return path
        
        return None

    def call_gemini(self, user_request, context_data, mode="default"):
        """
        다양한 작업 모드를 지원하는 통합 Gemini 호출 메서드.

        Args:
            user_request (str): 사용자의 원본 요청 문자열.
            context_data (str): AI가 참고할 주된 데이터 (선택된 텍스트, 문서 전체 등).
            mode (str): 작업 모드 ('default', 'template_analysis', 'template_apply').
        """
        
        # --- 1. 시스템 지침(Instruction) 결정 ---
        instruction_map = {
            "template_analysis": "instructions/template_analysis.md",
            "template_apply": "instructions/template_application.md",
            "default": "instructions/default_modification.md" # 기본 수정 지침
        }
        instruction_path = self._find_context_file(instruction_map.get(mode, "default_modification.md"))
        system_instruction = ""
        if instruction_path:
            try:
                with open(instruction_path, 'r', encoding='utf-8') as f:
                    system_instruction = f.read()
                print(f"✅ 시스템 지침 로드: {instruction_path}")
            except Exception as e:
                print(f"⚠️ 시스템 지침 파일 읽기 오류: {e}")
                
        # --- 2. 사용자 제공 추가 컨텍스트(@파일) 처리 ---
        context_files = re.findall(r'@([^\s]+)', user_request)
        user_context = ""
        if context_files:
            for filename in context_files:
                actual_path = self._find_context_file(filename)
                if actual_path:
                    try:
                        with open(actual_path, 'r', encoding='utf-8') as f:
                            user_context += f"\n--- 사용자 제공 컨텍스트: {os.path.basename(actual_path)} ---\n"
                            user_context += f.read()
                        print(f"📎 추가 컨텍스트 로드: {actual_path}")
                    except Exception as e:
                        print(f"⚠️ 컨텍스트 파일 읽기 오류: {e}")
        
        # --- 3. 최종 프롬프트 조합 ---
        prompt = f"""
    ### === 시스템 지침 ===
    {system_instruction}

    ### === 사용자 제공 컨텍스트 ===
    {user_context}

    ### === 작업 대상 데이터 ===
    {context_data}

    ### === 사용자 요청 ===
    {user_request}

    ---
    너의 임무는 위의 모든 정보를 종합하여, '시스템 지침'에 명시된 대로 **오직 최종 결과물만** 출력하는 것이다.
    """
        # --- 4. Gemini CLI 호출 ---
        try:
            command = 'gemini --model gemini-2.5-flash'
            result = subprocess.run(command, input=prompt, text=True, capture_output=True, encoding='utf-8', shell=True)
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                print(f"❌ Gemini 호출 실패: {result.stderr.strip()}"); return None
        except Exception as e:
            print(f"❌ Gemini 호출 오류: {e}"); return None


    
    def move_caret_right(self):
        """커서를 오른쪽으로 한 칸 이동 (블록 선택 해제 효과)"""
        try:
            return self.hwp.HAction.Run("MoveRight")
        except Exception as e:
            print(f"❌ 커서 오른쪽 이동 실패: {e}")
            return False

    def move_caret_down(self):
        """커서를 아래로 한 줄 이동 (블록 선택 해제 효과)"""
        try:
            return self.hwp.HAction.Run("MoveDown")
        except Exception as e:
            print(f"❌ 커서 아래 이동 실패: {e}")
            return False

    
    def _set_clip(self, text: str):
        """클립보드에 유니코드 텍스트 설정"""
        cb.OpenClipboard()
        cb.EmptyClipboard()
        cb.SetClipboardData(win32con.CF_UNICODETEXT, text)
        cb.CloseClipboard()


    def insert_table(self, markdown_table: str) -> bool:
        """마크다운 표를 HWP 문서에 삽입"""
        if not self.is_opened or not markdown_table:
            return False

        # 1) 마크다운 파싱
        lines = [line.strip() for line in markdown_table.strip().split('\n') if line.strip()]
        
        if len(lines) > 1 and lines[1].lstrip().startswith('|') and '-' in lines[1]:
            lines.pop(1)

        table_data = []
        for line in lines:
            if line.startswith('|') and line.endswith('|'):
                line = line[1:-1]
            cells = [cell.strip() for cell in line.split('|')]
            if any(cells):
                table_data.append(cells)

        rows = len(table_data)
        cols = max(len(r) for r in table_data) if rows > 0 else 0

        if rows * cols == 0:
            print("❌ 표 데이터 파싱 실패")
            return False

        try:
            self.move_caret_right()
            # 2) 표 생성 (self.hwp 사용!)
            act = self.hwp.CreateAction("TableCreate")
            pset = act.CreateSet()
            act.GetDefault(pset)
            
            pset.SetItem("Rows", rows)
            pset.SetItem("Cols", cols)
            pset.SetItem("WidthType", 2)
            pset.SetItem("HeightType", 0)
            
            act.Execute(pset)

            # 3) 행 단위 데이터 입력 (self.hwp 사용!)
            for r, row in enumerate(table_data):
                self.hwp.HAction.Run("TableCellBlockRow")
                
                padded_row = row + [""] * (cols - len(row))
                row_text = "\t".join(padded_row)
                
                self._set_clip(row_text)
                self.hwp.HAction.Run("Paste")
                
                if r < rows - 1:
                    self.hwp.HAction.Run("TableLowerCell")

            # 4) 표 편집 모드 종료 (self.hwp 사용!)
            self.hwp.HAction.Run("Cancel")
            print(f"✅ {rows}×{cols} 표 삽입 완료!")
            return True

        except Exception as e:
            print(f"❌ 표 삽입 실패: {e}")
            return False


    def analyze_document_for_template(self):
        """현재 문서를 분석하여 템플릿화 가능한 요소들을 추출"""
        if not self.is_opened:
            return None
        
        # 전체 텍스트 추출
        full_text = self.hwp.GetTextFile("TEXT", "")
        
        # 문서 구조 정보 수집
        structure_info = {
            "full_text": full_text,
            "paragraphs": full_text.split('\n'),
            "document_type": self._detect_document_type(full_text),
            "potential_variables": self._find_potential_variables(full_text)
        }
        
        return structure_info

    def _find_potential_variables(self, text):
        """템플릿화할 수 있는 변수들을 휴리스틱으로 찾기"""        
        potential_vars = []
        
        # 날짜 패턴
        date_patterns = re.findall(r'\d{4}년\s*\d{1,2}월\s*\d{1,2}일', text)
        # 이름 패턴 (직책 + 이름)
        name_patterns = re.findall(r'(과장|부장|팀장|대리|주임)\s*([가-힣]{2,4})', text)
        # 숫자 패턴
        number_patterns = re.findall(r'\d+(?:,\d{3})*(?:원|건|명|개)', text)
        
        return {
            "dates": date_patterns,
            "names": name_patterns, 
            "numbers": number_patterns
        }

    def create_template_from_current(self, template_name):
        """현재 문서를 템플릿으로 저장"""
        if not self.hwp:
            print("❌ HWP 객체가 초기화되지 않았습니다.")
            return False
            
        if not self.is_opened:
            print("❌ 열린 문서가 없습니다.")
            return False
        
        # 템플릿 저장 경로
        template_path = os.path.join(os.getcwd(), "templates", f"{template_name}.hwp")
        os.makedirs(os.path.dirname(template_path), exist_ok=True)
        
        try:
            self.hwp.Save()

            # 현재 문서를 템플릿으로 저장
            self.hwp.SaveAs(template_path)
            print(f"✅ 템플릿 저장 완료: {template_path}")
            return template_path
        except Exception as e:
            print(f"❌ 템플릿 저장 실패: {e}")
            return False

    def create_document_from_template(self, template_name, field_values):
        """템플릿을 바탕으로 새 문서 생성 (누름틀 제거 포함)"""
        template_path = os.path.join(os.getcwd(), "templates", f"{template_name}.hwp")
        
        if not os.path.exists(template_path):
            print(f"❌ 템플릿 파일이 없습니다: {template_path}")
            return False
        
        try:
            # 기존에 열린 파일이 있다면 닫기
            if self.is_opened:
                self.close_file()

            if not self.hwp:
                pythoncom.CoInitialize()
                self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
                self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
                self.hwp.XHwpWindows.Item(0).Visible = True

            # 템플릿 파일 열기
            if not self.open_file(template_path):
                print("❌ 템플릿 파일 열기 실패")
                return False
            
            # 1단계: 필드 값 적용
            print("🔄 누름틀에 값을 입력합니다...")
            for field_name, field_value in field_values.items():
                #merged_field_name = field_name+" 자동생성 필드"
                merged_field_name = field_name
                try:
                    # ✨ hwp 객체 상태 재확인
                    if not self.hwp:
                        raise Exception("HWP 객체가 None입니다")
                        
                    self.hwp.PutFieldText(merged_field_name, str(field_value))
                    print(f"✅ 필드 '{field_name}' -> '{field_value}' 적용 완료")
                except Exception as e:
                    print(f"⚠️ 필드 '{field_name}' 적용 실패: {e}")

            
            # 2단계: 모든 누름틀 제거 (텍스트는 유지)
            #print("🔄 모든 누름틀을 제거합니다...")
            #self._remove_all_fields()
            
            # 3단계: 새로운 파일로 저장
            import datetime
            output_dir = os.path.join(os.getcwd(), "output")
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, f"{template_name}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.hwp")
            
            if not self.hwp:
                raise Exception("저장 중 HWP 객체가 None입니다")
                
            self.hwp.SaveAs(output_path)
            print(f"📄 완성된 문서 저장: {output_path}")

            return True
        except Exception as e:
            print(f"❌ 템플릿 문서 생성 실패: {e}")
            return False

    # 모든 누름틀 삭제 예시
    def _remove_all_fields(self):
        """문서 내 모든 누름틀 제거 (텍스트는 유지) - 팝업 차단 강화"""
        try:
            # ✨ 강화된 팝업 차단 설정
            self.hwp.SetMessageBoxMode(0x00010001)  # 기본 팝업 차단
            self.hwp.SetMessageBoxMode(0x00010000)  # 추가 팝업 차단
            self.hwp.SetMessageBoxMode(0x10000000)  # 확인 대화상자 차단
            
            field_positions = []
            ctrl = self.hwp.HeadCtrl
            
            # 1) 모든 누름틀 위치 수집
            while ctrl:
                if ctrl.CtrlID == "%clk":  # 누름틀의 CtrlID
                    field_positions.append(ctrl.GetAnchorPos(0))
                ctrl = ctrl.Next
            
            # 2) 역순으로 누름틀 삭제
            for pos in reversed(field_positions):
                try:
                    self.hwp.SetPosBySet(pos)
                    # 누름틀 선택 후 삭제
                    self.hwp.Run("SelectCtrl")
                    self.hwp.Run("Delete")
                except Exception as e:
                    print(f"⚠️ 누름틀 삭제 중 오류: {e}")
            
            print(f"✅ 총 {len(field_positions)}개의 누름틀을 제거했습니다.")
            return len(field_positions) > 0
            
        except Exception as e:
            print(f"❌ 누름틀 제거 실패: {e}")
            return False
        finally:
            # ✨ 팝업 모드 원상 복구
            self.hwp.SetMessageBoxMode(0)



    def convert_text_to_field(self, search_text: str, field_name: str):
        """search_text를 찾아 CreateField()로 누름틀 변환 (가장 안정적인 방법)"""
        if not self.is_opened:
            return False
        
        try:
            # 팝업 자동 확인 처리
            self.hwp.SetMessageBoxMode(0x00010001)
            
            # 커서를 문서 맨 위로 이동
            self.hwp.HAction.Run("MoveTop")

            # 찾기 액션 실행
            find_act = self.hwp.CreateAction("RepeatFind")
            if not find_act:
                return False
            
            fset = find_act.CreateSet()
            find_act.GetDefault(fset)
            fset.SetItem("FindString", search_text)
            fset.SetItem("Direction", 1)
            
            if find_act.Execute(fset):
                # ✨ 핵심: CreateField() 직접 호출
                self.hwp.CreateField(
                    field_name,                    # 필드명 (PutFieldText에서 사용할 키)
                    f"{search_text}",             # 안내문 (사용자가 보는 텍스트)
                    f"{field_name} 자동생성 필드"  # 도움말
                )
                print(f"✅ '{search_text}' -> 누름틀 '{field_name}' 변환 완료")
                return True
            else:
                print(f"⚠️ '{search_text}' 찾기 실패")
                return False

        except Exception as e:
            print(f"❌ 누름틀 변환 실패: {e}")
            return False
        finally:
            self.hwp.SetMessageBoxMode(0)

    def get_field_list_from_file(self, template_name):
        """템플릿 파일에서 누름틀 필드 목록을 가져옵니다."""
        template_path = os.path.join(os.getcwd(), "templates", f"{template_name}.hwp")
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")
        
        # 현재 열린 파일이 있다면 상태 저장 후 닫기
        was_opened = self.is_opened
        original_file = self.current_file
        
        if was_opened:
            self.hwp.Save() # 혹시 모를 변경사항 저장
            self.hwp.Quit()
            self.is_opened = False

        # 임시로 템플릿 파일 열기
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        self.hwp.XHwpWindows.Item(0).Visible = False # 화면에 보이지 않게 처리
        self.hwp.Open(template_path)
        
        # 필드 목록 가져오기
        field_list_raw = self.hwp.GetFieldList(0, "")
        fields = [f.strip() for f in field_list_raw.split('\x02') if f.strip()]
        
        # 임시 파일 닫기
        self.hwp.Quit()
        
        # 원래 파일이 있었다면 다시 열기
        if was_opened:
            self.open_file(original_file)
        else:
            self.hwp = None # hwp 객체 초기화
            
        return fields


    def get_style_list(self):
        """'styles' 폴더에서 사용 가능한 스타일(.json) 목록을 반환합니다."""
        style_dir = os.path.join(os.getcwd(), "styles")
        if not os.path.exists(style_dir):
            os.makedirs(style_dir)
            return []
        
        try:
            styles = [f[:-5] for f in os.listdir(style_dir) if f.endswith('.json')]
            return styles
        except Exception as e:
            print(f"❌ 스타일 목록 로딩 실패: {e}")
            return []

    def apply_style_to_selection(self, style_data):
        """JSON 데이터를 바탕으로 선택 영역에 스타일을 적용합니다."""
        if not self.is_opened:
            print("❌ 스타일을 적용할 파일이 열려있지 않습니다.")
            return False
            
        try:
            # --- 1. 글자 모양 적용 (CharShape) ---
            if "CharShape" in style_data:
                char_action = self.hwp.CreateAction("CharShape")
                char_set = char_action.CreateSet()
                char_action.GetDefault(char_set)
                
                for key, value in style_data["CharShape"].items():
                    char_set.SetItem(key, value)
                    
                char_action.Execute(char_set)
                print("✅ 글자 모양 적용 완료")

            # --- 2. 문단 모양 적용 (ParaShape) ---
            if "ParaShape" in style_data:
                para_action = self.hwp.CreateAction("ParagraphShape")
                para_set = para_action.CreateSet()
                para_action.GetDefault(para_set)
                
                for key, value in style_data["ParaShape"].items():
                    para_set.SetItem(key, value)
                    
                para_action.Execute(para_set)
                print("✅ 문단 모양 적용 완료")
                
            return True
        except Exception as e:
            print(f"❌ 스타일 적용 실패: {e}")
            return False
 


    def analyze_document_structure(self):
        """문서 구조를 분석하여 스타일 적용 계획을 생성"""
        if not self.is_opened:
            return None
        
        try:
            # 전체 텍스트와 줄 정보 가져오기
            full_text = self.hwp.GetTextFile("TEXT", "")
            lines = full_text.split('\n')
            
            # 줄 번호와 함께 텍스트 정보 구성
            numbered_text = []
            for i, line in enumerate(lines, 1):
                if line.strip():  # 빈 줄 제외
                    numbered_text.append(f"줄 {i}: {line.strip()}")
            
            analysis_text = '\n'.join(numbered_text)
            
            # Gemini에게 구조 분석 요청
            analysis_request = "이 문서의 구조를 분석하여 각 부분에 적절한 스타일을 제안해줘."
            result = self.call_gemini(
                analysis_request, 
                analysis_text, 
                mode="document_style_analysis"
            )
            
            return result
        except Exception as e:
            print(f"❌ 문서 구조 분석 실패: {e}")
            return None

    def select_text_by_line_range(self, start_line, end_line):
        """지정된 줄 범위의 텍스트를 선택"""
        try:
            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")
            
            # 시작 줄로 이동
            for i in range(start_line - 1):
                self.hwp.HAction.Run("MoveDown")
            
            # 줄 선택 시작
            self.hwp.HAction.Run("MoveLineBegin")
            self.hwp.HAction.Run("SelectMode")
            
            # 끝 줄까지 선택
            for i in range(end_line - start_line):
                self.hwp.HAction.Run("MoveDown")
            self.hwp.HAction.Run("MoveLineEnd")
            
            return True
        except Exception as e:
            print(f"❌ 텍스트 선택 실패: {e}")
            return False

    def apply_smart_styles(self, style_plan, style_mapping):
        """스타일 계획에 따라 자동으로 스타일 적용"""
        try:
            success_count = 0
            
            for plan_item in style_plan:
                start_line = plan_item['start_line']
                end_line = plan_item['end_line']
                style_type = plan_item['style_type']
                
                # 해당 범위 선택
                if not self.select_text_by_line_range(start_line, end_line):
                    continue
                
                # 매핑된 스타일 적용
                if style_type in style_mapping:
                    style_file = style_mapping[style_type]
                    style_path = os.path.join(os.getcwd(), "styles", f"{style_file}.json")
                    
                    with open(style_path, 'r', encoding='utf-8') as f:
                        style_data = json.load(f)
                    
                    if self.apply_style_to_selection(style_data):
                        print(f"✅ {start_line}~{end_line}행에 '{style_type}' 스타일 적용 완료")
                        success_count += 1
                
                # 선택 해제
                self.hwp.HAction.Run("Cancel")
            
            print(f"🎉 총 {success_count}개 구간에 스타일이 적용되었습니다!")
            return success_count > 0
            
        except Exception as e:
            print(f"❌ 자동 스타일 적용 실패: {e}")
            return False

    def close_file(self):
        """안전한 파일 닫기"""
        if self.hwp and self.is_opened:
            try:
                self.hwp.Quit()
                print("📁 파일이 닫혔고, HWP 프로세스가 종료되었습니다.")
            except Exception as e:
                print(f"⚠️ 파일 닫기 중 오류: {e}")
            finally:
                self.hwp = None
                self.is_opened = False
                self.current_file = ""


def extract_json_from_markdown(text):
    """마크다운 코드 블록에서 JSON 부분만 추출"""
    # ```json ... ```
    json_match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if json_match:
        return json_match.group(1).strip()
    
    # ``` ... ```
    code_match = re.search(r'```\s*(.*?)\s*```', text, re.DOTALL)
    if code_match:
        return code_match.group(1).strip()
    
    # 코드 블록이 없으면 원본 반환
    return text.strip()

def strip_code_block(text: str) -> str:
    """
    ``````  또는  ``````  형식이면
    앞뒤 3글자를 잘라 순수 JSON 부분만 돌려준다.
    그밖엔 원본 그대로 반환.
    """
    text = text.strip()
    if text.startswith('``````'):
        return text[3:-3].strip()   # 앞 · 뒤 백틱 제거
    return text


def main():
    assistant = HWPAssistant()
    print("🤖 HWP AI 어시스턴트 v3.0 (템플릿 기능 탑재)이 시작되었습니다.")
    print("사용법:")
    print("  - 'open [파일경로]': 파일 열기")
    print("  - 'close' / 'quit': 닫기 / 종료")
    print("\n[수정 및 생성]")
    print("  - (텍스트 선택 후) [요청] @[스타일파일.md]: 선택 영역 수정")
    print("  - (텍스트 선택 후) 표로 만들어줘: 선택 영역을 표로 변환")
    print("\n[템플릿]")
    print("  - '템플릿생성 [템플릿이름]': 현재 문서를 템플릿으로 저장 시도")
    print("  - '템플릿사용 [이름] [내용]': 템플릿으로 새 문서 생성")
    
    while True:
        user_input = input("\n📝 명령어를 입력하세요: ").strip()
        
        # --- 기본 명령어 처리 ---
        if user_input.lower() == 'quit':
            assistant.close_file(); print("👋 어시스턴트를 종료합니다."); break
        elif user_input.lower() == 'close':
            assistant.close_file()
        elif user_input.startswith('open '):
            assistant.open_file(user_input[5:].strip().replace("\"", ""))
        
        # --- 템플릿 생성 명령어 처리 ---
        elif user_input.startswith('템플릿생성 '):
            if not assistant.is_opened:
                print("⚠️ 먼저 템플릿으로 만들 HWP 파일을 열어주세요.")
                continue

            template_name = user_input[6:].strip()
            print(f"🔄 '{template_name}' 템플릿 생성을 시작합니다...")
            
            # 1. 문서 분석
            structure = assistant.analyze_document_for_template()
            if not structure:
                print("❌ 문서 분석에 실패했습니다."); continue

            # 2. Gemini에게 템플릿화 요청
            print("🤖 Gemini에게 템플릿화 가능 영역 분석을 요청합니다...")
            analysis_request = "이 문서를 분석하여 템플릿으로 만들 변수들을 제안해줘."
            template_plan_str = assistant.call_gemini(analysis_request, json.dumps(structure, ensure_ascii=False, indent=2), mode="template_analysis")
            
            if not template_plan_str:
                print("❌ Gemini 분석에 실패했습니다."); continue
                
            print(f"📋 Gemini 분석 결과:\n{template_plan_str}")

            
            # 3. 사용자 확인 후 템플릿 생성
            try:
                 # JSON 추출 및 파싱
                clean_json = strip_code_block(extract_json_from_markdown(template_plan_str))
                template_plan = json.loads(clean_json)                
                fields_to_create = template_plan.get("template_fields", [])
                
                if not fields_to_create:
                    print("⚠️ 템플릿으로 만들 필드를 찾지 못했습니다."); continue

                print(f"✅ {len(fields_to_create)}개의 템플릿 필드를 발견했습니다:")
                for field in fields_to_create:
                    print(f"   - {field.get('field_name', 'unknown')}: {field.get('description', 'no description')}")

                confirm = input("이 분석 결과로 템플릿을 생성할까요? (y/n): ").lower()
                if confirm == 'y':
                    for field in fields_to_create:
                        assistant.convert_text_to_field(field["original_text"], field["field_name"])
                    
                    assistant.create_template_from_current(template_name)
                else:
                    print("❌ 템플릿 생성을 취소했습니다.")
            except (json.JSONDecodeError, KeyError) as e:
                print(f"❌ Gemini 분석 결과를 처리할 수 없습니다: {e}")

        # --- 템플릿 사용 명령어 처리 ---
        elif user_input.startswith('템플릿사용 '):
            parts = user_input[6:].split(' ', 1)
            if len(parts) < 2:
                print("⚠️ 사용법: 템플릿사용 [템플릿이름] [값 정보]"); continue

            template_name, user_values = parts[0], parts[1]
            print(f"🔄 '{template_name}' 템플릿을 사용하여 새 문서를 생성합니다...")
            
            # Gemini에게 사용자 입력 파싱 요청
            print("🤖 Gemini에게 값 파싱을 요청합니다...")
            parsing_request = f"다음 사용자 입력을 템플릿 값으로 파싱해줘: {user_values}"
            parsed_values_str = assistant.call_gemini(parsing_request, user_values, mode="template_apply")

            if not parsed_values_str:
                print("❌ Gemini 값 파싱에 실패했습니다."); continue
            
            print(f"📝 파싱된 값들: {parsed_values_str}")
            
            try:
                clean_json = strip_code_block(extract_json_from_markdown(parsed_values_str))
                field_values = json.loads(clean_json)
                assistant.create_document_from_template(template_name, field_values)
            except json.JSONDecodeError:
                print("❌ Gemini가 생성한 값(JSON)을 처리할 수 없습니다.")
        
        # --- 일반 수정 및 표 생성 처리 ---
        elif assistant.is_opened:
            selected_text = assistant.get_selected_text()
            if not selected_text and "표" not in user_input:
                print("⚠️ 먼저 HWP에서 텍스트를 선택하거나, 표 생성 요청을 해주세요."); continue
            
            print(f"📌 선택된 텍스트: '{selected_text[:50]}...'")
            print("🔄 Gemini에게 작업을 요청합니다...")
            
            modified_text = assistant.call_gemini(user_input, selected_text, mode="default")
            
            if modified_text:
                print(f"✨ Gemini 제안:\n{'-'*20}\n{modified_text}\n{'-'*20}")
                
                # 표 삽입
                if "표" in user_input and modified_text.strip().startswith('|'):
                    confirm = input("이 표를 현재 커서 위치에 삽입할까요? (y/n): ").lower()
                    if confirm == 'y': assistant.insert_table(modified_text)
                    else: print("❌ 표 삽입을 취소했습니다.")
                # 일반 텍스트 교체
                else:
                    confirm = input("이 내용으로 교체할까요? (y/n): ").lower()
                    if confirm == 'y':
                        if assistant.replace_selected_text(modified_text): print("✅ 성공적으로 교체되었습니다!")
                        else: print("❌ 교체에 실패했습니다.")
                    else: print("❌ 교체를 취소했습니다.")
        else:
            print("⚠️ 먼저 명령을 실행할 파일을 열어주세요. (예: open 파일경로)")

if __name__ == "__main__":
    main()
