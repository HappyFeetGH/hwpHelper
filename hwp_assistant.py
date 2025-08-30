import win32com.client as win32
import subprocess
import json
import sys
import os
import re
import win32clipboard as cb, win32con


class HWPAssistant:
    def __init__(self):
        self.hwp = None
        self.is_opened = False
        self.current_file = ""
        self.document_context = ""

    def open_file(self, file_path):
        if self.is_opened:
            print("⚠️  이미 파일이 열려있습니다. 'close' 명령으로 먼저 닫아주세요.")
            return False
        try:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            self.hwp.XHwpWindows.Item(0).Visible = True
            self.hwp.Open(file_path)
            self.is_opened = True
            self.current_file = file_path
            
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

    def call_gemini(self, user_request, selected_text):
        context_files = re.findall(r'@([^\s]+)', user_request)
        additional_context = ""
        
        if context_files:
            for filename in context_files:
                actual_path = self._find_context_file(filename)
                if actual_path:
                    try:
                        with open(actual_path, 'r', encoding='utf-8') as f:
                            additional_context += f"\n--- 추가 컨텍스트 파일: {os.path.basename(actual_path)} ---\n"
                            additional_context += f.read()
                        print(f"📎 추가 컨텍스트 로드: {actual_path}")
                    except Exception as e:
                        print(f"⚠️ 컨텍스트 파일 읽기 오류: {e}")
                else:
                    print(f"⚠️ 컨텍스트 파일을 찾을 수 없음: {filename}")
                    print(f"   시도한 경로들:")
                    print(f"   - 현재 디렉토리: {os.path.join(os.getcwd(), filename)}")
                    print(f"   - 스크립트 디렉토리: {os.path.join(os.path.dirname(__file__), filename)}")
    
        prompt = f"""
{self.document_context}
{additional_context}
---
### 작업 지시
- **사용자 선택 텍스트**:
{selected_text}
- **사용자 수정 요청**:
{user_request}

### === 너의 임무 ===
1. **지침 준수**: '추가 컨텍스트 파일'이 있다면, 그 파일의 어투, 형식, 스타일을 **반드시** 따라서 결과물을 생성해.
2. **결과물 생성**: '사용자 수정 요청'에 맞춰 '사용자 선택 텍스트'를 수정한 결과물을 만들어.
3. **형식 유지**: 만약 요청이 '표로 만들어줘'라면, 반드시 **마크다운 형식의 표**로 결과물을 출력해야 해. 그 외에는 일반 텍스트로 출력해.
4. **출력 정제**: 다른 설명, 인사말, 사과문 없이 **오직 수정된 결과물만** 출력해.
"""
        try:
            command = 'gemini --model models/gemini-2.5-flash'
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

    def close_file(self):
        if not self.is_opened: return
        try: self.hwp.Quit()
        except Exception: pass
        self.hwp, self.is_opened = None, False
        print("📁 파일이 닫혔고, HWP 프로세스가 종료되었습니다.")

def main():
    assistant = HWPAssistant()
    print("🤖 HWP AI 어시스턴트 v2.0 (맥락/표 지원)이 시작되었습니다.")
    print("사용법:")
    print("  - 'open [파일경로]': HWP 파일 열기")
    print("  - '[요청사항] @[컨텍스트파일.md]': 맥락 파일 참고하여 수정")
    print("  - '[선택된 텍스트를] 표로 만들어줘': 표 생성")
    print("  - 'close': 현재 파일 닫기")
    print("  - 'quit': 프로그램 종료")
    
    while True:
        user_input = input("\n📝 명령어를 입력하세요: ").strip()
        
        if user_input.lower() == 'quit':
            assistant.close_file()
            print("👋 어시스턴트를 종료합니다.")
            break
            
        elif user_input.lower() == 'close':
            assistant.close_file()
            
        elif user_input.startswith('open '):
            assistant.open_file(user_input[5:].strip().replace("\"", ""))
            
        elif assistant.is_opened:
            selected_text = assistant.get_selected_text()
            if not selected_text and "표" not in user_input:
                print("⚠️ 먼저 HWP에서 텍스트를 선택하거나, 표 생성 요청을 해주세요.")
                continue
            
            print(f"📌 선택된 텍스트: '{selected_text[:50]}...'")
            print("🔄 Gemini에게 작업을 요청합니다...")
            
            modified_text = assistant.call_gemini(user_input, selected_text)
            
            if modified_text:
                print(f"✨ Gemini 제안:\n{'-'*20}\n{modified_text}\n{'-'*20}")
                
                # 표 삽입 요청 처리
                if "표" in user_input and modified_text.strip().startswith('|'):
                    confirm = input("이 표를 현재 커서 위치에 삽입할까요? (y/n): ").lower()
                    if confirm == 'y': 
                        assistant.insert_table(modified_text)
                    else: 
                        print("❌ 표 삽입을 취소했습니다.")
                # 일반 텍스트 교체 처리
                else:
                    confirm = input("이 내용으로 교체할까요? (y/n): ").lower()
                    if confirm == 'y':
                        if assistant.replace_selected_text(modified_text): 
                            print("✅ 성공적으로 교체되었습니다!")
                        else: 
                            print("❌ 교체에 실패했습니다.")
                    else: 
                        print("❌ 교체를 취소했습니다.")
        else:
            print("⚠️ 먼저 'open [파일경로]' 명령으로 파일을 열어주세요.")

if __name__ == "__main__":
    main()
