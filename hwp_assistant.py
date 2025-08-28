import win32com.client as win32
import subprocess
import json
import sys
import os

class HWPAssistant:
    def __init__(self):
        self.hwp = None
        self.is_opened = False
        self.current_file = ""
        self.document_context = ""
        
    def open_file(self, file_path):
        """HWP 파일을 열고 **사용자에게 창을 보여준 뒤** 컨텍스트 생성"""
        if self.is_opened:
            print("⚠️ 이미 파일이 열려있습니다. 먼저 'close'를 실행해주세요.")
            return False

        try:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            
            # --- ✨ 핵심 수정 부분 시작 ✨ ---
            # 1. 창을 보이게 설정
            self.hwp.XHwpWindows.Item(0).Visible = True
            
            # 2. 파일 열기
            self.hwp.Open(file_path)
            
            # 3. 창을 최상단으로 활성화 (선택 사항, 사용자 편의성 증대)
            # self.hwp.XHwpWindows.Item(0).Activate() # 더 강력하게 창을 맨 위로 올림
            # --- ✨ 핵심 수정 부분 끝 ✨ ---

            self.is_opened = True
            self.current_file = file_path
            
            # 전체 문서 내용 추출 (컨텍스트용)
            full_text = self.hwp.GetTextFile("TEXT", "")
            self.document_context = f"""
현재 열린 파일: {file_path}
문서 유형: {self._detect_document_type(full_text)}
전체 내용 미리보기:
{full_text[:1000]}...
"""
            print(f"✅ 파일이 열렸습니다: {file_path}")
            print("🖥️  HWP 창이 화면에 표시되었습니다. 이제 텍스트를 선택하고 명령을 내리세요.")
            return True
            
        except Exception as e:
            print(f"❌ 파일 열기 실패: {e}")
            if self.hwp: # 실패 시 프로세스 정리
                self.hwp.Quit()
            return False
    
    def _detect_document_type(self, text):
        """문서 유형 자동 감지"""
        if "논문" in text or "연구" in text:
            return "학술논문"
        elif "보고서" in text:
            return "업무보고서"  
        elif "공문" in text or "시행" in text:
            return "공문서"
        else:
            return "일반문서"

    def get_selected_text(self):
        """사용자가 선택한 텍스트 추출"""
        if not self.hwp or not self.is_opened:
            return ""
        
        try:
            self.hwp.InitScan(0x01, 0x00ff)
            texts = []
            while True:
                status, txt = self.hwp.GetText()
                if status == 0:
                    break
                texts.append(txt)
            self.hwp.ReleaseScan()
            return "".join(texts).strip()
        except Exception as e:
            return ""

    def replace_selected_text(self, new_text):
        """선택된 영역을 새로운 텍스트로 교체"""
        if not self.hwp or not self.is_opened:
            return False
        
        try:
            pset = self.hwp.HParameterSet.HInsertText
            pset.Text = new_text
            self.hwp.HAction.Execute("InsertText", pset.HSet)
            return True
        except Exception as e:
            print(f"❌ 텍스트 교체 실패: {e}", file=sys.stderr)
            return False
    
    def call_gemini(self, user_request, selected_text):
        """Gemini CLI를 호출하여 텍스트 수정 요청 처리"""
        prompt = f"""
    {self.document_context}

    === 사용자 선택 텍스트 ===
    {selected_text}

    === 수정 요청 ===
    {user_request}

    === 지침 ===
    위 선택된 텍스트를 사용자의 요청에 맞게 수정해주세요.
    - 원본의 맥락과 스타일을 유지하되, 요청사항을 정확히 반영하세요.
    - 수정된 텍스트만 출력하고, 다른 설명은 붙이지 마세요.
    - 선택된 부분만 수정하고, 전체 문서 구조는 건드리지 마세요.
    """
        
        try:
            # Gemini CLI 호출 (정확한 모델명 사용)
            command = 'gemini --model models/gemini-2.5-flash'
            
            result = subprocess.run(
                command,
                input=prompt,
                text=True,
                capture_output=True,
                encoding='utf-8',
                shell=True
            )
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                print(f"❌ Gemini 호출 실패 (Return Code: {result.returncode}):")
                print(f"   - stderr: {result.stderr.strip()}")
                # 자동 대체 모델 시도 (선택적)
                print("🔄 기본 모델(gemini-pro)로 재시도합니다...")
                command = 'gemini --model gemini-pro' # 혹은 gemini-1.5-flash-latest 등 사용 가능한 모델
                result = subprocess.run(
                    command,
                    input=prompt, text=True, capture_output=True, encoding='utf-8', shell=True
                )
                if result.returncode == 0:
                    return result.stdout.strip()
                else:
                    print(f"❌ 재시도 실패: {result.stderr.strip()}")
                    return None

        except Exception as e:
            print(f"❌ Gemini 호출 중 예외 발생: {e}")
            return None



    def close_file(self):
        """HWP 파일 닫기 및 프로세스 종료"""
        if not self.is_opened:
            return
        
        try:
            # unsaved_prompt = "저장하지 않은 변경사항이 있습니다. 그래도 닫으시겠습니까?"
            # self.hwp.Quit(unsaved_prompt) # 사용자에게 저장 여부 묻기 (더 복잡한 구현 필요)
            self.hwp.Quit()
        except Exception as e:
            print(f"파일 닫기 중 오류 발생: {e}", file=sys.stderr)
        
        self.hwp = None
        self.is_opened = False
        self.current_file = ""
        self.document_context = ""
        print("📁 파일이 닫혔고, HWP 프로세스가 종료되었습니다.")

def main():
    assistant = HWPAssistant()
    print("🤖 HWP AI 어시스턴트가 시작되었습니다.")
    print("사용법:")
    print("1. 'open [파일경로]' - HWP 파일 열기")
    print("2. HWP 창에서 텍스트 선택 후, 터미널에 수정 요청 입력")
    print("3. 'close' - 현재 파일 닫기")
    print("4. 'quit' - 프로그램 종료")
    
    while True:
        user_input = input("\n📝 명령어를 입력하세요: ").strip()
        
        if user_input.lower() == 'quit':
            assistant.close_file()
            print("👋 어시스턴트를 종료합니다.")
            break
            
        elif user_input.lower() == 'close':
            assistant.close_file()

        elif user_input.startswith('open '):
            file_path = user_input[5:].strip().replace("\"", "") # 따옴표 제거
            assistant.open_file(file_path)
            
        elif assistant.is_opened:
            selected_text = assistant.get_selected_text()
            
            if not selected_text:
                print("⚠️ 먼저 HWP 창에서 텍스트를 선택한 후, 여기에 명령을 입력해주세요.")
                continue
                
            print(f"📌 선택된 텍스트: '{selected_text[:50]}...'")
            print("🔄 Gemini에게 수정을 요청합니다...")
            
            modified_text = assistant.call_gemini(user_input, selected_text)
            
            if modified_text:
                print(f"✨ Gemini 제안:\n{'-'*20}\n{modified_text}\n{'-'*20}")
                
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