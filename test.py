import win32com.client as win32
import os
import time

def test_putfieldtext():
    """
    PutFieldText의 안정성을 검증하는 강화된 테스트 코드.
    """
    print("🤖 PutFieldText 기능 테스트 시작...")
    
    # 테스트할 템플릿 파일 경로
    template_path = os.path.join(os.getcwd(), "templates", "알림장.hwp")
    
    if not os.path.exists(template_path):
        print(f"❌ 테스트 실패: '{template_path}' 파일이 없습니다.")
        return

    # 테스트할 필드명과 값
    field_to_test = "평가대상학년 필드입니다"
    value_to_insert = "테스트 성공!"
    
    hwp = None
    try:
        # HWP 실행 및 파일 열기
        print("🔄 HWP 프로그램 실행 중...")
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True
        
        print(f"📂 파일 열기: {template_path}")
        hwp.Open(template_path)
        
        # 1. PutFieldText 실행
        print(f"🔄 '{field_to_test}' 필드에 '{value_to_insert}' 값을 입력합니다...")
        hwp.PutFieldText(field_to_test, value_to_insert)
        
        # ✨ 핵심 수정: 상태 갱신을 위한 로직 추가
        print("⚙️  한/글 내부 상태 갱신 시도...")
        
        # 다른 필드로 포커스 이동 (문서의 첫 번째 필드 추천)
        all_fields = [f.strip() for f in hwp.GetFieldList(0, "").split('\x02') if f.strip()]
        if all_fields:
            first_field = all_fields[0]
            if first_field != field_to_test:
                hwp.MoveToField(first_field)
                print(f"   -> '{first_field}'(으)로 포커스 이동")
        
        # 다시 원래 필드로 포커스 이동하여 상태 재확인
        hwp.MoveToField(field_to_test)
        print(f"   -> 다시 '{field_to_test}'(으)로 포커스 이동")

        # 2. GetFieldText로 결과 재확인
        time.sleep(0.1) # 물리적 반응 시간 대기
        result_text = hwp.GetFieldText(field_to_test)
        
        print(f"📊 필드 값 재확인: '{result_text}'")
        
        if result_text == value_to_insert:
            print("✅ PutFieldText 실행 성공! 메모리상의 값 변경을 확인했습니다.")
        else:
            print("❌ PutFieldText 실행 실패! 값이 변경되지 않았습니다.")
            print("   (원인 추정: 필드 이름 오타 또는 문서 구조 문제)")
            return
            
        # 3. 변경사항 저장
        hwp.Save()
        print("💾 변경사항이 파일에 저장되었습니다.")
        
        print("\n🎉 테스트 성공! HWP 창에서 내용이 실제로 변경되었는지 확인하세요.")
        input("   확인 후 Enter 키를 누르면 프로그램이 종료됩니다.")

    except Exception as e:
        print(f"❌ 전체 프로세스 실패: {e}")
    finally:
        if hwp:
            hwp.Quit()

if __name__ == "__main__":
    test_putfieldtext()
