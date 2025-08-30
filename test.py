import win32com.client as win32
import win32clipboard as cb
import win32con
import os
import re

def _set_clip(text: str):
    """클립보드에 유니코드 텍스트 설정"""
    cb.OpenClipboard()
    cb.EmptyClipboard()
    cb.SetClipboardData(win32con.CF_UNICODETEXT, text)
    cb.CloseClipboard()

def insert_table(hwp, markdown_table: str) -> bool:
    """마크다운 표를 HWP 문서에 삽입"""
    if not markdown_table:
        print("❌ 표 데이터가 비어있습니다.")
        return False

    # 1) 마크다운 파싱
    lines = [line.strip() for line in markdown_table.strip().split('\n') if line.strip()]
    
    # 헤더 구분선 제거 (|---|---|)
    if len(lines) > 1 and lines[1].lstrip().startswith('|') and '-' in lines[1]:
        lines.pop(1)

    table_data = []
    for line in lines:
        if line.startswith('|') and line.endswith('|'):
            line = line[1:-1]
        cells = [cell.strip() for cell in line.split('|')]
        if any(cells):  # 빈 행 제외
            table_data.append(cells)

    rows = len(table_data)
    cols = max(len(r) for r in table_data) if rows > 0 else 0

    if rows * cols == 0:
        print("❌ 표 데이터 파싱 실패")
        return False

    print(f"📊 파싱 결과: {rows}행 {cols}열")
    for i, row in enumerate(table_data):
        print(f"   행 {i+1}: {row}")

    try:
        # 2) 표 생성 (CreateAction 패턴)
        act = hwp.CreateAction("TableCreate")
        pset = act.CreateSet()
        act.GetDefault(pset)
        
        pset.SetItem("Rows", rows)
        pset.SetItem("Cols", cols)
        pset.SetItem("WidthType", 2)  # 자동 너비
        pset.SetItem("HeightType", 0)  # 자동 높이
        
        act.Execute(pset)
        print("✅ 표 프레임 생성 완료")

        # 3) 행 단위 데이터 입력
        for r, row in enumerate(table_data):
            print(f"🔄 {r+1}행 데이터 입력: {row}")
            
            # 현재 행 전체 블록 선택
            hwp.HAction.Run("TableCellBlockRow")
            
            # 열 수를 맞춰 탭으로 구분된 텍스트 생성
            padded_row = row + [""] * (cols - len(row))
            row_text = "\t".join(padded_row)
            
            # 클립보드를 통해 붙여넣기
            _set_clip(row_text)
            hwp.HAction.Run("Paste")
            
            # 마지막 행이 아니면 다음 행으로 이동
            if r < rows - 1:
                hwp.HAction.Run("TableLowerCell")

        # 4) 표 편집 모드 종료
        hwp.HAction.Run("Cancel")
        print(f"✅ {rows}×{cols} 표 삽입 완료!")
        return True

    except Exception as e:
        print(f"❌ 표 삽입 실패: {e}")
        return False

def main():
    print("🤖 HWP 표 삽입 테스트 시작")
    
    # 테스트할 마크다운 표
    test_table = '''
| 항목 | 수량 |
|---|---|
| 사과 | 5개 |
| 바나나 | 10개 |
| 오렌지 | 3개 |
'''
    
    # HWP 파일 경로 (현재 폴더의 test.hwp)
    file_path = os.path.join(os.getcwd(), "test.hwp")
    
    # test.hwp 파일이 없으면 생성
    if not os.path.exists(file_path):
        print("📝 test.hwp 파일이 없어 새로 생성합니다...")
        try:
            hwp_temp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp_temp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            hwp_temp.XHwpWindows.Item(0).Visible = False  # 임시로 숨김
            hwp_temp.New()
            hwp_temp.SaveAs(file_path)
            hwp_temp.Quit()
            print("✅ 빈 test.hwp 파일 생성 완료")
        except Exception as e:
            print(f"❌ test.hwp 생성 실패: {e}")
            return

    try:
        # HWP 실행 및 파일 열기
        print("🔄 HWP 프로그램 실행 중...")
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True
        
        print(f"📂 파일 열기: {file_path}")
        hwp.Open(file_path)
        
        print("🔄 표 삽입 시작...")
        success = insert_table(hwp, test_table)
        
        if success:
            print("\n🎉 테스트 성공! HWP 창에서 표가 제대로 삽입되었는지 확인하세요.")
            print("📝 Enter 키를 누르면 프로그램이 종료됩니다.")
            input()
        else:
            print("\n❌ 테스트 실패!")
            
        # HWP 종료하지 않고 사용자가 결과 확인할 수 있도록 유지
        
    except Exception as e:
        print(f"❌ 전체 프로세스 실패: {e}")

if __name__ == "__main__":
    main()
