import win32com.client as win32
import json
import os
import sys

def get_char_shape(hwp_obj):
    """현재 커서 위치의 글자 모양(서식) 정보를 반환합니다."""
    act = hwp_obj.CreateAction("CharShape")
    p_set = act.CreateSet()
    act.GetDefault(p_set)
    
    # HWP 내부 단위(1/100 pt)를 pt 단위로 변환
    height = p_set.Item("Height") / 100.0 
    
    # HWP 폰트 ID를 실제 폰트 이름으로 변환
    face_name_id = p_set.Item("FaceNameUser")
    font_name = hwp_obj.HAction.GetFaceName(face_name_id)
    
    is_bold = p_set.Item("Bold")

    return {"font": font_name, "size": height, "bold": bool(is_bold)}

def extract_hwp_structure_with_style(file_path: str) -> dict:
    """
    HWP 문서의 구조, 내용, 핵심 서식 정보를 체계적으로 추출합니다.
    """
    # ... (파일 존재 확인 및 hwp 객체 생성 부분은 이전과 동일) ...
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.Open(file_path)

    result = {
        "document_path": file_path,
        "metadata": {"title": "", "fields": {}},
        "structure": [] # 문서 내용을 순서대로 담을 리스트
    }

    # --- 1. 누름틀 정보 추출 ---
    try:
        field_list_raw = hwp.GetFieldList(1, "누름틀")
        if field_list_raw:
            for field_name in field_list_raw.split("\x02"):
                if field_name:
                    result["metadata"]["fields"][field_name] = hwp.GetFieldText(field_name).strip()
    except Exception:
        pass
        
    # --- 2. 문서 전체를 순회하며 텍스트와 표, 서식 추출 ---
    hwp.InitScan() # 문서 전체 스캔 시작
    while True:
        ret, text = hwp.GetText()
        if ret == 0: # 문서 끝에 도달하면 종료
            break

        # 현재 커서 위치의 서식 정보 가져오기
        pos = hwp.GetPos()
        char_shape = get_char_shape(hwp)
        
        # 텍스트 블록 정보 추가
        result["structure"].append({
            "type": "paragraph",
            "text": text.strip(),
            "style": char_shape
        })
        
        # 현재 위치에 표가 있는지 확인
        if hwp.IsCtrlField("tbl"):
            table_data = {"type": "table", "cells": []}
            
            # 표 컨트롤 선택 및 정보 추출 로직 (이전 답변의 수정된 로직)
            hwp.FindCtrl()
            table_ctrl = hwp.Object
            rows = table_ctrl.RowCount
            cols = table_ctrl.ColCount
            
            for r in range(rows):
                row_data = []
                for c in range(cols):
                    cell_text = table_ctrl.GetCellText(r, c).replace("\r\n", " ").strip()
                    row_data.append(cell_text)
                table_data["cells"].append(row_data)
                
            result["structure"].append(table_data)

    hwp.ReleaseScan() # 스캔 종료
    hwp.Quit()
    
    # 첫 번째 유의미한 텍스트를 제목으로 설정
    for item in result["structure"]:
        if item["type"] == "paragraph" and item["text"]:
            result["metadata"]["title"] = item["text"]
            break

    return result

def extract_hwp_structure(file_path: str) -> dict:
    """
    HWP 문서의 양식 구조와 내용을 체계적으로 추출하여 JSON 호환 딕셔너리로 반환합니다.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(file_path)
        
        result = {
            "document_path": file_path,
            "document_title": "",
            "paragraphs": [],
            "fields": {},
            "tables": []
        }

        # --- 1. 문서 제목 추출 ---
        try:
            result["document_title"] = hwp.GetFieldText("제목").strip()
        except Exception:
            try:
                hwp.SetPos(2, 0, 0)
                act = hwp.CreateAction("GetPos")
                p_set = act.CreateSet()
                act.Execute(p_set)
                result["document_title"] = p_set.Item("ParaText").split('\r\n')[0].strip()
            except Exception:
                result["document_title"] = "제목을 찾을 수 없음"

        # --- 2. 누름틀 (Field) 정보 추출 ---
        try:
            field_list_raw = hwp.GetFieldList(1, "누름틀")
            if field_list_raw:
                for field_name in field_list_raw.split("\x02"):
                    if field_name:
                        try:
                            field_text = hwp.GetFieldText(field_name)
                            result["fields"][field_name] = field_text.strip()
                        except Exception:
                            result["fields"][field_name] = "[값 추출 오류]"
        except Exception:
            result["fields"] = {}
        
        # --- 3. 표(Table) 정보 추출 (수정된 로직) ---
        ctrl = hwp.HeadCtrl
        table_index = 0
        
        while ctrl:
            if ctrl.CtrlID == "tbl":
                table_data = { 
                    "table_index": table_index, 
                    "description": f"표 {table_index + 1}",
                    "cells": [] 
                }
                
                try:
                    # 표의 위치로 커서 이동
                    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    
                    # 표 선택하기
                    hwp.Run("ShapeObjSelect")
                    
                    # 표 속성 정보를 얻기 위한 액션 생성
                    act = hwp.CreateAction("TablePropertyDialog")
                    p_set = act.CreateSet()
                    act.GetDefault(p_set)
                    
                    # 행과 열 개수 추출
                    rows = p_set.Item("Rows") if p_set.Item("Rows") else 0
                    cols = p_set.Item("Cols") if p_set.Item("Cols") else 0
                    
                    # 간단한 표 정보만 기록 (실제 셀 내용은 전체 텍스트에서 파악 가능)
                    table_data["rows"] = rows
                    table_data["cols"] = cols
                    table_data["cells"] = f"표 크기: {rows}행 {cols}열"
                    
                    result["tables"].append(table_data)
                    table_index += 1
                    
                except Exception as e:
                    # 표 세부 정보 추출 실패 시, 최소한 표 존재 정보는 기록
                    table_data["rows"] = "알 수 없음"
                    table_data["cols"] = "알 수 없음"  
                    table_data["cells"] = f"표 {table_index + 1} 감지됨 (세부 정보 추출 실패)"
                    result["tables"].append(table_data)
                    table_index += 1

            ctrl = ctrl.Next

        # --- 4. 일반 문단 텍스트 추출 ---
        text_content = hwp.GetTextFile("TEXT", "")
        result["paragraphs"] = [p.strip() for p in text_content.split('\r\n') if p.strip()]

    finally:
        if hwp:
            hwp.Quit()
            
    return result

def extract_hwp_with_formatting(file_path: str) -> dict:
    """
    HWP 문서의 내용과 서식 정보를 모두 추출합니다.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    hwp = None
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(file_path)
        
        result = {
            "document_path": file_path,
            "document_title": "",
            "paragraphs": [],
            "fields": {},
            "tables": [],
            "formatting_info": {
                "fonts_used": [],
                "paragraph_formats": [],
                "character_formats": []
            }
        }

        # --- 1. 문서 제목 추출 ---
        try:
            result["document_title"] = hwp.GetFieldText("제목").strip()
        except Exception:
            try:
                hwp.SetPos(2, 0, 0)
                act = hwp.CreateAction("GetPos")
                p_set = act.CreateSet()
                act.Execute(p_set)
                result["document_title"] = p_set.Item("ParaText").split('\r\n')[0].strip()
            except Exception:
                result["document_title"] = "제목을 찾을 수 없음"

        # --- 2. 누름틀 (Field) 정보 추출 ---
        try:
            field_list_raw = hwp.GetFieldList(1, "누름틀")
            if field_list_raw:
                for field_name in field_list_raw.split("\x02"):
                    if field_name:
                        try:
                            field_text = hwp.GetFieldText(field_name)
                            result["fields"][field_name] = field_text.strip()
                        except Exception:
                            result["fields"][field_name] = "[값 추출 오류]"
        except Exception:
            result["fields"] = {}
        
        # --- 3. 표(Table) 정보 추출 (수정된 로직) ---
        ctrl = hwp.HeadCtrl
        table_index = 0
        
        while ctrl:
            if ctrl.CtrlID == "tbl":
                table_data = { 
                    "table_index": table_index, 
                    "description": f"표 {table_index + 1}",
                    "cells": [] 
                }
                
                try:
                    # 표의 위치로 커서 이동
                    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    
                    # 표 선택하기
                    hwp.Run("ShapeObjSelect")
                    
                    # 표 속성 정보를 얻기 위한 액션 생성
                    act = hwp.CreateAction("TablePropertyDialog")
                    p_set = act.CreateSet()
                    act.GetDefault(p_set)
                    
                    # 행과 열 개수 추출
                    rows = p_set.Item("Rows") if p_set.Item("Rows") else 0
                    cols = p_set.Item("Cols") if p_set.Item("Cols") else 0
                    
                    # 간단한 표 정보만 기록 (실제 셀 내용은 전체 텍스트에서 파악 가능)
                    table_data["rows"] = rows
                    table_data["cols"] = cols
                    table_data["cells"] = f"표 크기: {rows}행 {cols}열"
                    
                    result["tables"].append(table_data)
                    table_index += 1
                    
                except Exception as e:
                    # 표 세부 정보 추출 실패 시, 최소한 표 존재 정보는 기록
                    table_data["rows"] = "알 수 없음"
                    table_data["cols"] = "알 수 없음"  
                    table_data["cells"] = f"표 {table_index + 1} 감지됨 (세부 정보 추출 실패)"
                    result["tables"].append(table_data)
                    table_index += 1

            ctrl = ctrl.Next

        # --- 서식 정보 추출 ---
        
        # 1. 문서에 사용된 폰트 목록 추출
        try:
            fonts = set()
            for i in range(1, hwp.XHwpDocuments.Count + 1):
                doc = hwp.XHwpDocuments[i]
                for j in range(1, doc.XHwpXFont.Count + 1):
                    font = doc.XHwpXFont.Item(j).Name
                    if font:
                        fonts.add(font)
            result["formatting_info"]["fonts_used"] = list(fonts)
        except Exception:
            result["formatting_info"]["fonts_used"] = ["폰트 정보 추출 실패"]

        # 2. 문서를 순회하며 각 위치의 서식 정보 샘플링
        hwp.SetPos(2, 0, 0)  # 문서 시작으로 이동
        
        # 몇 개의 샘플 위치에서 서식 정보 추출
        sample_positions = [0, 100, 200, 500, 1000]  # 문자 위치 샘플
        
        for pos in sample_positions:
            try:
                hwp.SetPos(2, pos, pos)
                
                char_format = {
                    "position": pos,
                    "font_name": hwp.CharShape.Item("FaceNameUser"),
                    "font_size": hwp.CharShape.Item("Height") / 100.0,
                    "is_bold": hwp.CharShape.Item("Bold"),
                    "is_italic": hwp.CharShape.Item("Italic"),
                    "underline": hwp.CharShape.Item("Underline")
                }
                
                para_format = {
                    "position": pos,
                    "alignment": hwp.ParaShape.Item("Align"),
                    "left_margin": hwp.ParaShape.Item("LeftMargin"),
                    "line_spacing": hwp.ParaShape.Item("LineSpacing")
                }
                
                result["formatting_info"]["character_formats"].append(char_format)
                result["formatting_info"]["paragraph_formats"].append(para_format)
                
            except Exception:
                continue

        # 기존 텍스트 추출 코드
        text_content = hwp.GetTextFile("TEXT", "")
        result["paragraphs"] = [p.strip() for p in text_content.split('\r\n') if p.strip()]

    finally:
        if hwp:
            hwp.Quit()
            
    return result


if __name__ == '__main__':
    # 스크립트 실행 시 첫 번째 인자로 파일 경로를 받음
    if len(sys.argv) < 2:
        print("사용법: python extractor.py \"<HWP 파일 경로>\"", file=sys.stderr)
        sys.exit(1)
        
    # 첫 번째 인자(sys.argv[1])를 파일 경로로 사용
    hwp_file_path = sys.argv[1] 
    
    try:
        document_structure = extract_hwp_with_formatting(hwp_file_path)
        print(json.dumps(document_structure, ensure_ascii=False, indent=2))
        
    except FileNotFoundError as e:
        print(f"오류: {e}", file=sys.stderr)
    except Exception as e:
        print(f"HWP 파일을 처리하는 중 오류가 발생했습니다: {e}", file=sys.stderr)
