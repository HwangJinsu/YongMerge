import os
import time
import win32com.client
import pandas as pd
import tempfile
import image_utils
import pythoncom
import shutil
from PIL import Image

def get_word_instance(visible=False):
    """Word 인스턴스를 기존 작업과 분리하여 독립적으로 생성합니다."""
    print("DEBUG: 독립적인 Word 인스턴스 생성 중...")
    try:
        # DispatchEx를 사용하여 기존 창에 간섭하지 않는 새 프로세스 생성
        word = win32com.client.DispatchEx("Word.Application")
        print("DEBUG: Word 독립 프로세스 생성 완료.")
    except Exception as e:
        raise Exception(f"Word 실행 실패: {e}")
    
    try:
        word.Visible = visible
        word.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        word.DisplayAlerts = 0 # wdAlertsNone
    except Exception as e:
        print(f"DEBUG: Word 초기 설정 오류(무시): {e}")

    return word

def safe_open_doc(word, file_path, read_only=True):
    """파일 락에 대비하여 재시도 로직이 포함된 문서 열기"""
    abs_path = os.path.abspath(file_path)
    for attempt in range(3):
        try:
            doc = word.Documents.Open(abs_path, ReadOnly=read_only, AddToRecentFiles=False)
            if doc: return doc
        except Exception as e:
            print(f"DEBUG: 문서 열기 시도 {attempt+1} 실패: {e}")
            time.sleep(0.5)
    raise Exception(f"문서를 열 수 없습니다: {file_path}")

def insert_image_to_word(word_range, image_path, max_width_pt=450):
    """비율을 유지하며 Word에 이미지 삽입"""
    try:
        is_valid, message = image_utils.validate_image_path(image_path)
        if not is_valid: return False
            
        abs_path = os.path.abspath(image_path)
        shape = word_range.InlineShapes.AddPicture(FileName=abs_path, LinkToFile=False, SaveWithDocument=True)
        
        try:
            with Image.open(abs_path) as img:
                w, h = img.size
                ratio = h / w
                if shape.Width > max_width_pt:
                    shape.Width = max_width_pt
                    shape.Height = max_width_pt * ratio
        except: pass
        return True
    except Exception as e:
        print(f"ERROR: 이미지 삽입 오류: {e}")
        return False

def replace_text_in_story_ranges(doc, old_text, new_text):
    """문서 내 모든 영역(본문, 헤더, 푸터, 텍스트 상자 등)에서 텍스트/이미지 교체"""
    found_any = False
    is_image = image_utils.is_image_file(new_text)

    # 1. 모든 StoryRanges (본문, 헤더, 푸터 등) 처리
    for story in doc.StoryRanges:
        current_range = story
        while current_range:
            # 해당 영역 내의 텍스트/이미지 교체
            if _replace_in_range(doc, current_range, old_text, new_text, is_image):
                found_any = True
            
            # 해당 영역에 포함된 도형(Shapes) 처리 (텍스트 상자 등)
            try:
                if current_range.ShapeRange.Count > 0:
                    for shape in current_range.ShapeRange:
                        if shape.TextFrame.HasText:
                            if _replace_in_range(doc, shape.TextFrame.TextRange, old_text, new_text, is_image):
                                found_any = True
            except: pass

            current_range = current_range.NextStoryRange
    return found_any

def _replace_in_range(doc, target_range, old_text, new_text, is_image):
    """지정된 범위(Range) 내에서 텍스트 또는 이미지를 교체합니다."""
    found = False
    if is_image:
        search_range = target_range.Duplicate
        while True:
            find_obj = search_range.Find
            find_obj.ClearFormatting()
            find_obj.Text = old_text
            find_obj.Forward = True
            find_obj.Wrap = 0 # wdFindStop
            if find_obj.Execute():
                search_range.Text = ""
                if insert_image_to_word(search_range, new_text):
                    found = True
                # 다음 검색을 위해 범위 조정
                start_pos = search_range.End
                if start_pos >= target_range.End: break
                search_range = doc.Range(start_pos, target_range.End)
            else: break
    else:
        find_obj = target_range.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Text = old_text
        find_obj.Replacement.Text = str(new_text)
        if find_obj.Execute(Replace=2): # 2: wdReplaceAll
            found = True
    return found

def process_word_template(dataframe, template_file_path, output_type, progress_callback, save_path=None):
    """메인 프로세스"""
    # 작업 중에는 숨겨서 UI 부하 감소 및 포커스 충돌 방지
    word = get_word_instance(visible=False)

    try:
        if output_type == 'individual':
            return process_individual_word(word, dataframe, template_file_path, progress_callback)
        elif output_type == 'combined':
            return process_combined_word(word, dataframe, template_file_path, progress_callback, save_path)
    finally:
        if output_type == 'individual' or (output_type == 'combined' and not save_path):
            try: word.Quit()
            except: pass
        else:
            try: word.Visible = True # 결과물 확인용
            except: pass

def process_individual_word(word, dataframe, template_file_path, progress_callback):
    output_dir = os.path.dirname(template_file_path)
    base_name = os.path.splitext(os.path.basename(template_file_path))[0]
    ext = os.path.splitext(template_file_path)[1]
    total_rows = len(dataframe)

    for index, row in dataframe.iterrows():
        if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 100))
        doc = safe_open_doc(word, template_file_path)
        try:
            for col in dataframe.columns:
                val = str(row[col]) if pd.notna(row[col]) else ""
                for p in [f'{{{{{col}}}}}', f'{{{col}}}']:
                    replace_text_in_story_ranges(doc, p, val)
            
            out_path = os.path.join(output_dir, f"{base_name}_row_{index+1}{ext}")
            doc.SaveAs(os.path.abspath(out_path))
            doc.Close(0)
        except Exception as e:
            print(f"ERROR: 행 {index+1} 처리 실패: {e}")
            try: doc.Close(0)
            except: pass
    return f"완료: {output_dir}"

def process_combined_word(word, dataframe, template_file_path, progress_callback, save_path):
    total_rows = len(dataframe)
    temp_dir = tempfile.mkdtemp()
    temp_files = []
    
    try:
        for index, row in dataframe.iterrows():
            if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 50))
            doc = safe_open_doc(word, template_file_path)
            try:
                for col in dataframe.columns:
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    for p in [f'{{{{{col}}}}}', f'{{{col}}}']:
                        replace_text_in_story_ranges(doc, p, val)
                t_path = os.path.join(temp_dir, f"temp_{index:04d}.docx")
                doc.SaveAs(os.path.abspath(t_path))
                doc.Close(0)
                temp_files.append(t_path)
            except Exception as e:
                print(f"ERROR: 행 {index} 생성 실패: {e}")
                try: doc.Close(0)
                except: pass

        if not temp_files: raise Exception("생성된 파일 없음")
        combined_doc = safe_open_doc(word, temp_files[0], read_only=False)
        for i in range(1, len(temp_files)):
            if progress_callback: progress_callback.emit(50 + int((i / len(temp_files)) * 50))
            rng = combined_doc.Content
            rng.Collapse(0)
            rng.InsertBreak(7)
            rng = combined_doc.Content
            rng.Collapse(0)
            rng.InsertFile(os.path.abspath(temp_files[i]))
        
        combined_doc.SaveAs(os.path.abspath(save_path))
        return f"통합 완료: {save_path}"
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
