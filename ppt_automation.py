import os
import time
import win32com.client
import pandas as pd
import tempfile
import image_utils
import shutil
from PIL import Image

def get_ppt_instance():
    """PowerPoint 인스턴스를 기존 작업에 방해되지 않게 독립적으로 생성합니다."""
    print("DEBUG: PowerPoint 인스턴스 확보 중...")
    try:
        # DispatchEx를 사용하여 기존 창과 분리된 새 프로세스 생성
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        print("DEBUG: 독립적인 PowerPoint 프로세스를 생성했습니다.")
    except Exception as e:
        raise Exception(f"PowerPoint 실행 실패: {e}")
    
    ppt.Visible = True
    return ppt

def insert_image_to_ppt_from_shape(slide, rectangle_shape, image_path):
    """도형(사각형)의 위치와 크기에 맞춰 이미지를 삽입합니다."""
    try:
        is_valid, message = image_utils.validate_image_path(image_path)
        if not is_valid: return False

        abs_path = os.path.abspath(image_path)
        left = rectangle_shape.Left
        top = rectangle_shape.Top
        width = rectangle_shape.Width
        height = rectangle_shape.Height

        # 원본 비율 유지 계산
        try:
            with Image.open(abs_path) as img:
                img_w, img_h = img.size
                img_ratio = img_w / img_h
                rect_ratio = width / height

                if img_ratio > rect_ratio:
                    final_w = width
                    final_h = width / img_ratio
                else:
                    final_h = height
                    final_w = height * img_ratio

                # 중앙 정렬
                final_left = left + (width - final_w) / 2
                final_top = top + (height - final_h) / 2
        except:
            final_left, final_top, final_w, final_h = left, top, width, height

        slide.Shapes.AddPicture(
            FileName=abs_path, LinkToFile=0, SaveWithDocument=-1,
            Left=final_left, Top=final_top, Width=final_w, Height=final_h
        )
        return True
    except Exception as e:
        print(f"ERROR: PPT 이미지 삽입 오류: {e}")
        return False

def process_ppt_template(dataframe, template_file_path, output_type, progress_callback, save_path=None, image_width=None, image_height=None, debug_mode=False):
    """PPT 자동화 메인 로직"""
    ppt = get_ppt_instance()
    
    try:
        ppt.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        ppt.DisplayAlerts = 0

        if output_type == 'individual':
            return process_individual_ppt(ppt, dataframe, template_file_path, progress_callback)
        elif output_type == 'combined':
            return process_combined_ppt(ppt, dataframe, template_file_path, progress_callback, save_path)
    finally:
        # 개별 작업 시에는 프로세스 종료, 통합본일 경우 사용자가 볼 수 있게 유지할지 여부 판단
        if output_type == 'individual':
            try: ppt.Quit()
            except: pass
        else:
            # 통합본 저장 후에는 Visible 유지 (사용자 확인용)
            pass

def process_individual_ppt(ppt, dataframe, template_file_path, progress_callback):
    output_dir = os.path.dirname(template_file_path)
    base_name = os.path.splitext(os.path.basename(template_file_path))[0]
    total_rows = len(dataframe)

    for index, row in dataframe.iterrows():
        if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 100))
        
        abs_path = os.path.abspath(template_file_path)
        pres = ppt.Presentations.Open(abs_path, Untitled=-1, WithWindow=False)
        
        try:
            for slide in pres.Slides:
                shapes_to_delete = []
                for shape in slide.Shapes:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        txt = shape.TextFrame.TextRange.Text
                        new_txt = txt
                        for col in dataframe.columns:
                            field_val = str(row[col]) if pd.notna(row[col]) else ""
                            for p in [f'{{{{{col}}}}}', f'{{{col}}}']:
                                if p in txt:
                                    if image_utils.is_image_file(field_val):
                                        insert_image_to_ppt_from_shape(slide, shape, field_val)
                                        shapes_to_delete.append(shape)
                                    else:
                                        new_txt = new_txt.replace(p, field_val)
                        
                        if shape not in shapes_to_delete:
                            shape.TextFrame.TextRange.Text = new_txt
                
                for s in shapes_to_delete:
                    try: s.Delete()
                    except: pass
            
            output_file = os.path.join(output_dir, f"{base_name}_row_{index+1}.pptx")
            pres.SaveAs(os.path.abspath(output_file))
            pres.Close()
        except Exception as e:
            print(f"ERROR: 행 {index+1} 처리 중 오류: {e}")
            try: pres.Close()
            except: pass

    return f"INDIVIDUAL_DONE|{output_dir}|{total_rows}"

def process_combined_ppt(ppt, dataframe, template_file_path, progress_callback, save_path):
    total_rows = len(dataframe)
    temp_dir = tempfile.mkdtemp()
    temp_files = []

    try:
        # Stage 1: 임시 파일 생성
        for index, row in dataframe.iterrows():
            if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 50))
            abs_path = os.path.abspath(template_file_path)
            pres = ppt.Presentations.Open(abs_path, Untitled=-1, WithWindow=False)
            
            # (텍스트/이미지 교체 로직은 위와 동일)
            for slide in pres.Slides:
                shapes_to_delete = []
                for shape in slide.Shapes:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        txt = shape.TextFrame.TextRange.Text
                        new_txt = txt
                        for col in dataframe.columns:
                            field_val = str(row[col]) if pd.notna(row[col]) else ""
                            for p in [f'{{{{{col}}}}}', f'{{{col}}}']:
                                if p in txt:
                                    if image_utils.is_image_file(field_val):
                                        insert_image_to_ppt_from_shape(slide, shape, field_val)
                                        shapes_to_delete.append(shape)
                                    else:
                                        new_txt = new_txt.replace(p, field_val)
                        if shape not in shapes_to_delete:
                            shape.TextFrame.TextRange.Text = new_txt
                for s in shapes_to_delete:
                    try: s.Delete()
                    except: pass
            
            t_path = os.path.join(temp_dir, f"temp_{index:04d}.pptx")
            pres.SaveAs(os.path.abspath(t_path))
            pres.Close()
            temp_files.append(t_path)

        # Stage 2: 병합
        combined_pres = ppt.Presentations.Add(WithWindow=True)
        for i, f in enumerate(temp_files):
            if progress_callback: progress_callback.emit(50 + int((i / len(temp_files)) * 50))
            src = ppt.Presentations.Open(os.path.abspath(f), ReadOnly=True, WithWindow=False)
            for s in src.Slides:
                s.Copy()
                combined_pres.Slides.Paste(combined_pres.Slides.Count + 1)
            src.Close()
        
        combined_pres.SaveAs(os.path.abspath(save_path))
        return f"COMBINED_DONE|{save_path}|{len(temp_files)}"
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)