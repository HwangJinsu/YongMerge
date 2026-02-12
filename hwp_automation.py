import os
import time
import win32com.client
from win32com.client import dynamic
import pythoncom
import pandas as pd
import tempfile
import uuid
import traceback
import shutil
import image_utils
from PIL import Image

def ensure_hwp_app():
    """기존 HWP 인스턴스를 얻거나 새로 띄운다."""
    try:
        return win32com.client.GetActiveObject("HWPFrame.HwpObject")
    except Exception:
        try:
            # 새로 생성 시 독립 프로세스 고려 (DispatchEx 지원 안 되는 버전 대비)
            return win32com.client.Dispatch("HWPFrame.HwpObject")
        except Exception as err:
            raise Exception(f"HWP 실행 실패: {err}")

def get_hwp_instance():
    """HWP 인스턴스 초기 설정"""
    hwp = ensure_hwp_app()
    if hwp:
        try:
            hwp.Visible = True
            hwp.SetMessageBoxMode(0x00010000)
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except: pass
    return hwp

def get_file_format(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    return 'HWPX' if ext == '.hwpx' else 'HWP'

def insert_image_to_hwp(hwp, image_path):
    """현재 위치에 이미지 삽입 (비율 유지)"""
    try:
        is_valid, message = image_utils.validate_image_path(image_path)
        if not is_valid: return False

        abs_path = os.path.abspath(image_path)
        width_mm, height_mm = 0, 0
        try:
            with Image.open(abs_path) as img:
                dpi_x, dpi_y = img.info.get("dpi", (96, 96))
                width_mm = int((img.width / (dpi_x or 96)) * 25.4)
                height_mm = int((img.height / (dpi_y or 96)) * 25.4)
        except: pass

        return hwp.InsertPicture(abs_path, 1, 3, 0, 0, 0, width_mm, height_mm)
    except Exception as e:
        print(f"ERROR: HWP 이미지 삽입 중 오류: {e}")
        return False

def fill_fields(hwp, dataframe_row):
    """PutFieldText 기반 필드 채우기"""
    for column in dataframe_row.index:
        val = str(dataframe_row[column]) if pd.notna(dataframe_row[column]) else ""
        if val and image_utils.is_image_file(val):
            # 이미지 필드 처리 로직 (기존 _fill_image_field 활용)
            placeholder = f"__IMG_{uuid.uuid4().hex}__"
            hwp.PutFieldText(column, placeholder)
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = placeholder
            if hwp.HAction.Execute("RepeatFind", pset.HSet):
                hwp.HAction.Run("Delete")
                insert_image_to_hwp(hwp, val)
        else:
            hwp.PutFieldText(column, val)

def remove_all_fields(hwp):
    """문서 내의 모든 누름틀 필드를 삭제합니다. (내용 유지)"""
    try:
        field_list = hwp.GetFieldList(0, 2)
        if not field_list: return
        for field in [f for f in field_list.split("\x02") if f]:
            hwp.DeleteField(field, 2)
    except: pass

def process_hwp_template(dataframe, template_file_path, output_type, progress_callback, save_path=None):
    hwp = get_hwp_instance()
    if not hwp: raise Exception("HWP 실행 실패")

    try:
        if output_type == 'individual':
            return process_individual_hwp(hwp, dataframe, template_file_path, progress_callback)
        elif output_type == 'combined':
            return process_combined_hwp(hwp, dataframe, template_file_path, progress_callback, save_path)
    finally:
        # 작업 완료 후 창을 활성화하여 사용자에게 보여줌 (통합본의 경우)
        if output_type == 'combined':
            try: hwp.Visible = True
            except: pass
        else:
            try: hwp.Quit()
            except: pass

def process_individual_hwp(hwp, dataframe, template_file_path, progress_callback):
    output_dir = os.path.dirname(template_file_path)
    base_name = os.path.splitext(os.path.basename(template_file_path))[0]
    ext = os.path.splitext(template_file_path)[1]
    total_rows = len(dataframe)

    for index, row in dataframe.iterrows():
        if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 100))
        
        hwp.Clear(1)
        time.sleep(0.2)
        
        if not hwp.Open(os.path.abspath(template_file_path), get_file_format(template_file_path), ""):
            continue
            
        fill_fields(hwp, row)
        remove_all_fields(hwp) # 누름틀 제거
        
        out_path = os.path.join(output_dir, f"{base_name}_row_{index+1}{ext}")
        hwp.SaveAs(os.path.abspath(out_path), get_file_format(out_path), "")
        
    return f"HWP 개별 저장 완료: {output_dir}"

def process_combined_hwp(hwp, dataframe, template_file_path, progress_callback, save_path):
    total_rows = len(dataframe)
    temp_dir = tempfile.mkdtemp()
    file_paths = []

    try:
        for index, row in dataframe.iterrows():
            if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 50))
            hwp.Clear(1)
            time.sleep(0.1)
            if not hwp.Open(os.path.abspath(template_file_path), get_file_format(template_file_path), ""):
                continue
            fill_fields(hwp, row)
            t_path = os.path.join(temp_dir, f"temp_{index:04d}.hwp")
            hwp.SaveAs(os.path.abspath(t_path), "HWP", "")
            file_paths.append(t_path)

        if not file_paths: raise Exception("임시 파일 생성 실패")

        hwp.Clear(1)
        time.sleep(0.2)
        hwp.Open(os.path.abspath(file_paths[0]), "HWP", "")
        
        for i in range(1, len(file_paths)):
            if progress_callback: progress_callback.emit(50 + int((i / len(file_paths)) * 50))
            hwp.MovePos(3)
            hwp.HAction.Run("BreakPage")
            pset = hwp.HParameterSet.HInsertFile
            hwp.HAction.GetDefault("InsertFile", pset.HSet)
            pset.filename = os.path.abspath(file_paths[i])
            pset.KeepSection = 1
            hwp.HAction.Execute("InsertFile", pset.HSet)
        
        remove_all_fields(hwp) # 최종본 누름틀 제거
        hwp.SaveAs(os.path.abspath(save_path), get_file_format(save_path), "")
        return f"HWP 통합 완료: {save_path}"
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)