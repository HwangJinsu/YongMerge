import os
import time
import win32com.client
import pandas as pd
import tempfile
import image_utils

def get_ppt_instance():
    print("DEBUG: Attempting to get PowerPoint instance...")
    try:
        os.system("taskkill /f /im POWERPNT.EXE")
        time.sleep(1)
    except Exception as e:
        print(f"DEBUG: Could not kill PowerPoint process (this is often fine): {e}")

    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        print("DEBUG: Successfully created a new PowerPoint instance via Dispatch.")
    except Exception as e:
        raise Exception(f"PowerPoint 인스턴스 생성 실패: {e}")
    
    ppt.Visible = True
    print("DEBUG: PowerPoint instance is now visible.")
    return ppt

def insert_image_to_ppt(slide, placeholder_shape, image_path, width_cm=None, height_cm=None):
    """
    PowerPoint 슬라이드의 placeholder 위치에 이미지를 삽입합니다.

    Args:
        slide: PowerPoint Slide 객체
        placeholder_shape: 이미지를 삽입할 위치의 Shape 객체
        image_path: 삽입할 이미지 파일 경로
        width_cm: 이미지 너비 (cm), None이면 기본값 10cm 사용
        height_cm: 이미지 높이 (cm), None이면 기본값 7.5cm 사용

    Returns:
        bool: 성공 여부
    """
    try:
        # 이미지 파일 유효성 검증
        is_valid, message = image_utils.validate_image_path(image_path)
        if not is_valid:
            print(f"WARNING: 이미지 삽입 실패 - {message}")
            return False

        abs_path = os.path.abspath(image_path)

        # placeholder의 위치와 크기 정보 가져오기
        left = placeholder_shape.Left
        top = placeholder_shape.Top

        # None이면 기본값 사용
        if width_cm is None:
            width_cm = 10  # 기본값 10cm
        if height_cm is None:
            height_cm = 7.5  # 기본값 7.5cm

        # cm를 포인트로 변환 (1cm = 28.35 포인트)
        width_pt = width_cm * 28.35
        height_pt = height_cm * 28.35

        print(f"DEBUG: 이미지 크기: {width_cm}cm x {height_cm}cm ({width_pt}pt x {height_pt}pt)")

        # 이미지 삽입
        picture = slide.Shapes.AddPicture(
            FileName=abs_path,
            LinkToFile=0,  # 링크 안 함 (문서에 포함)
            SaveWithDocument=-1,  # 문서와 함께 저장
            Left=left,
            Top=top,
            Width=width_pt,
            Height=height_pt
        )

        print(f"DEBUG: 이미지 삽입 성공: {os.path.basename(image_path)} at ({left}, {top}), {width_cm}cm x {height_cm}cm")
        return True

    except Exception as e:
        print(f"ERROR: 이미지 삽입 중 오류: {e}")
        import traceback
        traceback.print_exc()
        return False

def insert_image_to_ppt_from_shape(slide, rectangle_shape, image_path):
    """
    PowerPoint 슬라이드의 사각형 shape를 기반으로 이미지를 삽입합니다.
    사각형의 크기와 위치를 읽어 동일한 위치에 같은 크기로 이미지를 삽입합니다.
    (원본 이미지 비율 유지하면서 사각형 내부에 맞춤)

    Args:
        slide: PowerPoint Slide 객체
        rectangle_shape: 기준이 되는 사각형 Shape 객체
        image_path: 삽입할 이미지 파일 경로

    Returns:
        bool: 성공 여부
    """
    try:
        # 이미지 파일 유효성 검증
        is_valid, message = image_utils.validate_image_path(image_path)
        if not is_valid:
            print(f"WARNING: 이미지 삽입 실패 - {message}")
            return False

        abs_path = os.path.abspath(image_path)

        # 사각형의 위치와 크기 정보 가져오기 (포인트 단위)
        left_pt = rectangle_shape.Left
        top_pt = rectangle_shape.Top
        width_pt = rectangle_shape.Width
        height_pt = rectangle_shape.Height

        # 포인트를 cm로 변환하여 로그 출력 (1cm = 28.35 포인트)
        left_cm = left_pt / 28.35
        top_cm = top_pt / 28.35
        width_cm = width_pt / 28.35
        height_cm = height_pt / 28.35

        print(f"DEBUG: 사각형 기반 이미지 삽입")
        print(f"DEBUG: 사각형 위치 및 크기: ({left_cm:.2f}cm, {top_cm:.2f}cm), {width_cm:.2f}cm x {height_cm:.2f}cm")

        # 이미지 원본 비율을 유지하면서 사각형 내부에 맞추기
        from PIL import Image
        try:
            with Image.open(abs_path) as img:
                img_width, img_height = img.size
                img_ratio = img_width / img_height
                rect_ratio = width_pt / height_pt

                # 비율 비교하여 최적 크기 계산
                if img_ratio > rect_ratio:
                    # 이미지가 더 넓음 → 너비를 사각형에 맞추고 높이 조정
                    final_width_pt = width_pt
                    final_height_pt = width_pt / img_ratio
                else:
                    # 이미지가 더 높음 → 높이를 사각형에 맞추고 너비 조정
                    final_height_pt = height_pt
                    final_width_pt = height_pt * img_ratio

                # 중앙 정렬을 위한 위치 조정
                final_left_pt = left_pt + (width_pt - final_width_pt) / 2
                final_top_pt = top_pt + (height_pt - final_height_pt) / 2

                print(f"DEBUG: 원본 이미지 비율: {img_ratio:.2f}, 사각형 비율: {rect_ratio:.2f}")
                print(f"DEBUG: 최종 이미지 크기: {final_width_pt / 28.35:.2f}cm x {final_height_pt / 28.35:.2f}cm")

        except Exception as e:
            # PIL 실패 시 사각형 크기 그대로 사용
            print(f"WARNING: 이미지 비율 계산 실패, 사각형 크기 그대로 사용: {e}")
            final_left_pt = left_pt
            final_top_pt = top_pt
            final_width_pt = width_pt
            final_height_pt = height_pt

        # 이미지 삽입
        picture = slide.Shapes.AddPicture(
            FileName=abs_path,
            LinkToFile=0,  # 링크 안 함 (문서에 포함)
            SaveWithDocument=-1,  # 문서와 함께 저장
            Left=final_left_pt,
            Top=final_top_pt,
            Width=final_width_pt,
            Height=final_height_pt
        )

        print(f"DEBUG: 사각형 기반 이미지 삽입 성공: {os.path.basename(image_path)}")
        return True

    except Exception as e:
        print(f"ERROR: 사각형 기반 이미지 삽입 중 오류: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_ppt_template(dataframe, template_file_path, output_type, progress_callback, save_path=None, image_width=None, image_height=None, debug_mode=True):
    """PowerPoint 템플릿을 처리합니다.

    Args:
        dataframe: 데이터프레임
        template_file_path: 템플릿 파일 경로
        output_type: 출력 타입 ('individual' 또는 'combined')
        progress_callback: 진행률 콜백
        save_path: 저장 경로 (combined일 때 필수)
        image_width: 이미지 너비 (cm), None이면 기본값 사용
        image_height: 이미지 높이 (cm), None이면 기본값 사용
        debug_mode: 디버그 모드 활성화 여부
    """
    print(f"DEBUG: Starting process_ppt_template. Output type: {output_type}, Image size: {image_width}cm x {image_height}cm")
    ppt = get_ppt_instance()
    if ppt is None: raise Exception("PowerPoint COM 객체를 가져오는 데 실패했습니다.")

    try:
        # msoAutomationSecurityForceDisable = 3
        # This is a critical step to prevent PowerPoint's "Protected View"
        # from blocking the automation of untrusted/downloaded files.
        print("DEBUG: Setting AutomationSecurity to Force-Disable...")
        ppt.AutomationSecurity = 3
        # ppAlertsNone = 0. This prevents any pop-up dialogs from blocking automation.
        print("DEBUG: Disabling PowerPoint display alerts...")
        ppt.DisplayAlerts = 0
    except Exception as e:
        print(f"WARNING: Could not set security/alert settings (this might cause issues): {e}")

    try:
        if output_type == 'individual':
            return process_individual_ppt(ppt, dataframe, template_file_path, progress_callback, image_width, image_height)
        elif output_type == 'combined':
            if not save_path:
                raise ValueError("통합 저장 경로가 지정되지 않았습니다.")
            return process_combined_safe_ppt(ppt, dataframe, template_file_path, progress_callback, save_path, image_width, image_height, debug_mode=debug_mode)
    finally:
        print("DEBUG: PPT automation finished. Quitting is not performed by default.")


def process_individual_ppt(ppt, dataframe, template_file_path, progress_callback, image_width=None, image_height=None):
    """개별 PPT 파일로 저장합니다."""
    output_dir = os.path.dirname(template_file_path)
    base_name = os.path.splitext(os.path.basename(template_file_path))[0]
    total_rows = len(dataframe)
    print(f"DEBUG: Starting individual PPT creation for {total_rows} rows.")

    for index, row in dataframe.iterrows():
        print(f"DEBUG: Processing row {index + 1}/{total_rows}")
        if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 100))

        print(f"DEBUG: Opening template: {template_file_path}")
        try:
            print(f"DEBUG: Opening template for row {index+1}: {template_file_path}")
            # Normalize path and open as a copy (Untitled=-1) for robustness
            normalized_path = os.path.abspath(template_file_path)
            print(f"DEBUG: Normalized path: {normalized_path}")
            presentation = ppt.Presentations.Open(normalized_path, Untitled=-1, WithWindow=False)
            print(f"DEBUG: Template opened successfully as a copy for row {index+1}.")
        except Exception as e:
            print(f"\nFATAL: Failed to open the template file: {template_file_path}")
            print(f"FATAL: This is often caused by PowerPoint's 'Protected View' or file corruption.")
            print(f"FATAL: The error was: {e}")
            raise e

        print("DEBUG: Replacing text and images in slides...")
        for slide in presentation.Slides:
            shapes_to_delete = []  # 이미지로 대체할 shape 목록
            images_to_insert = []  # 삽입할 이미지 정보 목록

            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text_frame = shape.TextFrame
                    original_text = text_frame.TextRange.Text
                    new_text = original_text

                    for column in dataframe.columns:
                        # 두 가지 패턴 시도: {{필드명}}, {필드명}
                        patterns = [
                            f'{{{{{column}}}}}',  # {{필드명}}
                            f'{{{column}}}'       # {필드명}
                        ]

                        for placeholder in patterns:
                            if placeholder in original_text:
                                field_value = str(row[column]) if pd.notna(row[column]) else ""

                                # 이미지 파일인지 확인
                                if image_utils.is_image_file(field_value):
                                    # {{이미지}} 패턴인 경우 사각형 기반 삽입 사용
                                    use_shape_size = (column == "이미지" and placeholder == "{{이미지}}")

                                    images_to_insert.append({
                                        'shape': shape,
                                        'path': field_value,
                                        'use_shape_size': use_shape_size
                                    })
                                    shapes_to_delete.append(shape)

                                    if use_shape_size:
                                        print(f"DEBUG: '{placeholder}' 를 이미지로 대체 예약 (사각형 크기/위치 사용): {os.path.basename(field_value)}")
                                    else:
                                        print(f"DEBUG: '{placeholder}' 를 이미지로 대체 예약: {os.path.basename(field_value)}")
                                    break  # 이 shape는 삭제될 것이므로 더 이상 처리 안 함
                                else:
                                    # 텍스트 대체
                                    new_text = new_text.replace(placeholder, field_value)
                                    print(f"DEBUG: Replaced '{placeholder}' with '{field_value[:50] if len(field_value) > 50 else field_value}'")
                                break  # 패턴 찾았으면 다음 컬럼으로

                    # 텍스트만 변경된 경우 적용
                    if new_text != original_text and shape not in shapes_to_delete:
                        text_frame.TextRange.Text = new_text

            # 이미지 삽입 (사각형 크기/위치 또는 사용자 지정 크기 사용)
            for img_info in images_to_insert:
                if img_info.get('use_shape_size'):
                    # 사각형의 크기와 위치를 기반으로 삽입
                    insert_image_to_ppt_from_shape(slide, img_info['shape'], img_info['path'])
                else:
                    # 기존 방식: 사용자 지정 크기 사용
                    insert_image_to_ppt(slide, img_info['shape'], img_info['path'], width_cm=image_width, height_cm=image_height)

            # 플레이스홀더 shape 삭제
            for shape in shapes_to_delete:
                try:
                    shape.Delete()
                except Exception as e:
                    print(f"DEBUG: Shape 삭제 실패 (무시): {e}")
        
        output_path = os.path.join(output_dir, f"{base_name}_row_{index+1}{os.path.splitext(template_file_path)[1]}")
        print(f"DEBUG: Saving individual file to: {output_path}")
        presentation.SaveAs(output_path)
        presentation.Close()
    
    return f"PPT 개별 문서 처리가 완료되었습니다.\n저장 폴더: {output_dir}"

def process_combined_safe_ppt(ppt, dataframe, template_file_path, progress_callback, save_path, image_width=None, image_height=None, debug_mode=True):
    """통합 PPT 파일로 저장합니다."""
    temp_dir = tempfile.mkdtemp()
    total_rows = len(dataframe)
    file_paths = []
    failing_file = "N/A"

    print(f"DEBUG: Stage 1: Creating {total_rows} individual PPT files in temp folder: {temp_dir}")
    for index, row in dataframe.iterrows():
        if progress_callback: progress_callback.emit(int(((index + 1) / total_rows) * 50))

        try:
            print(f"DEBUG: Opening template for row {index+1}: {template_file_path}")
            # Normalize path and open as a copy (Untitled=-1) for robustness
            normalized_path = os.path.abspath(template_file_path)
            print(f"DEBUG: Normalized path: {normalized_path}")
            presentation = ppt.Presentations.Open(normalized_path, Untitled=-1, WithWindow=False)
            print(f"DEBUG: Template opened successfully as a copy for row {index+1}.")
        except Exception as e:
            print(f"\nFATAL: Failed to open the template file: {template_file_path}")
            print(f"FATAL: This is often caused by PowerPoint's 'Protected View' or file corruption.")
            print(f"FATAL: The error was: {e}")
            raise e
        for slide in presentation.Slides:
            shapes_to_delete = []
            images_to_insert = []

            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text_frame = shape.TextFrame
                    original_text = text_frame.TextRange.Text
                    new_text = original_text

                    for column in dataframe.columns:
                        patterns = [f'{{{{{column}}}}}', f'{{{column}}}']

                        found = False
                        for placeholder in patterns:
                            if placeholder in original_text:
                                field_value = str(row[column]) if pd.notna(row[column]) else ""
                                
                                if image_utils.is_image_file(field_value):
                                    use_shape_size = (column == "이미지" and placeholder == "{{이미지}}")
                                    images_to_insert.append({'shape': shape, 'path': field_value, 'use_shape_size': use_shape_size})
                                    shapes_to_delete.append(shape)
                                    if use_shape_size:
                                        print(f"DEBUG: '{placeholder}' image replacement scheduled (using shape size): {os.path.basename(field_value)}")
                                    else:
                                        print(f"DEBUG: '{placeholder}' image replacement scheduled: {os.path.basename(field_value)}")
                                    found = True
                                    break
                                else:
                                    new_text = new_text.replace(placeholder, field_value)
                                    print(f"DEBUG: Replaced '{placeholder}' with '{field_value[:50] if len(field_value) > 50 else field_value}'")
                                found = True
                                break
                        
                        if found and any(img['shape'] == shape for img in images_to_insert):
                            break

                    if new_text != original_text and shape not in shapes_to_delete:
                        text_frame.TextRange.Text = new_text

            for img_info in images_to_insert:
                if img_info.get('use_shape_size'):
                    insert_image_to_ppt_from_shape(slide, img_info['shape'], img_info['path'])
                else:
                    insert_image_to_ppt(slide, img_info['shape'], img_info['path'], width_cm=image_width, height_cm=image_height)

            for shape in shapes_to_delete:
                try:
                    shape.Delete()
                except Exception as e:
                    print(f"DEBUG: Shape deletion failed (ignored): {e}")
        
        temp_path = os.path.join(temp_dir, f"temp_{index}.pptx")
        presentation.SaveAs(temp_path)
        presentation.Close()
        file_paths.append(temp_path)

    print(f"DEBUG: Stage 2: Merging {len(file_paths)} files into one using slide copy/paste method.")
    
    combined_pres = ppt.Presentations.Add(WithWindow=False)
    
    try:
        for i, file_path in enumerate(file_paths):
            failing_file = file_path
            if progress_callback:
                progress_callback.emit(50 + int(((i + 1) / len(file_paths)) * 50))
            
            print(f"DEBUG: Opening {os.path.basename(file_path)} to copy slides...")
            source_pres = ppt.Presentations.Open(file_path, ReadOnly=True, WithWindow=False)
            
            print(f"DEBUG: Source presentation has {source_pres.Slides.Count} slides.")
            for slide_index, slide in enumerate(source_pres.Slides, 1):
                print(f"DEBUG: Copying slide {slide_index} from {os.path.basename(file_path)}...")
                slide.Copy()
                
                print(f"DEBUG: Pasting slide into combined presentation (current slide count: {combined_pres.Slides.Count})...")
                combined_pres.Slides.Paste(combined_pres.Slides.Count + 1)
                print(f"DEBUG: Paste successful. Combined presentation now has {combined_pres.Slides.Count} slides.")

            source_pres.Close()
    except Exception as e:
        print("\n" + "="*60)
        print("FATAL: An error occurred during the slide merge process.")
        print(f"FATAL: The operation failed while processing the temporary file:")
        print(f"FATAL: >>> {os.path.basename(failing_file)} <<<")
        print(f"FATAL: This file and others are located in the temporary directory:")
        print(f"FATAL: >>> {temp_dir} <<<")
        print("FATAL: Please inspect the failing file to identify problematic content.")
        print("="*60 + "\n")
        raise e

    try:
        # Ensure the path is absolute and in the correct format for the OS
        normalized_save_path = os.path.abspath(save_path)
        print(f"DEBUG: Saving final merged presentation to (normalized): {normalized_save_path}")
        
        # ppSaveAsOpenXMLPresentation = 24
        # Explicitly setting the file format can improve reliability.
        combined_pres.SaveAs(normalized_save_path, FileFormat=24)
        print(f"DEBUG: SaveAs command completed.")
        
        combined_pres.Close()
        print(f"DEBUG: Closed the combined presentation object.")

    except Exception as e:
        print(f"\nFATAL: Failed to save the final merged presentation to: {save_path}")
        print(f"FATAL: This can be caused by file permissions or an invalid path.")
        print(f"FATAL: The error was: {e}")
        raise e

    if not debug_mode:
        print(f"INFO: Cleaning up temporary directory: {temp_dir}")
        for path in file_paths:
            try:
                os.remove(path)
            except OSError as e:
                print(f"WARNING: Could not remove temp file {path}: {e}")
        try:
            os.rmdir(temp_dir)
        except OSError as e:
            print(f"WARNING: Could not remove temp dir {temp_dir}: {e}")
    else:
        print(f"INFO: Debug mode is ON. Temporary files are preserved in: {temp_dir}")

    return f"PPT 통합 문서 처리가 완료되었습니다.\n저장 경로: {save_path}"
