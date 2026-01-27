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

try:
    from PIL import Image
except ImportError:
    Image = None

def ensure_hwp_app():
    """기존 HWP 인스턴스를 얻거나 새로 띄운다."""
    try:
        return win32com.client.GetActiveObject("HWPFrame.HwpObject")
    except Exception:
        try:
            return win32com.client.DispatchEx("HWPFrame.HwpObject")
        except Exception as err:
            print(f"DEBUG: HWP 인스턴스 확보 실패: {err}")
            raise

def get_hwp_instance():
    """HWP 인스턴스를 가져오거나 생성합니다."""
    hwp = None
    try:
        hwp = ensure_hwp_app()
        print("DEBUG: 새 HWP 인스턴스를 생성했습니다.")
        time.sleep(2)
    except Exception as e:
        raise Exception(f"HWP 인스턴스 생성 실패: {e}")
    
    if hwp:
        try:
            hwp.Visible = True
            hwp.XHwpWindows.Active_XHwpWindow.Visible = True
        except Exception as e:
            print(f"DEBUG: HWP 창 표시 설정: {e}")

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except:
            pass
            
        try:
            hwp.SetMessageBoxMode(0x00010000)
            print("DEBUG: HWP 메시지 모드 설정 완료.")
        except Exception as e:
            print(f"DEBUG: HWP 메시지 모드 설정: {e}")

    return hwp


def get_file_format(file_path):
    """파일 확장자에 따라 HWP 형식을 반환합니다."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.hwpx':
        return 'HWPX'
    elif ext == '.hwp':
        return 'HWP'
    else:
        return ''


def _move_cursor_to_document_end(hwp):
    """문서 끝으로 커서를 이동합니다."""
    for action in ("MoveDocEnd", "MoveBottomLevelEnd", "MoveTopLevelEnd"):
        try:
            hwp.HAction.Run(action)
            return True
        except Exception as err:
            print(f"DEBUG: HAction.Run('{action}') 실패: {err}")
    print("WARNING: 문서 끝으로 이동하지 못했습니다.")
    return False


def _get_image_size_mm(image_path):
    """이미지 원본 크기를 mm 단위로 계산합니다."""
    if Image is None:
        print("DEBUG: Pillow 미설치 - 이미지 크기 계산 생략")
        return None, None

    try:
        with Image.open(image_path) as img:
            dpi_x, dpi_y = img.info.get("dpi", (96, 96))
            dpi_x = dpi_x or 96
            dpi_y = dpi_y or 96
            width_mm = int(round((img.width / dpi_x) * 25.4))
            height_mm = int(round((img.height / dpi_y) * 25.4))
            print(f"DEBUG: 원본 이미지 크기(mm) - {width_mm} x {height_mm}")
            return width_mm, height_mm
    except Exception as err:
        print(f"DEBUG: 원본 이미지 크기 계산 실패: {err}")
        return None, None


def insert_image_to_hwp(hwp, image_path):
    """HWP 문서의 현재 커서 위치에 이미지를 삽입합니다."""
    temp_path = None
    try:
        is_valid, message = image_utils.validate_image_path(image_path)
        if not is_valid:
            print(f"WARNING: 이미지 삽입 실패 - {message}")
            return False

        abs_path = os.path.abspath(image_path)
        fd, temp_path = tempfile.mkstemp(suffix=os.path.splitext(abs_path)[1] or ".img")
        os.close(fd)
        shutil.copy2(abs_path, temp_path)
        win_path = temp_path.replace('/', '\\')
        size_option = 3  # 표 비율 유지

        width_mm, height_mm = _get_image_size_mm(abs_path)
        if width_mm is None or height_mm is None:
            width_mm = 0
            height_mm = 0

        print(f"DEBUG: InsertPicture 호출 - sizeOption={size_option}, width_mm={width_mm}, height_mm={height_mm}")

        result = hwp.InsertPicture(
            win_path,
            1,  # embedded
            size_option,
            0,  # reverse
            0,  # watermark
            0,  # effect
            width_mm,
            height_mm,
        )

        if result:
            print(f"DEBUG: 이미지 삽입 성공 (sizeOption={size_option}) - {os.path.basename(image_path)}")
            return True

        print("WARNING: InsertPicture 메서드 실패")
        return False

    except Exception as e:
        print(f"ERROR: 이미지 삽입 중 오류: {e}")
        traceback.print_exc()
        return False
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

def _put_field_text(hwp, field_name, field_value):
    """PutFieldText 또는 ActiveDocument.PutFieldText로 값을 입력한다."""
    putters = []
    putters.append(("HwpObject", getattr(hwp, "PutFieldText", None)))
    active_doc = getattr(hwp, "ActiveDocument", None)
    if active_doc:
        putters.append(("ActiveDocument", getattr(active_doc, "PutFieldText", None)))

    for source, putter in putters:
        if putter is None:
            continue
        try:
            putter(field_name, field_value)
            return True
        except Exception as err:
            print(f"DEBUG: PutFieldText 실패({source}, field='{field_name}'): {err}")
    return False


def _fill_image_field(hwp, field_name, image_path):
    """필드 위치에서 이미지를 삽입한다."""
    placeholder = f"__HWP_IMAGE__{field_name}_{uuid.uuid4().hex}__"

    if not _put_field_text(hwp, field_name, placeholder):
        print(f"WARNING: 이미지 필드 '{field_name}'에 플레이스홀더 입력 실패")
        return False

    try:
        hwp.MovePos(2)
    except Exception:
        pass

    try:
        pset = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("RepeatFind", pset.HSet)
    except Exception as err:
        print(f"DEBUG: 이미지 플레이스홀더 찾기 준비 실패: {err}")
        return False

    pset.FindString = placeholder
    pset.Direction = 1
    pset.IgnoreMessage = 1

    try:
        found = hwp.HAction.Execute("RepeatFind", pset.HSet)
    except Exception as err:
        print(f"DEBUG: 이미지 플레이스홀더 검색 실패: {err}")
        return False

    if not found:
        print(f"WARNING: 이미지 필드 '{field_name}' 플레이스홀더를 찾지 못했습니다.")
        return False

    try:
        hwp.HAction.Run("Delete")
        time.sleep(0.05)
    except Exception as err:
        print(f"DEBUG: 이미지 플레이스홀더 삭제 실패: {err}")
        return False

    return insert_image_to_hwp(hwp, image_path)


def fill_fields_with_find_replace(hwp, dataframe_row):
    """PutFieldText 기반으로 필드를 채우고, 이미지 필드는 플레이스홀더 기반으로 삽입한다."""
    filled = 0
    print("DEBUG: PutFieldText 기반 필드 채우기 시작")
    print(f"DEBUG: 채울 컬럼: {list(dataframe_row.index)}")

    image_queue = []

    for column in dataframe_row.index:
        try:
            raw_value = dataframe_row[column]
            field_value = "" if pd.isna(raw_value) else str(raw_value)

            if field_value and image_utils.is_image_file(field_value):
                image_queue.append((column, field_value))
                continue

            if _put_field_text(hwp, column, field_value):
                display_value = field_value[:30] + "..." if len(field_value) > 30 else field_value
                print(f"DEBUG: '{column}' 필드 텍스트 입력 완료 → '{display_value}'")
                filled += 1
            else:
                print(f"WARNING: '{column}' 필드에 텍스트를 입력하지 못했습니다.")

        except Exception as err:
            print(f"ERROR: 필드 '{column}' 처리 중 오류: {err}")
            traceback.print_exc()

    for column, image_path in image_queue:
        try:
            if _fill_image_field(hwp, column, image_path):
                print(f"DEBUG: '{column}' 필드 이미지 삽입 완료")
                filled += 1
            else:
                print(f"WARNING: '{column}' 필드 이미지 삽입 실패")
        except Exception as err:
            print(f"ERROR: 이미지 필드 '{column}' 처리 중 오류: {err}")
            traceback.print_exc()

    print(f"DEBUG: 필드 채우기 완료 - {filled}개 필드 채움")
    return filled


def remove_all_fields(hwp, progress_callback=None):
    """문서 내의 모든 누름틀(Click-Here) 필드를 삭제합니다. (내용은 유지)"""
    try:
        # 모든 누름틀 필드 목록 가져오기 (고유 이름만)
        field_list = hwp.GetFieldList(0, 2)
        if not field_list:
            print("DEBUG: 삭제할 누름틀 필드가 없습니다.")
            return

        base_fields = [f for f in field_list.split("\x02") if f]
        print(f"DEBUG: 총 {len(base_fields)}종류의 누름틀 필드 삭제 시작")

        total_base = len(base_fields)
        for i, field_name in enumerate(base_fields):
            if progress_callback:
                progress_callback.emit(95 + int((i / total_base) * 4))

            # 1. 우선 DeleteField(이름, 타입) 메서드 시도 (가장 깔끔함)
            # 타입 2: 누름틀. 보통 이 호출로 해당 이름을 가진 모든 인스턴스가 삭제됨.
            try:
                hwp.DeleteField(field_name, 2)
            except Exception as e:
                print(f"DEBUG: DeleteField('{field_name}', 2) 실패: {e}")
                
                # 2. 폴백: MoveToField + DeleteField 액션 방식
                # 매개변수 개수 오류 방지를 위해 유연하게 시도
                count = 0
                while count < 100: # 무한루프 방지
                    found = False
                    try:
                        # 4개 매개변수 시도
                        found = hwp.MoveToField(field_name, True, True, True)
                    except:
                        try:
                            # 1개 매개변수 시도
                            found = hwp.MoveToField(field_name)
                        except:
                            found = False
                    
                    if not found: break
                    
                    try:
                        hwp.HAction.Run("DeleteField")
                        count += 1
                    except:
                        break
            
            # 중간중간 메시지 처리 시간 부여
            if i % 10 == 0:
                time.sleep(0.01)
        
        hwp.MovePos(0)
        print("DEBUG: 모든 누름틀 필드 삭제 작업 완료")
    except Exception as e:
        print(f"DEBUG: 전체 필드 삭제 로직 중 오류: {e}")


def process_hwp_template(dataframe, template_file_path, output_type, progress_callback, save_path=None):
    """HWP 템플릿을 처리합니다."""
    hwp = None

    try:
        hwp = get_hwp_instance()
        if hwp is None:
            raise Exception("한글 COM 객체를 가져오는 데 실패했습니다.")

        if not os.path.exists(template_file_path):
            raise Exception(f"템플릿 파일이 존재하지 않습니다: {template_file_path}")

        file_format = get_file_format(template_file_path)
        
        if output_type == 'individual':
            return process_individual(hwp, dataframe, template_file_path, progress_callback)
        elif output_type == 'combined':
            if not save_path:
                raise ValueError("통합 저장 경로가 지정되지 않았습니다.")
            return process_combined_safe(hwp, dataframe, template_file_path, progress_callback, save_path)
        else:
            raise ValueError(f"알 수 없는 출력 타입: {output_type}")
            
    except Exception as e:
        print(f"ERROR: 처리 중 오류 발생: {e}")
        print(traceback.format_exc())
        raise
    finally:
        if hwp:
            try:
                print("DEBUG: HWP 인스턴스 종료 시작...")
                # 1. 열려있는 모든 문서 닫기
                try:
                    count = hwp.XHwpDocuments.Count
                    for _ in range(count):
                        hwp.XHwpDocuments.Item(0).Close(1) # 무조건 저장 안 함
                        time.sleep(0.1)
                except:
                    pass
                
                # 2. 인스턴스 종료
                try:
                    hwp.Quit()
                except:
                    # Quit 실패 시 프로세스 강제 종료 고려 가능하지만 우선 무시
                    pass
                
                # 3. 파일 락 해제를 위한 충분한 대기
                time.sleep(1.5)
                print("DEBUG: HWP 인스턴스 종료 완료")
            except Exception as e:
                print(f"DEBUG: HWP 종료 시퀀스 중 오류 (무시): {e}")


def process_individual(hwp, dataframe, template_file_path, progress_callback):
    """개별 문서로 저장합니다."""
    output_dir = os.path.dirname(template_file_path)
    base_name = os.path.splitext(os.path.basename(template_file_path))[0]
    ext = os.path.splitext(template_file_path)[1]
    total_rows = len(dataframe)
    file_format = get_file_format(template_file_path)

    print(f"DEBUG: 개별 문서 {total_rows}개 생성 시작")

    for index, row in dataframe.iterrows():
        try:
            # 진행률 업데이트
            if progress_callback:
                progress_callback.emit(int(((index + 1) / total_rows) * 100))

            # 기존 문서 닫기 및 인스턴스 유효성 체크
            try:
                _ = hwp.XHwpWindows.Count
                hwp.Clear(1)
                time.sleep(0.2)
            except Exception as e:
                print(f"DEBUG: HWP 인스턴스 이상 감지 ({e}), 재할당 시도")
                try:
                    hwp = get_hwp_instance()
                except:
                    pass

            # 문서 열기 (재시도 로직 추가)
            abs_template = os.path.abspath(template_file_path)
            opened = False
            for attempt in range(3):
                try:
                    _ = hwp.XHwpWindows.Count
                    result = hwp.Open(abs_template, file_format, "")
                    if result:
                        opened = True
                        break
                except Exception as open_err:
                    print(f"DEBUG: HWP Open 시도 {attempt+1} 실패: {open_err}")
                    if "-2147417851" in str(open_err) or "RPC" in str(open_err):
                        try:
                            hwp = get_hwp_instance()
                        except:
                            pass
                time.sleep(0.8)

            if not opened:
                raise Exception(f"템플릿 파일을 열 수 없습니다 (3회 시도): {template_file_path}")

            time.sleep(0.3)

            filled = fill_fields_with_find_replace(hwp, row)
            
            if index == 0 and filled == 0:
                print("\n⚠️  첫 번째 문서에서 필드가 하나도 채워지지 않았습니다.")
                print("    템플릿 문서를 확인하고 수정한 후 다시 시도하세요.\n")
            
            # 저장 전 누름틀(필드) 삭제
            remove_all_fields(hwp, progress_callback)
            
            # 저장
            output_path = os.path.join(output_dir, f"{base_name}_row_{index+1}{ext}")
            abs_output_path = os.path.abspath(output_path)
            
            result = hwp.SaveAs(abs_output_path, file_format, "")
            if not result:
                raise Exception(f"문서 저장 실패: {abs_output_path}")
            
            print(f"DEBUG: 문서 저장 완료 ({index+1}/{total_rows})")
            
        except Exception as e:
            print(f"ERROR: 행 {index+1} 처리 중 오류: {e}")
            raise

    try:
        hwp.Clear(1)
    except:
        pass
    
    return f"HWP 개별 문서 처리가 완료되었습니다.\n저장 폴더: {output_dir}\n생성된 파일 수: {total_rows}"


def process_combined_safe(hwp, dataframe, template_file_path, progress_callback, save_path):
    """통합 문서로 저장합니다 (복사 붙여넣기 방식)."""
    temp_dir = tempfile.mkdtemp()
    total_rows = len(dataframe)
    file_paths = []
    file_format = get_file_format(template_file_path)

    print(f"DEBUG: Stage 1 - 임시 폴더에 {total_rows}개의 HWP 파일 생성: {temp_dir}")

    try:
        # Stage 1: 개별 파일 생성
        for index, row in dataframe.iterrows():
            try:
                # 진행률 업데이트 (0-50%)
                if progress_callback:
                    progress_callback.emit(int(((index + 1) / total_rows) * 50))

                # 기존 문서 닫기 및 인스턴스 유효성 체크
                try:
                    # 가벼운 호출로 인스턴스 생존 확인
                    _ = hwp.XHwpWindows.Count
                    hwp.Clear(1)
                    time.sleep(0.3)
                except Exception as e:
                    print(f"DEBUG: HWP 인스턴스 이상 감지 ({e}), 재할당 시도")
                    try:
                        hwp = get_hwp_instance()
                    except:
                        pass

                # 문서 열기 (재시도 로직 추가하여 안정성 확보)
                abs_template = os.path.abspath(template_file_path)
                opened = False
                for attempt in range(3):
                    try:
                        # 매 시도 전 인스턴스 확인
                        _ = hwp.XHwpWindows.Count
                        result = hwp.Open(abs_template, file_format, "")
                        if result:
                            opened = True
                            break
                        else:
                            print(f"DEBUG: HWP Open 시도 {attempt+1} 반환값 False")
                    except Exception as open_err:
                        print(f"DEBUG: HWP Open 시도 {attempt+1} 중 예외: {open_err}")
                        # 서버 예외 오류 발생 시 인스턴스 교체 시도
                        if "-2147417851" in str(open_err) or "RPC" in str(open_err):
                            try:
                                print("DEBUG: 치명적 COM 오류 감지 - HWP 인스턴스 재시작")
                                hwp = get_hwp_instance()
                            except:
                                pass
                    
                    time.sleep(0.8) # 실패 시 충분히 대기

                if not opened:
                    raise Exception(f"템플릿 파일을 열 수 없습니다 (3회 시도): {template_file_path}")

                time.sleep(0.3)

                filled = fill_fields_with_find_replace(hwp, row)
                
                if index == 0 and filled == 0:
                    print("\n⚠️  첫 번째 문서에서 필드가 하나도 채워지지 않았습니다.")
                    print("    템플릿 문서를 확인하고 수정한 후 다시 시도하세요.\n")
                
                # 임시 파일로 저장 (hwp 형식으로 저장)
                temp_path = os.path.join(temp_dir, f"temp_{index:04d}.hwp")
                abs_temp_path = os.path.abspath(temp_path)
                
                result = hwp.SaveAs(abs_temp_path, "HWP", "")
                if not result:
                    raise Exception(f"임시 파일 저장 실패: {abs_temp_path}")
                
                file_paths.append(abs_temp_path)
                
                if (index + 1) % 10 == 0 or index + 1 == total_rows:
                    print(f"DEBUG: 임시 파일 생성 ({index+1}/{total_rows})")
                
            except Exception as e:
                print(f"ERROR: 행 {index+1} 처리 중 오류: {e}")
                raise

        # Stage 2: 파일 병합
        print(f"DEBUG: Stage 2 - {len(file_paths)}개 파일 병합 시작")
        
        if not file_paths:
            raise Exception("생성된 임시 파일이 없습니다.")
        
        # 기존 문서 닫기
        try:
            hwp.Clear(1)
            time.sleep(0.2)
        except:
            pass
        
        # 첫 번째 파일 열기
        result = hwp.Open(os.path.abspath(file_paths[0]), "HWP", "")
        if not result:
            raise Exception(f"첫 번째 파일을 열 수 없습니다: {file_paths[0]}")
        time.sleep(0.3)
        
        # 나머지 파일들을 복사해서 붙여넣기
        for i in range(1, len(file_paths)):
            try:
                # 진행률 업데이트 (50-100%)
                if progress_callback:
                    progress_callback.emit(50 + int((i / len(file_paths)) * 50))

                print(f"DEBUG: 파일 {i}/{len(file_paths)-1} 병합 시작: {os.path.basename(file_paths[i])}")

                # Step 1: 첫 번째 문서가 활성화되어 있는지 확인
                try:
                    doc_count_before = hwp.XHwpDocuments.Count
                    print(f"DEBUG: 병합 전 문서 수: {doc_count_before}")
                    if doc_count_before > 1:
                        # 추가 문서가 있다면 모두 닫기
                        for _ in range(doc_count_before - 1):
                            try:
                                hwp.XHwpDocuments.Item(1).Close(1)
                                time.sleep(0.1)
                            except:
                                pass
                except Exception as e:
                    print(f"DEBUG: 문서 상태 확인 실패 (계속 진행): {e}")

                # Step 2: 첫 번째 문서 끝으로 이동
                if not _move_cursor_to_document_end(hwp):
                    print("WARNING: 문서 끝 이동 실패 - 병합 과정은 계속 진행됩니다.")
                time.sleep(0.15)

                # Step 3: 파일 내용을 직접 삽입 (InsertFile은 자동으로 페이지 나누기 처리)
                next_file = os.path.abspath(file_paths[i])
                print(f"DEBUG: 파일 삽입 시도: {next_file}")

                # 대안: InsertFile 메서드 사용
                try:
                    # InsertFile 파라미터 설정
                    pset = hwp.HParameterSet.HInsertFile
                    hwp.HAction.GetDefault("InsertFile", pset.HSet)
                    pset.filename = next_file
                    pset.KeepSection = 1  # 구역 유지
                    pset.KeepCharshape = 1  # 글자 모양 유지
                    pset.KeepParashape = 1  # 문단 모양 유지
                    pset.KeepStyle = 1  # 스타일 유지

                    result = hwp.HAction.Execute("InsertFile", pset.HSet)
                    if result:
                        print(f"DEBUG: 파일 삽입 완료 (InsertFile 방식)")
                        time.sleep(0.2)
                    else:
                        print(f"WARNING: InsertFile 방식 실패, 복사/붙여넣기 방식 시도")
                        raise Exception("InsertFile failed")

                except Exception as insert_error:
                    print(f"DEBUG: InsertFile 실패, 복사/붙여넣기 방식으로 대체: {insert_error}")

                    # 복사/붙여넣기 방식으로 폴백
                    try:
                        # 임시 파일을 새 창으로 열기
                        result = hwp.Open(next_file, "HWP", "")
                        if not result:
                            print(f"ERROR: 파일 {i} 열기 실패, 건너뜀")
                            continue

                        time.sleep(0.4)

                        # 문서 수 확인
                        doc_count = hwp.XHwpDocuments.Count
                        print(f"DEBUG: 현재 열린 문서 수: {doc_count}")

                        if doc_count < 2:
                            print(f"ERROR: 문서가 제대로 열리지 않음 (count: {doc_count})")
                            continue

                        # 방금 연 문서 활성화 (마지막 문서)
                        last_doc_index = doc_count - 1
                        hwp.XHwpDocuments.Item(last_doc_index).SetActive()
                        time.sleep(0.2)
                        print(f"DEBUG: 문서 {last_doc_index} 활성화 완료")

                        # 전체 선택 후 복사
                        hwp.HAction.Run("SelectAll")
                        time.sleep(0.15)
                        hwp.HAction.Run("Copy")
                        time.sleep(0.15)
                        print(f"DEBUG: 내용 복사 완료")

                        # 첫 번째 문서로 전환
                        hwp.XHwpDocuments.Item(0).SetActive()
                        time.sleep(0.2)
                        print(f"DEBUG: 첫 번째 문서로 전환 완료")

                        # 문서 끝으로 이동 후 붙여넣기
                        hwp.MovePos(3)
                        time.sleep(0.1)
                        hwp.HAction.Run("Paste")
                        time.sleep(0.2)
                        print(f"DEBUG: 내용 붙여넣기 완료")

                        # 방금 연 문서 닫기
                        if hwp.XHwpDocuments.Count >= 2:
                            hwp.XHwpDocuments.Item(last_doc_index).Close(1)
                            time.sleep(0.15)
                            print(f"DEBUG: 임시 문서 닫기 완료")

                    except Exception as copy_error:
                        print(f"ERROR: 복사/붙여넣기 방식도 실패: {copy_error}")
                        # 열려있는 추가 문서 정리
                        try:
                            while hwp.XHwpDocuments.Count > 1:
                                hwp.XHwpDocuments.Item(1).Close(1)
                                time.sleep(0.1)
                        except:
                            pass
                        continue

                # Step 11: 진행 상황 출력
                if i % 5 == 0 or i == len(file_paths) - 1:
                    print(f"✅ 파일 병합 진행: {i}/{len(file_paths)-1} ({int(i/len(file_paths)*100)}%)")

            except Exception as e:
                print(f"❌ ERROR: 파일 {i} 병합 중 치명적 오류: {e}")
                traceback.print_exc()
                # 오류 발생 시에도 계속 진행
                continue
        
        # 모든 작업 완료 후 누름틀(필드) 삭제 (최종본 깔끔하게 정리)
        remove_all_fields(hwp, progress_callback)
        
        # 최종 파일 저장
        abs_save_path = os.path.abspath(save_path)
        save_format = get_file_format(save_path)
        result = hwp.SaveAs(abs_save_path, save_format, "")
        if not result:
            raise Exception(f"최종 파일 저장 실패: {abs_save_path}")
        
        print(f"DEBUG: 통합 파일 저장 완료: {save_path}")
        
        return f"HWP 통합 문서 처리가 완료되었습니다.\n저장 경로: {save_path}\n병합된 문서 수: {len(file_paths)}"
        
    finally:
        # 임시 파일 정리
        print("DEBUG: 임시 파일 정리 시작")
        time.sleep(0.5)
        
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
                print(f"DEBUG: 임시 폴더 삭제 완료")
        except Exception as e:
            print(f"WARNING: 임시 폴더 삭제 실패: {temp_dir} - {e}")
