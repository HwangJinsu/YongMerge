import sys
import random
import platform
import pythoncom
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLineEdit, 
    QHBoxLayout, QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView, 
    QFileDialog, QMessageBox, QLabel, QSizePolicy, QScrollArea, QFrame, QInputDialog,
    QProgressDialog
)
from PyQt5.QtCore import Qt, QMimeData, QEvent, pyqtSignal, QThread
from PyQt5.QtGui import QDrag, QPixmap, QKeySequence, QFontDatabase, QFont, QPalette, QColor, QBrush
import openpyxl
import pandas as pd
import time
import os
import winreg
import win32com.client
from win32com.client import dynamic

# --- Custom Automation Modules ---
import hwp_automation
import ppt_automation
import image_utils

# --- Windows specific imports for UI interaction ---
is_windows = platform.system() == "Windows"
if is_windows:
    try:
        import win32gui
        import win32api
        import win32con
        print("pywin32 module for UI interaction imported successfully.")
    except ImportError:
        print("Warning: pywin32 module not found. HWP/PPT automation will not be available.")
        is_windows = False
else:
    print("Running on non-Windows OS. HWP/PPT automation is not supported.")

# --- Worker Thread for Asynchronous Automation ---
class AutomationWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, str, str) # Pass (success message, output_type, file_path)
    error = pyqtSignal(str)

    def __init__(self, doc_type, dataframe, template_path, output_type, save_path=None):
        super().__init__()
        self.doc_type = doc_type
        self.dataframe = dataframe
        self.template_path = template_path
        self.output_type = output_type
        self.save_path = save_path

    def run(self):
        try:
            pythoncom.CoInitialize()
            result_message = ""
            if self.doc_type == 'hwp':
                result_message = hwp_automation.process_hwp_template(
                    self.dataframe, self.template_path, self.output_type, self.progress, self.save_path
                )
            elif self.doc_type == 'ppt':
                result_message = ppt_automation.process_ppt_template(
                    self.dataframe, self.template_path, self.output_type, self.progress, self.save_path, debug_mode=True
                )

            # finished 시그널에 (메시지, 출력타입, 파일경로) 전달
            output_file = self.save_path if self.output_type == 'combined' else None
            self.finished.emit(result_message, self.output_type, output_file)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            pythoncom.CoUninitialize()

# List of pleasant colors for field buttons
FIELD_COLORS = [
    "#AEC6CF", "#77DD77", "#FDFD96", "#FFB347", "#B39EB5", "#FF6961", "#CFCFC4", "#8A9A5B",
    "#F49AC2", "#DEAEDC", "#FAB2E1", "#C1E1C1", "#FFD1A3", "#ADD8E6", "#F08080", "#E6E6FA"
]

# Custom button class for dragging
class DraggableButton(QPushButton):
    singleClicked = pyqtSignal(str)

    def __init__(self, text, color, parent=None):
        super().__init__(text, parent)
        self.setAcceptDrops(False)
        self.setStyleSheet("""
            QPushButton {{
                background-color: {color};
                border: 1px solid #555;
                border-radius: 10px;
                padding: 5px 10px;
                margin: 2px;
                color: #333;
                font-weight: bold;
                min-width: 50px;
            }}
            QPushButton:hover {{
                background-color: {color_hover};
            }}
        """.format(color=color, color_hover=color + "CC"))
        self.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Fixed)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.pos()
            self._is_dragging = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            distance = (event.pos() - self.drag_start_position).manhattanLength()
            if not self._is_dragging and distance > QApplication.startDragDistance():
                self._is_dragging = True
                self.perform_drag()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
         if event.button() == Qt.LeftButton and not self._is_dragging:
              # 싱글클릭으로 필드 삽입
              self.singleClicked.emit(self.text())
         self._is_dragging = False
         super().mouseReleaseEvent(event)

    def perform_drag(self):
        mime_data = QMimeData()
        mime_data.setText(f'{{{self.text()}}}')
        drag = QDrag(self)
        pixmap = QPixmap(self.size())
        self.render(pixmap)
        drag.setPixmap(pixmap)
        drag.setMimeData(mime_data)
        drag.exec_()

# Custom TableWidget for enhanced interaction
class EnhancedTableWidget(QTableWidget):
    cellDataChangedSignal = pyqtSignal(int, int, str)
    rowsChangedSignal = pyqtSignal()
    imageColumnDoubleClicked = pyqtSignal(int, int)  # row, column 시그널 추가
    pastedSignal = pyqtSignal() # 붙여넣기 완료 시그널 추가

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
        self.setTextElideMode(Qt.ElideNone)
        self.setWordWrap(False)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.horizontalHeader().setMinimumSectionSize(120)
        default_row_height = max(int(self.fontMetrics().height() * 3), 34)
        self.verticalHeader().setDefaultSectionSize(default_row_height)
        self.horizontalHeader().setStyleSheet("font-weight: bold;")
        self.verticalHeader().setStyleSheet("font-weight: bold;")
        self.cellChanged.connect(self._on_cell_changed)
        self.dataframe_ref = None
        self.cellDoubleClicked.connect(self._on_cell_double_clicked)

    def setDataFrame(self, dataframe):
         self.dataframe_ref = dataframe
         self.update_table_from_dataframe()

    def updateDataFrameRef(self, dataframe):
        """테이블 다시 그리기 없이 DataFrame 참조만 업데이트 (행 추가/삭제 시 사용)"""
        self.dataframe_ref = dataframe

    def selectionChanged(self, selected, deselected):
        super().selectionChanged(selected, deselected)
        black_brush = QBrush(QColor(0, 0, 0))
        for index in selected.indexes():
            item = self.item(index.row(), index.column())
            if item:
                item.setForeground(black_brush)
        for index in deselected.indexes():
            item = self.item(index.row(), index.column())
            if item:
                item.setForeground(black_brush)

    def update_table_from_dataframe(self):
         if self.dataframe_ref is None: return
         self.cellChanged.disconnect(self._on_cell_changed)
         self.setRowCount(self.dataframe_ref.shape[0])
         self.setColumnCount(self.dataframe_ref.shape[1])
         self.setHorizontalHeaderLabels(self.dataframe_ref.columns.tolist())
         self.horizontalHeader().update()
         for r in range(self.dataframe_ref.shape[0]):
             for c in range(self.dataframe_ref.shape[1]):
                 value = self.dataframe_ref.iloc[r, c]
                 col_name = self.dataframe_ref.columns[c]

                 # 이미지 열인 경우 표시 이름으로 변환
                 if col_name == "이미지" and pd.notna(value) and str(value).strip():
                     import image_utils
                     if image_utils.is_image_file(str(value)):
                         display_value = image_utils.get_image_display_name(str(value))
                     else:
                         display_value = str(value)
                 else:
                     display_value = str(value) if pd.notna(value) else ""

                 self.setItem(r, c, QTableWidgetItem(display_value))
         self.cellChanged.connect(self._on_cell_changed)

    def _on_cell_changed(self, row, column):
        """셀 변경 이벤트 - 이미지 열은 표시 텍스트가 아닌 실제 경로만 변경"""
        item = self.item(row, column)
        if item:
            # 이미지 열도 사용자가 직접 경로를 수정할 수 있도록 허용 (이전에는 차단됨)
            # 일반 텍스트 열 및 이미지 열 모두 시그널 발생
            self.cellDataChangedSignal.emit(row, column, item.text())

    def _on_cell_double_clicked(self, row, column):
        """셀 더블클릭 이벤트 핸들러"""
        if self.dataframe_ref is None:
            return

        # 이미지 열인지 확인
        col_name = self.horizontalHeaderItem(column).text()
        if col_name == "이미지":
            # 이미지 열은 편집 불가, 대신 이미지 선택 다이얼로그 실행
            item = self.item(row, column)
            if item:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # 편집 불가 설정
            # 이미지 선택 다이얼로그 시그널 발생
            self.imageColumnDoubleClicked.emit(row, column)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.delete_selected_cells()
        elif event.matches(QKeySequence.Copy):
            self.copy_selected_cells()
        elif event.matches(QKeySequence.Paste):
             self.paste_to_selected_cells()
        else:
            super().keyPressEvent(event)

    def delete_selected_cells(self):
        selected_items = self.selectedItems()
        if not selected_items: return
        self.cellChanged.disconnect(self._on_cell_changed)
        for item in selected_items:
            item.setText("")
            row, col = item.row(), item.column()
            if self.dataframe_ref is not None:
                 col_name = self.horizontalHeaderItem(col).text()
                 self.dataframe_ref.at[row, col_name] = None
        self.cellChanged.connect(self._on_cell_changed)

    def copy_selected_cells(self):
        selected_ranges = self.selectedRanges()
        if not selected_ranges: return
        all_rows_data = []
        for selected_range in selected_ranges:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                row_data = []
                for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    # 데이터프레임이 있으면 실제 값을 복사
                    if self.dataframe_ref is not None:
                        col_name = self.horizontalHeaderItem(col).text()
                        val = self.dataframe_ref.at[row, col_name]
                        if pd.isna(val):
                            val = ""
                        else:
                            val = str(val)
                        row_data.append(val)
                    else:
                        # 데이터프레임이 없으면 화면 텍스트 복사 (폴백)
                        item = self.item(row, col)
                        row_data.append(item.text() if item else "")
                all_rows_data.append("\t".join(row_data))
        QApplication.clipboard().setText("\n".join(all_rows_data))

    def paste_to_selected_cells(self):
        clipboard_text = QApplication.clipboard().text()
        if not clipboard_text: return
        rows_data = clipboard_text.split('\n')
        selected_items = self.selectedItems()
        if not selected_items: return
        
        top_row = min(item.row() for item in selected_items)
        left_col = min(item.column() for item in selected_items)
        
        # 시그널 차단 (대량 업데이트 효율성 및 중복 시그널 방지)
        self.cellChanged.disconnect(self._on_cell_changed)
        
        try:
            for r_offset, row_data in enumerate(rows_data):
                if not row_data: continue
                cells_data = row_data.split('\t')
                for c_offset, cell_data in enumerate(cells_data):
                    target_row, target_col = top_row + r_offset, left_col + c_offset
                    if target_row < self.rowCount() and target_col < self.columnCount():
                        col_name = self.horizontalHeaderItem(target_col).text()
                        
                        # 실제 저장할 값
                        actual_value = cell_data if cell_data else None
                        # 화면에 표시할 값
                        display_value = cell_data
                        
                        # 이미지 열 처리
                        if col_name == "이미지" and actual_value:
                            import image_utils
                            # 값이 이미지 파일 경로인 경우 표시 이름 변경
                            if image_utils.is_image_file(actual_value):
                                display_value = image_utils.get_image_display_name(actual_value)
                        
                        # 테이블 위젯 아이템 업데이트
                        new_item = QTableWidgetItem(display_value if display_value else "")
                        self.setItem(target_row, target_col, new_item)
                        
                        # 데이터프레임 업데이트
                        if self.dataframe_ref is not None:
                            self.dataframe_ref.at[target_row, col_name] = actual_value
        finally:
            self.cellChanged.connect(self._on_cell_changed)
            # 붙여넣기 완료 시그널 발생 (버튼 상태 갱신 등)
            self.pastedSignal.emit()

# Main app class
class MailMergeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.dataframe = pd.DataFrame()
        self.template_file_path = None
        self.worker = None
        self.hwp_app = None
        self.initUI()
        self.load_initial_data()
        self.data_table.cellDataChangedSignal.connect(self.update_dataframe_from_cell)
        self.data_table.rowsChangedSignal.connect(self.handle_table_rows_changed)
        self.data_table.imageColumnDoubleClicked.connect(self.on_image_cell_double_clicked)
        self.check_hwp_registry()

    def check_hwp_registry(self):
        if not is_windows: return
        try:
            winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "HWPFrame.HwpObject")
        except FileNotFoundError:
            try:
                winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "HWP.Application")
            except FileNotFoundError:
                print("⚠️ 한글 COM 등록 정보를 찾을 수 없습니다.")

    def initUI(self):
        self.setWindowTitle('✨ 용merge ✨')
        self.setGeometry(100, 100, 1180, 840)
        font_id = QFontDatabase.addApplicationFont("PretendardVariable.ttf")
        if font_id != -1:
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        else:
            font_family = "Segoe UI"
        base_font = QFont(font_family, 12)
        self.setFont(base_font)
        self.setStyleSheet(f"""
            QWidget {{ background-color: #FFFFFF; font-family: '{font_family}'; color: #1E1E1E; font-size: 15px; }}
            QLabel.title {{ font-size: 20px; font-weight: 600; color: #202020; }}
            QLabel.subtitle {{ font-size: 18px; font-weight: 600; color: #2D2F33; }}
            QLabel {{ font-size: 15px; color: #42454D; }}
            QLineEdit {{ border: 1px solid #C2C7CF; border-radius: 8px; padding: 12px; font-size: 15px; }}
            QTableWidget {{ background: #FFFFFF; gridline-color: #E1E4E8; font-size: 14px; selection-background-color: #E8F1FF; }}
            QTableWidget::item:selected {{ color: #000000; }}
        """)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)

        # 상단 메뉴
        menubar = self.menuBar()
        info_menu = menubar.addMenu("정보")
        license_action = info_menu.addAction("오픈소스 라이선스")
        license_action.triggered.connect(self.show_open_source_info)
        
        field_creation_layout = QHBoxLayout()
        self.field_name_input = QLineEdit(placeholderText="새 필드 이름 입력 후 Enter")
        self.field_name_input.setFixedHeight(48)
        self.field_name_input.returnPressed.connect(self.create_field)
        self.create_field_button = self._make_secondary_button("➕ 필드 생성")
        self.create_field_button.clicked.connect(self.create_field)
        field_creation_layout.addWidget(self._styled_label("필드 관리:", css_class="subtitle"))
        field_creation_layout.addWidget(self.field_name_input)
        field_creation_layout.addWidget(self.create_field_button)
        main_layout.addLayout(field_creation_layout)

        field_list_frame = QFrame()
        field_list_frame.setFrameShape(QFrame.StyledPanel)
        field_list_frame_layout = QVBoxLayout(field_list_frame)
        field_list_frame_layout.addWidget(self._styled_label("🏷️ 사용 가능한 필드 (클릭하여 문서에 삽입):", css_class="subtitle"))
        field_buttons_container = QWidget()
        self.available_fields_layout = QHBoxLayout(field_buttons_container)
        self.available_fields_layout.setAlignment(Qt.AlignLeft)
        self.available_fields_layout.setSpacing(6)
        self.available_fields_layout.setContentsMargins(12, 2, 12, 2)
        self.available_fields_layout.addStretch(1)
        field_buttons_scroll_area = QScrollArea()
        field_buttons_scroll_area.setWidgetResizable(True)
        field_buttons_scroll_area.setWidget(field_buttons_container)
        field_list_frame.setMaximumHeight(180)
        field_buttons_scroll_area.setFixedHeight(120)
        field_list_frame_layout.addWidget(field_buttons_scroll_area)
        main_layout.addWidget(field_list_frame)

        # 문서 및 데이터 조작 패널
        doc_ops_panel = QVBoxLayout()
        doc_ops_panel.setSpacing(12)

        template_row = QHBoxLayout()
        template_row.setSpacing(12)
        self.select_template_button = self._make_primary_button("📁 템플릿 파일 선택")
        self.select_template_button.clicked.connect(self.select_template_file)
        template_row.addWidget(self.select_template_button)
        self.template_path_display = QLineEdit(readOnly=True, placeholderText="템플릿 파일(한글, 파워포인트)을 선택하세요")
        self.template_path_display.setFixedHeight(48)
        template_row.addWidget(self.template_path_display)
        doc_ops_panel.addLayout(template_row)

        xlsx_row = QHBoxLayout()
        xlsx_row.setSpacing(12)
        self.upload_xlsx_button = self._make_primary_button("⬆️ XLSX 업로드")
        self.upload_xlsx_button.clicked.connect(self.upload_xlsx)
        xlsx_row.addWidget(self.upload_xlsx_button)
        self.xlsx_path_display = QLineEdit(readOnly=True, placeholderText="업로드한 엑셀 파일 경로가 여기에 표시됩니다")
        self.xlsx_path_display.setFixedHeight(48)
        xlsx_row.addWidget(self.xlsx_path_display)
        doc_ops_panel.addLayout(xlsx_row)

        control_row = QHBoxLayout()
        control_row.setSpacing(12)
        self.add_row_button = self._make_secondary_button("➕ 행 추가")
        self.add_row_button.clicked.connect(self.add_row)
        control_row.addWidget(self.add_row_button)

        self.delete_row_button = self._make_secondary_button("🗑️ 선택 [행] 삭제")
        self.delete_row_button.clicked.connect(self.delete_selected_rows)
        control_row.addWidget(self.delete_row_button)

        self.delete_col_button = self._make_secondary_button("🗑️ 선택 [열] 삭제")
        self.delete_col_button.clicked.connect(self.delete_selected_columns)
        control_row.addWidget(self.delete_col_button)

        self.add_image_button = self._make_secondary_button("📷 이미지 추가")
        self.add_image_button.clicked.connect(self.add_images)
        control_row.addWidget(self.add_image_button)

        self.download_template_button = self._make_secondary_button("⬇️ 양식 다운로드")
        self.download_template_button.clicked.connect(self.download_xlsx_template)
        control_row.addWidget(self.download_template_button)
        self.generate_button = self._make_primary_button("✨ 문서 생성 ✨")
        self.generate_button.clicked.connect(self.generate_document)
        self.generate_button.setEnabled(False)
        control_row.addStretch(1)
        control_row.addWidget(self.generate_button)
        doc_ops_panel.addLayout(control_row)

        main_layout.addLayout(doc_ops_panel)

        # 테이블
        self.data_table = EnhancedTableWidget(self)
        self.data_table.setDataFrame(self.dataframe)
        table_palette = self.data_table.palette()
        table_palette.setColor(QPalette.Highlight, QColor(232, 241, 255))
        table_palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
        self.data_table.setPalette(table_palette)
        main_layout.addWidget(self.data_table)
        
        # 붙여넣기 시그널 연결
        self.data_table.pastedSignal.connect(self.update_generate_button_state)

    def _styled_label(self, text, css_class=None):
        label = QLabel(text)
        if css_class == "title":
            label.setStyleSheet("font-size: 22px; font-weight: 700; color: #1F1F20;")
        elif css_class == "subtitle":
            label.setStyleSheet("font-size: 18px; font-weight: 600; color: #2C2C2E;")
        else:
            label.setStyleSheet("font-size: 16px; font-weight: 500; color: #45474D;")
        return label

    def _make_primary_button(self, text):
        btn = QPushButton(text)
        btn.setFixedHeight(48)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #2563EB;
                color: #FFFFFF;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 600;
                padding: 12px;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
            QPushButton:pressed {
                background-color: #1E40AF;
            }
            QPushButton:disabled {
                background-color: #A5B4FC;
            }
        """)
        return btn

    def _make_secondary_button(self, text):
        btn = QPushButton(text)
        btn.setFixedHeight(44)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #F3F4F6;
                color: #1F2937;
                border: 1px solid #CBD5F5;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 500;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #E5E7EB;
            }
            QPushButton:pressed {
                background-color: #D1D5DB;
            }
        """)
        return btn

    def _find_hwp_window_handle(self, template_path_lower):
        if not is_windows:
            return None
        base_name = os.path.basename(template_path_lower)
        matches = []

        def enum_handler(hwnd, extra):
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return
            title_lower = title.lower()
            if base_name in title_lower and ("한글" in title_lower or "hwp" in title_lower):
                extra.append(hwnd)

        win32gui.EnumWindows(enum_handler, matches)
        return matches[0] if matches else None

    def _bring_window_to_front(self, hwnd):
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as bring_err:
            print(f"DEBUG: 창 전환 실패: {bring_err}")

    def _send_ctrl_s(self):
        try:
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('S'), 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(ord('S'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception as key_err:
            print(f"DEBUG: Ctrl+S 키 전송 실패: {key_err}")

    def _ensure_hwp_visibility(self, hwp):
        """한글 COM 인스턴스의 창이 사용자에게 보이도록 강제합니다."""
        try:
            hwp.Visible = True
        except Exception as err:
            print(f"DEBUG: HWP Visible 설정 실패(무시): {err}")
        try:
            windows = getattr(hwp, "XHwpWindows", None)
            if windows:
                active_window = getattr(windows, "Active_XHwpWindow", None)
                if active_window:
                    try:
                        active_window.Visible = True
                    except Exception as active_err:
                        print(f"DEBUG: Active_XHwpWindow.Visible 설정 실패(무시): {active_err}")
                elif getattr(windows, "Count", 0):
                    # 첫 번째 창을 활성화 시도
                    for base in (0, 1):
                        try:
                            window = windows.Item(base)
                            window.Visible = True
                            break
                        except Exception:
                            continue
        except Exception as err:
            print(f"DEBUG: HWP 창 가시성 확보 실패(무시): {err}")

    def _enumerate_hwp_documents(self, hwp):
        try:
            docs = getattr(hwp, "XHwpDocuments", None)
        except Exception as e:
            print(f"DEBUG: HWP XHwpDocuments 접근 실패 (COM 연결 끊김 추정): {e}")
            return []

        if not docs:
            print("DEBUG: HWP XHwpDocuments 정보 없음")
            return []

        documents = []
        try:
            count = getattr(docs, "Count", 0) or 0
        except Exception:
            count = 0
        
        print(f"DEBUG: HWP 열린 문서 수 추정: {count}")

        def _try_item(index, note):
            try:
                doc = docs.Item(index)
                if doc and doc not in documents:
                    documents.append(doc)
            except Exception as doc_err:
                print(f"DEBUG: XHwpDocuments.Item({index}) 접근 실패({note}): {doc_err}")

        # 0-based 접근
        for idx in range(count):
            _try_item(idx, "0-based")

        # 1-based 접근
        for idx in range(1, count + 1):
            _try_item(idx, "1-based")

        # Enumerator 접근
        enum_provider = getattr(docs, "_NewEnum", None)
        if enum_provider:
            try:
                enum = enum_provider()
                if enum:
                    index = 0
                    while True:
                        variant = enum.Next(1)
                        if not variant:
                            break
                        doc = variant[0]
                        index += 1
                        if doc and doc not in documents:
                            documents.append(doc)
            except Exception as enum_err:
                print(f"DEBUG: HWP 문서 열거 실패(무시): {enum_err}")

        if not documents:
            print("DEBUG: HWP 문서를 찾지 못했습니다 (빈 목록)")
        else:
            for idx, doc in enumerate(documents, start=1):
                full = getattr(doc, "FullName", None)
                path = getattr(doc, "Path", None)
                name = getattr(doc, "Name", None)
                print(f"DEBUG: HWP 문서 후보[{idx}] - FullName='{full}', Path='{path}', Name='{name}'")

        return documents

    def _match_hwp_document(self, doc, target_path_lower, template_name_lower):
        if not doc:
            return False

        candidates = []
        for attr in ("FullName",):
            try:
                value = getattr(doc, attr, None)
                if value:
                    candidates.append(value)
            except Exception:
                continue

        try:
            path_value = getattr(doc, "Path", None)
            name_value = getattr(doc, "Name", None)
            if path_value and name_value:
                candidates.append(os.path.join(path_value, name_value))
        except Exception:
            pass

        for candidate in candidates:
            try:
                norm_candidate = os.path.normcase(os.path.normpath(os.path.abspath(candidate)))
                print(f"DEBUG: 문서 경로 비교 - candidate='{norm_candidate}', target='{target_path_lower}'")
                if norm_candidate == target_path_lower:
                    return True
            except Exception:
                continue

        try:
            doc_name = getattr(doc, "Name", None)
            if doc_name and doc_name.lower() == template_name_lower:
                return True
        except Exception:
            pass

        return False

    def _get_hwp_document(self, hwp, target_path_lower, template_name_lower):
        for doc in self._enumerate_hwp_documents(hwp):
            try:
                if self._match_hwp_document(doc, target_path_lower, template_name_lower):
                    return doc
            except Exception:
                continue
        return None

    def show_open_source_info(self):
        message = (
            "용머지(YongMerge)는 다음 오픈소스 소프트웨어를 사용하며, 각 라이선스 조건을 준수하여 배포됩니다:\n\n"
            "• Python 3 (PSF License)\n"
            "      https://www.python.org/\n"
            "• PyQt5 (GPL v3)\n"
            "      https://www.riverbankcomputing.com/software/pyqt/\n"
            "• python-pptx (MIT License)\n"
            "      https://python-pptx.readthedocs.io/\n"
            "• Pillow (HPND License)\n"
            "      https://python-pillow.org/\n"
            "• pandas (BSD 3-Clause License)\n"
            "      https://pandas.pydata.org/\n"
            "• PyInstaller (GPL v2 with exceptions)\n"
            "      https://pyinstaller.org/\n"
            "• pywin32 / win32com (PSF License)\n"
            "      https://github.com/mhammond/pywin32\n\n"
            "저작권, 라이선스 전문 및 상세 정보는 앱 소스코드 저장소에서 확인할 수 있습니다."
        )
        QMessageBox.information(self, "오픈소스 라이선스", message)


    def _insert_hwp_field(self, field_name):
        """활성 HWP 문서에 누름틀을 삽입합니다."""
        if not is_windows or not self.template_file_path:
            return False

        template_abs_path = os.path.abspath(self.template_file_path)
        target_path_lower = os.path.normcase(os.path.normpath(template_abs_path))
        template_name_lower = os.path.basename(template_abs_path).lower()

        hwnd = self._find_hwp_window_handle(target_path_lower)
        if hwnd:
            self._bring_window_to_front(hwnd)
            time.sleep(0.2)
        else:
            print("DEBUG: HWP 창 핸들을 찾지 못했습니다.")

        coinitialized = False
        try:
            pythoncom.CoInitialize()
            coinitialized = True
        except Exception:
            pass

        # 기존 HWP 인스턴스가 유효한지 확인
        if self.hwp_app is not None:
            try:
                # 가벼운 속성 접근으로 연결 상태 확인
                _ = self.hwp_app.Visible
            except Exception:
                print("DEBUG: 기존 HWP 인스턴스 연결 끊김 감지 - 참조 초기화")
                self.hwp_app = None

        hwp = self.hwp_app
        if hwp is None:
            try:
                hwp = hwp_automation.ensure_hwp_app()
                self.hwp_app = hwp
            except Exception as ensure_err:
                print(f"DEBUG: HWP 인스턴스 확보 실패: {ensure_err}")
                if coinitialized:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
                return False

        self._ensure_hwp_visibility(hwp)

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception as reg_err:
            print(f"DEBUG: HWP RegisterModule 실패(무시): {reg_err}")

        try:
            hwp.SetMessageBoxMode(0x00010000)
        except Exception as msg_err:
            print(f"DEBUG: HWP SetMessageBoxMode 실패(무시): {msg_err}")

        target_doc = self._get_hwp_document(hwp, target_path_lower, template_name_lower)

        if not target_doc:
            print("DEBUG: 템플릿과 일치하는 HWP 문서를 찾지 못했습니다.")
            if coinitialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            return False

        try:
            if hasattr(target_doc, "SetActive"):
                target_doc.SetActive()
        except Exception as active_err:
            print(f"DEBUG: 문서 활성화 실패(무시): {active_err}")

        refreshed_hwnd = self._find_hwp_window_handle(target_path_lower)
        if refreshed_hwnd:
            self._bring_window_to_front(refreshed_hwnd)
            time.sleep(0.1)
        else:
            if hwnd:
                self._bring_window_to_front(hwnd)
                time.sleep(0.1)

        first_value = ""
        if field_name in self.dataframe.columns and not self.dataframe.empty:
            first_cell = self.dataframe.iloc[0][field_name]
            if pd.notna(first_cell):
                first_value = str(first_cell)

        memo_value = ""
        if field_name in self.dataframe.columns:
            for value in self.dataframe[field_name].tolist():
                if pd.notna(value):
                    memo_value = str(value)
                    break

        try:
            result = hwp.CreateField(first_value, memo_value, field_name)
            if result:
                print(f"DEBUG: CreateField 메서드로 누름틀 생성 완료 - {field_name}")
                return True
            print(f"DEBUG: CreateField 메서드 실패 - {field_name}")
        except Exception as err:
            print(f"DEBUG: HWP CreateField 실패: {err}")
        finally:
            if coinitialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        return False

    def _close_template_if_open(self, doc_type):
        path = self.template_file_path
        if not path or not os.path.exists(path):
            return
        abs_path = os.path.abspath(path).lower()
        try:
            import win32com.client as com
        except ImportError:
            return

        if doc_type == 'ppt':
            try:
                ppt = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                return
            try:
                for presentation in list(ppt.Presentations):
                    try:
                        full = os.path.abspath(presentation.FullName).lower()
                    except Exception:
                        continue
                    if full == abs_path:
                        try:
                            if presentation.Saved == 0:
                                presentation.Save()
                        except Exception:
                            pass
                        presentation.Close()
                        print(f"DEBUG: 기존 PPT 템플릿 저장 후 닫기 - {path}")
                        break
            except Exception as err:
                print(f"DEBUG: PPT 템플릿 닫기 실패: {err}")
        elif doc_type == 'hwp':
            hwnd = self._find_hwp_window_handle(abs_path)
            if hwnd:
                print(f"DEBUG: HWP 창 핸들 확보({hwnd}), 저장 후 닫기 수행")
                self._bring_window_to_front(hwnd)
                time.sleep(0.2)
                self._send_ctrl_s()
                time.sleep(0.3)
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    print(f"DEBUG: WM_CLOSE 전송 완료 - {path}")
                except Exception as close_err:
                    print(f"DEBUG: WM_CLOSE 전송 실패: {close_err}")
                return
            try:
                hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
            except Exception as dispatch_err:
                print(f"DEBUG: 활성 HWP 인스턴스 없음: {dispatch_err}")
                return
            try:
                hwp.HAction.Run("FileSave")
            except Exception as save_err:
                print(f"DEBUG: HWP FileSave 실패 (무시): {save_err}")
            try:
                hwp.HAction.Run("FileClose")
                print(f"DEBUG: HWP FileClose 실행 - {path}")
            except Exception as err:
                print(f"DEBUG: HWP 템플릿 닫기 실패: {err}")

    def load_initial_data(self):
        if self.dataframe.empty:
             self.dataframe = pd.DataFrame(index=range(5))
        self.data_table.update_table_from_dataframe()

    def create_field(self, field_name=None, from_input=True):
        if from_input:
            field_name = self.field_name_input.text().strip()
        if not field_name or field_name in self.dataframe.columns: return

        # DataFrame에 열 추가 (행이 없으면 최소 5개 행 생성)
        if len(self.dataframe) == 0:
            self.dataframe = pd.DataFrame({field_name: [None] * 5})
        else:
            self.dataframe[field_name] = [None] * len(self.dataframe)

        # 테이블 업데이트 (setDataFrame을 사용하여 완전히 새로고침)
        self.data_table.setDataFrame(self.dataframe)

        print(f"DEBUG: '{field_name}' 필드 생성 완료. DataFrame columns: {list(self.dataframe.columns)}, shape: {self.dataframe.shape}")

        color = random.choice(FIELD_COLORS)
        field_button = DraggableButton(field_name, color)
        field_button.singleClicked.connect(self.on_field_button_single_clicked)
        remove_button = QPushButton("X")
        remove_button.setFixedSize(20, 20)
        remove_button.clicked.connect(lambda _, name=field_name: self.remove_field(name))
        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.addWidget(field_button)
        btn_layout.addWidget(remove_button)
        btn_layout.setContentsMargins(0,0,0,0)
        self.available_fields_layout.insertWidget(self.available_fields_layout.count() - 1, btn_widget)
        if from_input: self.field_name_input.clear()
        self.update_generate_button_state()

    def remove_field(self, field_name):
        """필드 삭제 (DataFrame과 UI에서 모두 제거)"""
        # DataFrame에서 열 삭제
        if field_name in self.dataframe.columns:
            self.dataframe = self.dataframe.drop(columns=[field_name])
            print(f"DEBUG: DataFrame에서 '{field_name}' 열 삭제 완료")
            print(f"DEBUG: 남은 DataFrame columns: {list(self.dataframe.columns)}")

            # 테이블 업데이트 (완전히 다시 그리기)
            self.data_table.setDataFrame(self.dataframe)
            print(f"DEBUG: 테이블 업데이트 완료 - 테이블 열 개수: {self.data_table.columnCount()}")

        # 필드 버튼 UI에서 제거
        for i in range(self.available_fields_layout.count()):
             item = self.available_fields_layout.itemAt(i)
             if item and item.widget():
                  button = item.widget().findChild(DraggableButton)
                  if button and button.text() == field_name:
                       item.widget().deleteLater()
                       print(f"DEBUG: 필드 버튼 '{field_name}' UI에서 삭제 완료")
                       break

        self.update_generate_button_state()

    def delete_selected_columns(self):
        selected_indexes = self.data_table.selectedIndexes()
        if not selected_indexes:
            QMessageBox.warning(self, "경고", "삭제할 열을 선택해주세요.")
            return

        # 선택된 열 인덱스 추출 (중복 제거)
        selected_columns = set(index.column() for index in selected_indexes)

        # 유효한 열 이름만 필터링 (DataFrame과 테이블 헤더 둘 다 확인)
        valid_column_names = set()
        for col_idx in selected_columns:
            # 테이블 헤더에서 열 이름 가져오기
            header_item = self.data_table.horizontalHeaderItem(col_idx)
            if header_item:
                col_name = header_item.text()
                # DataFrame에 해당 열이 존재하는지 확인
                if col_name in self.dataframe.columns:
                    valid_column_names.add(col_name)

        if not valid_column_names:
            QMessageBox.warning(self, "경고", "삭제할 유효한 열이 없습니다.\n(DataFrame에 존재하지 않는 열입니다)")
            return

        # 사용자에게 확인
        col_list = ", ".join(valid_column_names)
        reply = QMessageBox.question(
            self,
            "열 삭제 확인",
            f"다음 {len(valid_column_names)}개 열을 삭제하시겠습니까?\n\n{col_list}",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # 열 삭제 실행
        for field_name in valid_column_names:
            self.remove_field(field_name)

        print(f"DEBUG: {len(valid_column_names)}개 열 삭제 완료: {col_list}")

    def update_generate_button_state(self):
         enabled = bool(self.template_file_path) and not self.dataframe.columns.empty and not self.dataframe.dropna(how='all').empty
         self.generate_button.setEnabled(enabled)

    def update_dataframe_from_cell(self, row, column, value):
        """셀 값이 변경되었을 때 DataFrame 업데이트"""
        # 행 범위 체크
        if row >= len(self.dataframe):
            print(f"WARNING: 행 인덱스 {row}가 DataFrame 범위({len(self.dataframe)})를 벗어났습니다.")
            return

        # 열 범위 체크
        if column >= len(self.dataframe.columns):
            print(f"WARNING: 열 인덱스 {column}가 DataFrame 열 개수({len(self.dataframe.columns)})를 벗어났습니다.")
            print(f"DEBUG: DataFrame columns: {list(self.dataframe.columns)}")
            print(f"DEBUG: 테이블 열 개수: {self.data_table.columnCount()}")
            return

        # DataFrame 업데이트
        col_name = self.dataframe.columns[column]
        
        # 이미지 열의 경우, 표시 텍스트(📷 ...)가 DataFrame에 저장되지 않도록 방어
        if col_name == "이미지" and isinstance(value, str) and value.startswith("📷 "):
             # 현재 저장된 값의 표시 이름과 같다면 (즉, 사용자가 내용 변경 없이 엔터만 친 경우) 무시
             current_val = self.dataframe.at[row, col_name]
             if current_val and image_utils.get_image_display_name(current_val) == value:
                 return
             
             # 내용이 다르더라도 "📷 "로 시작하면 유효한 파일 경로가 아닐 확률이 높으므로
             # 실제 파일이 존재하지 않는 한 업데이트를 무시하거나 경고
             if not os.path.exists(value):
                 print(f"DEBUG: 이미지 열의 표시 텍스트 업데이트 무시: {value}")
                 return

        self.dataframe.at[row, col_name] = value if value else None
        self.update_generate_button_state()

    def handle_table_rows_changed(self):
         self.sync_dataframe_with_table_rows()
         self.update_generate_button_state()

    def sync_dataframe_with_table_rows(self):
        table_rows = self.data_table.rowCount()
        df_rows = len(self.dataframe)
        if table_rows > df_rows:
            new_rows = pd.DataFrame([([None] * len(self.dataframe.columns))] * (table_rows - df_rows), columns=self.dataframe.columns)
            self.dataframe = pd.concat([self.dataframe, new_rows], ignore_index=True)
        elif table_rows < df_rows:
            self.dataframe = self.dataframe.iloc[:table_rows].reset_index(drop=True)
        
        # DataFrame 객체가 변경되었으므로 테이블 위젯의 참조도 업데이트
        self.data_table.updateDataFrameRef(self.dataframe)

    def add_row(self):
        insert_pos = self.data_table.currentRow() + 1 if self.data_table.selectedIndexes() else self.data_table.rowCount()
        self.data_table.insertRow(insert_pos)
        self.sync_dataframe_with_table_rows()
        self.update_generate_button_state()

    def delete_selected_rows(self):
        rows_to_delete = sorted(list(set(index.row() for index in self.data_table.selectedIndexes())), reverse=True)
        if not rows_to_delete: return
        for row in rows_to_delete:
            self.data_table.removeRow(row)
        self.sync_dataframe_with_table_rows()
        self.update_generate_button_state()

    def upload_xlsx(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "XLSX 파일 업로드", "", "Excel Files (*.xlsx)")
        if not file_path: return
        try:
            uploaded_df = pd.read_excel(file_path).astype(object).where(pd.notna, None)
            for col_name in uploaded_df.columns:
                if col_name not in self.dataframe.columns:
                    self.create_field(field_name=col_name, from_input=False)
            self.dataframe = uploaded_df.reindex(columns=self.dataframe.columns)
            self.data_table.setDataFrame(self.dataframe)
            self.update_generate_button_state()
            self.xlsx_path_display.setText(file_path)
            QMessageBox.information(self, "완료", "XLSX 파일 업로드 및 데이터/필드 등록 완료.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"XLSX 파일 로드 중 오류: {e}")

    def download_xlsx_template(self):
        if self.dataframe.columns.empty: return
        file_path, _ = QFileDialog.getSaveFileName(self, "XLSX 양식 다운로드", "mailmerge_template.xlsx", "Excel Files (*.xlsx)")
        if file_path:
            pd.DataFrame(columns=self.dataframe.columns).to_excel(file_path, index=False)

    def _open_hwp_template_via_com(self, template_path):
        template_abs_path = os.path.abspath(template_path)
        target_path_lower = os.path.normcase(os.path.normpath(template_abs_path))
        template_name_lower = os.path.basename(template_abs_path).lower()

        coinitialized = False
        try:
            pythoncom.CoInitialize()
            coinitialized = True
        except Exception:
            pass

        # 기존 HWP 인스턴스가 유효한지 확인
        if self.hwp_app is not None:
            try:
                # 가벼운 속성 접근으로 연결 상태 확인
                _ = self.hwp_app.Visible
            except Exception:
                print("DEBUG: 기존 HWP 인스턴스 연결 끊김 감지 - 참조 초기화")
                self.hwp_app = None

        hwp = self.hwp_app
        if hwp is None:
            try:
                hwp = hwp_automation.ensure_hwp_app()
                self.hwp_app = hwp
            except Exception as ensure_err:
                print(f"DEBUG: HWP 템플릿 연결 실패: {ensure_err}")
                if coinitialized:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
                return False

        self._ensure_hwp_visibility(hwp)

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception as reg_err:
            print(f"DEBUG: HWP RegisterModule 실패(무시): {reg_err}")

        try:
            hwp.SetMessageBoxMode(0x00010000)
        except Exception as msg_err:
            print(f"DEBUG: HWP SetMessageBoxMode 실패(무시): {msg_err}")

        doc = self._get_hwp_document(hwp, target_path_lower, template_name_lower)

        opened = False
        if not doc:
            file_format = hwp_automation.get_file_format(template_abs_path) or "HWP"
            open_attempts = [
                lambda: hwp.Open(template_abs_path, file_format, "forceopen:true"),
                lambda: hwp.Open(template_abs_path, file_format, ""),
            ]
            for attempt in open_attempts:
                try:
                    result = attempt()
                    if result:
                        opened = True
                        break
                except Exception as open_err:
                    print(f"DEBUG: HWP Open 실패(무시): {open_err}")

            if not opened:
                docs = getattr(hwp, "XHwpDocuments", None)
                if docs and hasattr(docs, "Open"):
                    for option in ("forceopen:true", ""):
                        try:
                            docs.Open(template_abs_path, file_format, option)
                            opened = True
                            break
                        except Exception as xopen_err:
                            print(f"DEBUG: XHwpDocuments.Open 실패(옵션={option}): {xopen_err}")

            if opened:
                time.sleep(0.3)
                doc = self._get_hwp_document(hwp, target_path_lower, template_name_lower)
                if not doc:
                    print("DEBUG: HWP Open 후에도 템플릿 문서를 찾지 못했습니다.")

        if doc:
            try:
                doc.SetActive()
            except Exception as active_err:
                print(f"DEBUG: HWP 템플릿 활성화 실패(무시): {active_err}")
            hwnd = self._find_hwp_window_handle(target_path_lower)
            if hwnd:
                self._bring_window_to_front(hwnd)
                time.sleep(0.1)

            # 기본 새 문서가 남아 있으면 닫기
            for extra_doc in self._enumerate_hwp_documents(hwp):
                if extra_doc is doc:
                    continue
                try:
                    extra_path = getattr(extra_doc, "Path", None)
                    extra_name = getattr(extra_doc, "Name", None)
                except Exception:
                    continue
                if not extra_path and extra_name and extra_name.startswith("새 문서"):
                    try:
                        extra_doc.Close(0)
                        print("DEBUG: 기본 새 문서 닫기 완료")
                    except Exception as close_err:
                        print(f"DEBUG: 기본 새 문서 닫기 실패(무시): {close_err}")

            success = True
        else:
            print("DEBUG: COM으로 템플릿 문서를 제어하지 못했습니다.")
            success = False

        if coinitialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        return success

    def select_template_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "템플릿 파일 선택", "", "Document Files (*.hwp *.hwpx *.ppt *.pptx)")
        if file_path:
            self.template_path_display.setText(file_path)
            self.template_file_path = file_path
            self.update_generate_button_state()
            if not is_windows or not os.path.exists(file_path):
                return

            extension = os.path.splitext(file_path)[1].lower()
            handled = False
            if extension in (".hwp", ".hwpx"):
                handled = self._open_hwp_template_via_com(file_path)
            if not handled:
                try:
                    os.startfile(file_path)
                    print(f"DEBUG: 템플릿 파일 실행 (폴백) - {file_path}")
                except Exception as open_err:
                    print(f"DEBUG: 템플릿 파일 실행 실패 (무시): {open_err}")

    def generate_document(self):
        if not self.template_file_path: return
        valid_dataframe = self.dataframe.dropna(how='all').reset_index(drop=True)
        if valid_dataframe.empty: return

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("출력 방식 선택")
        individual_button = msg_box.addButton("개별 파일로 저장", QMessageBox.ActionRole)
        combined_button = msg_box.addButton("통합 파일로 저장", QMessageBox.ActionRole)
        msg_box.addButton("취소", QMessageBox.RejectRole)
        msg_box.exec_()

        clicked = msg_box.clickedButton()
        if clicked == individual_button: output_type = 'individual'
        elif clicked == combined_button: output_type = 'combined'
        else: return

        file_extension = os.path.splitext(self.template_file_path)[1].lower()
        doc_type = 'hwp' if file_extension in ['.hwp', '.hwpx'] else 'ppt'
        save_path = None

        if output_type == 'combined':
            output_dir = os.path.dirname(self.template_file_path)
            base_name = os.path.splitext(os.path.basename(self.template_file_path))[0]
            suggested_path = os.path.join(output_dir, f"{base_name}_통합본{file_extension}")
            save_path, _ = QFileDialog.getSaveFileName(self, f"통합 {doc_type.upper()} 파일 저장", suggested_path, f"{doc_type.upper()} Files (*{file_extension})")
            if not save_path: return

        if not is_windows: return

        self._close_template_if_open(doc_type)

        self.progress_dialog = QProgressDialog("문서 생성 중...", "취소", 0, 100, self)
        self.progress_dialog.canceled.connect(self.cancel_automation)

        self.worker = AutomationWorker(doc_type, valid_dataframe, self.template_file_path, output_type, save_path)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_automation_complete)
        self.worker.error.connect(self.on_automation_error)

        self.generate_button.setEnabled(False)
        self.worker.start()
        self.progress_dialog.show()

    def update_progress(self, value):
        self.progress_dialog.setValue(value)

    def on_automation_complete(self, message, output_type, output_file):
        """병합 완료 시 호출되는 콜백

        Args:
            message: 완료 메시지
            output_type: 'individual' 또는 'combined'
            output_file: 통합 파일 경로 (combined인 경우), None이면 개별 파일
        """
        self.progress_dialog.setValue(100)
        self.generate_button.setEnabled(True)

        # 완료 메시지 표시
        QMessageBox.information(self, "완료", message)

        # 통합 파일인 경우 자동으로 열기
        if output_type == 'combined' and output_file and os.path.exists(output_file):
            try:
                print(f"DEBUG: 통합 파일 열기: {output_file}")

                # 운영체제에 맞는 파일 열기 명령 실행
                if is_windows:
                    os.startfile(output_file)
                else:
                    # macOS나 Linux의 경우
                    import subprocess
                    if platform.system() == 'Darwin':  # macOS
                        subprocess.call(['open', output_file])
                    else:  # Linux
                        subprocess.call(['xdg-open', output_file])

                print(f"DEBUG: 통합 파일 열기 완료")

            except Exception as e:
                print(f"WARNING: 통합 파일 열기 실패 (무시 가능): {e}")

    def on_automation_error(self, message):
        self.progress_dialog.close()
        QMessageBox.critical(self, "자동화 오류", f"오류가 발생했습니다: {message}")
        self.generate_button.setEnabled(True)

    def cancel_automation(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self.generate_button.setEnabled(True)

    def add_images(self):
        """이미지 파일을 선택하고 시트에 추가합니다."""
        # 다중 이미지 파일 선택 다이얼로그
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "이미지 파일 선택 (다중 선택 가능)",
            "",
            "Image Files (*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif *.webp)"
        )

        if not file_paths:
            return

        # 선택된 파일 검증
        valid_images = []
        invalid_images = []

        for file_path in file_paths:
            is_valid, message = image_utils.validate_image_path(file_path)
            if is_valid:
                valid_images.append(file_path)
            else:
                invalid_images.append((file_path, message))

        # 유효하지 않은 이미지가 있으면 경고
        if invalid_images:
            error_msg = "다음 파일을 불러올 수 없습니다:\n\n"
            for path, reason in invalid_images[:5]:  # 최대 5개만 표시
                error_msg += f"• {os.path.basename(path)}: {reason}\n"
            if len(invalid_images) > 5:
                error_msg += f"\n... 외 {len(invalid_images) - 5}개"
            QMessageBox.warning(self, "이미지 검증 실패", error_msg)

        if not valid_images:
            return

        # Step 1: '이미지' 필드가 없으면 자동 생성
        image_field_name = "이미지"
        if image_field_name not in self.dataframe.columns:
            self.create_field(field_name=image_field_name, from_input=False)
            print(f"DEBUG: '{image_field_name}' 필드 자동 생성 완료")

        # Step 2: 이미지 열의 마지막 데이터가 있는 행 찾기
        image_col_idx = self.dataframe.columns.get_loc(image_field_name)
        last_data_row = -1  # 데이터가 없으면 -1

        for idx in range(len(self.dataframe) - 1, -1, -1):
            cell_value = self.dataframe.at[idx, image_field_name]
            if pd.notna(cell_value) and str(cell_value).strip():
                last_data_row = idx
                break

        # 다음 행부터 시작 (마지막 데이터 행 + 1)
        start_row = last_data_row + 1
        print(f"DEBUG: 이미지 열의 마지막 데이터 행: {last_data_row}, 입력 시작 행: {start_row}")

        # Step 3: 필요한 행 수 계산 및 추가
        required_rows = start_row + len(valid_images)
        current_row_count = len(self.dataframe)

        # 부족한 행 추가
        if required_rows > current_row_count:
            rows_to_add = required_rows - current_row_count
            for _ in range(rows_to_add):
                self.data_table.insertRow(self.data_table.rowCount())
            self.sync_dataframe_with_table_rows()

        # Step 4: 마지막 데이터 다음 행부터 이미지 경로 순차 입력

        for idx, img_path in enumerate(valid_images):
            row_idx = start_row + idx
            normalized_path = image_utils.normalize_image_path(img_path)

            # DataFrame에 저장
            self.dataframe.at[row_idx, image_field_name] = normalized_path

            # 테이블에 표시 (아이콘 + 파일명)
            display_text = image_utils.get_image_display_name(img_path)
            item = QTableWidgetItem(display_text)
            
            # 테이블 업데이트 시 시그널 차단 (DataFrame에 표시 텍스트가 덮어씌워지는 것 방지)
            self.data_table.blockSignals(True)
            self.data_table.setItem(row_idx, image_col_idx, item)
            self.data_table.blockSignals(False)

            print(f"DEBUG: 행 {row_idx + 1}에 이미지 추가: {os.path.basename(img_path)}")

        # Step 4: UI 업데이트
        self.update_generate_button_state()

        # 성공 메시지
        QMessageBox.information(
            self,
            "이미지 추가 완료",
            f"{len(valid_images)}개의 이미지가 '{image_field_name}' 필드에 추가되었습니다.\n\n"
            f"시작 행: {start_row + 1}\n"
            f"종료 행: {start_row + len(valid_images)}"
        )

    def on_image_cell_double_clicked(self, row, column):
        """이미지 열 셀 더블클릭 시 이미지 파일 선택 다이얼로그"""
        # 단일 이미지 파일 선택
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "이미지 파일 선택",
            "",
            "Image Files (*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif *.webp)"
        )

        if not file_path:
            return

        # 선택된 파일 검증
        is_valid, message = image_utils.validate_image_path(file_path)
        if not is_valid:
            QMessageBox.warning(self, "이미지 검증 실패", message)
            return

        # 이미지 경로 정규화
        normalized_path = image_utils.normalize_image_path(file_path)

        # DataFrame에 저장
        col_name = self.dataframe.columns[column]
        self.dataframe.at[row, col_name] = normalized_path

        # 테이블에 표시 (아이콘 + 파일명)
        display_text = image_utils.get_image_display_name(file_path)
        item = QTableWidgetItem(display_text)
        
        # 테이블 업데이트 시 시그널 차단 (DataFrame에 표시 텍스트가 덮어씌워지는 것 방지)
        self.data_table.blockSignals(True)
        self.data_table.setItem(row, column, item)
        self.data_table.blockSignals(False)

        print(f"DEBUG: 행 {row + 1}, 열 '{col_name}'에 이미지 추가: {os.path.basename(file_path)}")

        # UI 업데이트
        self.update_generate_button_state()

    def on_field_button_single_clicked(self, field_name):
        """필드 버튼 싱글클릭 시 {{필드명}} 형식으로 문서에 삽입하고 자동 저장

        PPT에서 '이미지' 필드인 경우: {{이미지}} 텍스트가 포함된 사각형 삽입
        그 외: 일반 텍스트로 {{필드명}} 삽입
        삽입 후 문서 자동 저장
        """
        if not is_windows: return

        hwp_ppt_windows = []
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd) and self.winId() != hwnd:
                window_title = win32gui.GetWindowText(hwnd)
                if ("HWP" in window_title.upper() or "한글" in window_title) or ("PowerPoint" in window_title):
                    results.append(hwnd)
        win32gui.EnumWindows(enum_windows_callback, hwp_ppt_windows)

        if not hwp_ppt_windows:
            QMessageBox.warning(self, "경고", "열려있는 HWP 또는 PowerPoint 창을 찾을 수 없습니다.")
            return

        hwnd = hwp_ppt_windows[0]
        window_title = win32gui.GetWindowText(hwnd)
        doc_type = 'HWP' if "HWP" in window_title.upper() or "한글" in window_title else 'PPT'

        try:
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
            print(f"DEBUG: {doc_type} 창 활성화 완료")

            # PPT에서 '이미지' 필드인 경우 사각형 삽입
            if doc_type == 'PPT' and field_name == "이미지":
                print("DEBUG: PowerPoint 이미지 필드 삽입 시작")

                # 방법 1: COM API 시도 (단, 실패 시 방법 2로 폴백)
                com_success = False
                try:
                    print("DEBUG: PowerPoint COM 준비 대기 시작 (1.5초)")
                    time.sleep(1.5)
                    print("DEBUG: PowerPoint COM 방식 사각형 삽입 시도")
                    com_success = self._insert_ppt_image_rectangle()
                except Exception as e:
                    print(f"DEBUG: COM 방식 실패: {e}")

                # 방법 2: COM 실패 시 키보드 자동화로 사각형 삽입
                if not com_success:
                    print("DEBUG: COM 방식 실패, 키보드 자동화 방식으로 전환")
                    keyboard_success = self._insert_ppt_rectangle_by_keyboard()
                    if keyboard_success:
                        print("DEBUG: 키보드 자동화 방식으로 사각형 삽입 성공")
                        time.sleep(0.3)
                        self._save_with_keyboard()
                    else:
                        QMessageBox.warning(
                            self,
                            "경고",
                            "PowerPoint 이미지 사각형 삽입에 실패했습니다.\n\n"
                            "수동으로 사각형을 그리고 그 안에 '{{이미지}}' 텍스트를 입력해주세요."
                        )
                else:
                    # COM 방식 성공
                    print(f"DEBUG: PPT 이미지 사각형 삽입 성공 (COM), Ctrl+S로 저장")
                    time.sleep(0.3)
                    self._save_with_keyboard()
                return  # PPT 이미지는 여기서 종료
            else:
                if doc_type == 'HWP':
                    if self._insert_hwp_field(field_name):
                        self._auto_save_document(doc_type)
                        return
                    else:
                        print("DEBUG: HWP 누름틀 생성 실패, 기존 붙여넣기 방식 사용")
                field_placeholder = f'{{{{{field_name}}}}}'
                QApplication.clipboard().setText(field_placeholder)
                win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                win32api.keybd_event(ord('V'), 0, 0, 0)
                time.sleep(0.05)
                win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.2)
                print(f"DEBUG: '{field_placeholder}' 문서에 삽입 완료 (문서 타입: {doc_type})")

            # 필드 삽입 후 문서 자동 저장
            self._auto_save_document(doc_type)

        except Exception as e:
            QMessageBox.critical(self, "오류", f"필드 삽입 중 오류 발생: {e}")

    def _auto_save_document(self, doc_type):
        """템플릿 문서 자동 저장 (COM API 또는 Ctrl+S)"""
        try:
            import win32com.client as com
            time.sleep(0.3)  # 필드 삽입이 완전히 끝날 때까지 대기

            if doc_type == 'HWP':
                hwnd = self._find_hwp_window_handle(self.template_file_path.lower() if self.template_file_path else "")
                if hwnd:
                    print(f"DEBUG: HWP 창 핸들 확보({hwnd}), Ctrl+S 수행")
                    self._bring_window_to_front(hwnd)
                    time.sleep(0.2)
                    self._send_ctrl_s()
                    time.sleep(0.3)
                else:
                    try:
                        hwp = com.GetActiveObject("HWPFrame.HwpObject")
                    except Exception:
                        try:
                            hwp = com.DispatchEx("HWPFrame.HwpObject")
                        except Exception:
                            hwp = dynamic.Dispatch("HWPFrame.HwpObject")
                    if not hwp:
                        print("WARNING: HWP 인스턴스를 가져올 수 없음, Ctrl+S 시도")
                        raise Exception("HWP instance not found")
                    try:
                        result = hwp.Save()
                        if result:
                            print("DEBUG: HWP 문서 자동 저장 완료 (COM API)")
                        else:
                            print("WARNING: HWP Save() 반환값 False, Ctrl+S 시도")
                            raise Exception("HWP Save failed")
                    except Exception as e:
                        print(f"DEBUG: HWP 저장 실패, Ctrl+S로 대체: {e}")
                        self._save_with_keyboard()

            elif doc_type == 'PPT':
                try:
                    ppt = com.GetActiveObject("PowerPoint.Application")
                    if ppt and ppt.ActivePresentation:
                        ppt.ActivePresentation.Save()
                        print(f"DEBUG: PPT 문서 자동 저장 완료 (COM API)")
                    else:
                        print(f"WARNING: PPT 인스턴스를 가져올 수 없음, Ctrl+S 시도")
                        raise Exception("PPT instance not found")
                except Exception as e:
                    print(f"DEBUG: PPT COM 저장 실패, Ctrl+S로 대체: {e}")
                    self._save_with_keyboard()

            time.sleep(0.2)  # 저장 완료 대기

        except Exception as e:
            print(f"WARNING: 문서 자동 저장 중 오류 (무시 가능): {e}")
            import traceback
            traceback.print_exc()

    def _save_with_keyboard(self):
        """키보드 입력으로 Ctrl+S 실행"""
        try:
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('S'), 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(ord('S'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            print(f"DEBUG: Ctrl+S 키 입력 완료")
        except Exception as e:
            print(f"WARNING: Ctrl+S 입력 실패: {e}")

    def _insert_ppt_rectangle_by_keyboard(self):
        """키보드 자동화로 PowerPoint에 사각형 삽입 (COM 대안)

        PowerPoint에서:
        1. Alt+N, S, H: 삽입 → 도형 → 사각형 (단축키)
        2. 마우스 드래그로 사각형 그리기
        3. 텍스트 입력: {{이미지}}

        Returns:
            bool: 성공 여부
        """
        try:
            print("DEBUG: 키보드 자동화로 PPT 사각형 삽입 시작")

            # PowerPoint 창이 활성화되어 있는 상태에서 시작
            time.sleep(0.5)

            # 방법 1: Alt + N (삽입) → S (도형) → H (사각형)
            try:
                print("DEBUG: PowerPoint 도형 메뉴 접근 시도")

                # ESC로 기존 선택 해제 (2번)
                for _ in range(2):
                    win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
                    time.sleep(0.05)
                    win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.1)
                time.sleep(0.3)
                print("DEBUG: 기존 선택 해제 완료")

                # Alt 키 누르고 바로 N 키 (삽입 탭)
                win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(ord('N'), 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(ord('N'), 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.1)
                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.5)  # 삽입 리본이 활성화될 때까지 대기
                print("DEBUG: 삽입 탭 활성화 완료")

                # S 키 (도형 메뉴)
                win32api.keybd_event(ord('S'), 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(ord('S'), 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.5)  # 도형 메뉴가 열릴 때까지 대기
                print("DEBUG: 도형 메뉴 활성화 완료")

                # H 키 (사각형 선택)
                win32api.keybd_event(ord('H'), 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(ord('H'), 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.6)  # 사각형 커서 모드로 전환될 때까지 충분히 대기
                print("DEBUG: 사각형 그리기 모드 활성화 완료")

            except Exception as menu_err:
                print(f"DEBUG: 메뉴 접근 실패: {menu_err}")
                return False

            # 방법 2: 마우스로 사각형 그리기 (화면 중앙에)
            # PowerPoint 창의 중심 좌표를 가져와서 사각형 그리기
            try:
                # PowerPoint 창 핸들 찾기
                ppt_windows = []
                def enum_callback(hwnd, results):
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if "PowerPoint" in title:
                            results.append(hwnd)
                win32gui.EnumWindows(enum_callback, ppt_windows)

                if not ppt_windows:
                    print("DEBUG: PowerPoint 창을 찾을 수 없음")
                    return False

                # 첫 번째 PowerPoint 창의 좌표 가져오기
                hwnd = ppt_windows[0]
                rect = win32gui.GetWindowRect(hwnd)
                left, top, right, bottom = rect

                # 창 중앙 계산
                center_x = (left + right) // 2
                center_y = (top + bottom) // 2

                # 사각형 크기 (픽셀)
                rect_width = 300
                rect_height = 200

                # 사각형 시작/끝 좌표
                start_x = center_x - rect_width // 2
                start_y = center_y - rect_height // 2
                end_x = center_x + rect_width // 2
                end_y = center_y + rect_height // 2

                print(f"DEBUG: 사각형 그리기 시작: ({start_x}, {start_y}) → ({end_x}, {end_y})")

                # 마우스 이동 및 드래그
                import ctypes

                # 시작 위치로 이동
                ctypes.windll.user32.SetCursorPos(start_x, start_y)
                time.sleep(0.2)
                print(f"DEBUG: 마우스 시작 위치 이동 완료: ({start_x}, {start_y})")

                # 마우스 왼쪽 버튼 다운
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                time.sleep(0.2)
                print("DEBUG: 마우스 버튼 다운")

                # 끝 위치로 이동 (천천히)
                ctypes.windll.user32.SetCursorPos(end_x, end_y)
                time.sleep(0.3)
                print(f"DEBUG: 마우스 끝 위치 이동 완료: ({end_x}, {end_y})")

                # 마우스 왼쪽 버튼 업
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                time.sleep(0.8)  # 사각형 생성 완료까지 충분히 대기

                print("DEBUG: 사각형 그리기 완료")

            except Exception as draw_err:
                print(f"DEBUG: 사각형 그리기 실패: {draw_err}")
                return False

            # 방법 3: 사각형에 텍스트 입력
            try:
                # 사각형을 그리면 자동으로 선택 상태가 됨
                # 바로 타이핑하거나 F2로 편집 모드 진입

                # "{{이미지}}" 텍스트를 클립보드에 복사
                text_to_type = "{{이미지}}"
                QApplication.clipboard().setText(text_to_type)
                time.sleep(0.2)
                print(f"DEBUG: 클립보드에 텍스트 복사 완료: {text_to_type}")

                # 사각형이 선택된 상태에서 바로 타이핑 (F2 대신)
                # 일부 PowerPoint 버전에서는 바로 입력 가능
                print("DEBUG: 텍스트 붙여넣기 시작")
                win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(ord('V'), 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.1)
                win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.5)  # 붙여넣기 완료 대기

                print("DEBUG: 사각형 텍스트 입력 완료: {{이미지}}")

                # ESC 키로 선택 해제 (2번)
                print("DEBUG: 편집 모드 종료 및 선택 해제")
                for _ in range(2):
                    win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
                    time.sleep(0.1)
                    win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.2)

                print("DEBUG: 사각형 삽입 및 편집 완료")

                return True

            except Exception as text_err:
                print(f"DEBUG: 텍스트 입력 실패: {text_err}")
                import traceback
                traceback.print_exc()
                return False

        except Exception as e:
            print(f"ERROR: 키보드 자동화 사각형 삽입 실패: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _insert_ppt_image_rectangle(self):
        """PowerPoint에 {{이미지}} 텍스트가 포함된 사각형 삽입

        COM 상태 안정화를 위해 재시도 로직 포함
        """
        try:
            import win32com.client
            try:
                from win32com.client import constants
            except ImportError:
                constants = None

            # PowerPoint 인스턴스 가져오기 (재시도 로직)
            ppt = None
            max_retries = 3
            retry_delay = 0.5

            for attempt in range(max_retries):
                try:
                    print(f"DEBUG: PowerPoint Dispatch 시도 {attempt + 1}/{max_retries}")
                    ppt = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
                    ppt.Visible = True
                    try:
                        ppt.Activate()
                    except Exception:
                        pass
                    print(f"DEBUG: PowerPoint 인스턴스 준비 완료 (시도 {attempt + 1})")
                    break
                except Exception as e:
                    print(f"DEBUG: Dispatch 실패 (시도 {attempt + 1}): {e}")
                    if attempt < max_retries - 1:
                        print(f"DEBUG: {retry_delay}초 후 재시도...")
                        time.sleep(retry_delay)
                        retry_delay *= 1.5  # 점진적으로 대기 시간 증가
                    else:
                        # 모든 시도 실패
                        QMessageBox.warning(
                            self,
                            "경고",
                            "PowerPoint가 실행되지 않았거나 응답하지 않습니다.\n\n"
                            "다음 단계를 따라주세요:\n"
                            "1. PowerPoint를 실행합니다.\n"
                            "2. 프레젠테이션 파일을 엽니다.\n"
                            "3. 이미지를 삽입할 슬라이드를 선택합니다.\n"
                            "4. 잠시 기다린 후 다시 '이미지' 버튼을 클릭합니다."
                        )
                        return False

            try:
                slide = ppt.ActiveWindow.View.Slide
            except Exception as slide_err:
                print(f"DEBUG: 활성 슬라이드 확인 실패: {slide_err}")
                QMessageBox.warning(
                    self,
                    "경고",
                    "사각형을 삽입할 슬라이드를 찾을 수 없습니다.\n\n슬라이드를 선택한 뒤 다시 시도해주세요."
                )
                return False

            if not slide:
                QMessageBox.warning(
                    self,
                    "경고",
                    "사각형을 삽입할 슬라이드를 찾을 수 없습니다.\n\n슬라이드를 선택한 뒤 다시 시도해주세요."
                )
                return False

            # 사각형 삽입
            shape = slide.Shapes.AddShape(1, 100, 100, 200, 100)
            print("DEBUG: 사각형 삽입 성공")

            # 플레이스홀더 및 스타일 적용
            shape.TextFrame.TextRange.Text = "{{이미지}}"
            try:
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = 0xFFFFFF
                shape.Line.ForeColor.RGB = 0x000000
                shape.Line.DashStyle = getattr(constants, "msoLineDash", 4)
                shape.Line.Weight = 1.5
            except Exception as border_err:
                print(f"DEBUG: 사각형 테두리 설정 실패 (무시): {border_err}")
            shape.TextFrame.TextRange.Font.Size = 14
            shape.TextFrame.TextRange.Font.Bold = True
            shape.TextFrame.TextRange.Font.Color.RGB = 0x000000
            try:
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = getattr(constants, "ppAlignCenter", 2) if constants else 2
                shape.TextFrame.VerticalAnchor = getattr(constants, "msoAnchorMiddle", 3) if constants else 3
            except Exception as style_err:
                print(f"DEBUG: 텍스트 정렬 설정 실패 (무시): {style_err}")

            print("DEBUG: PPT에 이미지 사각형 삽입 완료")
            return True  # 성공 시 True 반환

        except Exception as e:
            print(f"ERROR: PPT 사각형 삽입 중 오류: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"사각형 삽입 중 오류 발생: {e}")
            return False  # 실패 시 False 반환

    def _resolve_active_ppt_slide(self, ppt, constants=None):
        """현재 사용자가 보고 있는 PowerPoint 슬라이드를 반환합니다."""
        try:
            # 슬라이드 쇼 모드 우선
            try:
                if ppt.SlideShowWindows.Count > 0:
                    slide = ppt.SlideShowWindows(1).View.Slide
                    if slide:
                        print(f"DEBUG: SlideShowWindow에서 활성 슬라이드 획득 - index {slide.SlideIndex}")
                        return slide
            except Exception as slideshow_err:
                print(f"DEBUG: SlideShowWindow 확인 실패: {slideshow_err}")

            window = ppt.ActiveWindow
            if not window:
                print("DEBUG: PowerPoint ActiveWindow가 없습니다.")
                return None

            # 보기 유형 보정 (슬라이드 정렬 보기 등)
            if constants:
                try:
                    view_type = getattr(window, "ViewType", None)
                    pp_slide_sorter = getattr(constants, "ppViewSlideSorter", None)
                    pp_notes_page = getattr(constants, "ppViewNotesPage", None)
                    if view_type in {pp_slide_sorter, pp_notes_page}:
                        selection = getattr(window, "Selection", None)
                        if selection and getattr(selection, "SlideRange", None) and selection.SlideRange.Count > 0:
                            slide_index = selection.SlideRange(1).SlideIndex
                            window.View.GotoSlide(slide_index)
                            time.sleep(0.1)
                        else:
                            window.View.GotoSlide(1)
                            time.sleep(0.1)
                except Exception as view_type_err:
                    print(f"DEBUG: ViewType 보정 실패: {view_type_err}")

            # 일반 보기에서 View.Slide 시도
            try:
                slide = getattr(window.View, "Slide", None)
                if slide:
                    return slide
            except Exception as view_err:
                print(f"DEBUG: ActiveWindow.View.Slide 접근 실패: {view_err}")

            # SlideRange 기반 접근
            try:
                slide_range = getattr(window.View, "SlideRange", None)
                if slide_range and slide_range.Count > 0:
                    return slide_range(1)
            except Exception as range_err:
                print(f"DEBUG: View.SlideRange 접근 실패: {range_err}")

            # Selection 기반 슬라이드 추출
            try:
                selection = getattr(window, "Selection", None)
                if selection and getattr(selection, "SlideRange", None):
                    if selection.SlideRange.Count > 0:
                        return selection.SlideRange(1)
            except Exception as sel_err:
                print(f"DEBUG: Selection 기반 슬라이드 확인 실패: {sel_err}")

            # View에 SlideIndex만 있을 수도 있음
            try:
                slide_index = getattr(window.View, "SlideIndex", None)
                if slide_index and ppt.ActivePresentation:
                    return ppt.ActivePresentation.Slides(slide_index)
            except Exception as index_err:
                print(f"DEBUG: SlideIndex 기반 슬라이드 확인 실패: {index_err}")

            # 마지막 폴백: 활성 프레젠테이션의 첫 슬라이드
            try:
                if ppt.ActivePresentation and ppt.ActivePresentation.Slides.Count > 0:
                    print("DEBUG: 폴백으로 첫 번째 슬라이드를 반환")
                    return ppt.ActivePresentation.Slides(1)
            except Exception as pres_err:
                print(f"DEBUG: 폴백 슬라이드 확인 실패: {pres_err}")
        except Exception as err:
            print(f"DEBUG: 활성 슬라이드 확인 중 오류: {err}")

        return None

def main():
    app = QApplication(sys.argv)
    ex = MailMergeApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
