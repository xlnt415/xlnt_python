# 라이브러리
import os, sys, pandas as pd, json, re, logging
# from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QTableWidget, QTableWidgetItem, QLineEdit, QHBoxLayout, QAction, QMessageBox, QInputDialog, QShortcut, QCompleter, QComboBox, QStyledItemDelegate, QTabWidget, QLabel, QUndoStack, QUndoCommand, QDialog, QDialogButtonBox, QGridLayout
from PyQt5.QtGui import QKeySequence, QDrag
from PyQt5.QtCore import Qt, QEvent, QMimeData


# 매체 자동완성 Class
class CompleterDelegate(QStyledItemDelegate):
    def __init__(self, completer, parent=None):
        super().__init__(parent)
        self.completer = completer


    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setCompleter(self.completer)
        self.completer.setFilterMode(Qt.MatchContains)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        return editor
    
    def setEditorData(self, editor, index):
        # The text that is currently in the cell
        text = index.model().data(index, Qt.EditRole)
        editor.setText(text)  # Add current item
        # editor.setCurrentText(text)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text(), Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)
        

class UndoRedoCommand(QUndoCommand):
    def __init__(self, table, old_value, new_value, row, col):
        super().__init__()
        self.table = table
        self.old_value = old_value
        self.new_value = new_value
        self.row = row
        self.col = col

    def undo(self):
        self.table.blockSignals(True)
        self.table.item(self.row, self.col).setText(self.old_value)
        self.table.blockSignals(False)

    def redo(self):
        self.table.blockSignals(True)
        self.table.item(self.row, self.col).setText(self.new_value)
        self.table.blockSignals(False)


###
class ExcelLikeApp(QWidget):
    # global headers
    # headers = ["매체", "상품", "광고비", "노출량", "클릭수", "조회수"]  # 클래스 변수로 headers 정의
    
    def __init__(self):
        super().__init__()
        logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)
        self.logger.info("Application started")
    
        self.logger.info("Undo stack initialized")
        self.selected_folder = None
        self.headers = ["매체", "상품", "광고비", "노출량", "클릭수", "조회수"]
        self.headers_1 = ["매체", "상품", "광고비", "노출량", "클릭수", "조회수"]
        self.initUI()
        self.ctrl_pressed = False
        self.product_completer = None
        self.setup_completers()
        # 딕셔너리 로드
        self.load_media_product_dict()
        # self.folder_name_input = None
        # self.setup_completers()

    def initUI(self):
        self.layout = QVBoxLayout(self)
        # Tab widget
        self.tab_widget = QTabWidget(self)
        self.layout.addWidget(self.tab_widget)

        # Folder selection section
        folder_layout = QHBoxLayout()
        self.folder_button = QPushButton('폴더 선택', self)
        self.folder_button.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.folder_button)

        self.folder_name_input = QLineEdit(self)
        self.folder_name_input.setPlaceholderText('저장할 파일 이름 입력')
        folder_layout.addWidget(self.folder_name_input)
        self.layout.addLayout(folder_layout)

        # Table display section
        table_tab = QWidget()
        self.undo_stack = QUndoStack(self)
        table_tab_layout = QVBoxLayout(table_tab)
        self.table = QTableWidget(50, 6)
        self.table.itemChanged.connect(self.onItemChanged)
        headers_1 = ["매체", "상품", "광고비", "노출량", "클릭수", "조회수"]
        self.table.setHorizontalHeaderLabels(headers_1)
        table_tab_layout.addWidget(self.table)
        self.tab_widget.addTab(table_tab, "재적재")
        
        # Setup budget section code
        budget_tab = QWidget()
        budget_tab_layout = QVBoxLayout(budget_tab)
        self.budget_table = QTableWidget(15, 2)
        headers_2 = ["노출량", "광고비"]
        self.budget_table.setHorizontalHeaderLabels(headers_2)
        budget_tab_layout.addWidget(self.budget_table)

        budget_input_layout = QHBoxLayout()
        self.budget_input = QLineEdit(self)
        self.budget_input.setPlaceholderText('광고비 입력')
        self.budget_button = QPushButton('광고비 비율', self)
        self.budget_button.clicked.connect(self.calculate_budget_percentage)
        
        self.reset_button = QPushButton('배분초기화', self)  # Create the new button
        self.reset_button.clicked.connect(self.clear_budget_table)
        budget_input_layout.addWidget(self.budget_input)
        budget_input_layout.addWidget(self.budget_button)
        budget_input_layout.addWidget(self.reset_button)
        budget_tab_layout.addLayout(budget_input_layout)
        
        self.tab_widget.addTab(budget_tab, "광고비배분")
        
        self.table.itemChanged.connect(self.onItemChanged)
        self.budget_table.itemChanged.connect(self.onBudgetItemChanged)   
        
        self.media_product_dict = {}
        self.load_media_product_dict()
        self.setup_completers()
 
        control_layout = QHBoxLayout()
        self.add_indicator_button = QPushButton('지표수정', self)
        self.add_indicator_button.clicked.connect(self.edit_headers)
        control_layout.addWidget(self.add_indicator_button)

        self.total_button = QPushButton('총합계', self)
        self.total_button.clicked.connect(self.sum_entries)
        control_layout.addWidget(self.total_button)

        self.layout.addLayout(control_layout)

        # Shortcuts and actions
        total_shortcut = QShortcut(QKeySequence('Ctrl++'), self)
        total_shortcut.activated.connect(self.sum_entries)
        
        self.setupActions()
        
        # Save and clear buttons
        self.save_button = QPushButton('저장', self)
        self.save_button.clicked.connect(self.save_data)
        self.save_button.setShortcut('Ctrl+S')
        control_layout.addWidget(self.save_button)

        self.clear_button = QPushButton('초기화', self)
        self.clear_button.clicked.connect(self.clear_data)
        self.clear_button.setShortcut('Ctrl+Delete')
        control_layout.addWidget(self.clear_button)

        # Undo/Redo 기능 추가
        self.undoStack = QUndoStack(self)

        self.undo_shortcut = QShortcut(QKeySequence('Ctrl+Z'), self)
        self.undo_shortcut.activated.connect(self.undoStack.undo)
        self.logger.info("Undo shortcut set up")

        self.redo_shortcut = QShortcut(QKeySequence('Ctrl+Shift+Z'), self)
        self.redo_shortcut.activated.connect(self.undoStack.redo)
        self.logger.info("Redo shortcut set up")


# 복붙 기능
    def setupActions(self):
        # Setup clipboard actions without assigning specific shortcuts
        self.paste_action = QAction('붙여넣기', self)
        self.paste_action.triggered.connect(lambda: self.paste_from_clipboard(QApplication.focusWidget()))
        
        self.copy_action = QAction('복사', self)
        self.copy_action.triggered.connect(self.copy_to_clipboard)

        # Add actions to the application; they will check for the focused widget
        self.addAction(self.paste_action)
        self.addAction(self.copy_action)

        # Assign shortcuts dynamically based on focus changes
        self.installEventFilter(self)

# 복사 붙여넣기, 오려두기, 삭제 기능 추가 완료
    def eventFilter(self, source, event):
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_V and event.modifiers() == Qt.ControlModifier:
                if isinstance(QApplication.focusWidget(), QTableWidget):
                    self.paste_from_clipboard(QApplication.focusWidget())
                    return True
            elif event.key() == Qt.Key_C and event.modifiers() == Qt.ControlModifier:
                if isinstance(QApplication.focusWidget(), QTableWidget):
                    self.copy_to_clipboard()
                
            elif event.key() == Qt.Key_X and event.modifiers() == Qt.ControlModifier:  # 오려두기
                if isinstance(QApplication.focusWidget(), QTableWidget):
                    self.cut_to_clipboard()
                    return True
                return True
                
            # 삭제 기능 처리
            if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
                if source is self.table or source is self.budget_table:
                    selected_items = source.selectedItems()
                    for item in selected_items:
                        item.setText("")  # 선택된 항목의 텍스트를 지움
                    return True  # 키 이벤트가 처리됨
        return super().eventFilter(source, event)

# 오려두기 기능
    def cut_to_clipboard(self):
        focus_widget = QApplication.focusWidget()
        if isinstance(focus_widget, QTableWidget):
            selected_ranges = focus_widget.selectedRanges()
            if not selected_ranges:
                return
            clipboard = QApplication.clipboard()
            data = ""
            selected = selected_ranges[0]
            for row in range(selected.topRow(), selected.bottomRow() + 1):
                row_data = []
                for col in range(selected.leftColumn(), selected.rightColumn() + 1):
                    item = focus_widget.item(row, col)
                    row_data.append(item.text() if item else '')
                    if item:
                        item.setText('')  # 원래 위치에서 제거
                data += '\t'.join(row_data) + '\n'
            clipboard.setText(data.strip())

    def paste_from_clipboard(self, table_widget):
        if not isinstance(table_widget, QTableWidget):
            return
        clipboard = QApplication.clipboard()
        data = clipboard.text()
        selected_row = table_widget.currentRow()
        selected_col = table_widget.currentColumn()
        for i, row_data in enumerate(data.splitlines()):
            cols = row_data.split('\t')
            for j, cell_data in enumerate(cols):
                if selected_row + i < table_widget.rowCount() and selected_col + j < table_widget.columnCount():
                    item = QTableWidgetItem(cell_data)
                    table_widget.setItem(selected_row + i, selected_col + j, item)

    def copy_to_clipboard(self):
        focus_widget = QApplication.focusWidget()
        if isinstance(focus_widget, QTableWidget):
            selected_ranges = focus_widget.selectedRanges()
            if not selected_ranges:
                return
            clipboard = QApplication.clipboard()
            data = ""
            selected = selected_ranges[0]
            for row in range(selected.topRow(), selected.bottomRow() + 1):
                row_data = []
                for col in range(selected.leftColumn(), selected.rightColumn() + 1):
                    item = focus_widget.item(row, col)
                    row_data.append(item.text() if item else '')
                data += '\t'.join(row_data) + '\n'
            clipboard.setText(data.strip())
# 저장 경로 설정

    def save_data(self):
        if self.selected_folder and self.folder_name_input.text():
            # 파일 이름에서 '\n'이나 '\t'를 '_'로 교체
            sanitized_folder_name = self.folder_name_input.text().replace('\n', '_').replace('\t', '_')
            file_path = f"{self.selected_folder}/{sanitized_folder_name}.xlsx"
            
            
            if os.path.exists(file_path):
                QMessageBox.warning(self, "파일 저장 오류", "동일한 파일명이 존재합니다. 다른 이름을 사용해 주세요.")
                return

            data = {header: [] for header in self.headers_1}  # headers_1을 사용하도록 수정
            for row in range(self.table.rowCount()):
                for col in range(min(self.table.columnCount(), len(self.headers_1))):  # headers_1 길이에 맞춤
                    item = self.table.item(row, col)
                    if item and item.text():  # 값이 있고, 비어있지 않은 경우
                        text = item.text()
                        if col > 1:  # 2열 이후 숫자 변환 처리
                            try:
                                numeric_value = int(text)
                            except ValueError:
                                numeric_value = text  # 숫자 변환 실패 시 원본 텍스트 저장
                        else:
                            numeric_value = text
                    else:
                        numeric_value = ""  # 셀이 비어있는 경우 빈 문자열 저장

                    data[self.headers_1[col]].append(numeric_value)

            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "저장 성공", f"데이터가 {file_path}에 성공적으로 저장되었습니다.")
        elif not self.select_folder:
            QMessageBox.warning(self, '입력 오류', '파일 경로를 설정해주세요.')

        else:
            QMessageBox.warning(self, "입력 오류", "폴더와 파일 이름을 모두 입력해주세요.")
    def select_folder(self):
        try:
            self.selected_folder = QFileDialog.getExistingDirectory(self, "폴더 선택")
            if self.selected_folder:
                QMessageBox.information(self, '폴더 선택 완료', f"선택된 폴더: {self.selected_folder}")
            else:
                raise Exception("폴더를 선택하지 않았습니다.")
        except Exception as e:
            QMessageBox.warning(self, "오류", str(e))

    # def edit_headers(self):
    #     current_headers = ", ".join(self.headers_1)
    #     text, ok = QInputDialog.getText(self, '지표수정', '현재 지표들(,구분):', QLineEdit.Normal, current_headers)
        
    #     if ok and text:
    #         self.headers_1 = [header.strip() for header in text.split(',') if header.strip()]
    #         self.table.setColumnCount(len(self.headers_1))
    #         self.table.setHorizontalHeaderLabels(self.headers_1)
    
    def edit_headers(self):
        current_headers = ", ".join(self.headers_1)
        dialog = QDialog(self)
        dialog.setWindowTitle("지표수정")
        dialog_layout = QVBoxLayout(dialog)
        
        label = QLabel("현재 지표들(,구분):", dialog)
        dialog_layout.addWidget(label)
        
        line_edit = QLineEdit(dialog)
        line_edit.setText(current_headers)
        dialog_layout.addWidget(line_edit)
        
        # 미리 정의된 헤더를 추가하는 버튼
        button_grid = QGridLayout()
        button_labels = [
            "SNS 액션", "유입", "세션", "사용자",
            "이벤트참여", "회원가입", "장바구니", "구매",
            "매출액", "인스톨", "앱실행", "기타 전환"
        ]
        
        positions = [(i, j) for i in range(4) for j in range(3)]
        for position, label in zip(positions, button_labels):
            button = QPushButton(label, dialog)
            button.clicked.connect(lambda _, text=label: self.add_text(line_edit, text))
            button_grid.addWidget(button, *position)
        
        dialog_layout.addLayout(button_grid)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        dialog_layout.addWidget(button_box)

        if dialog.exec_() == QDialog.Accepted:
            text = line_edit.text()
            if text:
                self.headers_1 = [header.strip() for header in text.split(',') if header.strip()]
                self.table.setColumnCount(len(self.headers_1))
                self.table.setHorizontalHeaderLabels(self.headers_1)
    
    def add_text(self, line_edit, text):
        current_text = line_edit.text()
        if current_text:
            line_edit.setText(f"{current_text}, {text}")
        else:
            line_edit.setText(text)
        
# 총합계_4 _ 테이블 크기가 다시 줄어듦 _ 해결 완료
    '''
    def sum_entries(self):
        summary = {}
        any_valid_data = False

        headers = self.headers_1  # headers_1을 사용하도록 업데이트
        for row in range(self.table.rowCount()):
            media_item = self.table.item(row, 0)
            product_item = self.table.item(row, 1)

            if media_item and media_item.text().strip() and product_item and product_item.text().strip():
                any_valid_data = True
                media = media_item.text().strip()
                product = product_item.text().strip()
                key = (media, product)

                if key not in summary:
                    summary[key] = {'sum': [None] * (len(headers) - 2), 'row_index': []}
                summary[key]['row_index'].append(row)

                for col in range(2, len(headers)):
                    item = self.table.item(row, col)
                    if item and item.text().strip():
                        try:
                            value = int(float(item.text().replace(',', '').strip()))
                            if summary[key]['sum'][col - 2] is None:
                                summary[key]['sum'][col - 2] = value
                            else:
                                summary[key]['sum'][col - 2] += value
                        except ValueError:
                            continue

        if not any_valid_data:
            QMessageBox.warning(self, "Missing Information", "모든 행에서 유효한 매체 및 상품 데이터를 찾을 수 없습니다.")
            return

        self.table.clearContents()
        self.table.setRowCount(len(summary) + self.table.rowCount())

        current_row = 0
        for key, data in summary.items():
            self.table.setItem(current_row, 0, QTableWidgetItem(key[0]))
            self.table.setItem(current_row, 1, QTableWidgetItem(key[1]))
            for i, value in enumerate(data['sum']):
                cell_value = "" if value is None else str(value)
                self.table.setItem(current_row, i + 2, QTableWidgetItem(cell_value))
            current_row += 1

        QMessageBox.information(self, ".")
    '''    
    def sum_entries(self):
        summary = {}
        any_valid_data = False

        headers = self.headers_1  # headers_1을 사용하도록 업데이트
        for row in range(self.table.rowCount()):
            media_item = self.table.item(row, 0)
            product_item = self.table.item(row, 1)

            if media_item and media_item.text().strip() and product_item and product_item.text().strip():
                any_valid_data = True
                media = media_item.text().strip()
                product = product_item.text().strip()
                key = (media, product)

                if key not in summary:
                    summary[key] = {'sum': [None] * (len(headers) - 2), 'row_index': []}
                summary[key]['row_index'].append(row)

                for col in range(2, len(headers)):
                    item = self.table.item(row, col)
                    if item and item.text().strip():
                        try:
                            value = int(float(item.text().replace(',', '').strip()))
                            if summary[key]['sum'][col - 2] is None:
                                summary[key]['sum'][col - 2] = value
                            else:
                                summary[key]['sum'][col - 2] += value
                        except ValueError:
                            continue

        if not any_valid_data:
            QMessageBox.warning(self, "Missing Information", "모든 행에서 유효한 매체 및 상품 데이터를 찾을 수 없습니다.")
            return

        self.table.clearContents()
        self.table.setRowCount(len(summary) + self.table.rowCount())

        current_row = 0
        summary_info = []

        for key, data in summary.items():
            self.table.setItem(current_row, 0, QTableWidgetItem(key[0]))
            self.table.setItem(current_row, 1, QTableWidgetItem(key[1]))
            for i, value in enumerate(data['sum']):
                cell_value = "" if value is None else str(value)
                self.table.setItem(current_row, i + 2, QTableWidgetItem(cell_value))
            total_sum = sum([v for v in data['sum'] if v is not None])
            summary_info.append(f"{key[0]} - {key[1]}: {total_sum}")
            current_row += 1

        total_count = sum([sum([v for v in data['sum'] if v is not None]) for data in summary.values()])
        QMessageBox.information("합산완료")



    def calculate_budget_percentage(self):
        try:
            # raw_budget = int(self.budget_input.text())  # 예산 입력 받기
            raw_budget = self.budget_input.text()  # 예산 입력 받기
            sanitized_budget = re.sub(r'[^\d]', '', raw_budget)
            budget = int(sanitized_budget)

            values = []
            for row in range(self.budget_table.rowCount()):
                item = self.budget_table.item(row, 0)
                if item is not None:
                    try:
                        # value = int(item.text())  # 값 가져오기
                        raw_value = item.text().replace(',', '').replace('\\', '').replace('₩', '')
                        value = int(raw_value)
                        values.append(value)
                    except ValueError:
                        # 값이 숫자로 변환될 수 없는 경우 무시
                        continue

            total_value = sum(values)
            if total_value == 0:
                QMessageBox.warning(self, "오류", "총합이 0입니다. 유효한 숫자를 입력해주세요.")
                return

            for row, value in enumerate(values):
                # 각 값의 비율에 따라 예산을 계산
                proportion = value / total_value
                budget_for_cell = int(budget * proportion)  # 결과를 정수로 변환
                # 결과를 두 번째 테이블의 두 번째 열에 표시
                if row < self.budget_table.rowCount():
                    self.budget_table.setItem(row, 1, QTableWidgetItem(str(budget_for_cell)))

        except ValueError:
            QMessageBox.warning(self, "입력 오류", "입력한 광고비용이 숫자가 아닙니다. 유효한 숫자를 입력해주세요.")
    def clear_data(self):
        self.table.setRowCount(30)  # 테이블 행 수를 30으로 설정
        self.table.setColumnCount(len(self.headers))    
        
        # Block signals for both tables to prevent triggering events
        self.table.blockSignals(True)
        self.budget_table.blockSignals(True)
        # self.table.blockSignals(False)
        # self.budget_table.blockSignals(False)

        # Clear the main data table
        for row in range(self.table.rowCount()):
            for column in range(self.table.columnCount()):
                self.table.setItem(row, column, QTableWidgetItem(""))  # Clear any existing text items
                # Check for QComboBox in the main table and reset it
                if column == 1:  # Assuming column 1 has the QComboBox
                    combo_box = self.table.cellWidget(row, column)
                    if combo_box:
                        combo_box.setCurrentIndex(0)  # Reset to the first item (default/placeholder)
                        combo_box.clear()  # Optionally clear all items if they are to be re-populated later

        # Clear the budget data table
        # for row in range(self.budget_table.rowCount()):
        #     for column in range(self.budget_table.columnCount()):
        #         self.budget_table.setItem(row, column, QTableWidgetItem(""))  # Clear any existing text items

        self.headers_1 = self.headers.copy()  # 클래스 변수 headers를 headers_1에 복사
        self.table.setHorizontalHeaderLabels(self.headers_1)  # 테이블 헤더를 원래 headers로 업데이트하는 방법 구상
        # print(len(self.headers_1))
        
        # Re-enable signals after modifications
        self.table.blockSignals(False)
        # self.budget_table.blockSignals(False)
        # self.table.blockSignals(True)
        # self.budget_table.blockSignals(True)

        # Reset headers to original headers
        # self.table.setHorizontalHeaderLabels(self.headers)  # Assuming self.headers is defined and holds original headers
        # 파일 이름 입력 필드 초기화

        # Clear the filename input field
        self.folder_name_input.clear()

    def clear_budget_table(self):
        """Clears the budget table, resetting it to its default state."""
        self.budget_table.blockSignals(True)  # Block signals to prevent unnecessary events
        for row in range(self.budget_table.rowCount()):
            for column in range(self.budget_table.columnCount()):
                self.budget_table.setItem(row, column, QTableWidgetItem(""))  # Clear text items

        self.budget_table.blockSignals(False)


    
    # 테이블 자동 정수 변환
    '''
    def sanitize_data(self, item):
        if item is None:
            return
        col = item.column()
        if col > 1:  # Assuming columns 2 and beyond need numeric processing
            text = item.text().replace(',', '').replace('\\', '').replace('₩', '')
            try:
                # Set the item text to an integer converted value if possible
                item.setText(str(int(text)))
            except ValueError:
                # If conversion fails, reset the text to just the sanitized text without commas
                item.setText(text)
                
    def sanitize_budget_data(self, item):
        if item is None:
            return
        text = item.text().replace(',', '').replace('\\', '').replace('₩', '')
        try:
            # Convert the cleaned text to integer and set it back to the item
            item.setText(str(int(text)))
        except ValueError:
            # Reset the text to just the sanitized text if conversion fails
            item.setText(text)
    
    '''

        # QMessageBox.information(self, "합산 완료", "데이터가 합산되었습니다.")
# re사용 테이블 자동 정수 변환
    
    # def sanitize_data(self, item):
    #     if item is None:
    #         return
    #     col = item.column()
    #     if col > 1:  # 2번째 컬럼 이후는 숫자 처리
    #         text = item.text()
            
    #         if int(text) - 1 == 0:
    #             item.setText('')
    #         # 정규 표현식을 사용하여 숫자 이외의 문자를 제거
    #         sanitized_text = re.sub(r'[^\d]', '', text)
    #         try:
    #             # 가능한 경우 정수로 변환된 값을 텍스트로 설정
    #             item.setText(str(int(sanitized_text)))
    #         except ValueError:
    #             # 변환에 실패하면 콤마 없이 정리된 텍스트로 설정
    #             item.setText(sanitized_text)
                

    def onItemChanged(self, item):
        if item is None:
            return
        
        row = item.row()
        col = item.column()
        new_value = item.text()
        
        old_value = getattr(self, 'last_item_text', None)

        # 값이 실제로 변경되었는지 확인
        if old_value != new_value:
            sanitized_value = self.sanitize_data(new_value, col)
            if sanitized_value != new_value:
                self.table.blockSignals(True)
                item.setText(sanitized_value)
                self.table.blockSignals(False)
                new_value = sanitized_value
            
            self.handle_item_changed(old_value, new_value, row, col)

        self.last_item_text = new_value
        
    def sanitize_data(self, value, col):
        # 값을 자동으로 수정하는 함수
        if col > 1:  # 2번째 열 이후의 숫자들을 처리합니다
            text = value
            if text == '':
                return ''
            sanitized_text = re.sub(r'[^0-9.]', '', text)
            sanitized_text = sanitized_text.split(".")[0]
            if sanitized_text == '':
                return ''
            try:
                int_value = int(sanitized_text)
                if int_value == 0:
                    return ''
                else:
                    return str(int_value)
            except ValueError:
                return sanitized_text
        return value

    def handle_item_changed(self, old_value, new_value, row, col):
        # 이전 값과 새로운 값을 인식하고 Undo/Redo 명령을 처리하는 함수
        if old_value != new_value:
            command = UndoRedoCommand(self.table, old_value, new_value, row, col)
            self.undoStack.push(command)

## 자동완성
### 딕셔너리 업로드    
    def load_media_product_dict(self):
        try:
            with open('media_product_data.json', 'r', encoding='utf-8') as file:
                data_dict = json.load(file)
            
            self.media_product_dict = data_dict
        except FileNotFoundError:
            # Properly handle the error, e.g., by logging it or modifying the UI to show an error message
            print("The file was not found. Please check the filename and try again.")

    # 자동완성
    
    # def setup_completers(self):
    #     self.media_completer = QCompleter(list(self.media_product_dict.keys()))
    #     self.media_completer.setCaseSensitivity(Qt.CaseInsensitive)
    #     self.media_completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
    #     self.table.setItemDelegateForColumn(0, self.media_completer)

        # self.product_completer = QCompleter([])
        # self.product_completer.setCaseSensitivity(Qt.CaseInsensitive)
        # self.product_completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        # self.table.setItemDelegateForColumn(1, self.product_completer)
        # self.table.cellChanged.connect(self.update_product_completer)

    # 자동완성2
    def setup_completers(self):
        if hasattr(self, 'media_product_dict'):
            media_completer = QCompleter(list(self.media_product_dict.keys()))
            media_completer.setCaseSensitivity(Qt.CaseInsensitive)
            media_completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
            media_delegate = CompleterDelegate(media_completer, self.table)
            self.table.setItemDelegateForColumn(0, media_delegate)
            
            
        ## 상품 자동완성
            self.product_completer = QCompleter([])
            self.product_completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.product_completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
            product_delegate = CompleterDelegate(self.product_completer, self.table)
            self.table.setItemDelegateForColumn(1, product_delegate)       
            
            # Connect media cell changes to update product completer
            self.table.cellChanged.connect(self.update_product_completer)

    def enhanced_autocomplete_setup(self):
        """
        Enhance the autocomplete setup to dynamically update suggestions based on user input.
        """
        # Assume the media_product_dict is already loaded as a dictionary where keys are media types.
        if hasattr(self, 'media_product_dict'):
            media_completer = QCompleter(list(self.media_product_dict.keys()))
            media_completer.setCaseSensitivity(Qt.CaseInsensitive)
            media_completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
            media_delegate = CompleterDelegate(media_completer, self.table)
            
            # Connecting text changes to update suggestions dynamically
            self.media_line_edit.textChanged.connect(self.update_media_completer)

            self.table.setItemDelegateForColumn(0, media_delegate)

    def create_undo_action(self, old_value, new_value, row, col):
        self.logger.info(f"Creating undo action: old={old_value}, new={new_value}, row={row}, col={col}")
        command = UndoRedoCommand(self.table, old_value, new_value, row, col)
        self.undo_stack.push(command)
        self.logger.info(f"Undo action created and pushed to stack. Stack count: {self.undo_stack.count()}")

    def update_product_completer(self, row, col):
        if col == 0:  # Media column
            media_value = self.table.item(row, col).text()
            if media_value in self.media_product_dict:
                products = self.media_product_dict[media_value]
                self.product_completer.model().setStringList(products)
            else:
                self.product_completer.model().setStringList([])
                

    def onBudgetItemChanged(self, item):
        if not hasattr(self, 'last_budget_item_text'):
            self.last_budget_item_text = None
        if self.last_budget_item_text is not None:
            old_value = self.last_budget_item_text
            new_value = item.text()
            row = item.row()
            col = item.column()
            command = UndoRedoCommand(self.budget_table, old_value, new_value, row, col)
            self.undoStack.push(command)
        self.last_budget_item_text = item.text()

    def keyPressEvent(self, event):
        self.logger.info(f"Key pressed: {event.key()}")
        if event.key() == Qt.Key_Z and event.modifiers() == Qt.ControlModifier:
            self.logger.info("Ctrl+Z detected")
            self.undo_action()
        elif event.key() == Qt.Key_Z and event.modifiers() == (Qt.ControlModifier | Qt.ShiftModifier):
            self.logger.info("Ctrl+Shift+Z detected")
            self.redo_action()
        elif event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            for item in self.table.selectedItems():
                item.setText('')
        super().keyPressEvent(event)

    def setup_ui(self):
        self.table.setRowCount(10)
        self.table.setColumnCount(2)
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        self.setLayout(layout)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelLikeApp()
    ex.show()
    sys.exit(app.exec_())