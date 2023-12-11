import os

from PyQt5.QtGui import QIcon, QPixmap, QFont
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QLabel, QPushButton, QMessageBox, QFileDialog
from order_processor import OrderProcessor
from datetime import datetime

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('For Chika Note')

        self.resize(700, 500)

        # 배경색 설정
        self.setStyleSheet("background-color: white;")

        #폰트 설정
        font = QFont()
        font.setPointSize(12)  # 폰트 크기 설정
        font.setBold(True)  # 폰트를 볼드체로 설정

        # 파일 선택 버튼에 스타일 적용
        file_button_style = '''
                    QPushButton {
                        background-color: #fdc3d1; /* 배경색 */
                        color: black; /* 텍스트 색상 */
                        border: none;
                        
                        border-radius: 15px; /* 버튼을 동그랗게 만듦 */
                        
                    }
                    
                    

                    QPushButton:hover {
                        background-color: #fab6c7; /* 마우스 호버 시 배경색 변경 */
                    }
                '''

        #실행 버튼 스타일
        # 실행 버튼에 스타일 및 폰트 적용
        run_button_style = '''
                    QPushButton {
                        background-color: #fdc3d1; /* 배경색 */
                        color: black; /* 텍스트 색상 */
                        border: none;
                        border-radius: 15px; /* 버튼을 동그랗게 만듦 */

                    }

                    QPushButton:hover {
                        background-color: #fab6c7; /* 마우스 호버 시 배경색 변경 */
                    }
                '''
        # 이미지를 표시할 QLabel 생성
        label = QLabel(self)

        image_path = resource_path('./static/chikaIsCute.jpeg')

        pixmap = QPixmap(image_path)
        label.setPixmap(pixmap)
        label.setGeometry(220, 10, 700, 300)

        # 파일 선택란
        left_tag_select = QLabel(self)
        left_tag_select.setText('파일 경로 :')
        left_tag_select.setGeometry(85, 250, 100, 30)
        left_tag_select.setFont(font)

        self.excel_file = QLineEdit(self)
        self.excel_file.setGeometry(150, 250, 330, 30)
        self.excel_file.setReadOnly(True)

        file_button = QPushButton('파일 선택', self)
        file_button.clicked.connect(self.showDialog)
        file_button.setGeometry(500, 250, 100, 30)
        file_button.setStyleSheet(file_button_style)
        file_button.setFont(font)

        # 경로 선택란
        left_tag_route = QLabel(self)
        left_tag_route.setText('저장할 경로 :')
        left_tag_route.setGeometry(75, 300, 100, 30)
        left_tag_route.setFont(font)

        self.file_path = QLineEdit(self)
        self.file_path.setGeometry(150, 300, 330, 30)
        self.file_path.setReadOnly(True)

        file_button = QPushButton('경로 선택', self)
        file_button.clicked.connect(self.showDialogRoute)
        file_button.setGeometry(500, 300, 100, 30)
        file_button.setStyleSheet(file_button_style)
        file_button.setFont(font)

        #실행 버튼
        run_button = QPushButton('실행', self)
        run_button.setGeometry(300, 370, 100, 30)
        run_button.clicked.connect(self.run_processing)
        run_button.setStyleSheet(run_button_style)
        run_button.setFont(font)

        self.show()

    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self, '파일 선택', '', '엑셀 파일 (*.xlsx);;모든 파일 (*)')

        if fname[0]:  # 사용자가 파일을 선택한 경우
            self.excel_file.setText(fname[0])

    def showDialogRoute(self):
        dirname = QFileDialog.getExistingDirectory(self, 'Open Directory', '/home')

        if dirname:
            self.file_path.setText(dirname)

    def run_processing(self):
        msg_box = QMessageBox()

        if not self.excel_file.text():
            msg_box.warning(self, '경고', '파일을 선택하세요.')
            return

        if not self.file_path.text():
            msg_box.warning(self, '경고', '저장할 경로를 선택하세요.')
            return

        input_file_path = self.excel_file.text()
        current_date = datetime.now().strftime("%Y%m%d")
        output_file_path = f"{self.file_path.text()}/{current_date}_발주 수량.xlsx"

        order_processor = OrderProcessor(input_file_path)
        order_processor.process_orders()
        order_processor.save_result_to_excel(output_file_path)

        msg_box.information(self, '완료', '작업이 완료되었습니다.')

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
