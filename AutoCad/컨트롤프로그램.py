import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QProgressBar
from PyQt5.QtGui import QFont
from New_folder_creater import create_folder
from PyQt5.QtCore import QTimer


class PythonRunner(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python Runner")
        self.setGeometry(100, 100, 400, 300)

        layout = QVBoxLayout()

        # Create button to create folder
        self.create_folder_button = QPushButton("폴더 생성하기")
        self.create_folder_button.clicked.connect(self.create_new_folder)
        self.create_folder_button.setFixedSize(200, 100)  # 버튼의 크기를 조정함

        # 버튼의 폰트 설정
        font = QFont()
        font.setPointSize(10)  # 텍스트 크기를 20으로 설정
        font.setBold(True)  # 텍스트를 굵게 설정
        self.create_folder_button.setFont(font)

        layout.addWidget(self.create_folder_button)

        self.progress_bar = QProgressBar(self)
        layout.addWidget(self.progress_bar)
        self.progress_bar.setVisible(False)  # 처음에는 보이지 않도록 설정

        self.setLayout(layout)

    def create_new_folder(self):
        # Select folder
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")

        if folder_path:
            # Show progress bar
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)  # 초기화
            self.progress_bar.setMaximum(100)  # 최대값 설정

            # Call create_folder function with selected folder path
            create_folder(folder_path)

            # Hide progress bar after completion with a delay
            QTimer.singleShot(1000, self.hide_progress_bar)

    def hide_progress_bar(self):
        self.progress_bar.setVisible(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PythonRunner()
    window.show()
    sys.exit(app.exec_())
