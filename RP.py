import os
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QLabel, QGridLayout, QWidget


class RP_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RP")
        self.setGeometry(100, 100, 800, 600)
        self.initUI()

    def initUI(self):
        # 创建一个有菜单栏的窗口
        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)

        # 创建一个网格布局
        gridLayout = QGridLayout()
        self.centralWidget.setLayout(gridLayout)

        # 创建一个输入框
        self.input_text = QLineEdit(self)
        self.input_text.setPlaceholderText("请输入内容")
        gridLayout.addWidget(self.input_text, 0, 0, 1, 2)

        # 创建一个按钮
        self.button = QPushButton("转换", self)
        gridLayout.addWidget(self.button, 1, 0)

        # 创建一个输出标签
        self.output_label = QLabel(self)
        self.output_label.setText("输出结果")
        gridLayout.addWidget(self.output_label, 1, 1)

        # 按钮点击事件
        self.button.clicked.connect(self.convert)

    def convert(self):
        # 获取输入框的内容
        input_text = self.input_text.text()
        # 转换为大写
        output_text = input_text.upper()
        # 设置输出标签的内容
        self.output_label.setText(output_text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = RP_MainWindow()
    MainWindow.show()
    sys.exit(app.exec())