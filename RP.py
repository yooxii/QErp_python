import os
import sys
import json
from rich import inspect
import openpyxl as xl
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from RPMainWindow import Ui_MainWindow

import DealTxt as dt

def find_tests_name(sheet, report):
    """找到测试项目名称和起始位置"""
    res = {}
    start = None
    
    for col in sheet.columns:
        for cell in col:
            if cell.value == report['flag_data_start_row']:
                r = cell.row
                for c in range(cell.column, sheet.max_column + 1):
                    if sheet.cell(row=r, column=c).value == report['flag_data_start_col']:
                        start = {'row': r, 'col': c}
                        break
                if start: break

    if start:
        for row in range(start['row'] + 1, sheet.max_row + 1):
            cellValue = sheet.cell(row=row, column=start['col']).value
            if cellValue is not None and sheet.cell(row=row, column=start['col'] + 1).value is None:
                res[cellValue] = {'row': row, 'col': start['col']}

    return res

class RPMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(RPMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.action_openfolder.triggered.connect(self.open_folder)
        self.action_openreport.triggered.connect(self.open_report)
        self.action_quitapp.triggered.connect(self.close)
        
        if not self.centralwidget.layout():
            self.centralwidget.setLayout(QVBoxLayout())
        
        self.load_config()

    def load_config(self):
        try:
            cfgPath = QFileDialog.getOpenFileName(self, '选择配置文件', filter='配置文件(*json)')[0]
            if not cfgPath:
                raise FileNotFoundError("未选择配置文件")
            inspect(cfgPath)
            with open(cfgPath, 'r', encoding='utf-8') as f:
                self.qerp = json.load(f)
                # inspect(self.qerp)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.Warning("错误", f"加载配置文件失败: {str(e)}")
            self.close()

    def open_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, u"选择根目录", os.path.expanduser("~"))

    def open_report(self):
        self.report_path = QFileDialog.getOpenFileName(self,"打开报告",self.folder_path)[0]
        # inspect(self.report_path)
        self.wb = xl.load_workbook(self.report_path)
        self.show_tests()
        
    def show_tests(self):
        report = self.qerp['Report']
        sheet = self.wb[report['sheet_name']]
        res = find_tests_name(sheet, report)
        
        inspect(res)
        
        self.testTitles_layout = QVBoxLayout()
        self.testTitles_layout.setContentsMargins(0, 0, 0, 0)
        self.testTitles_layout.setSpacing(0)
        self.testTitles_layout.setAlignment(Qt.AlignTop)
        self.testTitles_layout.setSizeConstraint(QLayout.SetFixedSize)

        for test_name, pos in res.items():
            test_title = QLabel(test_name)
            test_title.setAlignment(Qt.AlignCenter)
            test_title.setStyleSheet(u"QLabel { font-size: 18px; }")
            self.testTitles_layout.addWidget(test_title)

        self.centralwidget.layout().insertLayout(0, self.testTitles_layout)
        

    def show_about(self):
        self.about_win = QWidget()
        self.about_win.setWindowTitle("关于")
        icon = QIcon()
        icon.addFile(u":/emipdf/acbel -1.jpg", QSize(), QIcon.Normal, QIcon.Off)
        self.about_win.setWindowIcon(icon)
        self.about_win.resize(300, 200)
        self.about_win.setStyleSheet(u"QLabel { font-size: 15px; }")
        self.about_win.setLayout(QVBoxLayout())

        label_about = QLabel(text="EMI-Report\n\n版本：1.1.0\n\n作者：Lucas Li\n\n邮箱：Lucas_Li@acbel.com", alignment=Qt.AlignCenter)
        label_about.setWordWrap(True)
        self.about_win.layout().addWidget(label_about)

        self.about_win.show()
        self.windows.append(self.about_win)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message',
            "确定退出吗?", QMessageBox.Yes | 
            QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = RPMainWindow()
    MainWindow.show()
    sys.exit(app.exec_())