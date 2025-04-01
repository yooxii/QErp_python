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

def find_tests_name(sheet, rp_flag):
    """找到测试项目名称和起始位置"""
    res = {}
    start = None
    
    for col in sheet.columns:
        for cell in col:
            if cell.value == rp_flag['flag_data_start_row']:
                r = cell.row
                for c in range(cell.column, sheet.max_column + 1):
                    if sheet.cell(row=r, column=c).value == rp_flag['flag_data_start_col']:
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
        self.action_openreport.triggered.connect(self.open_report)
        self.action_opendatafile.triggered.connect(self.open_data_file)
        self.action_loadselect.triggered.connect(self.load_selects)
        self.action_savereport.triggered.connect(self.save_report)
        self.action_saveselects.triggered.connect(self.save_selects)
        self.action_sets_import.triggered.connect(self.import_sets)
        self.action_sets_export.triggered.connect(self.export_sets)
        self.action_quitapp.triggered.connect(self.close)
        
        if not self.centralwidget.layout():
            self.centralwidget.setLayout(QVBoxLayout())
        
        self.load_config()
        
        self.testTitles_layout = None
        self.auto_load_select = True

    def load_config(self):
        try:
            self.cfgPath = QFileDialog.getOpenFileName(self, '选择配置文件', filter='配置文件(*json)')[0]
            if not self.cfgPath:
                raise FileNotFoundError("未选择配置文件")
            inspect(self.cfgPath)
            with open(self.cfgPath, 'r', encoding='utf-8') as f:
                self.qerp = json.load(f)
                # inspect(self.qerp)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.warning(self, "错误", "配置文件加载失败：\n" + str(e))

    def load_selects(self):
        selects = self.qerp['Select']
        # 选取txt['select']的每一个元素的第一个，如果是字符串，则设置select_box的对应位置的值
        warnings = []
        for i, select_key in enumerate(selects):
            if select_key not in self.select_box:
                warnings.append(select_key)
                continue
            if isinstance(selects[select_key], str):
                self.select_box[select_key][0].setCurrentText(selects[select_key])
            else:
                self.select_box[select_key][0].setCurrentText(selects[select_key][0])
                self.select_box[select_key][1].setCurrentText(selects[select_key][1])
                self.select_box[select_key][2].setCurrentText(str(selects[select_key][2]))
        if warnings:
            QMessageBox.warning(self, "警告", "配置文件中不存在以下选择项：\n" + "\n".join(warnings))

    def open_report(self):
        self.report_path = QFileDialog.getOpenFileName(self,"打开报告",os.path.expanduser("~"), filter='Excel(*.xlsx *.xls)')[0]
        # inspect(self.report_path)
        self.rootpath = os.path.dirname(self.report_path)
        self.wb = xl.load_workbook(self.report_path)
        self.show_tests()
        
    def save_report(self):
        if not self.test_datas:
            QMessageBox.warning(self, "错误", "没有数据，无法保存")
            return
        save_path = QFileDialog.getSaveFileName(self,"保存报告",self.rootpath, filter='Excel(*.xlsx *.xls)')[0]
        if not save_path:
            return
        
        seqs = self.test_datas
        tests_cell = self.test_cell
        tests_box = self.select_box

        for i in range(len(seqs)):
            for _, select_key in enumerate(tests_cell.keys()):
                select = tests_box[select_key][0].currentText()
                if select == "NA":
                    continue
                if select_key in tests_cell:
                    row = tests_cell[select_key]['row']
                    col = tests_cell[select_key]['col'] + i + 1
                    value = seqs[i].get(select, "")
                    if not value:
                        continue
                    data_type = tests_box[select_key][1].currentText()
                    data_col = int(tests_box[select_key][2].currentText())
                    if data_type == "NA" or data_col == "NA":
                        continue
                    Cvalue = value[data_type][data_col]
                    # 保留三位小数
                    Cell = self.st.cell(row=row, column=col)
                    Cell.number_format = '0.000'
                    Cell.value = float(Cvalue)

        self.wb.save(save_path)
        QMessageBox.information(self, "成功", "保存成功")
        
    def save_selects(self):
        selects = {}
        for key,test in self.select_box.items():
            select = test[0].currentText()
            if select == "NA":
                continue
            selects[key] = [test[0].currentText(), test[1].currentText(), test[2].currentText()]
        self.qerp['Select'] = selects
        with open(self.cfgPath, 'w', encoding='utf-8') as f:
            json.dump(self.qerp, f, ensure_ascii=False, indent=4)
            
    def reset_window(self, layout: QLayout = None):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                if item.layout() is not None:
                    self.reset_window(item.layout())
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    layout.removeItem(item)

    def show_tests(self):
        self.report = self.qerp['Report']
        self.st = self.wb[self.report['sheet_name']]
        res = find_tests_name(self.st, self.report)
        self.test_cell = res
        
        inspect(res)
        
        # 清空窗口
        self.reset_window(self.centralwidget.layout())
        self.testTitles_layout = None
            
        self.testTitles_layout = QVBoxLayout()
        self.testTitles_layout.setContentsMargins(0, 0, 0, 0)
        self.testTitles_layout.setSpacing(0)
        self.testTitles_layout.setAlignment(Qt.AlignTop)
        self.testTitles_layout.setSizeConstraint(QLayout.SetFixedSize)

        self.select_box = {}
        
        test_layout = QGridLayout()
        test_layout.setContentsMargins(0, 0, 0, 0)
        test_layout.setSpacing(0)
        test_layout.setAlignment(Qt.AlignTop)
        test_layout.setSizeConstraint(QLayout.SetFixedSize)
        
        title_col1 = QLabel("报告项目")
        title_col1.setAlignment(Qt.AlignCenter)
        title_col1.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col1, 0, 0, 1, 1)
        
        title_col2 = QLabel("测试数据项目")
        title_col2.setAlignment(Qt.AlignCenter)
        title_col2.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col2, 0, 1, 1, 1)
        
        title_col3 = QLabel("数据读取类别")
        title_col3.setAlignment(Qt.AlignCenter)
        title_col3.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col3, 0, 2, 1, 1)
        
        title_col4 = QLabel("数据在第几行")
        title_col4.setAlignment(Qt.AlignCenter)
        title_col4.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col4, 0, 3, 1, 1)
        
        test_layout.setColumnStretch(0, 1)
        test_layout.setColumnStretch(1, 1)
        test_layout.setColumnStretch(2, 1)
        test_layout.setColumnStretch(3, 1)
        
        row = 1
        for test_name, pos in res.items():
            # 测试数据选择框
            data_boxes = self.create_test_box(test_name, test_layout, row)
            self.select_box[test_name] = data_boxes
            row += 1
        
        self.testTitles_layout.addLayout(test_layout)

        self.centralwidget.layout().insertLayout(0, self.testTitles_layout)

    def create_test_box(self, test_name: str, test_layout: QGridLayout, row: int):
        test_title = QLabel(test_name+u"：")
        test_title.setAlignment(Qt.AlignLeft)
        test_title.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_title.setMinimumHeight(30)
        test_title.setMinimumWidth(250)
        test_title.setMaximumWidth(350)
        test_layout.addWidget(test_title, row, 0, 1, 1)
            
        test_box = []
        data_name = QComboBox()
        data_name.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_name.setMinimumHeight(25)
        data_name.setMinimumWidth(300)
        data_name.setMaximumWidth(450)
        test_layout.addWidget(data_name, row, 1, 1, 1)
        test_box.append(data_name)
        
        data_type = QComboBox()
        data_type.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_type.setMinimumHeight(25)
        data_type.setMinimumWidth(80)
        data_type.setMaximumWidth(150)
        test_layout.addWidget(data_type, row, 2, 1, 1)
        test_box.append(data_type)
        
        data_col = QComboBox()
        data_col.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_col.setMinimumHeight(25)
        data_col.setMinimumWidth(50)
        data_col.setMaximumWidth(150)
        test_layout.addWidget(data_col, row, 3, 1, 1)
        test_box.append(data_col)
        
        return test_box
        
    def open_data_file(self):
        DT = dt.DealTxt()
        self.test_datas = DT.deal_data(qerp=self.qerp)
        test_data = self.test_datas[0]
        inspect(test_data)
        
        for test_name, test in self.select_box.items():
            test_select_data = test[0]
            test_select_data.clear()
            test_select_data.addItem("NA")
            test_select_data.addItems(list(test_data.keys()))
            
            test_select_type = test[1]
            test_select_type.clear()
            test_select_type.addItem("NA")
            test_select_type.addItems(self.qerp["TXT"]["read"])
            
            test_select_col = test[2]
            test_select_col.clear()
            test_select_col.addItem("NA")
            test_select_col.addItems("0123456789")
            
        if self.auto_load_select:
            self.load_selects()

    def import_sets(self):
        self.load_config()
        self.reset_window(self.centralwidget.layout())
        self.testTitles_layout = None

    def export_sets(self):
        save_path = QFileDialog.getSaveFileName(self,"保存配置文件",self.cfgPath, filter='配置文件(*json)')[0]
        if not save_path:
            return
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(self.qerp, f, ensure_ascii=False, indent=4)

    def show_about(self):
        self.about_win = QWidget()
        self.about_win.setWindowTitle("关于")
        icon = QIcon()
        icon.addFile(u":/emipdf/acbel -1.jpg", QSize(), QIcon.Normal, QIcon.Off)
        self.about_win.setWindowIcon(icon)
        self.about_win.resize(300, 200)
        self.about_win.setStyleSheet(u"QLabel { font-size: 15px; }")
        self.about_win.setLayout(QVBoxLayout())

        label_about = QLabel(text="QErp\n\n版本：1.1.0\n\n作者：Lucas Li\n\n邮箱：Lucas_Li@acbel.com", alignment=Qt.AlignCenter)
        label_about.setWordWrap(True)
        self.about_win.layout().addWidget(label_about)

        self.about_win.show()
        self.windows.append(self.about_win)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '退出',
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