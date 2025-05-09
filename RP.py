import os
import select
import sys
import json
from rich import inspect
import openpyxl as xl
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from RPMainWindow import Ui_MainWindow

import DealTxt as dt
import DealXlsx as dx

def find_tests_name(sheet, rp_flag):
    """找到测试项目名称和起始位置"""
    res = {}
    start = None
    table_start_col = 1
    
    for col in sheet.columns:
        for cell in col:
            if cell.value == rp_flag['flag_data_start_row']:
                r = cell.row
                table_start_col = cell.column
                for c in range(cell.column, sheet.max_column + 1):
                    if sheet.cell(row=r, column=c).value == rp_flag['flag_data_start_col']:
                        start = {'row': r, 'col': c}
                        break
                if start: break
    
    if start:
        for row in range(start['row'] + 1, sheet.max_row + 1):
            if str(rp_flag['flag_data_end_row']) in str(sheet.cell(row=row,column=table_start_col).value):
                break
            cellValue = sheet.cell(row=row, column=start['col']).value
            if cellValue is not None and sheet.cell(row=row, column=start['col'] + 1).value is None:
                if cellValue in res:
                    continue
                res[cellValue] = {'row': row, 'col': start['col']}

    return res

class RPMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(RPMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.action_openreport.triggered.connect(self.open_report)
        self.action_opendatatxt.triggered.connect(self.open_data_txt)
        self.action_opendataexcel.triggered.connect(self.open_data_excel)
        self.action_loadselect.triggered.connect(self.load_selects)
        self.action_savereport.triggered.connect(self.save_report)
        self.action_saveselects.triggered.connect(self.save_selects)
        self.action_sets_import.triggered.connect(self.import_sets)
        self.action_sets_export.triggered.connect(self.export_sets)
        self.action_quitapp.triggered.connect(self.close)
        
        if not self.centralwidget.layout():
            self.centralwidget.setLayout(QVBoxLayout())
        
        self.title1 = u"QErp-曙光报告辅助工具" # 主标题
        self.title2 = "" # 副标题
        self.rootpath = "" # 根路径
        self.select_box = {} # 保存了选择框的字典
        self.Model = None # 机种型号
        
        if self.load_config():
            self.show_start_window()
        
        self.testTitles_layout = None
        self.auto_load_select = True
        self.TXTDATALOADED = False
        self.EXCELDATALOADED = False
        

    def show_start_window(self):
        """显示开始界面，提示用户选择机种型号"""
        self.start_window = QDialog()
        self.start_window.setWindowTitle(self.title1)
        self.start_window.resize(400, 300)
        self.start_window.setWindowModality(Qt.ApplicationModal)
        self.start_window.setWindowIcon(QIcon(".\\logo\\acbel-1.jpg"))

        try:
            layout = QVBoxLayout()
            self.start_window.setLayout(layout)

            label = QLabel(text="请选择机种型号：")
            label.setMaximumHeight(20)
            label.setFont(QFont("Microsoft YaHei", 10))
            layout.addWidget(label)

            combo = QComboBox(self.start_window)
            combo.setMinimumHeight(30)
            combo.setFont(QFont("Microsoft YaHei", 10))
            combo.addItems(self.qerp["Select"].keys())
            layout.addWidget(combo)

            btn = QPushButton("确定", self.start_window)
            btn.clicked.connect(self.start_window.accept)
            layout.addWidget(btn)

            if self.start_window.exec_():
                self.Model = combo.currentText()
                self.setWindowTitle(self.title1 + " : " + self.Model)
                
        except Exception as e:
            QMessageBox.warning(self, "错误", f"显示开始界面时发生错误：{str(e)}")


    def load_config(self):
        """加载配置文件"""
        try:
            self.cfgPath = QFileDialog.getOpenFileName(self, '选择配置文件', filter='配置文件(*json)')[0]
            if not self.cfgPath:
                raise FileNotFoundError("未选择配置文件")
            # inspect(self.cfgPath)
            with open(self.cfgPath, 'r', encoding='utf-8') as f:
                self.qerp = json.load(f)
                # inspect(self.qerp)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.warning(self, "错误", "配置文件加载失败：\n" + str(e))
            return False
        return True

    def load_selects(self):
        """加载txt选择项"""
        selects = self.qerp['Select'][self.Model] # 从配置文件中获取选择项
        # 选取txt['select']的每一个元素的第一个，如果是字符串，则设置select_box的对应位置的值
        warnings = []
        # if not self.TXTDATALOADED:
        #     QMessageBox.warning(self, "警告", "请先打开txt数据文件")
        # if not self.EXCELDATALOADED:
        #     QMessageBox.warning(self, "警告", "请先打开excel数据文件")
        
        for i, select_key in enumerate(selects):
            if select_key not in self.select_box:
                warnings.append(select_key)
                continue
            if isinstance(selects[select_key], str):
                self.select_box[select_key][1].setCurrentText(selects[select_key])
            else:
                if selects[select_key][0] == "TXT" and self.TXTDATALOADED:
                    for i in range(len(selects[select_key])):
                        self.select_box[select_key][i].setCurrentText(str(selects[select_key][i]))
                    
                elif selects[select_key][0] in self.qerp["EXCEL"]["is_datafile_flag"] and self.EXCELDATALOADED:
                    for i in range(len(selects[select_key])):
                        self.select_box[select_key][i].setCurrentText(str(selects[select_key][i]))
        if warnings:
            QMessageBox.warning(self, "警告", "配置文件中不存在以下选择项：\n" + "\n".join(warnings))

    def open_report(self):
        """打开报告文件"""
        if not self.Model:
            self.show_start_window()
        self.report_path = QFileDialog.getOpenFileName(self,"打开报告",os.path.expanduser("~"), filter='Excel(*.xlsx *.xls)')[0]
        # inspect(self.report_path)
        self.rootpath = os.path.dirname(self.report_path)
        self.wb = xl.load_workbook(self.report_path)
        self.show_tests()

    def save_report(self):
        """保存报告文件"""
        if not self.txt_datas:
            QMessageBox.warning(self, "错误", "没有数据，无法保存")
            return
        save_path = QFileDialog.getSaveFileName(self,"保存报告",self.rootpath, filter='Excel(*.xlsx *.xls)')[0]
        if not save_path:
            return
        
        tests_cell = self.test_cell
        tests_box = self.select_box

        for i in range(3):
            for _, select_key in enumerate(tests_cell.keys()):
                data_src = tests_box[select_key][0].currentText()
                select = tests_box[select_key][1].currentText()
                if select == "NA":
                    continue
                if select_key in tests_cell:
                    row = tests_cell[select_key]['row']
                    col = tests_cell[select_key]['col'] + i + 1
                    
                    if data_src == "TXT" and self.TXTDATALOADED:
                        seqs_txt = self.txt_datas
                        value = seqs_txt[i].get(select, "")
                        if not value:
                            continue
                        data_type = tests_box[select_key][2].currentText()
                        data_col = int(tests_box[select_key][3].currentText())
                        if data_type == "NA" or data_col == "NA":
                            continue
                        elif data_type == "Reading/+" or data_type == "Min":
                            Cvalue = float(value["Reading/+"][data_col])+float(value["Min"][data_col])
                        else:
                            Cvalue = value[data_type][data_col]
                    
                    elif data_src in self.qerp["EXCEL"]["is_datafile_flag"] and self.EXCELDATALOADED:
                        excel_datas = list(self.excel_datas[tests_box[select_key][0].currentText()].values())
                        value = excel_datas[i].get(select, "")
                        if not value:
                            continue
                        data_type = tests_box[select_key][2].currentText()
                        if data_type == "NA":
                            continue
                        Cvalue = value[data_type]
                    
                    # 保留三位小数
                    Cell = self.st.cell(row=row, column=col)
                    Cell.number_format = '0.000'
                    tmp = float(Cvalue)
                    Cell.value = tmp if tmp > 0 else -tmp # 取绝对值

        self.wb.save(save_path)
        QMessageBox.information(self, "成功", "保存成功")

    def save_selects(self):
        """保存选择项到配置文件"""
        selects = {}
        for key,test in self.select_box.items():
            select = test[0].currentText()
            if select == "NA":
                continue
            selects[key] = [test[0].currentText(), test[1].currentText(), test[2].currentText(), test[3].currentText()]
        self.qerp['Select'][self.Model] = selects
        with open(self.cfgPath, 'w', encoding='utf-8') as f:
            json.dump(self.qerp, f, ensure_ascii=False, indent=4)

    def reset_window(self, layout: QLayout = None):
        """重置窗口布局"""
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
        """显示测试项目"""
        self.report = self.qerp['Report']
        self.st = self.wb[self.report['sheet_name']]
        res = find_tests_name(self.st, self.report)
        self.test_cell = res
        self.txt_seltype = self.qerp["TXT"]["read"].copy()
        if "Min" in self.txt_seltype:
            self.txt_seltype.remove("Min")
        
        # inspect(res)
        
        # 清空窗口
        self.reset_window(self.centralwidget.layout())
        self.testTitles_layout = None
            
        self.testTitles_layout = QVBoxLayout()
        self.testTitles_layout.setContentsMargins(0, 0, 0, 0)
        self.testTitles_layout.setSpacing(0)
        self.testTitles_layout.setAlignment(Qt.AlignTop)
        self.testTitles_layout.setSizeConstraint(QLayout.SetFixedSize)
        
        test_layout = QGridLayout()
        test_layout.setContentsMargins(0, 0, 0, 0)
        test_layout.setSpacing(0)
        test_layout.setAlignment(Qt.AlignTop)
        test_layout.setSizeConstraint(QLayout.SetFixedSize)
        
        title_col1 = QLabel("报告项目")
        title_col1.setAlignment(Qt.AlignCenter)
        title_col1.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col1, 0, 0, 1, 1)
        
        title_col5 = QLabel("数据来源")
        title_col5.setAlignment(Qt.AlignCenter)
        title_col5.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col5, 0, 1, 1, 1)
        
        title_col2 = QLabel("测试数据项目")
        title_col2.setAlignment(Qt.AlignCenter)
        title_col2.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col2, 0, 2, 1, 1)
        
        title_col3 = QLabel("数据读取类别")
        title_col3.setAlignment(Qt.AlignCenter)
        title_col3.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col3, 0, 3, 1, 1)
        
        title_col4 = QLabel("数据在第几行")
        title_col4.setAlignment(Qt.AlignCenter)
        title_col4.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_layout.addWidget(title_col4, 0, 4, 1, 1)
        
        test_layout.setColumnStretch(0, 1)
        test_layout.setColumnStretch(1, 1)
        test_layout.setColumnStretch(2, 1)
        test_layout.setColumnStretch(3, 1)
        test_layout.setColumnStretch(4, 1)
        
        row = 1
        for test_name, pos in res.items():
            # 测试数据选择框
            data_boxes = self.create_test_box(test_name, test_layout, row)
            self.select_box[test_name] = data_boxes
            row += 1
        
        self.testTitles_layout.addLayout(test_layout)

        self.centralwidget.layout().insertLayout(0, self.testTitles_layout)

    def create_test_box(self, test_name: str, test_layout: QGridLayout, row: int):
        """创建测试项目的选择框"""
        test_title = QLabel(test_name+u"：")
        test_title.setAlignment(Qt.AlignLeft)
        test_title.setStyleSheet(u"QLabel { font-size: 16px; }")
        test_title.setMinimumHeight(30)
        test_title.setMinimumWidth(250)
        test_title.setMaximumWidth(350)
        test_layout.addWidget(test_title, row, 0, 1, 1)
            
        test_box = []
        data_src = QComboBox()
        data_src.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_src.setMinimumHeight(25)
        data_src.setMinimumWidth(50)
        data_src.setMaximumWidth(100)
        test_layout.addWidget(data_src, row, 1, 1, 1)
        test_box.append(data_src)
        
        data_name = QComboBox()
        data_name.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_name.setMinimumHeight(25)
        data_name.setMinimumWidth(300)
        data_name.setMaximumWidth(450)
        test_layout.addWidget(data_name, row, 2, 1, 1)
        test_box.append(data_name)
        
        data_type = QComboBox()
        data_type.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_type.setMinimumHeight(25)
        data_type.setMinimumWidth(80)
        data_type.setMaximumWidth(150)
        test_layout.addWidget(data_type, row, 3, 1, 1)
        test_box.append(data_type)
        
        data_col = QComboBox()
        data_col.setStyleSheet(u"QComboBox { font-size: 12px; }")
        data_col.setMinimumHeight(25)
        data_col.setMinimumWidth(50)
        data_col.setMaximumWidth(150)
        test_layout.addWidget(data_col, row, 4, 1, 1)
        test_box.append(data_col)
        
        return test_box

    def init_select_box(self):
        if not self.TXTDATALOADED and not self.EXCELDATALOADED:
            return
        for test_name, test in self.select_box.items():
            test_select_src = test[0]
            test_select_src.clear()
            test_select_src.addItem("NA")
            test_select_src.addItem("TXT")
            test_select_src.addItems(self.qerp["EXCEL"]["is_datafile_flag"])
            test_select_src.currentIndexChanged.connect(self.on_data_src_changed)
            
            test_select_data = test[1]
            test_select_data.clear()
            test_select_data.addItem("NA")
            
            test_select_type = test[2]
            test_select_type.clear()
            test_select_type.addItem("NA")
            
            test_select_col = test[3]
            test_select_col.clear()
            test_select_col.addItem("NA")

    def open_data_txt(self):
        """打开并处理数据txt文件"""
        DT = dt.DealTxt()
        self.txt_datas = DT.deal_data(qerp=self.qerp)
        self.txt_selects = list(self.txt_datas[0])
        self.TXTDATALOADED = True
        
        self.init_select_box()

        if self.auto_load_select:
            self.load_selects()

    def open_data_excel(self):
        DX = dx.DealXlsx(self.qerp, self.rootpath)
        self.excel_datas = DX.deal_data_folder()
        self.excel_types = list(self.excel_datas)

        excel_key = {}
        excel_seq = {}
        for key, value in self.excel_datas.items():
            excel_key[key] = list(list(value.values())[0])
            tmp  = {}
            for seq, data in list(value.values())[0].items():
                tmp[seq] = list(data)
            excel_seq[key] = tmp

        self.excel_selects_key = excel_key # 保存多种excel数据文件的选择项，两层字典
        self.excel_selects = excel_seq # 保存多种excel数据文件的选择项，三层字典
        
        self.EXCELDATALOADED = True
        
        if self.auto_load_select:
            self.load_selects()

    def on_data_src_changed(self, index):
        """数据来源选择框改变事件"""
        src_select = self.sender()
        src_text = self.sender().currentText()
        if src_text == "TXT" and self.TXTDATALOADED:
            data_type = self.txt_seltype
            data_col_tmp = self.qerp["TXT"]["data_Max_cols"]
            data_col = [str(i) for i in list(range(data_col_tmp))]
            for test_name, test in self.select_box.items():
                if test[0] is src_select:
                    test[1].clear()
                    test[1].addItem("NA")
                    test[1].addItems(self.txt_selects)
                    test[2].clear()
                    test[2].addItem("NA")
                    test[2].addItems(data_type)
                    test[3].clear()
                    test[3].addItem("NA")
                    test[3].addItems(data_col)
            
        elif src_text in self.qerp["EXCEL"]["is_datafile_flag"]:
            for test_name, test in self.select_box.items():
                if test[0] is src_select:
                    test[1].clear()
                    test[1].addItem("NA")
                    test[1].addItems(self.excel_selects_key[src_text])
                    test[1].currentIndexChanged.connect(self.on_data_select_changed)
        else:
            return
    
    def on_data_select_changed(self):
        """数据选择框改变事件"""
        seq_select = self.sender()
        seq_text = self.sender().currentText()
        if seq_text == "NA":
            return
        for test_name, test in self.select_box.items():
            if test[1] is seq_select:
                if not test[0].currentText() in self.excel_selects_key.keys():
                    continue
                data_type = list(self.excel_selects[test[0].currentText()][seq_text])
                test[2].clear()
                test[2].addItem("NA")
                test[2].addItems(data_type)
                test[3].clear()
                test[3].addItem("NA")

    def import_sets(self):
        """导入配置文件"""
        if self.load_config():
            self.show_start_window()
        self.reset_window(self.centralwidget.layout())
        self.testTitles_layout = None

    def export_sets(self):
        """导出配置文件"""
        save_path = QFileDialog.getSaveFileName(self,"保存配置文件",self.cfgPath, filter='配置文件(*json)')[0]
        if not save_path:
            return
        with open(save_path+".json", 'w', encoding='utf-8') as f:
            json.dump(self.qerp, f, ensure_ascii=False, indent=4)

    def show_about(self):
        """显示关于界面"""
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
        self.about_win.activateWindow()

    def closeEvent(self, event):
        """关闭事件处理"""
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