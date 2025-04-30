import re
import os
import sys
import json
import openpyxl as xl
from PySide2.QtWidgets import QApplication, QFileDialog 

def load_config():
    """加载配置文件"""
    try:
        cfgPath = QFileDialog.getOpenFileName(None, "选择配置文件", "", "JSON Files (*.json)")[0]
        if not cfgPath:
            raise FileNotFoundError("未选择配置文件")
        with open(cfgPath, 'r', encoding='utf-8') as f:
            qerp = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"错误: 加载配置文件失败 - {str(e)}")
        sys.exit()
    finally:
        print(f"配置文件加载成功: {cfgPath}")
    
    return cfgPath, qerp


class DealXlsx:
    def __init__(self, _qerp:dict):
        self.qerp = _qerp
        self.excel_path = self.qerp["initialdir"]
        
    def open_data_folder(self):
        """
        在当前路径找不到数据文件的情况下，提示打开数据文件夹
        """
        self.excel_path = QFileDialog.getOpenFileUrl(dir=self.excel_path)[0].path()
        if not self.excel_path:
            raise FileNotFoundError("请选择数据文件夹")
        return self.excel_path
        
    def filter_data_files(self):
        """
        
        """
        # 找到excel_path下所有的数据文件
        res = []
        data_files = os.listdir(self.excel_path)
        flags = self.qerp["EXCEL"]["is_datafile_flag"]
        for file in data_files:
            if not file.startswith("~$") and (re.match(r"^.*\.xlsx$", file) or re.match(r"^.*\.xls$", file)): # 排除临时文件和非excel文件
                for flag in flags:
                    if flag in file: # 匹配数据文件标志
                        res.append(file)
                        break
        return res
        
    def deal_data_file(self, file_name):
        wb = xl.load_workbook(os.path.join(self.excel_path, file_name))
        st = wb.active
        """找到测试项目名称和起始位置"""
        res = {}
        excel_cfg = self.qerp["EXCEL"]
        seq_flag = excel_cfg["seq_flag"]
        
        for row in range(1, st.max_row+1):
            cell = st[f"{excel_cfg['seq_col']}{row}"].value
            if not isinstance(cell, str):
                continue
            if seq_flag in cell:
                print(cell)
        # TODO

        return res
    
    def deal_data_folder(self):
        data_files = self.filter_data_files()
        for file in data_files:
            self.deal_data_file(file)
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    _, qerp = load_config()
    dx = DealXlsx(qerp)
    # print(dx.filter_data_files())
    dx.deal_data_folder()
    