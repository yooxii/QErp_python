import re
import os
import sys
import json
import rich
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
    def __init__(self, _qerp:dict,rootpath=None):
        self.qerp = _qerp
        self.excel_path = rootpath
        
    def open_data_folder(self):
        """
        在当前路径找不到数据文件的情况下，提示打开数据文件夹
        """
        self.excel_path = QFileDialog.getExistingDirectory(caption="选择数据文件夹", options=QFileDialog.ShowDirsOnly, dir=self.excel_path)
        if not self.excel_path:
            raise FileNotFoundError("请选择数据文件夹")
        return self.excel_path
        
    def filter_data_files(self, path):
        """
        
        """
        # 找到excel_path下所有的数据文件
        res = {}
        data_files = os.listdir(path)
        flags = self.qerp["EXCEL"]["is_datafile_flag"]
        for flag in flags:
            res[flag] = []
        for file in data_files:
            if not file.startswith("~$") and (re.match(r"^.*\.xlsx$", file) or re.match(r"^.*\.xls$", file)): # 排除临时文件和非excel文件
                for flag in flags:
                    if flag in file: # 匹配数据文件标志
                        res[flag].append(file)
                        break
        
        # 删除空数据文件
        del_keys = []
        for key in res.keys():
            if res[key] == []:
                del_keys.append(key)
        for key in del_keys:
            del res[key]
        
        # 排序数据文件
        for key in res.keys():
            res[key] = sorted(res[key], key=lambda x: x.split("_")[-1])
        return res
        
    def deal_data_file(self, file_name):
        wb = xl.load_workbook(os.path.join(self.excel_path, file_name))
        st = wb.active
        """找到测试项目名称和起始位置"""
        res = {}
        excel_cfg = self.qerp["EXCEL"]
        seq_flag = excel_cfg["seq_flag"]
        
        for row in range(1, st.max_row+1):
            cell_cr = f"{excel_cfg['seq_col']}{row}"
            cell = st[cell_cr].value
            if not isinstance(cell, str):
                continue
            if seq_flag in cell:
                seq = cell.replace("*", "").strip()
                seq_data = {}
                for row2 in range(row+1, st.max_row+1):
                    cell_cr2 = f"{excel_cfg['seq_col']}{row2}"
                    cell2 = st[cell_cr2].value
                    if not isinstance(cell, str):
                        continue
                    if seq_flag in cell2:
                        row = row2-1
                        break
                    if cell2 in excel_cfg["end_flag"]:
                        break
                    seq_data[cell2] = st[f"{excel_cfg['data_col']}{row2}"].value
                res[seq] = seq_data

        return res
    
    def deal_data_folder(self):
        path = self.open_data_folder()
        data_files = self.filter_data_files(path)
        res = {}
        for key, data_files in data_files.items():
            tmp = {}
            for file in data_files:
                tmp[file] = self.deal_data_file(file)
            res[key] = tmp
        rich.inspect(res)
        return res
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    _, qerp = load_config()
    dx = DealXlsx(qerp)
    # print(dx.filter_data_files())
    res = dx.deal_data_folder()
    with open("res_excel.json", "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=4)
    # sys.exit(app.exec_())