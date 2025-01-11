from tkinter import ttk
import openpyxl as xl
import tkinter as tk
from tkinter import filedialog
import json

import DealTxt as dt


with open('QErp.json', 'r', encoding='utf-8') as f:
    qerp = json.load(f)
    report = qerp['Report']

# 一些全局变量
# global txt_seqs

def cfg_excel():
    """配置Excel窗口"""
    root = tk.Tk()
    root.title("配置Excel")
    root.geometry("600x600")

    input_box = []

    input_frame = tk.Frame(root)
    input_frame.pack(pady=2)

    for key, value in report.items():
        tk.Label(input_frame, text=f"{key}:").pack(anchor='w', pady=2)
        input_excel = tk.Entry(input_frame, width=40, justify='center')
        input_excel.insert(0, value)
        input_excel.pack(pady=2)
        input_box.append(input_excel)

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=20)

    btn_save = tk.Button(btn_frame, text="保存", command=lambda: [save_cfg(input_box), root.destroy()])
    btn_save.pack(side=tk.LEFT, padx=10)

    btn_exit = tk.Button(btn_frame, text="不保存", command=root.destroy)
    btn_exit.pack(side=tk.LEFT, padx=10)

    root.mainloop()


def save_cfg(input_box):
    """保存配置"""
    for i, input_excel in enumerate(input_box):
        report[list(report.keys())[i]] = input_excel.get()
    qerp['Report'] = report
    with open('QErp.json', 'w', encoding='utf-8') as f:
        json.dump(qerp, f, ensure_ascii=False, indent=4)


def show_datas():
    """显示数据窗口"""
    excel = Excel()
    tests = excel.read_excel()
    # 打开一个数据窗口
    root = tk.Tk()
    root.title("数据窗口")
    root.geometry("800x600")

    data_box = []

    data_canvas = tk.Canvas(root, width=600, height=400)
    data_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 滚动条
    scrollbar = tk.Scrollbar(data_canvas, orient=tk.VERTICAL, command=data_canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    data_canvas.configure(yscrollcommand=scrollbar.set)
    data_canvas.bind(
        "<Configure>",
        lambda e: data_canvas.configure(scrollregion=data_canvas.bbox("all"))
    )
    # 创建一个 Frame 作为内容容器
    data_frame = tk.Frame(data_canvas)
    data_canvas.create_window((0, 0), window=data_frame, anchor="nw")

    # 在 Frame 中添加内容
    for key, value in tests.items():
        small_frame = tk.Frame(data_frame)
        small_frame.pack(pady=2, side=tk.TOP)
        tk.Label(small_frame, text=key + ":", justify='left', width=30).pack(anchor='w', pady=2, side=tk.LEFT)  # 标签左对齐，添加垂直间距
        Comb_data = ttk.Combobox(small_frame, width=40, justify='center')
        Comb_data.insert(0, value)
        Comb_data.pack(pady=2, side=tk.LEFT)
        data_box.append(Comb_data)

    # 按钮框架
    btn_frame = tk.Frame(root)
    btn_frame.pack(side=tk.BOTTOM, pady=10)  # 添加按钮框架并调整外边距

    # 加载
    btn_load = tk.Button(btn_frame, text="加载", command=lambda: load_tests(data_box))
    btn_load.pack(side=tk.LEFT, padx=2)

    # 保存
    btn_save = tk.Button(btn_frame, text="另存为", command=lambda: { excel.save_excel(tests, data_box, txt_seqs), root.quit() })
    btn_save.pack(side=tk.LEFT, padx=2)

    # 退出
    btn_exit = tk.Button(btn_frame, text="退出", command=lambda: { root.quit() })
    btn_exit.pack(side=tk.LEFT, padx=2)

    # 启动主循环
    root.mainloop()


def load_tests(select_box):
    """
    加载数据
    :param root: 数据窗口
    :return:
    """
    global txt_seqs
    txt_seqs = dt.deal_data1(dt.open_file())
    txt_seqs = dt.deal_data2(txt_seqs)
    values = txt_seqs[0]
    value = ["NA"] + list(values.keys())
    for i in range(len(select_box)):
        select_box[i]['values'] = value
        select_box[i].set(value[0])

def find_tests(sheet):
    """
    找到测试数据
    :param sheet: Excel sheet
    :return: 测试数据字典
    """
    res = {}
    for col in sheet.columns:
        for cell in col:
            if cell.value is None:
                break
            if cell.value == report['flag_data_start_row']:
                r = cell.row
                for c in range(cell.column, sheet.max_column+1):
                    if sheet.cell(row=r, column=c).value == report['flag_data_start_col']:
                        break
                start = {'row': r, 'col': c}
                break

    for row in range(start['row']+1, sheet.max_row+1):
        cellValue = sheet.cell(row=row, column=start['col']).value
        if cellValue is not None and sheet.cell(row=row, column=start['col']+1).value is None:
            res[cellValue] = {'row':row,'col':start['col']}

    return res

class Excel:
    """Excel操作类"""

    def __init__(self):
        self.file_path = None
        self.open_file()
        self.wb = xl.load_workbook(self.file_path)

    def open_file(self):
        root = tk.Tk()
        root.withdraw()
        self.file_path = filedialog.askopenfilename(
            title="Open Excel file",
            initialdir=qerp['initialdir'],
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

    def read_excel(self):
        if not self.file_path:
            self.open_file()
        sheet = self.wb[report['sheet_name']]
        print(sheet.title)      # 打印sheet名称
        res = find_tests(sheet)
        print(res)              # 打印测试数据
        return res

    def save_excel(self, tests, data_box, seqs):
        if not seqs:
            # 弹出警告窗口，提示用户先加载数据
            root = tk.Tk()
            root.title("警告")
            root.geometry("300x100")
            tk.Label(root, text="请先加载数据！").pack(pady=20)
            btn_ok = tk.Button(root, text="确定", command=root.destroy)
            btn_ok.pack(pady=20)
            root.mainloop()
            return
        save_file = filedialog.asksaveasfilename(
            title="另存为Excel文件",
            initialdir=qerp['initialdir'],
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        save_selects = {}
        # for j, data in enumerate(data_box):
        #     save_selects[list(tests.keys())[j]] = data.get()
        # save_select(save_selects)
        with open('Select.json', 'r', encoding='utf-8') as f:
            save_selects = json.load(f)

        for i in range(len(seqs)):
            for select_key, select in save_selects.items():
                if select == "NA":
                    continue
                row = tests[select_key]['row']
                col = tests[select_key]['col']+i+1
                value = seqs[i][select]
                self.wb[report['sheet_name']].cell(row=row, column=col).value = str(value)

        # dt.save_file(save_file, save_seqs)
        save_file = save_file.replace('.xlsx', '')
        self.wb.save(save_file+'.xlsx')

def save_select(save_selects):
    """
    保存选择
    :param save_selects: 选择字典
    :return:
    """
    with open('Select.json', 'w', encoding='utf-8') as f:
        json.dump(save_selects, f, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    # file_path = open_excel()
    cfg_excel()
    show_datas()
    # write_excel(file_path, data)