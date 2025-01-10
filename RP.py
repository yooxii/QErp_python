from tkinter import ttk
import openpyxl as xl
import tkinter as tk
from tkinter import filedialog
import json

import DealTxt as dt


with open('QErp.json', 'r', encoding='utf-8') as f:
    qerp = json.load(f)
    RP = qerp['Report']

file_path = ''


def cfg_excel():
    """
    配置Excel窗口
    """
    # 打开一个配置窗口
    root = tk.Tk()
    root.title("配置Excel")
    root.geometry("600x600")

    input_box = []

    # 输入框框架
    input_frame = tk.Frame(root)
    input_frame.pack(pady=2)

    # 输入框
    for key, value in RP.items():
        tk.Label(input_frame, text=key+':').pack(anchor='w', pady=2)  # 标签左对齐，添加垂直间距
        input_excel = tk.Entry(input_frame, width=40, justify='center')
        input_excel.insert(0, value)
        input_excel.pack(pady=2)
        input_box.append(input_excel)

    # 按钮框架
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=20)

    # 保存配置
    btn_save = tk.Button(btn_frame, text="保存", command=lambda: { save_cfg(input_box), root.destroy()})
    btn_save.pack(side=tk.LEFT, padx=10)

    # 退出窗口
    btn_exit = tk.Button(btn_frame, text="不保存", command=lambda: { root.destroy() })
    btn_exit.pack(side=tk.LEFT, padx=10)

    # 运行窗口
    root.mainloop()


def save_cfg(input_box):
    """
    保存配置
    :param input_box: 输入框列表
    :return:
    """
    for i in range(len(input_box)):
        RP[list(RP.keys())[i]] = input_box[i].get()
    qerp['Report'] = RP
    with open('QErp.json', 'w', encoding='utf-8') as f:
        json.dump(qerp, f, ensure_ascii=False, indent=4)

def show_datas():
    """
    显示数据窗口
    """
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
    btn_save = tk.Button(btn_frame, text="保存", command=lambda: { excel.save_excel(data_box), root.quit() })
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
    values = dt.deal_data(dt.open_file())[0]
    value = ["NA"] + list(values.keys())
    for i in range(len(select_box)):
        select_box[i]['values'] = value
        select_box[i].set(value[0])

def find_tests(sheet):
    res = {}
    for col in sheet.columns:
        for cell in col:
            if cell.value is None:
                break
            if cell.value == RP['flag_data_start_row']:
                r = cell.row
                for c in range(cell.column, sheet.max_column+1):
                    if sheet.cell(row=r, column=c).value == RP['flag_data_start_col']:
                        break
                start = {'row': r, 'col': c}
                break

    for row in range(start['row']+1, sheet.max_row+1):
        cellValue = sheet.cell(row=row, column=start['col']).value
        if cellValue is not None and sheet.cell(row=row, column=start['col']+1).value is None:
            res[cellValue] = {'row':row,'col':start['col']}

    return res

class excel:
    """
    Excel操作类
    """

    def __init__(self):
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title="Open Excel file",initialdir=qerp['initialdir'],filetypes=[("Excel files", "*.xlsx *.xls")])
        return file_path

    def read_excel():
        wb = xl.load_workbook(file_path)
        # 打开指定名称的工作表
        sheet = wb[RP['sheet_name']]
        print(sheet.title)
        res = find_tests(sheet)
        print(res)
        return res

    def save_excel(data):
        save_file = filedialog.asksaveasfilename(title="保存Excel文件", initialdir=qerp['initialdir'], filetypes=[("Excel files", "*.xlsx *.xls")])
        if save_file:
            wb = xl.load_workbook(file_path)
            sheet = wb[RP['sheet_name']]
            for i in range(len(data)):
                sheet.cell(row=data[i]['row'], column=data[i]['col']).value = data[i]['value']
            wb.save(save_file)


    def write_excel(data):
        wb = xl.load_workbook(file_path)
        sheet = wb.active[RP['sheet_name']]


        wb.save(file_path)


if __name__ == '__main__':
    # file_path = open_excel()
    cfg_excel()
    show_datas()
    # write_excel(file_path, data)