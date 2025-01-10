import openpyxl as xl
import tkinter as tk
from tkinter import filedialog
import json

with open('QErp.json', 'r', encoding='utf-8') as f:
    qerp = json.load(f)
    RP = qerp['Report']


def open_excel():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Open Excel file",initialdir=qerp['initialdir'],filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def cfg_excel():
    """
    配置Excel窗口
    """
    # 打开一个配置窗口
    root = tk.Tk()
    root.title("配置Excel")
    root.geometry("600x400")

    input_box = []

    # 输入框框架
    input_frame = tk.Frame(root)
    input_frame.pack(pady=5)  # 添加垂直间距

    # 输入框
    for key, value in RP.items():
        tk.Label(input_frame, text=key+':').pack(anchor='w', pady=5)  # 标签左对齐，添加垂直间距
        input_excel = tk.Entry(input_frame, width=40, justify='center')
        input_excel.insert(0, value)
        input_excel.pack(pady=5)  # 添加垂直间距
        input_box.append(input_excel)

    # 按钮框架
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=20)  # 添加垂直间距

    # 保存配置
    btn_save = tk.Button(btn_frame, text="保存修改", command=lambda: save_cfg(input_box))
    btn_save.pack(side=tk.LEFT, padx=10)  # 添加水平间隔

    # 退出窗口
    btn_exit = tk.Button(btn_frame, text="下一步", command=lambda: { root.destroy(), show_datas() })
    btn_exit.pack(side=tk.LEFT, padx=10)  # 添加水平间隔

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
    tests = read_excel()
    # 打开一个数据窗口
    root = tk.Tk()
    root.title("数据窗口")
    root.geometry("600x400")

    # 滚动条
    scrollbar = tk.Scrollbar(root)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(root, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE, width=30)
    for key, value in tests.items():
        listbox.insert(tk.END, key)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)  # 添加一些内边距
    scrollbar.config(command=listbox.yview)

    # 按钮框架
    btn_frame = tk.Frame(root)
    btn_frame.pack(side=tk.BOTTOM, pady=10)  # 添加按钮框架并调整外边距

    # 加载
    btn_load = tk.Button(btn_frame, text="加载", command=lambda: load_tests(root))
    btn_load.pack(side=tk.LEFT, padx=5)  # 添加水平间隔

    # 退出
    btn_exit = tk.Button(btn_frame, text="退出", command=lambda: { root.quit() })
    btn_exit.pack(side=tk.LEFT, padx=5)  # 添加水平间隔

    # 启动主循环
    root.mainloop()


def load_tests(root,select_box):
    """
    加载数据
    :param root: 数据窗口
    :return:
    """

def read_excel():
    file_path = open_excel()
    wb = xl.load_workbook(file_path)
    # 打开指定名称的工作表
    sheet = wb[RP['sheet_name']]
    print(sheet.title)
    start = find_test_start(sheet)
    print(start)
    res = find_tests(sheet, start)
    print(res)
    return res

def find_test_start(sheet):
    for col in sheet.columns:
        for cell in col:
            if cell.value is None:
                break
            if cell.value == RP['flag_data_start_row']:
                r = cell.row
                for c in range(cell.column, sheet.max_column+1):
                    if sheet.cell(row=r, column=c).value == RP['flag_data_start_col']:
                        break
                res = {'row': r, 'col': c}
                break
    return res

def find_tests(sheet, start):
    res = {}
    for row in range(start['row']+1, sheet.max_row+1):
        cellValue = sheet.cell(row=row, column=start['col']).value
        if cellValue is not None and sheet.cell(row=row, column=start['col']+1).value is None:
            res[cellValue] = {'row':row,'col':start['col']}

    return res

def write_excel(file_path, data):
    wb = xl.load_workbook(file_path)
    sheet = wb.active[RP['sheet_name']]


    wb.save(file_path)


if __name__ == '__main__':
    # file_path = open_excel()
    cfg_excel()
    # write_excel(file_path, data)