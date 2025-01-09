import openpyxl as xl
import tkinter as tk
from tkinter import filedialog
import json

with open('QErp.json', 'r', encoding='utf-8') as f:
    qerp = json.load(f)
    EXCEL = qerp['EXCEL']


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
    root.geometry("400x400")

    input_box = []
    # 输入框
    for key, value in EXCEL.items():
        tk.Label(root, text=key).pack()
        input_excel = tk.Entry(root, width=40,justify='center')
        input_excel.insert(0, value)
        input_excel.pack()
        input_box.append(input_excel)
    # 保存配置
    btn_save = tk.Button(root, text="保存修改", command=lambda: save_cfg(input_box))
    btn_save.pack()
    # 退出窗口
    btn_exit = tk.Button(root, text="下一步", command=lambda: {root.destroy() , show_datas()})
    btn_exit.pack()
    # 运行窗口
    root.mainloop()

def save_cfg(input_box):
    """
    保存配置
    :param input_box: 输入框列表
    :return:
    """
    for i in range(len(input_box)):
        EXCEL[list(EXCEL.keys())[i]] = input_box[i].get()
    qerp['EXCEL'] = EXCEL
    with open('QErp.json', 'w', encoding='utf-8') as f:
        json.dump(qerp, f, ensure_ascii=False, indent=4)

def show_datas():
    """
    显示数据窗口
    """
    # 打开一个数据窗口
    root = tk.Tk()
    root.title("数据窗口")
    root.geometry("400x400")

    # 显示数据
    tk.Label(root, text="数据").pack()
    data_text = tk.Text(root, width=40, height=20)
    data_text.pack()
    data_text.insert(tk.INSERT, "数据显示区")

    # 退出窗口
    btn_exit = tk.Button(root, text="退出", command=lambda: root.destroy())
    btn_exit.pack()

    # 运行窗口
    root.mainloop()

def read_excel(file_path):
    wb = xl.load_workbook(file_path)
    # 打开指定名称的工作表
    sheet = wb[EXCEL['sheet_name']]
    print(sheet.title)
    data = []
    
    return data


def write_excel(file_path, data):
    wb = xl.load_workbook(file_path)
    sheet = wb.active[EXCEL['sheet_name']]


    wb.save(file_path)


if __name__ == '__main__':
    file_path = open_excel()
    cfg_excel()
    data = read_excel(file_path)
    print(data)
    # write_excel(file_path, data)