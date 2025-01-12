from tkinter import ttk, messagebox
import openpyxl as xl
import tkinter as tk
from tkinter import filedialog
import json
import DealTxt as dt

# 加载配置文件
def load_config():
    with open('QErp.json', 'r', encoding='utf-8') as f:
        return json.load(f)

qerp = load_config()
report = qerp['Report']

def initialize_window(root):
    """初始化窗口布局和变量"""
    for widget in root.winfo_children():
        if not isinstance(widget, tk.Menu):
            widget.destroy()

    global data_box, txt_seqs, excelrp, tests
    data_box = []
    txt_seqs = []
    excelrp = ExcelRp()
    tests = excelrp.read_excel()
    
    data_canvas = tk.Canvas(root, width=600, height=400)
    data_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(data_canvas, orient=tk.VERTICAL, command=data_canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    data_canvas.configure(yscrollcommand=scrollbar.set)
    data_canvas.bind("<Configure>", lambda e: data_canvas.configure(scrollregion=data_canvas.bbox("all")))
    
    data_frame = tk.Frame(data_canvas)
    data_canvas.create_window((0, 0), window=data_frame, anchor="nw")

    for key, value in tests.items():
        small_frame = tk.Frame(data_frame)
        small_frame.pack(pady=2, side=tk.TOP)
        tk.Label(small_frame, text=f"{key}:", justify='left', width=30).pack(anchor='w', pady=2, side=tk.LEFT)
        Comb_data = ttk.Combobox(small_frame, width=40, justify='center')
        Comb_data.insert(0, value)
        Comb_data.pack(pady=2, side=tk.LEFT)
        data_box.append(Comb_data)

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
    
    btn_save = tk.Button(btn_frame, text="保存", command=lambda: (save_cfg(input_box), root.destroy()))
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

def load_tests(select_box):
    """加载数据"""
    global txt_seqs
    try:
        txt_seqs = dt.deal_data2(dt.deal_data1(dt.open_file()))
        if not txt_seqs or not txt_seqs[0]:
            raise ValueError("未找到有效的数据值")

        values = ["NA"] + list(txt_seqs[0].keys())
        for select in select_box:
            select['values'] = values
            select.set(values[0])

    except Exception as e:
        messagebox.showerror("错误", f"加载选择时出错: {str(e)}")

def show_datas():
    """显示数据窗口"""
    root = tk.Tk()
    root.title("QErp")
    root.geometry("800x600")
    
    menubar = tk.Menu(root)
    
    data_menu = tk.Menu(menubar, tearoff=0)
    data_menu.add_command(label="选择数据", command=lambda: load_tests(data_box))
    data_menu.add_command(label="读取选择", command=lambda: load_select(data_box))
    menubar.add_cascade(label="数据", menu=data_menu)

    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="打开", command=lambda: initialize_window(root))
    file_menu.add_command(label="另存为", command=lambda: excelrp.save_excel(data_box, txt_seqs))
    menubar.add_cascade(label="文件", menu=file_menu)

    menubar.add_command(label="退出", command=root.quit)
    root.config(menu=menubar)

    initialize_window(root)
    root.mainloop()

def load_select(select_box):
    """加载选择"""
    try:
        with open('Select.json', 'r', encoding='utf-8') as f:
            save_selects = json.load(f)
        for i, select_key in enumerate(save_selects.keys()):
            select_box[i].set(save_selects[select_key])
    except FileNotFoundError:
        messagebox.showwarning("警告", "选择文件不存在，无法加载选择。")

def find_tests(sheet):
    """找到测试数据"""
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

class ExcelRp:
    """ExcelRp类，用于操作Excel文件"""

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
        res = find_tests(sheet)
        return res

    def save_excel(self, data_box, seqs):
        if not seqs:
            messagebox.showwarning("警告", "请先加载数据！")
            return

        save_file = filedialog.asksaveasfilename(
            title="另存为Excel文件",
            initialdir=qerp['initialdir'],
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        saveSelects = {list(tests.keys())[j]: data.get() for j, data in enumerate(data_box)}
        save_select(saveSelects)

        for i in range(len(seqs)):
            for select_key, select in saveSelects.items():
                if select == "NA":
                    continue
                if select_key in tests:
                    row = tests[select_key]['row']
                    col = tests[select_key]['col'] + i + 1
                    value = seqs[i].get(select, "")
                    self.wb[report['sheet_name']].cell(row=row, column=col).value = str(value)

        self.wb.save(f"{save_file.replace('.xlsx', '')}.xlsx")

def save_select(saveSelects):
    """保存选择"""
    root = tk.Tk()
    root.title("保存选择")
    root.geometry("300x100")
    
    tk.Label(root, text="是否保存选择？").pack(pady=20)

    tk.Button(root, text="是", command=lambda: (save_select1(saveSelects), root.destroy())).pack(side=tk.LEFT, padx=20)
    tk.Button(root, text="否", command=root.destroy).pack(side=tk.LEFT, padx=20)
    
    root.mainloop()

def save_select1(saveSelects):
    """辅助函数，保存选择"""
    with open('Select.json', 'w', encoding='utf-8') as f:
        json.dump(saveSelects, f, ensure_ascii=False, indent=4)

if __name__ == '__main__':
    cfg_excel()
    show_datas()
