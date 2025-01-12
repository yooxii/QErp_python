import openpyxl as xl
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import json
import DealTxt as dt

# 加载配置文件
def load_config():
    try:
        with open('QErp.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
        return None

qerp = load_config()
if qerp is None:
    exit(1)

report = qerp['Report']
txt = qerp['TXT']

def initialize_window(root):
    """初始化窗口布局和变量"""
    for widget in root.winfo_children():
        if not isinstance(widget, tk.Menu):
            widget.destroy()

    global data_box, txt_seqs, excelrp, tests_name
    data_box = []
    txt_seqs = []
    excelrp = ExcelRp()
    tests_name = excelrp.read_excel()
    
    data_canvas = tk.Canvas(root, width=600, height=400)
    data_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(data_canvas, orient=tk.VERTICAL, command=data_canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    data_canvas.configure(yscrollcommand=scrollbar.set)
    data_canvas.bind("<Configure>", lambda e: data_canvas.configure(scrollregion=data_canvas.bbox("all")))
    
    data_frame = tk.Frame(data_canvas)
    data_canvas.create_window((0, 0), window=data_frame, anchor="nw")

    for key, value in tests_name.items():
        small_frame = tk.Frame(data_frame)
        small_frame.pack(pady=2, side=tk.TOP)
        tk.Label(small_frame, text=f"{key}:", justify='left', width=30).pack(anchor='w', pady=2, side=tk.LEFT)
        Comb_data = ttk.Combobox(small_frame, width=40, justify='center')
        Comb_data.insert(0, value)
        Comb_data.pack(pady=2, side=tk.LEFT)
        data_box.append(Comb_data)

def cfg_excel(root_main):
    """配置Excel窗口"""
    root = tk.Tk()
    root.title("配置Excel")
    root.geometry("600x600")
    
    input_box = create_input_box(root, report)
    
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=20)
    
    btn_save = tk.Button(btn_frame, text="应用", command=lambda: (save_cfg(input_box), root.destroy(), initialize_window(root_main)))
    btn_save.pack(side=tk.LEFT, padx=10)

    btn_exit = tk.Button(btn_frame, text="返回", command=root.destroy)
    btn_exit.pack(side=tk.LEFT, padx=10)

    root.mainloop()

def create_input_box(root, report):
    """创建输入框"""
    input_box = []
    input_frame = tk.Frame(root)
    input_frame.pack(pady=2)
    
    for key, value in report.items():
        tk.Label(input_frame, text=f"{key}:").pack(anchor='w', pady=2)
        input_excel = tk.Entry(input_frame, width=40, justify='center')
        input_excel.insert(0, value)
        input_excel.pack(pady=2)
        input_box.append(input_excel)
    
    return input_box

def save_cfg(input_box):
    """保存配置"""
    for i, input_excel in enumerate(input_box):
        report[list(report.keys())[i]] = input_excel.get()
    qerp['Report'] = report
    try:
        with open('QErp.json', 'w', encoding='utf-8') as f:
            json.dump(qerp, f, ensure_ascii=False, indent=4)
    except IOError as e:
        messagebox.showerror("错误", f"保存配置文件失败: {str(e)}")

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

    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="打开新报告", command=lambda: initialize_window(root))
    file_menu.add_command(label="另存为", command=lambda: excelrp.save_excel(data_box, txt_seqs))
    menubar.add_cascade(label="文件", menu=file_menu)
    
    data_menu = tk.Menu(menubar, tearoff=0)
    data_menu.add_command(label="选择数据", command=lambda: load_tests(data_box))
    data_menu.add_command(label="读取选择", command=lambda: load_select(data_box))
    data_menu.add_command(label="保存选择", command=lambda: save_select(data_box))
    menubar.add_cascade(label="数据", menu=data_menu)

    setting_menu = tk.Menu(menubar, tearoff=0)
    setting_menu.add_command(label="配置Excel", command=lambda: cfg_excel(root))
    menubar.add_cascade(label="设置", menu=setting_menu)

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
    except json.JSONDecodeError as e:
        messagebox.showerror("错误", f"加载选择时出错: {str(e)}")

def find_tests_name(sheet):
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

class ExcelRp:
    """ExcelRp类，用于操作Excel文件"""

    def __init__(self):
        self.file_path = None
        self.wb = None
        self.open_file()

    def open_file(self):
        root = tk.Tk()
        root.withdraw()
        self.file_path = filedialog.askopenfilename(
            title="Open Excel file",
            initialdir=qerp['initialdir'],
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_path:
            try:
                self.wb = xl.load_workbook(self.file_path)
            except Exception as e:
                messagebox.showerror("错误", f"打开Excel文件失败: {str(e)}")
                self.wb = None

    def read_excel(self):
        if not self.file_path or not self.wb:
            self.open_file()
        if not self.wb:
            return {}
        sheet = self.wb[report['sheet_name']]
        res = find_tests_name(sheet)
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
        if not save_file:
            return
        
        saveSelects = {list(tests_name.keys())[j]: data.get() for j, data in enumerate(data_box)}

        for i in range(len(seqs)):
            for select_key, select in saveSelects.items():
                if select == "NA":
                    continue
                if select_key in tests_name:
                    row = tests_name[select_key]['row']
                    col = tests_name[select_key]['col'] + i + 1
                    value = seqs[i].get(select, "")
                    flag = False
                    for st in txt['select']:
                        if select_key.find(st) != -1:
                            Cvalue = value[txt['select'][st]]
                            flag = True
                            break
                    if not flag:
                        Cvalue = value[1]
                    # 保留三位小数
                    self.wb[report['sheet_name']].cell(row=row, column=col).number_format = '0.000'
                    self.wb[report['sheet_name']].cell(row=row, column=col).value = float(Cvalue)

        try:
            self.wb.save(f"{save_file.replace('.xlsx', '')}.xlsx")
            disp_save_select(data_box)
        except Exception as e:
            messagebox.showerror("错误", f"保存Excel文件失败: {str(e)}")

def save_select(data_box):
    """辅助函数，保存选择"""
    saveSelects = {list(tests_name.keys())[j]: data.get() for j, data in enumerate(data_box)}
    try:
        with open('Select.json', 'w', encoding='utf-8') as f:
            json.dump(saveSelects, f, ensure_ascii=False, indent=4)
    except IOError as e:
        messagebox.showerror("错误", f"保存选择失败: {str(e)}")

def disp_save_select(data_box):
    """保存选择窗口"""
    root = tk.Tk()
    root.title("保存选择")
    root.geometry("300x100")
    
    tk.Label(root, text="是否保存选择？").pack(pady=20)

    tk.Button(root, text="是", command=lambda: (save_select(data_box), root.destroy())).pack(side=tk.LEFT, padx=20)
    tk.Button(root, text="否", command=root.destroy).pack(side=tk.LEFT, padx=20)
    
    root.mainloop()

if __name__ == '__main__':
    show_datas()
