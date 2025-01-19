import openpyxl as xl
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import json
import DealTxt as dt

# 加载所有配置文件
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

def initialize_window(root, WidthMain, HeightMain):
    """初始化窗口布局和变量"""
    for widget in root.winfo_children():
        if not isinstance(widget, tk.Menu):
            widget.destroy()

    global data_box, txt_seqs, excelrp, tests_name
    data_box = []
    txt_seqs = []
    excelrp = ExcelRp()
    tests_name = excelrp.read_excel()
    
    data_canvas = tk.Canvas(root, width=WidthMain-20, height=HeightMain)
    data_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=data_canvas.yview)
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

    WidthCfgExcel = 500
    HeightCfgExcel = 300
    
    # 计算窗口左上角的位置
    x = (WidthScreen / 2) - (WidthCfgExcel / 2)
    y = (HeightScreen / 2) - (WidthCfgExcel / 2)
    
    root.geometry(f"{WidthCfgExcel}x{HeightCfgExcel}+{int(x)}+{int(y)}")
    
    input_box = create_report_inputbox(root)
    
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=20)
    
    btn_save = tk.Button(btn_frame, text="应用", command=lambda: (save_reportcfg(input_box), root.destroy(), initialize_window(root_main)))
    btn_save.pack(side=tk.LEFT, padx=10)

    btn_exit = tk.Button(btn_frame, text="返回", command=root.destroy)
    btn_exit.pack(side=tk.LEFT, padx=10)

    root.mainloop()

def cfg_txt(txt_config):
    """配置TXT窗口"""
    root = tk.Tk()
    root.title("配置TXT")
    
    WidthCfgTxt = 500
    HeightCfgTxt = 500
    
    # 计算窗口左上角的位置
    x = (WidthScreen / 2) - (WidthCfgTxt / 2)
    y = (HeightScreen / 2) - (WidthCfgTxt / 2)
    
    root.geometry(f"{WidthCfgTxt}x{HeightCfgTxt}+{int(x)}+{int(y)}")
    
    txtcfg_canvas = tk.Canvas(root, width=WidthCfgTxt-20, height=HeightCfgTxt-20)
    txtcfg_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=txtcfg_canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    txtcfg_canvas.configure(yscrollcommand=scrollbar.set)
    txtcfg_canvas.bind("<Configure>", lambda e: txtcfg_canvas.configure(scrollregion=txtcfg_canvas.bbox("all")))

    data_frame = tk.Frame(txtcfg_canvas)
    txtcfg_canvas.create_window((0, 0), window=data_frame, anchor="nw")

    input_frame = tk.Frame(data_frame)
    input_frame.grid(row=0, column=0, padx=10, pady=10)

    input_box = create_txt_inputbox(input_frame)
    
    btn_frame = tk.Frame(data_frame)
    btn_frame.grid(row=1, column=0, padx=10, pady=10)

    btn_save = tk.Button(btn_frame, text="应用", command=lambda: (save_txtcfg(input_box), root.destroy()))
    btn_save.pack(side=tk.LEFT, padx=10)

    btn_exit = tk.Button(btn_frame, text="返回", command=root.destroy)
    btn_exit.pack(side=tk.LEFT, padx=10)

    root.mainloop()

def create_report_inputbox(root):
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

def create_txt_inputbox(root):
    """创建输入框"""
    input_box = []
    input_frame = tk.Frame(root)
    input_frame.pack(pady=2)
    
    i=0
    for key, value in txt.items():
        tk.Label(input_frame, text=f"{key}:").grid(row=i, column=0, padx=10, pady=10)
        # 如果value是字符串直接输出
        if isinstance(value, str):
            input_excel = tk.Entry(input_frame, width=40, justify='center')
            input_excel.insert(0, value)
        # 否则用json.dumps输出
        else:
            heigh = len(value)+2 if isinstance(value, list) else 10
            input_excel = tk.Text(input_frame, width=40, height=heigh, wrap='word')
            input_excel.insert(tk.INSERT, json.dumps(value, ensure_ascii=False, indent=4))
        input_excel.grid(row=i, column=1, padx=10, pady=10)
        input_box.append(input_excel)
        i+=1
    
    return input_box

def save_reportcfg(input_box):
    """保存报告配置"""
    for i, input_excel in enumerate(input_box):
        report[list(report.keys())[i]] = input_excel.get()
    qerp['Report'] = report
    savecfg()

def save_txtcfg(input_box):
    """保存数据文件配置"""
    for i, input_txt in enumerate(input_box):
        # 如果input_txt是ENTRY，则直接获取值
        if isinstance(txt[list(txt.keys())[i]], str):
            txt[list(txt.keys())[i]] = input_txt.get()
        # 否则用json.loads解析
        else:
            txt[list(txt.keys())[i]] = json.loads(input_txt.get(1.0, tk.END))
    qerp['TXT'] = txt
    savecfg()

def savecfg():
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

    global HeightScreen, WidthScreen

    WidthScreen = root.winfo_screenwidth()
    HeightScreen = root.winfo_screenheight()

    WidthMain = 600
    HeightMain = 600
    
    # 计算窗口左上角的位置
    x = (WidthScreen / 2) - (WidthMain / 2)
    y = (HeightScreen / 2) - (HeightMain / 2)
    
    root.geometry(f"{WidthMain}x{HeightMain}+{int(x)}+{int(y)}")
    
    menubar = tk.Menu(root)

    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="选择文件夹", command=lambda: load_rootpath())
    file_menu.add_command(label="打开报告", command=lambda: initialize_window(root, WidthMain, HeightMain))
    file_menu.add_command(label="报告另存为", command=lambda: excelrp.save_excel(data_box, txt_seqs))
    menubar.add_cascade(label="报告", menu=file_menu)
    
    data_menu = tk.Menu(menubar, tearoff=0)
    data_menu.add_command(label="打开数据文件", command=lambda: load_tests(data_box))
    data_menu.add_command(label="打开选择文件", command=lambda: load_select_path())
    data_menu.add_command(label="加载选择", command=lambda: load_selects(data_box, qerp['selectfile']))
    data_menu.add_command(label="保存选择", command=lambda: save_select(data_box))
    menubar.add_cascade(label="数据处理", menu=data_menu)

    setting_menu = tk.Menu(menubar, tearoff=0)
    setting_menu.add_command(label="配置Report", command=lambda: cfg_excel(root))
    setting_menu.add_command(label="配置TXT", command=lambda: cfg_txt(txt))
    menubar.add_cascade(label="设置", menu=setting_menu)

    menubar.add_command(label="退出", command=root.quit)
    root.config(menu=menubar)

    # initialize_window(root)
    root.mainloop()

def load_rootpath():
    """加载文件夹"""
    # 打开路径选择框
    rootpath = filedialog.askdirectory(
        title="选择文件夹",
        initialdir=qerp['initialdir']
    )
    qerp['initialdir'] = rootpath
    return rootpath

def load_select_path():
    """加载选择"""
    # 打开路径选择框
    selectFile_path = filedialog.askopenfilename(
        title="选择选择文件",
        initialdir=qerp['initialdir'],
        filetypes=[("保存的选择文件", "*.json")]
    )
    qerp['selectfile'] = selectFile_path
    savecfg() # 保存选择文件路径
    return selectFile_path

def load_selects(select_box, selectFile_path):
    if not selectFile_path:
        selectFile_path = load_select_path()
    try:
        with open(selectFile_path, 'r', encoding='utf-8') as f:
            save_selects = json.load(f)
        for i, select_key in enumerate(save_selects.keys()):
            select_box[i].set(save_selects[select_key])
    except FileNotFoundError:
        messagebox.showwarning("警告", "选择文件不存在，无法加载选择。")
    except json.JSONDecodeError as e:
        messagebox.showerror("错误", f"加载选择时出错: {str(e)}")

def save_select_path():
    """保存选择路径"""
    # 打开路径选择框
    selectFile_path = filedialog.asksaveasfilename(
        title="保存选择文件",
        initialdir=qerp['initialdir'],
        filetypes=[("保存的选择文件", "*.json")]
    )
    qerp['selectfile'] = selectFile_path
    savecfg() # 保存选择文件路径
    return selectFile_path

def save_select(data_box):
    """保存选择"""

    if data_box is None:
        messagebox.showwarning("警告", "请先加载数据！")
        return

    saveSelects = {list(tests_name.keys())[j]: data.get() for j, data in enumerate(data_box)}

    saveSelectPath = qerp['selectfile']
    if not saveSelectPath:
        saveSelectPath = save_select_path()

    if not saveSelectPath:
        return
    try:
        with open(saveSelectPath, 'w', encoding='utf-8') as f:
            json.dump(saveSelects, f, ensure_ascii=False, indent=4)
    except IOError as e:
        messagebox.showerror("错误", f"保存选择失败: {str(e)}")

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
                    for st_key, st_value in txt['select'].items():
                        if select_key.find(st_key) != -1:
                            Cvalue = value[st_value[0]][st_value[1]]
                            flag = True
                            break
                    if not flag:
                        # 未找到匹配的选择项，使用第一个值
                        Cvalue = value[list(value.keys())[0]][0]
                    # 保留三位小数
                    Cell = self.wb[report['sheet_name']].cell(row=row, column=col)
                    Cell.number_format = '0.000'
                    Cell.value = float(Cvalue)

        try:
            self.wb.save(f"{save_file.replace('.xlsx', '')}.xlsx")
            disp_save_select(data_box)
        except Exception as e:
            messagebox.showerror("错误", f"保存Excel文件失败: {str(e)}")

def disp_save_select(data_box):
    """保存选择窗口"""
    root = tk.Tk()
    root.title("保存选择")

    WidthScreen = root.winfo_screenwidth()
    HeightScreen = root.winfo_screenheight()

    WidthSaveselect = 250
    HeightSaveselect = 150
    
    # 计算窗口左上角的位置
    x = (WidthScreen / 2) - (WidthSaveselect / 2)
    y = (HeightScreen / 2) - (HeightSaveselect / 2)
    
    root.geometry(f"{WidthSaveselect}x{HeightSaveselect}+{int(x)}+{int(y)}")

    save_frame = tk.Frame(root)
    save_frame.pack(pady=20)
    
    tk.Label(save_frame, text="是否覆盖之前保存的选择文件？").pack(pady=20)

    tk.Button(save_frame, text="是", command=lambda: (save_select(data_box), root.destroy())).pack(side=tk.LEFT, padx=50)
    tk.Button(save_frame, text="否", command=root.destroy).pack(side=tk.LEFT, padx=50)
    root.mainloop()

if __name__ == '__main__':
    show_datas()
