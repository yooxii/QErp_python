import DealTxt
import RP
import tkinter as tk
from tkinter import ttk

def disp_data():
    # 这里可以添加显示数据的逻辑
    pass

def main():
    # 创建主窗口
    root = tk.Tk()
    root.title('QErp')
    root.geometry('400x200')

    # 设置窗口背景颜色
    root.configure(bg='#f0f0f0')

    # 创建一个框架来组织内容
    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # 设置框架的背景颜色
    frame.configure(style='TFrame')

    # 创建标签
    label = ttk.Label(frame, text='欢迎使用 QErp', font=('Helvetica', 16))
    label.grid(row=0, column=0, pady=10)

    # 创建按钮
    button = ttk.Button(frame, text='显示数据', command=disp_data)
    button.grid(row=1, column=0, pady=5)

    button2 = ttk.Button(frame, text='退出', command=root.quit)
    button2.grid(row=2, column=0, pady=5)

    # 设置按钮样式
    style = ttk.Style()
    style.configure('TButton', font=('Helvetica', 12), padding=10)
    style.configure('TFrame', background='#f0f0f0')

    # 运行主循环
    root.mainloop()


def disp_data():
    root = tk.Tk()
    root.title('QErp')
    root.geometry('800x400')

    # 创建文本框
    text_box = tk.Text(root, width=50, height=20, wrap='word', yscrollcommand=True)
    text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 创建滚动条
    scrollbar = tk.Scrollbar(root, command=text_box.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 将文本框和滚动条关联
    text_box.config(yscrollcommand=scrollbar.set)

    # 处理数据
    res = DealTxt.deal_data1(DealTxt.open_file())
    res = DealTxt.show_data(res)
    
    # 显示数据
    for i in res:
        text_box.insert(tk.END, i + '\n')
    root.mainloop()

if __name__ == '__main__':
    main()