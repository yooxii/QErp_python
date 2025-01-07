import DealTxt
import tkinter as tk

def main():
    root = tk.Tk()
    root.title('QErp')
    root.geometry('300x100')
    label = tk.Label(root, text='请选择文件')
    label.pack()
    button = tk.Button(root, text='显示数据', command=disp_data)
    button.pack()
    button2 = tk.Button(root, text='退出', command=root.quit)
    button2.pack()
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
    res = DealTxt.deal_data(DealTxt.open_file())
    res = DealTxt.show_data(res)
    
    # 显示数据
    for i in res:
        text_box.insert(tk.END, i + '\n')
    root.mainloop()

if __name__ == '__main__':
    main()