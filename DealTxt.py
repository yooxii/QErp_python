
import json
import openpyxl as px
from tkinter import filedialog

with open('QErp.json', 'r', encoding='utf-8') as f:
    qerp = json.load(f)
    FLAGS = qerp['FLAGS']


def open_file():
    # 打开路径选择框
    file_path = filedialog.askopenfilename(initialdir=qerp['initialdir'], title='选择文件', filetypes=[('文本文件', '*.txt')])
    # 打开文件
    file = open(file_path, 'r', encoding='utf-8')
    # 读取文件内容
    content = file.read()
    res = content.split('\n')
    file.close()
    return res

def deal_txt(datas: list[str]):
    res = []
    # 处理文件内容
    for i in range(len(datas)):
        seq = {}
        if datas[i].find(FLAGS['seq_start']) != -1:
            tmp = datas[i].replace(FLAGS['seq_start'], '')
            seqs = []
            for j in range(i+1, len(datas)):
                if datas[j].find(FLAGS['seq_end']) != -1:
                    break
                seqs.append(datas[j])
            i = j
            seq[tmp.replace(FLAGS['pass'], '')] = seqs
            res.append(seq)
    return res
            

if __name__ == '__main__':
    data = open_file()
    print(deal_txt(data))