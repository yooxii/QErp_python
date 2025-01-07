import re
import json
from typing import List
from tkinter import filedialog

with open('QErp.json', 'r', encoding='utf-8') as f:
    qerp = json.load(f)
    FLAGS = qerp['FLAGS']


def open_file():
    """
        打开文件并读取内容
        :return: 文件内容
    """
    # 打开路径选择框
    file_path = filedialog.askopenfilename(initialdir=qerp['initialdir'], title='选择文件', filetypes=[('文本文件', '*.txt')])
    # 打开文件
    file = open(file_path, 'r', encoding='utf-8')
    # 读取文件内容
    content = file.read()
    res = content.split('\n')
    file.close()
    return res

def deal_data(datas: List[str]):
    res = []
    seqs = {}
    # 处理数据内容
    i = 0
    while i < len(datas):
        if datas[i].find(FLAGS['seq_start']) != -1:
            tmp = datas[i].replace(FLAGS['seq_start'], '')
            reads = []
            for j in range(i+1, len(datas)):
                if datas[j].find(FLAGS['seq_end']) != -1:
                    break
                is_read = 0
                for flag in FLAGS['read']:
                    if datas[j].find(flag) != -1:
                        is_read = 1
                        break
                for flag in FLAGS['noread']:
                    if datas[j].find(flag) != -1:
                        is_read = 0
                        break
                if is_read == 1:
                    try:
                        while datas[j].replace(' ', '') != '':
                            reads.append(re.split(r'\s{2,}', datas[j]))
                            j += 1
                        break
                    except IndexError:
                        break
            i = j
            if len(reads) != 0:
                seqs[tmp.replace(FLAGS['pass'], '')] = reads
        try:
            if datas[i].find(FLAGS['info_start']) != -1:
                if len(seqs) != 0:
                    res.append(seqs)
                seqs = {}
        except IndexError:
            res.append(seqs)
            break
        i += 1
    return res

def show_data(data):
    res = []
    for seq in data:
        for key in seq:
            res.append(key)
            for read in seq[key]:
                tmp = ''
                for item in read:
                    tmp += item +'\t'
                res.append(tmp)
            res.append('\n')
    return res

if __name__ == '__main__':
    f = open('res.txt', 'w', encoding='utf-8')
    f.write(str(deal_data(open_file())))
    f.close()
    # print(deal_txt(data))
