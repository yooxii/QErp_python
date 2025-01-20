import re
import sys
import json
from typing import List
from tkinter import filedialog, messagebox
import tkinter as tk

def load_config():
    """加载配置文件"""
    root = tk.Tk()
    root.withdraw()

    try:
        cfgPath = filedialog.askopenfilename(title='选择配置文件', filetypes=[('JSON', '*.json')])
        if not cfgPath:
            raise FileNotFoundError("未选择配置文件")
        with open(cfgPath, 'r', encoding='utf-8') as f:
            qerp = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
        sys.exit()
    finally:
        root.destroy()
    
    return cfgPath, qerp

def open_file(qerp):
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

def deal_data1(datas: List[str], qerp):
    res = []
    seqs = {}

    txt = qerp['TXT']

    # 处理数据内容
    i = 0
    while i < len(datas):
        if datas[i].find(txt['seq_start']) != -1:  # 找到测试项目开始标志
            seqName = datas[i].replace(txt['seq_start'], '')  # 测试项目名称，去除多余字符
            reads = []
            for j in range(i + 1, len(datas)):  # 遍历测试项目内容
                if datas[j].find(txt['seq_end']) != -1:  # 找到测试项目结束标志
                    break
                
                is_read = any(flag in datas[j] for flag in txt['read'])

                if any(flag in datas[j] for flag in txt['noread']):
                    is_read = False

                if is_read:
                    try:
                        while datas[j].replace(' ', '') != '':
                            reads.append(re.split(r'\s{2,}', datas[j]))
                            j += 1
                        reads.append([txt['read_end']]) # 插入结束标识
                    except IndexError:
                        break
            i = j
            if len(reads) != 0:
                seqs[seqName.replace(txt['pass'], '')] = reads
        try:
            if datas[i].find(txt['info_start']) != -1:
                if len(seqs) != 0:
                    res.append(seqs)
                seqs = {}
        except IndexError:
            seqs[seqName.replace(txt['pass'], '')].append([txt['read_end']])
            res.append(seqs)
            return res
        i += 1
    if len(seqs) != 0:
        res.append(seqs)
    return res

def deal_data2(datas: List, qerp):
    # datas = [{seq:{read:[value]}}]
    txt = qerp['TXT']
    res = []
    for uut in datas:
        UUT = {}
        for seqName, reads in uut.items():
            SeqRead = {}
            ReadNo = []
            readTmp = []
            readFlag = txt['read']
            try:
                # 遍历读取行
                for SingleLine in reads:
                    if len(ReadNo) == 0:
                        for item in SingleLine:
                            if item in readFlag:
                                ReadNo.append(SingleLine.index(item))  # 找到符合条件的索引
                    ReadNo = list(set(ReadNo))  # 去重
                    # 追加符合条件的读取
                    if len(ReadNo) != 0:
                        for i in ReadNo:
                            if i < len(SingleLine):
                                readTmp.append(SingleLine[i])   # 这里因为是按行执行，
                    if SingleLine[0] == txt['read_end']:
                        for i in range(len(ReadNo)):    # 所以需要将结果拆分为多个列表，按len(ReadNo)个一行分割按列读取
                            SeqRead[readTmp[i]] = readTmp[i+len(ReadNo)::len(ReadNo)]
                        ReadNo.clear()  # 重置索引
                        readTmp.clear()  # 重置临时列表
            except IndexError as e:
                print(f"错误: 读取过程中发生索引错误 - {str(e)}")
            UUT[seqName] = SeqRead
        res.append(UUT)
    return res

def deal_data(qerp):
    datas = open_file(qerp)
    data1 = deal_data1(datas, qerp)
    data2 = deal_data2(data1, qerp)
    # data = show_data(data2)
    return data2

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
    path, qerp = load_config()
    data = deal_data(qerp)
    f = open('res.txt', 'w', encoding='utf-8')
    f.write(json.dumps(data, indent=4, ensure_ascii=False))
    f.close()
