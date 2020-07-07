# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 11:26:31 2020

@author: daviszhang
"""

import tkinter as tk
import os,re
import json,time
import requests
import threading
import webbrowser
#from zipfile import ZipFile
from bs4 import BeautifulSoup

window = tk.Tk()
window.title('DAX & M 批量格式化工具')
windowWidth = 410
windowHeight = 260
positionRight = int(window.winfo_screenwidth()/2 - windowWidth/2)
positionDown = int(window.winfo_screenheight()/2 - windowHeight/2)
window.geometry('410x260+{}+{}'.format(positionRight, positionDown)) 
window.configure(background='white')
var = tk.StringVar(None,"D")   
l = tk.Label(window, bg='#FFFFFF', width=20, font=('Arial', 12), text='格式化选项：')
l.pack()

def _selection():
    return var.get()
    
r1 = tk.Radiobutton(window, text='仅DAX', font=('Arial', 9), variable=var, value='D', bg='#FFFFFF', command=_selection)
r1.pack()
r2 = tk.Radiobutton(window, text='仅M', font=('Arial', 9), variable=var, value='M', bg='#FFFFFF', command=_selection)
r2.pack()
r3 = tk.Radiobutton(window, text='全部', font=('Arial', 9), variable=var, value='A', bg='#FFFFFF', command=_selection)
r3.pack()

def _formatter():
    directory = directory_entry.get()
    text = test()
    mode = _selection()
    #window.text.insert(text)
    result_label.configure(text=text)
    try:        
        with open(os.path.join(directory,'DataModelSchema'), encoding='utf-16-le') as source:
            content = source.read()     
    except:
        result_label.configure(text="不存在该文件或路径不正确")
        
    with open(os.path.join(directory,'DataModelSchema.json'),mode='w',encoding='utf-8') as sourceUTF8:
        sourceUTF8.write(content)
    with open(os.path.join(directory,'DataModelSchema.json')) as f:
        data = json.load(f)

    baseurl = "https://www.daxformatter.com/?embed=1&fx="

    endurl = "&r=US"

    def cleanhtml(raw_html):
        cleanr = re.compile('<.*?>')
        cleantext = re.sub(cleanr, '', raw_html)
        return cleantext

    def daxformatter(name,exp):
        item = name + "=" + exp
        url = baseurl + item + endurl
        r = requests.post(url)
        soup = BeautifulSoup(r.text, 'html.parser')
        html_result = soup.body.div.div
        if html_result.div is None:   #prevent error code in DAX
            result = str(html_result).replace("<br/>","\n")
            result = result.replace("\xa0"," ")
            result = cleanhtml(result)
            result = result.split("=",1)[1]
        else:
            result = exp
        return result
    
    def mformatter(exp):
        baseurl_m = 'https://m-formatter.azurewebsites.net/api/format/v1'
        requestBody = {
            'code': exp
        }

        headers = {
            'Content-Type': 'application/json'
        }
        r = requests.post(baseurl_m,json=requestBody,headers=headers)
        soup = BeautifulSoup(r.text, 'html.parser')
        html_result = soup.div
        if html_result is None:
            result = exp
        else:
            result = str(html_result).replace("<br/>","\n")
            result = result.replace("\xa0"," ")
            result = cleanhtml(result)
        return result
    
    if  mode == "D" or mode == "A":
        for n in range(0,len(data['model']['tables'])):
            try:
                data['model']['tables'][n]['isHidden']
            except:
                try:
                    for i in range(0,len(data['model']['tables'][n]['measures'])):
                        item_m_name = data['model']['tables'][n]['measures'][i]['name']
                        item_m_exp = data['model']['tables'][n]['measures'][i]['expression']
                        data['model']['tables'][n]['measures'][i]['expression'] = daxformatter(item_m_name,item_m_exp)
                except:
                    pass
        
                for i in range(1,len(data['model']['tables'][n]['columns'])):
                    try:
                        item_c_name = data['model']['tables'][n]['columns'][-i]['name']
                        item_c_exp = data['model']['tables'][n]['columns'][-i]['expression']
                        data['model']['tables'][n]['columns'][-i]['expression'] = daxformatter(item_c_name,item_c_exp)
                    except:
                        break
    if mode == "M" or mode =="A":
        for m in range(0,len(data['model']['tables'])):
            exp_item = data['model']['tables'][m]['partitions'][0]['source']['expression']
            try:
                if data['model']['tables'][m]['partitions'][0]['source']['type'] == 'm':
                    data['model']['tables'][m]['partitions'][0]['source']['expression'] = mformatter(exp=exp_item)
            except:
                pass
     
    try:
        with open(os.path.join(directory,'DataModelSchema'), mode='w',encoding='utf-16-le') as output:
            json.dump(data,output,indent=4)
        result_label.configure(text="已完成！")
    except:
        result_label.configure(text="写入失败，请检查文件格式并重试")


def callback(url):
    webbrowser.open_new(url)

thread = threading.Thread(target = _formatter)
while thread.is_alive():
    window.update()
    time.sleep(0.001)

header_label = tk.Label(window, text='在此输入DataModelSchema文件路径:', bg='#FFFFFF')
header_label.pack()

directory_frame = tk.Frame(window,bg='#F8F8FF')
directory_frame.pack(side=tk.TOP)
directory_label = tk.Label(directory_frame, bg='#FFFFFF')
directory_label.pack(side=tk.LEFT)
directory_entry = tk.Entry(directory_frame)
directory_entry.pack(side=tk.LEFT)


result_label = tk.Label(window,
                        bg='#FFFFFF')
result_label.pack()

calculate_btn = tk.Button(window, text='执行!', command=_formatter)
calculate_btn.pack()

version_label = tk.Label(window, text='版本号：测试版',font=('Arial', 8),
                        bg='#FFFFFF')
version_label.pack()

link1 = tk.Label(window, text="支持文档", fg="blue", cursor="hand2")
link1.pack()
link1.bind("<Button-1>", lambda e: callback("https://d-bi.gitee.io/"))

window.mainloop()

'''【2020.7.3】The project has been suspended. Modifying the datamodelschema will not affect the format of the M code, which may be determined by the datamashup.
Datamashup is composed of several fat format files, which can be decompressed and extracted (using 7z).
But after decompression, it can't be compressed into the original format again, at least the implementation of this compression technology is not clear.
Therefore, even if section 1.m is modified, it cannot be applied back to PBIX or PBIT for the time being'''
