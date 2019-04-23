import tkinter as tk
import os
from tkinter import messagebox
from tkinter import filedialog
from docx import Document

# 第1步，实例化object，建立窗口window
window = tk.Tk()
 
# 第2步，给窗口的可视化起名字
window.title('替换文本')
 
# 第3步，设定窗口的大小(长 * 宽)
window.geometry('600x600')  # 这里的乘是小x

file_path_label = tk.Label(window,text = '')
file_path_label.pack()

filenames = ""
def chose_file_path():
    # 请求选择文件夹/目录
    my_filetypes = [('all files', '.*')]
    # answer = filedialog.askdirectory(parent=window,
    #                                 initialdir=os.getcwd(),
    #                                 title="Please select a folder:")
                                    # 请求选择一个或多个文件
    global filenames
    filenames = filedialog.askopenfilenames(parent=window,
                                        initialdir=os.getcwd(),
                                        title="Please select one or more files:",
                                        filetypes=my_filetypes)
    if len(filenames) != 0:
        string_filename =""
        for i in range(0,len(filenames)):
            string_filename += str(filenames[i])+"\n"
        file_path_label.config(text = string_filename)
                             

file_path_button = tk.Button(window, text="选择文件夹", command=chose_file_path)
file_path_button.pack()
 
e1 = tk.Entry(window, show=None, font=('Arial', 14))  # 显示成明文形式
e1.pack()

e2 = tk.Entry(window, show=None, font=('Arial', 14))  # 显示成明文形式
e2.pack()

def replace_string(filename):
    doc = Document(filename)
    # for p in doc.paragraphs:
    #     if 'old text' in p.text:
    #         inline = p.runs
    #         # Loop added to work with runs (strings with same style)
    #         for i in range(len(inline)):
    #             if 'old text' in inline[i].text:
    #                 text = inline[i].text.replace('old text', 'new text')
    #                 inline[i].text = text
    # doc.save('dest1.docx')
    return 1

def hit_me():   # 在文本框内容最后接着插入输入内容
    replace_string(filenames[0])

b = tk.Button(window, text="全部替换", command=hit_me)
b.pack()


# 第6步，主窗口循环显示
window.mainloop()
