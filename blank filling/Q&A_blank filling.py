from docx import Document
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import random
import os

# 全局變量
doc_entry = None
ico_entry = None
doc_button = None
ico_button = None
doc_var = None
ico_var = None
change_button = None
doc = None
question = []
A = []
B = []
C = []
D = []
ans = []
win = None
pq = None
pans = None
entry = None
button1 = None
button2 = None
num = None


def get_doc():
    """取得文件"""
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title="Select question set",
                                          filetypes=(("word document",
                                                      "*.docx*"),
                                                     ("all files",
                                                      "*.*")))
    doc_var.set(filename)


def get_ico():
    """取得ICON"""
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title="Select icon",
                                          filetypes=(("icon file",
                                                      "*.ico*"),
                                                     ("all files",
                                                      "*.*")))
    ico_var.set(filename)


def change():
    """選擇問題和ICON跳到答題介面"""
    global doc, win, pans, pq, que, entry, button1, button2, doc_var, ico_var
    doc = Document(str(doc_var.get()))
    doc_entry.destroy()
    ico_entry.destroy()
    doc_button.destroy()
    ico_button.destroy()
    change_button.destroy()
    pq = tk.StringVar()  # 問題
    pans = tk.StringVar()  # 答案
    width = win.winfo_screenwidth()
    height = win.winfo_screenheight()
    win.geometry(f"{width}x{height}")
    get_question()
    win.iconbitmap(str(ico_var.get()))
    pans = tk.StringVar()  # 答案
    pq = tk.StringVar()  # 問題
    update()
    que = tk.Label(win, height=10, width=1045, wraplength=1000, font=15, textvariable=pq, bg="#FFFFFF")
    que.pack()
    entry = tk.Entry(win, textvariable=pans, width=100, font=13)
    entry.pack()
    button1 = tk.Button(win, height=3, width=10, text="Submit", command=update, font=("Times New Roman", 13))
    button1.pack()
    button2 = tk.Button(win, height=3, width=10, text="Show Ans", command=ShowAns, font=("Times New Roman", 13))
    button2.pack()


def get_question():
    """取得問題"""
    for t in doc.paragraphs:
        q = ""
        a = ""
        flag = False  # 判斷是否答案
        for run in t.runs:
            if not run.underline:
                q += run.text
                flag = True
            else:
                q += "____"
                a += run.text + " "
        if flag:
            a = a[:-1]
            question.append(q)
            ans.append(a)
    i = 0
    while i < len(ans):
        if ans[i] == "":
            question.pop(i)
            ans.pop(i)
        i += 1


def update():
    global num, pans, pq
    """更新問題"""
    if num is None:  # 判斷是否第一次
        num = random.randrange(len(question))
        pq.set(question[num])
        pans.set("")
    else:
        if pans.get() == ans[num]:  # 判斷答案是否正確
            num = random.randrange(len(question))
            pq.set(question[num])
            pans.set("")
        else:
            messagebox.showwarning(title="Q&A", message="你錯了")


def ShowAns():
    """顯示答案"""
    global pans, num
    pans.set(ans[num])


if __name__ == "__main__":
    win = tk.Tk()
    win.title("Q&A")
    win.configure(bg="#FFFFFF")
    win.geometry("580x180")
    """選擇問題和ICON"""
    doc_var = tk.StringVar()  # 文件路徑
    ico_var = tk.StringVar()  # ICON路徑
    doc_entry = tk.Entry(win, textvariable=doc_var, width=70, font=("Times New Roman", 10))
    doc_button = tk.Button(win, text="Browse Question Set", command=get_doc,
                           font=("Times New Roman", 10))
    ico_entry = tk.Entry(win, textvariable=ico_var, width=70, font=("Times New Roman", 10))
    ico_button = tk.Button(win, text="Browse Icon", command=get_ico, font=("Times New Roman", 10))
    change_button = tk.Button(win, text="Submit", command=change, width=25, height=2)
    doc_entry.place(x=10, y=10)
    doc_button.place(x=440, y=7)
    ico_entry.place(x=10, y=50)
    ico_button.place(x=440, y=47)
    change_button.place(x=220, y=100)
    win.mainloop()
