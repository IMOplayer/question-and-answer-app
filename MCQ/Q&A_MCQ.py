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
ans_A = None
ans_B = None
ans_C = None
ans_D = None
que = None
buttonA = None
buttonB = None
buttonC = None
buttonD = None
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
    global doc, win, ans_A, ans_B, ans_C, ans_D, pq, que, entry, button1, button2, doc_var, ico_var
    doc = Document(str(doc_var.get()))
    doc_entry.destroy()
    ico_entry.destroy()
    doc_button.destroy()
    ico_button.destroy()
    change_button.destroy()
    get_question()
    win.iconbitmap(str(ico_var.get()))
    width = win.winfo_screenwidth()
    height = win.winfo_screenheight()
    win.geometry(f"{width}x{height}")
    pq = tk.StringVar()  # 問題
    ans_A = tk.StringVar()  # 答案A
    ans_B = tk.StringVar()  # 答案B
    ans_C = tk.StringVar()  # 答案C
    ans_D = tk.StringVar()  # 答案D
    update()
    que = tk.Label(win, height=10, width=1045, wraplength=1000, font=15, textvariable=pq, bg="#FFFFFF")
    que.pack()
    buttonA = tk.Button(win, height=3, width=100, textvariable=ans_A, command=check_A, font=("Times New Roman", 13))
    buttonA.pack()
    buttonB = tk.Button(win, height=3, width=100, textvariable=ans_B, command=check_B, font=("Times New Roman", 13))
    buttonB.pack()
    buttonC = tk.Button(win, height=3, width=100, textvariable=ans_C, command=check_C, font=("Times New Roman", 13))
    buttonC.pack()
    buttonD = tk.Button(win, height=3, width=100, textvariable=ans_D, command=check_D, font=("Times New Roman", 13))
    buttonD.pack()


def get_question():
    """取得問題"""
    for t in doc.paragraphs:
        temp = ""
        flag = False  # 判斷是否答案
        for run in t.runs:
            temp += run.text
            if run.underline:
                flag = True
        if temp[0] == "A":  # 判斷取得文字是題目還是A,B,C,D
            A.append(temp)
        elif temp[0] == "B":
            B.append(temp)
        elif temp[0] == "C":
            C.append(temp)
        elif temp[0] == "D":
            D.append(temp)
        else:
            question.append(temp)
            flag = False
        if flag:
            ans.append(temp[0])


def update():
    """更新問題"""
    global ans_A, ans_B, ans_C, ans_D, pq, num
    num = random.randrange(len(question))
    pq.set(question[num])
    ans_A.set(A[num])
    ans_B.set(B[num])
    ans_C.set(C[num])
    ans_D.set(D[num])


def check_A():
    """檢查答案是否A"""
    if ans[num] == "A":
        update()
    else:
        messagebox.showwarning(title="Q&A", message="你錯了")


def check_B():
    """檢查答案是否B"""
    if ans[num] == "B":
        update()
    else:
        messagebox.showwarning(title="Q&A", message="你錯了")


def check_C():
    """檢查答案是否C"""
    if ans[num] == "C":
        update()
    else:
        messagebox.showwarning(title="Q&A", message="你錯了")


def check_D():
    """檢查答案是否D"""
    if ans[num] == "D":
        update()
    else:
        messagebox.showwarning(title="Q&A", message="你錯了")


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