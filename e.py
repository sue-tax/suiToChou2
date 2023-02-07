'''
Created on 2021/02/09
エラーメッセージ用メソッド
@author: sue-t
'''


import c


def eprint(title, msg):
    if c.__eprint_type__ == 1:
        print(title,msg)
    elif c.__eprint_type__ == 2:
        from tkinter import messagebox
        messagebox.showinfo(title, msg)
    else:
        pass
