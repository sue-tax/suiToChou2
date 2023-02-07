'''
Created on 2021/02/09
デバックプリント用メソッド
@author: sue
'''

import c


def dprint(msg):
    if c.__dprint_type__ == 1:
        print(msg)
    elif c.__dprint_type__ == 2:
        from tkinter import messagebox
        messagebox.showinfo("dprint", msg)
    else:
        pass

def dprint_w(title, msg):
    if c.__dprint_type__ == 1:
        print(title + " " + msg)
    elif c.__dprint_type__ == 2:
        from tkinter import messagebox
        messagebox.showinfo(title, msg)
    else:
        pass

def dprint_method_start():
    import inspect
    full_name = str(inspect.stack()[1].filename)
    file_name = full_name.split("\\")[-1]
    dprint_w(
            file_name,
            inspect.stack()[1].function + " start"
        )

def dprint_method_end():
    import inspect
    full_name = str(inspect.stack()[1].filename)
    file_name = full_name.split("\\")[-1]
    dprint_w(
            file_name,
            inspect.stack()[1].function + " end"
        )

def dprint_data(data):
    dprint_w(
            str(type(data)),
            data.__str__()
        )

def dprint_name(name, data):
    dprint_w(
            name,
            data.__str__()
        )
