from tkinter import *
from tkinter.messagebox import showinfo


def getInput(title):
    def return_callback(event):
        print('quit...')
        root.quit()

    def close_callback():
        showinfo('message', 'no click...')



    root = Tk(className=title)
    root.wm_attributes('-topmost', 1)
    screenwidth, screenheight = root.maxsize()
    width = 300
    height = 100
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root.geometry(size)
    root.resizable(0, 0)
    entry = Entry(root)
    entry.bind('<Return>', return_callback)
    entry.place(x=50, y=30)
    entry.focus_set()
    root.protocol("WM_DELETE_WINDOW", close_callback)
    root.mainloop()
    str = entry.get()

    root.destroy()
    return str


print(getInput('请输入SN'))
