#!/use/bin/env python
# _*_ coding:utf-8 _*_


import huitie_gateway
from tkinter.messagebox import showinfo
import tkinter, threading, ctypes, inspect
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showinfo


def getInput(title):
    def return_callback(event):
        # print('quit...')
        root.quit()

    def close_callback():
        showinfo('提示', '请按Enter键...')

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
    entry.insert(0, 'HTKJ000200')
    root.protocol("WM_DELETE_WINDOW", close_callback)
    root.mainloop()
    str = entry.get()

    root.destroy()
    return str

def test_start():
    sn =getInput('请输入SN')
    with open('.\\config', 'w') as f:
        f.write(sn)
    huitie_gateway.excel_set()
    huitie_gateway.log_info()
    try:
        if len(selected) == 0:
            askback = tkinter.messagebox.askyesno(title='请确认', message='未选择任何测试项目......')
            if askback == FALSE:
                sys.exit()
        # selected = ['ping', 'iperf3', 'rtc', 'temp', 'eeprom', 'check_ssd', 'usb_s', 'mac', '5g', 'poe']
        if 'ping' in selected:
            huitie_gateway.test_ping()  # 网口通断测试
        if 'iperf3' in selected:
            huitie_gateway.test_iperf3()  # 网口速率测试
        if 'rtc' in selected:
            huitie_gateway.test_rtc()  # RTC测试
        if 'temp' in selected:
            huitie_gateway.test_temp()  # 温度传感器测试
        if 'eeprom' in selected:
            huitie_gateway.test_eeprom()  # eeprom测试
        if 'check_ssd' in selected:
            huitie_gateway.test_usb_m_and_check_ssd()  # USB_M接口测试，固态硬盘检测
        if 'usb_s' in selected:
            huitie_gateway.test_usb_s()  # USB_S口虚拟U盘测试
        if 'mac' in selected:
            huitie_gateway.test_mac()  # MAC地址测试
        if '5g' in selected:
            huitie_gateway.test_5g()  # 5G网口业务测试，PoE带载测试
        if 'poe' in selected:
            huitie_gateway.test_poe()  # PoE带载测试
        huitie_gateway.check_test_report()  # 检查测试报告是否通过

    except KeyboardInterrupt:
        sys.exit()



class myStdout():	# 重定向类
    def __init__(self):
        # 将其备份
        self.stdoutbak = sys.stdout
        self.stderrbak = sys.stderr
        # 重定向
        sys.stdout = self
        sys.stderr = self

    def write(self, info):
        # info信息即标准输出sys.stdout和sys.stderr接收到的输出信息
        result_print.insert('end', info)	# 在多行文本控件最后一行插入print信息
        result_print.update()	# 更新显示的文本，不加这句插入的信息无法显示
        result_print.see(tkinter.END)	# 始终显示最后一行，不加这句，当文本溢出控件最后一行时，不会自动显示最后一行

    def restoreStd(self):
        # 恢复标准输出
        sys.stdout = self.stdoutbak
        sys.stderr = self.stderrbak

#自定义的线程函数类
def thread_it(func, *args):
  '''将函数放入线程中执行'''
  # 创建线程
  global t
  t = threading.Thread(target=func, args=args)
  # 守护线程
  t.setDaemon(True)
  # 启动线程
  t.start()

def _async_raise(tid, exctype):
  """raises the exception, performs cleanup if needed"""
  tid = ctypes.c_long(tid)
  if not inspect.isclass(exctype):
    exctype = type(exctype)
  res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
  if res == 0:
    raise ValueError("invalid thread id")
  elif res != 1:
    # """if it returns a number greater than one, you're in trouble,
    # and you should call it again with exc=NULL to revert the effect"*斜体样式*""
    ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
    raise SystemError("PyThreadState_SetAsyncExc failed")

def stop_thread(thread):
    try:
        _async_raise(thread.ident, SystemExit)
    except ValueError:
        tkinter.messagebox.showerror(title='测试已终止', message='请勿重复点击！！！')

#------------------------GUI函数------------------------#




def my_GUI():
    mystd = myStdout()  # 实例化重定向类
    win = tkinter.Tk()
    win.title('慧铁测试程序v1.0')
    win.geometry('517x600')

    global result_print
    result_print = tkinter.Text(win, width=71, height=28, relief="solid")  # 创建多行文本控件
    result_print.place(x=5, y=222)    # 布局在窗体上

    cv = Canvas(win, width=2000, height=190)
    cv.place(x=0, y=20)
    cv.create_line(5, 2, 350, 2, 350, 190, 5, 190, 5, 2,  fill='grey')
    cv.create_line(360, 2, 508, 2, 508, 190, 360, 190, 360, 2, fill='grey')
    cv.create_line(360, 71, 508, 71, fill='grey')
    # cv.create_line(360, 94, 508, 94, fill='grey')
    cv.create_line(360, 132, 508, 132, fill='grey')

    # ----------------菜单栏-------------------#


    def _quit():
        win.quit()
        win.destroy()
        exit()

    # Display a Message Box


    def _msgBox1():
        showinfo('软件信息', '网址：http://www.digitgate.cn\n开发者: wangliang\n电话:18321313457\n地址：南京市玄武大道108号徐庄软件园聚慧园1号楼8F')

    def clear_text():
        result_print.delete('1.0', 'end')

    def getselect(item):
        print(item, 'selected')

    # 反选
    def unselectall():
        for index, item in enumerate(list1):
            v[index].set('')

    # 全选
    def selectall():
        for index, item in enumerate(list1):
            v[index].set(item)

    # 获取选择项
    def showselect():
        global selected
        selected = [i.get() for i in v if i.get()]
        print('>> 测试项：', selected)


    # Creating a Menu Bar
    menuBar = Menu(win)
    win.config(menu=menuBar)

    # Add menu items
    fileMenu = Menu(menuBar, tearoff=0)
    # fileMenu.add_command(label="设置")

    fileMenu.add_separator()
    fileMenu.add_command(label="退出", command=_quit)
    menuBar.add_cascade(label="文件", menu=fileMenu)

    # Add another Menu to the Menu Bar and an item
    msgMenu = Menu(menuBar, tearoff=0)
    msgMenu.add_command(label="软件信息", command=_msgBox1)
    msgMenu.add_separator()
    menuBar.add_cascade(label="关于", menu=msgMenu)
    # ----------------菜单栏-------------------#

    # ----------------GUI内容-------------------#
    ttk.Label(win, text="测试项目").place(x=15, y=12)
    # ttk.Label(win, text="SN").place(x=371, y=12)

    frame1 = tkinter.Frame(win, pady=2, padx=2)
    frame1.place(x=29, y=48)
    # 全选反选
    opt = tkinter.IntVar()
    ttk.Radiobutton(frame1, text='全选', variable=opt, value=1, command=selectall).grid(row=0, column=0, sticky='w')
    ttk.Radiobutton(frame1, text='反选', variable=opt, value=0, command=unselectall).grid(row=0, column=1, sticky='w')
    list1 = ['ping', 'iperf3', 'rtc', 'temp', 'eeprom', 'check_ssd', 'usb_s', 'mac', '5g', 'poe']
    v = []
    # 设置勾选框，每四个换行
    for index, item in enumerate(list1):
        v.append(tkinter.StringVar())
        ttk.Checkbutton(frame1, text=item, variable=v[-1], onvalue=item, offvalue='',
                        command=lambda item=item: getselect(item)).grid(row=index // 4 + 1, column=index % 4,
                                                                        sticky='wn')
    ttk.Button(frame1, text="获取测试项", command=showselect).grid(row=index // 4 + 2, column=0)

    # # SA_IP 写入
    # global sn
    # sn = tkinter.StringVar()
    # sn_num = ttk.Entry(win, width=14, textvariable=sn)
    # sn_num.insert(0, 'HTKJ000200')
    # sn_num.place(x=380, y=33)

    # 测试按钮
    action = tkinter.Button(win, text="开始测试", font=("黑体", 13), command=lambda: thread_it(test_start))  # 创建一个按钮, text：显示按钮上面显示的文字, command：关闭窗口，程序继续运行
    action.place(x=388, y=45)  # 设置其在界面中出现的位置  column代表列   row 代表行

    # 停止按钮
    action = tkinter.Button(win, text="停止测试", font=("黑体", 13), command=lambda: stop_thread(t))
    action.place(x=388, y=108)

    # 清空按钮
    action = tkinter.Button(win, text="清空显示", font=("黑体", 13), command=lambda: thread_it(clear_text))
    action.place(x=388, y=165)

    win.mainloop()
    mystd.restoreStd()  # 恢复标准输出

#---------------程序运行入口---------------#


if __name__ == '__main__':
    my_GUI()
