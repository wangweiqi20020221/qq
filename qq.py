# 原理是先将需要发送的文本放到剪贴板中，然后将剪贴板内容发送到qq窗口
# 之后模拟按键发送enter键发送消息

import win32gui
import win32con
import win32clipboard as w
import time
from tkinter import *
from tkinter import messagebox

def getText():
    """获取剪贴板文本"""
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return d

def setText(aString):
    """设置剪贴板文本"""
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    w.CloseClipboard()

def send_qq(to_who, msg):
    """发送qq消息
    to_who：qq消息接收人
    msg：需要发送的消息
    """
    # 将消息写到剪贴板
    setText(msg)
    # 获取qq窗口句柄
    qq = win32gui.FindWindow(None, to_who)
    # 投递剪贴板消息到QQ窗体
    win32gui.SendMessage(qq, 258, 22, 2080193)
    win32gui.SendMessage(qq, 770, 0, 0)
    # 模拟按下回车键
    win32gui.SendMessage(qq, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.SendMessage(qq, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

def timing():
    """定时模块"""
    time_input = []                            # 定义time_input是一个列表，下面的代码是向这个列表中添加信息
    time_input.append(year.get())
    time_input.append(month.get().zfill(2))
    time_input.append(day.get().zfill(2))
    time_input.append(hour.get().zfill(2))
    time_input.append(minute.get().zfill(2))
    time_input.append("00")

    """因为"%s-%s-%s %s:%s:%s" % time_set"中的time_set需要使用元组，而元组不可更改，所以先将所有内容依次输入到列表中，再将列表转换为元组。"""
    time_set = (time_input[0], time_input[1], time_input[2], time_input[3], time_input[4], time_input[5])

    """不停的更新现在的时间，直到现在的时间和发送消息的时间相同时timing模块结束。"""
    while True:

        """获取当前时间"""
        time_now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))

        """如果当前的时间和输入的时间相同，则结束timing模块"""
        if "%s-%s-%s %s:%s:%s" % time_set == time_now:
            return "时间到，正在发送消息。"

        """为防止循环时过多占用计算机资源，在更新一次当前时间后1秒再次更新当前时间"""
        time.sleep(1)

"""发送"""
def send():
    if var.get() is True:
        timing()
        time.sleep(2)  # 系统的时间
    try:
        send_qq(user.get(), msg.get())
        messagebox.showinfo("successful","发送成功")
    except:
        messagebox.showerror("error", "发送失败")

"""
print("此程序只能定时发送qq消息，并且需要打开接受消息的人的会话窗口。")
time.sleep(1)
print("此程序只能在windows环境下运行。")
time.sleep(1)
to_who = input("输入发送给谁：")
msg = input("输入发送的内容：")
a = input("是否定时发送，如果是按1：")
if a == "1":
    timing()
send_qq(to_who, msg)
"""

win = Tk()
Label(text="欢迎使用qq小工具").pack()
Label(text="此程序只能定时发送qq消息，并且需要打开接受消息的人的会话窗口。").pack()
Label(text="运行时程序会出现死机状态，但此情况正常，在发送成功后将恢复正常状态。").pack()
Label(text="接收的人").pack()
user = Entry()
user.pack()
Label(text="发送的消息").pack()
msg = Entry()
msg.pack()
Label(text="发送年份").pack()
year = Entry()
year.pack()
Label(text="发送月份").pack()
month = Entry()
month.pack()
Label(text="发送日期").pack()
day = Entry()
day.pack()
Label(text="发送小时（24小时制）").pack()
hour = Entry()
hour.pack()
Label(text="发送分钟").pack()
minute = Entry()
minute.pack()
var = BooleanVar()
checkbutton = Checkbutton(text="是否定时发送", onvalue="True", variable=var)
checkbutton.pack()
Button(text="发送", command=send).pack()
win.mainloop()
