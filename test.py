import os
import time
import tkinter as tk
from tkinter import ttk, messagebox as msg
import pymouse, pykeyboard, pyperclip
import winreg

# 目前已知的BUG：好友列表的最后一个人不能设置备注，否则会死循环
# 如何强制停止：运行时甩动鼠标即可

# 创建用户窗口
mainForm = tk.Tk()
mainForm.title('获取你的微信好友列表')
mainForm.wm_attributes('-topmost', 1)  # 用于设置窗体永远置顶
mainForm.resizable(0, 0)  # 用于设置窗体不能被修改大小

mouseNameXY = (-1, -1)  # 用于存放好友详情里的用户名坐标
mouseNoteXY = (-1, -1)  # 用于存放备注坐标
mouseUserXY = (-1, -1)  # 用于存放好友列表窗口底部的用户名的坐标


def onClickMouseButton():
    global mouseNameXY, mouseNoteXY, mouseUserXY
    msg.showinfo('提示', '请先打开微信，然后打开联系人列表中的第一个联系人，再按确认键。')
    m = pymouse.PyMouse()
    msg.showinfo('提示', '请在点击确认后2秒内将鼠标移动到详情里好友的用户名处，2秒后自动获取鼠标位置。')
    time.sleep(2)
    mouseNameXY = m.position()
    msg.showinfo('成功', '成功，请在点击确认后2秒内将鼠标移动到好友的备注，2秒后自动获取鼠标位置。')
    time.sleep(2)
    mouseNoteXY = m.position()
    msg.showinfo('成功', '成功，请在点击确认后2秒内将鼠标移动到列表里的第一个好友处，2秒后自动获取鼠标位置。')
    time.sleep(2)
    mouseUserXY = m.position()
    msg.showinfo('成功', '任务成功完成')


mouseButton = ttk.Button(mainForm, text='点击获取鼠标位置', width=20, command=onClickMouseButton)
mouseButton.grid(row=0, padx=5, pady=5)


def onClickStartButton():
    def get_desktop():
        # 获取桌面地址
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        return winreg.QueryValueEx(key, "Desktop")[0]
    # 判断是否已经读取鼠标位置
    if mouseNameXY == (-1, -1) or mouseNoteXY == (-1, -1) or mouseUserXY == (-1, -1):
        msg.showerror('错误', '请先获取鼠标位置哦哦哦')
        return
    # 开始群发
    if not msg.askokcancel('提示', '开始干活喽，在我说完成之前别动鼠标哦'):
        return
    m = pymouse.PyMouse()  # 获取鼠标对象
    k = pykeyboard.PyKeyboard()  # 获取键盘对象
    f = open(get_desktop()+os.sep+'微信好友列表.txt', 'w', encoding='utf-8')  # 在桌面创建一个名为“微信好友列表.txt”的文件
    while True:
        # 这里面的time.sleep(0.05)试电脑配置而定，如果设置得太小可能电脑反应不过来
        oriContent = pyperclip.paste()  # oriContent保存复制好友备注前，剪贴板的信息
        m.click(mouseNoteXY[0], mouseNoteXY[1])
        time.sleep(0.05)
        m.click(mouseNoteXY[0], mouseNoteXY[1])  # 经过测试需要按两下备注框才能生效
        time.sleep(0.05)
        k.press_key(k.control_key)
        k.tap_key('a')  # 按下Ctrl+A组合键，全选
        time.sleep(0.05)
        k.tap_key('c')  # 按下Ctrl+C组合键，复制
        k.release_key(k.control_key)
        time.sleep(0.05)
        content = pyperclip.paste()  # content保存复制好友备注后，剪贴板的信息
        if content == oriContent:  # 如果剪贴板的信息没有改变，可能是没有设置备注，则复制好友的用户名
            m.click(mouseNameXY[0], mouseNameXY[1], 2)  # 第三个参数为2，即点击鼠标右键，会弹出“复制”选项
            time.sleep(0.05)
            m.click(mouseNameXY[0] + 5, mouseNameXY[1] + 5)  # 点击弹出的“复制”选项。看不懂的话手动在微信框试一试就知道了
            time.sleep(0.05)
            content = pyperclip.paste()  # content保存复制好友用户名后，剪贴板的信息
            if content == oriContent:  # 如果剪贴板的信息还是没有改变，说明已经到了最后一个人
                break
            else:
                f.write(content+'\n')  # 写入文件
        else:
            f.write(content+'\n')  # 写入文件
        m.click(mouseUserXY[0],mouseUserXY[1])
        time.sleep(0.05)
        k.tap_key(k.down_key)  # 按下鼠标的↓键
        time.sleep(0.05)
    f.close()
    msg.showinfo('提示', '任务成功完成！请到桌面查看！')


startButton = ttk.Button(mainForm, text='启动', width=20, command=onClickStartButton)
startButton.grid(row=1, padx=5, pady=5)

mainForm.mainloop()
