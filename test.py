from tkinter import *
import time

# 创建窗口
root = Tk()

# 设置窗口大小和位置
root.geometry('200x100+500+300')
root.title('启动中')

# 创建标签显示提示信息
label = Label(root, text='程序启动中，请稍等...')
label.pack(pady=20)

# 更新窗口显示
root.update()

# 模拟程序加载过程
time.sleep(3)

# 关闭提示窗口
root.destroy()

# 这里可以继续编写你的程序代码
# ...

# 启动程序主循环
mainloop()