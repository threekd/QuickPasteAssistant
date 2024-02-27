from tkinter import Tk, Label, Toplevel
from tkinter.ttk import Progressbar

class StartBox(Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title('启动中')
        self.geometry("200x100+500+300")
        self.overrideredirect(True)  # 使窗口没有边框
        Label(self, text='程序启动中，请稍候...').pack(expand=True)
        self.progressbar = Progressbar(self, orient='horizontal', length=100, mode='indeterminate')
        self.progressbar.pack(expand=True)
        self.progressbar.start(10)

    def end(self):
        self.progressbar.stop()
        self.destroy()

class MainWindow(Tk):
    def __init__(self):
        super().__init__()
        self.title('主程序')
        self.geometry('800x600')

        # 假设这里有些初始化代码
        # 例如载入程序配置，初始化用户界面等等

        # 初始化完成，关闭启动窗口
        app_start.end()  # 假设 `app_start` 是 StartBox 实例

if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # 隐藏根窗口

    app_start = StartBox(root)  # 显示启动页
    app_main = MainWindow()  # 初始化主程序

    root.mainloop()