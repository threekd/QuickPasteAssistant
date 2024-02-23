import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTextEdit, QVBoxLayout, QWidget

class EmittingStream:
    def __init__(self, text_edit):
        self.text_edit = text_edit

    def write(self, text):
        self.text_edit.append(text)  # 在文本编辑框中添加文字

    def flush(self):
        pass  # 如果需要，可以在这里实现必要的缓冲清理

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 创建一个 QTextEdit 来显示标准输出
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)  # 将文本编辑设置为只读

        # 使用自定义的 EmittingStream 类来重定向 sys.stdout
        sys.stdout = EmittingStream(self.text_edit)

        # 设置布局
        vbox = QVBoxLayout()
        vbox.addWidget(self.text_edit)

        # 将布局设置到 QWidget 中
        central_widget = QWidget()
        central_widget.setLayout(vbox)

        # 设置固定的 QWidget 作为 QMainWindow 的 central widget
        self.setCentralWidget(central_widget)

        # 显示信息
        print("这条信息会在 QTextEdit 控件中显示。")

def main():
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()