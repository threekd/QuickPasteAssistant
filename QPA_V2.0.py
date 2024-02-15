import sys, os, time

import pandas as pd
import pyperclip
import pyautogui
from PyQt6.QtGui import QAction, QPainter, QColor, QFont
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QSettings, QSize
from PyQt6.QtWidgets import (QWidget, QMainWindow, QFileDialog, QApplication, QLabel,QComboBox,
    QPushButton, QLineEdit, QHBoxLayout, QVBoxLayout, QProgressBar,QSpinBox, QMessageBox)


is_mainWindow_active = True

class ReadFileThread(QThread):

    notifyProgress = pyqtSignal(int)

    def __init__(self,filepath):
        super().__init__()
        self.filepath = filepath

    def run(self):

        self.sheets_data = pd.read_excel(self.filepath, sheet_name=None)
        self.notifyProgress.emit(1)

class RunLoopThread(QThread):

    completed = pyqtSignal(int)
    checkActiveSignal = pyqtSignal()

    def __init__(self, data, start_row, end_row, delay, next_key):
        super().__init__()
        self.data = data
        self.start_row = start_row
        self.end_row = end_row
        self.delay = delay
        self.next_key = next_key
        self.is_running = True
    
    def wait_click_inputWindow(self):
        global is_mainWindow_active

        while is_mainWindow_active and self.is_running:
            time.sleep(1)
            self.checkActiveSignal.emit()
        time.sleep(2)

    def pause(self):

        self.is_running = False
    
    def run(self):

        self.wait_click_inputWindow()

        for i in range(self.start_row, self.end_row):
            if not self.is_running:
                break

            pyperclip.copy(str(self.data[i]))
            time.sleep(0.5)
            pyautogui.hotkey('ctrl','v')
            time.sleep(self.delay)
            pyautogui.press(self.next_key)

            self.completed.emit(i+1)

class CustomProgressBar(QProgressBar):
    def __init__(self, parent=None):
        super(CustomProgressBar, self).__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Set alignment to center
        self.setTextVisible(False)  # Hide the default percentage text

    def paintEvent(self, e):
        super(CustomProgressBar, self).paintEvent(e)

        # Create a QPainter object for drawing
        painter = QPainter(self)
        painter.setPen(QColor(0, 0, 0))  # Set the pen color to white
        painter.setFont(QFont("Helvetica", 10))  # Set the font and size

        # Calculate the text for the percentage
        progress = self.value() * 100 // self.maximum()
        text = "{}%".format(progress)  # Format the text

        # Calculate the offset position for the text
        # You might need to adjust the offset to fit the progress bar height and font size
        offset = QSize(5, 0)  # Assuming there is a 5px margin to the right side
        
        # Set the text rectangle for drawing, with right alignment
        textRect = self.rect().adjusted(0, 0, -offset.width(), -offset.height())
        painter.drawText(textRect, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, text)
        painter.end()  # End the drawing session
        

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.initUI()
        self.settings = QSettings("Myorganization","MyApp")



    def initUI(self):

        menubar = self.menuBar()
        NextKeyMenu = menubar.addMenu('&about')

        self.labelFilePath = QLabel('Select an Excel: ')
        self.labelSheetName = QLabel('Select a sheet:')
        self.labelColumnName = QLabel('Select a column:')
        self.labelIndex = QLabel('Select a range:')

        self.btn_OpenFile = QPushButton('Open File', self)
        self.btn_Start = QPushButton('Start')
        self.btn_Pause = QPushButton('Pause')

        self.combo_SheetList = QComboBox(self)
        self.combo_ColumnList = QComboBox(self)
        self.combo_ArrowList = QComboBox(self)
        self.combo_ArrowList.addItems(['Move next: →','Move next: ↓','Move next: tab'])

        self.qle_StartRow = QLineEdit()
        self.qle_Lastbegin_row = QLineEdit()

        self.progress = CustomProgressBar(self)

        self.btn_OpenFile.clicked.connect(self.showDialog)
        self.btn_Start.clicked.connect(self.btn_Start_Clicked)
        self.btn_Pause.clicked.connect(self.btn_Pause_Clicked)

        self.combo_SheetList.textActivated.connect(self.combo_SheetList_onActive)
        self.combo_ColumnList.textActivated.connect(self.combo_ColumnList_onActive)

        self.progress.setMinimum(0)
        self.progress.setValue(0)

        hbox_file = QHBoxLayout()
        hbox_file.addWidget(self.labelFilePath, stretch=1)
        hbox_file.addWidget(self.btn_OpenFile, stretch=2)

        hbox_sheet = QHBoxLayout()
        hbox_sheet.addWidget(self.labelSheetName, stretch=1)
        hbox_sheet.addWidget(self.combo_SheetList, stretch=2)

        hbox_column = QHBoxLayout()
        hbox_column.addWidget(self.labelColumnName, stretch=1)
        hbox_column.addWidget(self.combo_ColumnList, stretch=2)

        hbox_range = QHBoxLayout()
        hbox_range.addWidget(self.labelIndex, stretch=1)
        hbox_range.addWidget(self.qle_StartRow, stretch=1)
        hbox_range.addWidget(self.qle_Lastbegin_row, stretch=1)

        hbox_progress = QHBoxLayout()
        hbox_progress.addWidget(self.combo_ArrowList, stretch=1)
        hbox_progress.addWidget(self.progress, stretch=2)

        hbox_buttun = QHBoxLayout()
        hbox_buttun.addStretch(1)
        hbox_buttun.addWidget(self.btn_Start, stretch=1)
        hbox_buttun.addSpacing(50)
        hbox_buttun.addWidget(self.btn_Pause, stretch=1)
        hbox_buttun.addStretch(1)

        self.statusBar = self.statusBar()

        self.labelDelay = QLabel('Delay (ms): ')
        self.spin_Delay = QSpinBox(self)
        self.spin_Delay.setRange(1, 99999)
        self.spin_Delay.setValue(2000)

        self.statusBar.insertWidget(0, self.spin_Delay)
        self.statusBar.insertWidget(0, self.labelDelay) 

        self.status_Label = QLabel('Version 2.0')
        self.status_Label.setStyleSheet('color: red')
        self.statusBar.addPermanentWidget(self.status_Label)

        vbox = QVBoxLayout()
        vbox.addLayout(hbox_file)
        vbox.addLayout(hbox_sheet)
        vbox.addLayout(hbox_column)
        vbox.addLayout(hbox_range)
        vbox.addLayout(hbox_progress)
        vbox.addLayout(hbox_buttun)

        main_widget = QWidget()
        main_widget.setLayout(vbox)
        self.setCentralWidget(main_widget)

        self.resize(320,250)
        self.center()
        self.setWindowTitle('Input Bot')
        self.show()

    def showDialog(self):

        home_dir = self.settings.value('last_dir')
        if home_dir:
            home_dir = home_dir

        FilePath = QFileDialog.getOpenFileName(self,'Open file',home_dir, "Excel(*.xlsx)")

        if FilePath[0]:
            self.btn_OpenFile.setText(os.path.basename(FilePath[0]))
            self.combo_SheetList.clear()
            self.combo_ColumnList.clear()

            self.combo_SheetList.addItem('Reading file...')
            self.combo_ColumnList.addItem('Please wait a few seconds...')
            self.update_status('Loading...')

            self.data = ReadFileThread(FilePath[0])
            self.data.notifyProgress.connect(self.onFileRead)
            self.data.start()

            self.settings.setValue('last_dir', FilePath[0])

    def onFileRead(self,C_progress):

        if C_progress == 1:
            self.Sht_List = list(self.data.sheets_data.keys())
            self.combo_SheetList.clear()
            self.combo_SheetList.addItems(self.Sht_List)
            self.combo_SheetList_onActive(self.combo_SheetList.currentText())
            self.update_status('Ready')

    def combo_SheetList_onActive(self, text):

        self.login_data = self.data.sheets_data[text]
        self.columns = self.login_data.columns.tolist()
        self.combo_ColumnList.clear()
        self.combo_ColumnList.addItems(self.columns)

        self.Lastbegin_row = self.login_data.shape[0]
        self.LastLineNum = self.login_data.shape[1]

        self.qle_StartRow.setText('1')
        self.qle_Lastbegin_row.setText(str(self.Lastbegin_row))

    def combo_ColumnList_onActive(self):
        
        self.qle_StartRow.setText('1')
        self.update_status('Ready')

    def center(self):

        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def btn_Start_Clicked(self):

        next_key_mapping = {
            'Move next: →': 'right',
            'Move next: ↓': 'down',
            'Move next: tab': 'tab'
        }

        selected_key_text = self.combo_ArrowList.currentText()
        next_key = next_key_mapping.get(selected_key_text, 'right')
        if not hasattr(self, 'login_data'):
            self.msg_info('Please load data to get started...')
            return

        if hasattr(self, 'thread_runloop') and self.thread_runloop.isRunning():
            return
        
        self.update_status('Running...')

        begin_row = int(self.qle_StartRow.text())-1
        end_row = int(self.qle_Lastbegin_row.text())

        self.progress.setMaximum(end_row)

        self.thread_runloop = RunLoopThread(self.login_data[self.combo_ColumnList.currentText()],
                                    begin_row,
                                    end_row,
                                    self.spin_Delay.value()/1000,
                                    next_key)
        self.thread_runloop.completed.connect(self.onLoopCompleted)
        self.thread_runloop.checkActiveSignal.connect(self.checkActive)
        self.thread_runloop.start()

    def btn_Pause_Clicked(self):
        global is_mainWindow_active

        if not hasattr(self, 'login_data'):
            self.msg_info('Please load data to get started...')
            return

        is_mainWindow_active = True

        self.thread_runloop.pause()
        self.update_status('Pause...')

    def onLoopCompleted(self,progress):

        self.qle_StartRow.setText(str(progress))
        self.progress.setValue(progress)
        
        if self.progress.value() == self.progress.maximum():
            self.update_status('Completed!')

    def closeEvent(self, event):
        exit()

    def update_status(self,message):
        self.status_Label.setText(message)

    def checkActive(self):

        global is_mainWindow_active

        if self.isActiveWindow():
            is_mainWindow_active = True
            self.update_status('Pause...')
        else:
            is_mainWindow_active = False
            self.update_status('Running...')
    
    def msg_info(self, msg):
        
        msgBox = QMessageBox()
        msgBox.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        msgBox.setIcon(QMessageBox.Icon.Information)
        msgBox.setText(msg)
        msgBox.setWindowTitle("Warning")
        msgBox.exec()
        


def main():

    app = QApplication(sys.argv)
    InputDialog = MainWindow()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
