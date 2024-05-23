QPA_Name = 'QuickPasteAssistant_V3.1.pyw'
QPA_Version = 'Version 3.1'

required_packages = ("openpyxl","Pyarrow", "pandas", "pyautogui", "PyQt6")

try:
    import ctypes
    from tkinter import messagebox
    import sys, os, time, subprocess
    import traceback
    import pandas as pd
    import pyperclip
    import pyautogui
    from PyQt6.QtGui import QAction, QPainter, QColor, QFont, QDesktopServices
    from PyQt6.QtCore import QThread, pyqtSignal, Qt, QSettings, QSize, QUrl, QItemSelectionModel
    from PyQt6.QtWidgets import (QWidget, QMainWindow, QFileDialog, QApplication, QLabel,QComboBox, QListWidgetItem, QListWidget,
    QPushButton, QLineEdit, QHBoxLayout, QVBoxLayout, QProgressBar,QSpinBox, QMessageBox)

except ImportError:

    message = "Some required Python libraries are missing. Do you want to install them?"
    answer = messagebox.askyesno("Install Dependencies", message)
    
    if answer:

        ctypes.windll.kernel32.AllocConsole()

        sys.stdout = open('CONOUT$', 'w', buffering=1)
        sys.stderr = open('CONOUT$', 'w', buffering=1)

        print("Install Packages...")
        print("It may take several minutes...")
        for package in required_packages:
            try:
                print(f"Installing {package}...")
                process = subprocess.Popen([sys.executable, "-m", "pip", "install", package], 
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                out, err = process.communicate()
                print(out.decode())
                print(err.decode(), file=sys.stderr)
            except subprocess.CalledProcessError as e:
                print(f"Failed to install {package}. Error: {e}")

        print("Completed!")
        ctypes.windll.kernel32.FreeConsole()
        subprocess.Popen([sys.executable.replace('python.exe', 'pythonw.exe'), QPA_Name])
    else:
        sys.exit()

is_mainWindow_active = True

class ReadFileThread(QThread):

    notifyread_finished = pyqtSignal(int)

    def __init__(self,filepath):
        super().__init__()
        self.filepath = filepath

    def run(self):

        self.sheets_data = pd.read_excel(self.filepath, sheet_name=None)
        self.notifyread_finished.emit(1)

class RunLoopThread(QThread):

    signal_Loop = pyqtSignal(int)
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
            self.checkActiveSignal.emit()
            time.sleep(1)
        time.sleep(2)

    def pause(self):

        self.is_running = False
    
    def run(self):

        self.wait_click_inputWindow()

        for i in range(self.start_row, self.end_row):
            if not self.is_running:
                break

            pyautogui.typewrite(str(self.data[i]), interval=0.1)
            time.sleep(self.delay)
            pyautogui.press(self.next_key)

            self.signal_Loop.emit(i+1)

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
        self.settings = QSettings("Myorganization","QPA")

        selection_style = """
        QListWidget::item:selected {
            background-color: green;
            color: white;
        }
        """
        self.SelectedItem_ListWidget.setStyleSheet(selection_style)



    def initUI(self):

        menubar = self.menuBar()
        aboutMenu = menubar.addMenu('&About')
        GithubWebsiteAction = QAction('Visit Github', self)
        GithubWebsiteAction.triggered.connect(self.openGithubWebsite)
        aboutMenu.addAction(GithubWebsiteAction)

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
        self.combo_ArrowList.addItems(['Next: ↓','Next: →','Next: tab','Next: enter'])

        self.qle_StartRow = QLineEdit()
        self.qle_Lastbegin_row = QLineEdit()

        self.progressbar_Loop = CustomProgressBar(self)

        self.btn_OpenFile.clicked.connect(self.showDialog)
        self.btn_Start.clicked.connect(self.btn_Start_Clicked)
        self.btn_Pause.clicked.connect(self.btn_Pause_Clicked)

        self.combo_SheetList.textActivated.connect(self.combo_SheetList_onActive)
        self.combo_ColumnList.textActivated.connect(self.combo_ColumnList_onActive)

        self.qle_StartRow.textChanged.connect(self.updateListWidget)
        self.qle_Lastbegin_row.textChanged.connect(self.updateListWidget)

        self.progressbar_Loop.setMinimum(0)
        self.progressbar_Loop.setValue(0)

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

        hbox_progressbar_Loop = QHBoxLayout()
        hbox_progressbar_Loop.addWidget(self.combo_ArrowList, stretch=1)
        hbox_progressbar_Loop.addWidget(self.progressbar_Loop, stretch=2)

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

        self.status_Label = QLabel(QPA_Version)
        self.status_Label.setStyleSheet('color: red')
        self.statusBar.addPermanentWidget(self.status_Label)

        self.vbox1 = QVBoxLayout()
        self.vbox1.addLayout(hbox_file)
        self.vbox1.addLayout(hbox_sheet)
        self.vbox1.addLayout(hbox_column)
        self.vbox1.addLayout(hbox_range)
        self.vbox1.addLayout(hbox_progressbar_Loop)
        self.vbox1.addLayout(hbox_buttun)

        self.SelectedItem_ListWidget = QListWidget()

        self.SelectedItem_ListWidget.hide()  # Initially hide this QWidget

        mainLayout = QHBoxLayout()
        mainLayout.addLayout(self.vbox1)
        mainLayout.addWidget(self.SelectedItem_ListWidget)  # Add the wrapper QWidget to the main layout

        main_widget = QWidget()
        main_widget.setLayout(mainLayout)
        self.setCentralWidget(main_widget)

        self.resize(320,250)
        self.center()
        self.setWindowTitle('QuickPasteAssistant')
        self.show()
    
    def setAllControlsEnabled(self, enabled):
        # List all the widgets you want to enable/disable
        control_list = [
            self.btn_OpenFile,
            self.combo_SheetList,
            self.combo_ColumnList,
            self.combo_ArrowList,
            self.qle_StartRow,
            self.qle_Lastbegin_row,
            self.btn_Start,
            self.spin_Delay
        ]
        for control in control_list:
            control.setEnabled(enabled)

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
            self.data.notifyread_finished.connect(self.onFileRead)
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

        self.updateListWidget()

    def combo_ColumnList_onActive(self):
        
        self.qle_StartRow.setText('1')
        self.qle_Lastbegin_row.setText(str(self.Lastbegin_row))
        self.updateListWidget()
        self.update_status('Ready')

    def updateListWidget(self):
        # first, check if data is loaded and if the text is convertible to integer
        if not hasattr(self, 'login_data') or not self.qle_StartRow.text().isdigit() or not self.qle_Lastbegin_row.text().isdigit():
            return

        begin_row = int(self.qle_StartRow.text()) - 1
        end_row = int(self.qle_Lastbegin_row.text())

        # check if the selected range is valid
        if begin_row >= end_row or end_row > self.Lastbegin_row:
            msg_info('Warning','Invalid row range selected.')
            return

        # clear the existing list
        self.SelectedItem_ListWidget.clear()

        # get the selected column name
        column_name = self.combo_ColumnList.currentText()
        column_name_str = "0) " + column_name
        self.SelectedItem_ListWidget.addItem(QListWidgetItem(column_name_str))

        # loop through the selected range and add the items to the list
        for row in range(begin_row, end_row):
            value = self.login_data.at[row, column_name]  # use at() to get value at a particular cell
            value_str= str(row+1) + ") " + str(value)
            item = QListWidgetItem(value_str)  # convert the value to string
            self.SelectedItem_ListWidget.addItem(item)

        max_width = 0
        for index in range(self.SelectedItem_ListWidget.count()):
            max_width = max(max_width, self.SelectedItem_ListWidget.sizeHintForColumn(index))
        
        self.SelectedItem_ListWidget.setFixedWidth(max_width + 20)
        self.resize(320 + max_width + 20,250)

        self.SelectedItem_ListWidget.show()

    def center(self):

        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def btn_Start_Clicked(self):

        next_key_mapping = {
            'Next: ↓': 'down',
            'Next: →': 'right',
            'Next: tab': 'tab',
            'Next: enter': 'enter'
        }

        selected_key_text = self.combo_ArrowList.currentText()
        next_key = next_key_mapping.get(selected_key_text, 'right')
        if not hasattr(self, 'login_data'):
            msg_info('Warning','Please load data to get started...')
            return

        if hasattr(self, 'thread_runloop') and self.thread_runloop.isRunning():
            return
        
        self.update_status('Running...')

        begin_row = int(self.qle_StartRow.text())-1
        end_row = int(self.qle_Lastbegin_row.text())

        self.progressbar_Loop.setMaximum(end_row)

        self.thread_runloop = RunLoopThread(self.login_data[self.combo_ColumnList.currentText()],
                                    begin_row,
                                    end_row,
                                    self.spin_Delay.value()/1000,
                                    next_key)
        self.thread_runloop.signal_Loop.connect(self.onLoopProgress)
        self.thread_runloop.checkActiveSignal.connect(self.checkActive)

        self.setAllControlsEnabled(False)

        self.thread_runloop.start()

    def btn_Pause_Clicked(self):
        global is_mainWindow_active

        if not hasattr(self, 'login_data'):
            msg_info('Warning','Please load data to get started...')
            return
        if hasattr(self, 'thread_runloop') and self.thread_runloop.isRunning():
            self.thread_runloop.pause()
            self.update_status('Pause...')
        else:
            msg_info('Warning', 'There is no loop running to pause.')

        self.setAllControlsEnabled(True)
        is_mainWindow_active = True

    def onLoopProgress(self,index):

    #    self.qle_StartRow.setText(str(progress))
        begin_row = int(self.qle_StartRow.text())-1

        self.progressbar_Loop.setValue(index)

        self.SelectedItem_ListWidget.clearSelection()
        self.SelectedItem_ListWidget.setCurrentRow(index - begin_row + 1, QItemSelectionModel.SelectionFlag.Select)
        
        if self.progressbar_Loop.value() == self.progressbar_Loop.maximum():
            self.completeLoop()

    def completeLoop(self):
        # Enable text input boxes
        self.setAllControlsEnabled(True)
        self.update_status('Completed!')

    def closeEvent(self, event):
        exit()

    def update_status(self,message):
        self.status_Label.setText(message)

    def checkActive(self):

        global is_mainWindow_active

        if self.isActiveWindow():
            is_mainWindow_active = True
            self.update_status('Please click another input box...')
        else:
            is_mainWindow_active = False
            self.update_status('Running...')

    def openGithubWebsite(self):
        url = QUrl("https://github.com/threekd/QuickPasteAssistant")
        QDesktopServices.openUrl(url)
        
def excepthook(exc_type, exc_value, exc_traceback):
    print("Exception hook called")
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    ErrorMsgBox = QMessageBox()
    ErrorMsgBox.setWindowTitle('An error occurred')
    ErrorMsgBox.setText(error_msg)
    ErrorMsgBox.setIcon(QMessageBox.Icon.Critical)
    ErrorMsgBox.setWindowFlags(ErrorMsgBox.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
    ErrorMsgBox.exec()


def msg_info(title, msg):
    msgBox = QMessageBox()
    msgBox.setWindowFlags(msgBox.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
    msgBox.setIcon(QMessageBox.Icon.Information)
    msgBox.setText(msg)
    msgBox.setWindowTitle(title)
    msgBox.exec()

sys.excepthook = excepthook

def main():
    app = QApplication(sys.argv)
    InputDialog = MainWindow()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
