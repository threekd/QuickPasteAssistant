import ctypes
import tkinter as tk
from tkinter import ttk
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