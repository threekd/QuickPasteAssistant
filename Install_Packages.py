import os

os.system(f"python -m pip install -U pip")

libs = ("openpyxl","Pyarrow", "pandas", "pyautogui", "PyQt6")

print("Libraries are being installed …")
try:
    for lib in libs:
        os.system(f"pip install {lib}")
    for lib in libs:
        os.system(f"pip install {lib}")
    print("All libraries have been installed successfully")
except:
    print("Something wrong")
    print("Please run this file to try again.")

os.system("pause")