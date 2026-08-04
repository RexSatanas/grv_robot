import datetime
import os
import subprocess
import py_libPath, py_sap, py_keyring, py_70, py_common
import logging
import time
import excel_work
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from pywinauto import Desktop
from Robot.Scripts.sap_work import SapWork

path = os.getcwd()
py_common.read_config()


def open_folder_and_popup(folder_path):
    subprocess.Popen(['explorer', folder_path])

    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo('Внимание', 'Пожалуйста удалите старые excel файлы и положите новые'
                                    ' в папку и нажмите ОК')
    root.destroy()

    f_name = os.path.basename(folder_path)
    windows = Desktop(backend="uia").windows()
    for w in windows:
        if f_name in w.window_text():
            w.close()


if __name__ == '__main__':
    logging.info('Начало работы робота')
    open_folder_and_popup(excel_work.folder_path)
    sap = SapWork()
    sap.main('H17200')
    logging.info('Окончание работы робота')
