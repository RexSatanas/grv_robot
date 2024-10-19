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

path = os.getcwd()
py_common.read_config()
sap = py_sap.SAP()

zhr_pa61_tm = {'transaction': 'ZHR_PA61_TM', 'info_type': '9700', 'info_type_2': '9705'}


def sap_work(system):
    ex_files_lst = excel_work.get_file_list()  # список всех excel доументов

    list_count = 0
    for _ in ex_files_lst:  # считаем кол-во документов
        list_count += 1

    session = sap.get_session(system)
    try:    # если появятся доп окна при логине
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception:
        pass

    session.findById("wnd[0]").sendVKey(0)

    w = 0  # счетчик
    while w < list_count:
        worker_data = excel_work.get_worker_data(ex_files_lst[w])

        session.findById("wnd[0]/tbar[0]/okcd").text = zhr_pa61_tm['transaction']  # вводим транзакцию
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell[1]").topNode = "          1"
        session.findById("wnd[0]/usr/ctxtRP50G-PERNR").text = worker_data['tab_num']  # ввод таб номера
        session.findById("wnd[0]/usr/ctxtRP50G-PERNR").caretPosition = 6
        session.findById("wnd[0]").sendVKey(0)
        result = session.findById("wnd[0]/sbar/pane[0]").text
        if 'Лицо' in result:
            logging.info(result)
        else:
            session.findById(
                "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350"
                "/ctxtRP50G-CHOIC").text = zhr_pa61_tm['info_type']  # вводим Инфо-тип
            session.findById(
                "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                "/ctxtRP50G-BEGDA").text = worker_data['month_start']  # вводим дату начала периода
            session.findById(
                "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                "/ctxtRP50G-ENDDA").text = worker_data['month_end']  # вводим дату конца периода
            session.findById(
                "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                "/ctxtRP50G-ENDDA").setFocus()
            session.findById(
                "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                "/ctxtRP50G-ENDDA").caretPosition = 10
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[5]").press()  # нажимаем создать
            time.sleep(2)
            session.findById("wnd[0]/usr/ctxtP9700-BEGDA").text = worker_data['month_start']
            session.findById("wnd[0]/usr/ctxtP9700-ENDDA").text = worker_data['month_end']
            session.findById("wnd[0]/usr/ctxtP9700-PERIOD").text = "МЕС"
            session.findById("wnd[0]/usr/ctxtP9700-PERIOD").caretPosition = 3
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            result = session.findById("wnd[0]/sbar/pane[0]").text
            if result == 'При вводе этих данных будет удалена запись данных' or 'Запись, действит.' in result:
                logging.info('у пользователя {0} уже был инфотип {1}.'.format(
                    worker_data['fio'],
                    zhr_pa61_tm['info_type']))
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
            else:
                logging.info('{0} инфотип {1} введен'.format(worker_data['fio'], zhr_pa61_tm['info_type']))

        w += 1

        # начался табельщик

    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZHR_PT_SCHED_EDITOR"  # переход в тарнзакцию
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkP_M_SUM").selected = -1

    w = 0  # счетчик
    while w < list_count:

        worker_data = excel_work.get_worker_data(ex_files_lst[w])  # словарь с данными о работнике
        dates_and_time = excel_work.get_grv_dates(ex_files_lst[w])  # все грв даты
        amount = excel_work.get_amount_of_dates(ex_files_lst[w])

        session.findById("wnd[0]/usr/ctxtSO_PERNR-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtSO_PERNR-LOW").text = worker_data['tab_num']  # вводим таб номер
        session.findById("wnd[0]/usr/txtP_YEAR").text = ""
        session.findById("wnd[0]/usr/txtP_YEAR").text = worker_data['year']  # вводим расчетный год
        session.findById("wnd[0]/usr/txtP_M_N_B").text = ""
        session.findById("wnd[0]/usr/txtP_M_N_B").text = worker_data['month']  # вводим месяц с
        session.findById("wnd[0]/usr/txtP_M_N_E").text = ""
        session.findById("wnd[0]/usr/txtP_M_N_E").text = worker_data['month']  # вводим месяц по
        session.findById("wnd[0]/usr/txtP_M_N_E").setFocus()
        session.findById("wnd[0]/usr/txtP_M_N_E").caretPosition = 2

        session.findById("wnd[0]/tbar[1]/btn[8]").press()  # нажимаем выполнить

        result = session.findById("wnd[0]/sbar/pane[0]").text
        try:
            if 'Лицо' in result:
                logging.info(result)
            else:
                session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").currentCellColumn = ""
                session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").selectedRows = "1"
                logging.info('Заношу график {0}'.format(worker_data['fio']))
                i = 0
                try:
                    while i < amount:
                        personal_date = dates_and_time[i].split(' ')[0]
                        day = personal_date.split('.')[0]
                        work_s = dates_and_time[i].split(' ')[1]
                        work_e = dates_and_time[i].split(' ')[2]
                        rest_s = dates_and_time[i].split(' ')[3]
                        rest_e = dates_and_time[i].split(' ')[4]
                        time.sleep(1)
                        # установить видимость ячейки
                        session.findById(
                            "wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").firstVisibleColumn = f"DAY{day}"
                        try:
                            if work_s == 'МП':
                                session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").modifyCell(
                                    1, f"DAY{day}", "МП")
                                session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").triggerModified()
                            else:
                                session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/"
                                                 "shell").setCurrentCell(1, f"DAY{day}")  # выбрать ячейку
                                session.findById("wnd[0]/"
                                                 "usr/cntlCONTAINER0100/shellcont/shell").doubleClickCurrentCell()
                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/cntlCONTAINER2000/"
                                    "shellcont/shell").insertRows("0")  # добавить запись

                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                    "cntlCONTAINER2000/shellcont/shell").triggerModified()
                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                    "cntlCONTAINER2000/shellcont/shell").modifyCell(0, "BEGUZ", work_s)
                                time.sleep(1)
                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                    "cntlCONTAINER2000/shellcont/shell").currentCellColumn = "ENDUZ"
                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                    "cntlCONTAINER2000/shellcont/shell").selectedRows = "0"
                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                    "cntlCONTAINER2000/shellcont/shell").triggerModified()

                                if rest_s != 'Без_ОП' and rest_e != 'Без_ОП':
                                    time_s_obj = datetime.strptime(rest_s, '%H:%M')
                                    time_e_obj = datetime.strptime(rest_e, '%H:%M')
                                    work_e_obj = datetime.strptime(work_e, '%H:%M')
                                    obed_duration = time_e_obj - time_s_obj
                                    new_work_e = work_e_obj - obed_duration
                                    new_work_e_str = new_work_e.strftime('%H:%M')

                                    session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").modifyCell(0, "ENDUZ", new_work_e_str)
                                else:
                                    session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").modifyCell(0, "ENDUZ", work_e)

                                session.findById(
                                    "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                    "cntlCONTAINER2000/shellcont/shell").pressToolbarButton("SAVE_DAY")
                        except Exception:
                            pass

                        logging.info('Добавлена информация по {0} дню'.format(day))

                        i += 1

                    session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").pressToolbarButton(
                        "SAVE")  # сохраним табель
                    session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").currentCellColumn = "STATUS_NAME"
                    session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").clearSelection()
                    session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").contextMenu()
                    session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").selectContextMenuItem(
                        "CHECK_ONE")  # проверяем
                    result = session.findById("wnd[0]/sbar/pane[0]").text
                    if result == 'При проверке графиков были ошибки':
                        logging.info('{0} {1}'.format(worker_data['fio'], result))
                    else:
                        session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").setCurrentCell(
                            1, "STATUS_NAME")
                        session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").clearSelection()
                        session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").contextMenu()
                        session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").selectContextMenuItem(
                            "APPR_ONE")  # утверждаем
                        session.findById("wnd[1]/usr/btnBUTTON_1").press()
                        logging.info('{0} график утвержден'.format(worker_data['fio']))
                    session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    time.sleep(1)
                except Exception as e:
                    logging.error(e)
        except Exception as e:
            logging.error(e)
        w += 1

    session.findById("wnd[0]").close()  # закрываем sap
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  # подтверждаем


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
    sap_work('H17200')
    logging.info('Окончание работы робота')
