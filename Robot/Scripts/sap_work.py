import datetime
import os
import py_libPath, py_sap, py_keyring, py_common
import logging
import time
import excel_work
from datetime import datetime

path = os.getcwd()
py_common.read_config()


class SapWork:
    def __init__(self):
        self.sap = py_sap.SAP()
        self.session = None

    def main(self, system):
        list_count, ex_files_lst = self.get_files_amount()
        self.sap.get_session(system)
        self.add_info_type_to_worker(
            list_count,
            ex_files_lst
        )
        self.add_dates_to_worksheet()
        self.sap.close_session()


    @staticmethod
    def get_files_amount():
        ex_files_lst = excel_work.get_file_list()
        list_count = 0
        for _ in ex_files_lst:  # считаем кол-во документов
            list_count += 1
        return list_count, ex_files_lst

    def add_info_type_to_worker(self, list_count: int, ex_files_lst: list):
        info_type = "9700"
        w = 0  # счетчик
        while w < list_count:
            worker_data = excel_work.get_worker_data(ex_files_lst[w])

            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nZHR_PA61_TM"  # вводим транзакцию
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell[1]").topNode = "          1"
            self.session.findById("wnd[0]/usr/ctxtRP50G-PERNR").text = worker_data['tab_num']  # ввод таб номера
            self.session.findById("wnd[0]/usr/ctxtRP50G-PERNR").caretPosition = 6
            self.session.findById("wnd[0]").sendVKey(0)
            result = self.session.findById("wnd[0]/sbar/pane[0]").text
            if 'Лицо' in result:
                logging.info(result)
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_ITKEYS:SAPMP50A:0350"
                    "/ctxtRP50G-CHOIC").text = info_type  # вводим Инфо-тип
                self.session.findById(
                    "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                    "/ctxtRP50G-BEGDA").text = worker_data['month_start']  # вводим дату начала периода
                self.session.findById(
                    "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                    "/ctxtRP50G-ENDDA").text = worker_data['month_end']  # вводим дату конца периода
                self.session.findById(
                    "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                    "/ctxtRP50G-ENDDA").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsMENU_TABSTRIP/tabpTAB01/ssubSUBSCR_MENU:SAPMP50A:0400/subSUBSCR_TIME:SAPMP50A:0330"
                    "/ctxtRP50G-ENDDA").caretPosition = 10
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById("wnd[0]/tbar[1]/btn[5]").press()  # нажимаем создать
                self.session.findById("wnd[0]/usr/ctxtP9700-BEGDA").text = worker_data['month_start']
                self.session.findById("wnd[0]/usr/ctxtP9700-ENDDA").text = worker_data['month_end']
                self.session.findById("wnd[0]/usr/ctxtP9700-PERIOD").text = "МЕС"
                self.session.findById("wnd[0]/usr/ctxtP9700-PERIOD").caretPosition = 3
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                result = self.session.findById("wnd[0]/sbar/pane[0]").text
                if result == 'При вводе этих данных будет удалена запись данных' or 'Запись, действит.' in result:
                    logging.info('у пользователя {0} уже был инфотип {1}.'.format(
                        worker_data['fio'],
                        info_type))
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    self.session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
                else:
                    logging.info('{0} инфотип {1} введен'.format(worker_data['fio'], info_type))

            w += 1


    def add_dates_to_worksheet(self, list_count: int, ex_files_lst: list):
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZHR_PT_SCHED_EDITOR"  # переход в тарнзакцию
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/chkP_M_SUM").selected = -1

        w = 0  # счетчик
        while w < list_count:

            worker_data = excel_work.get_worker_data(ex_files_lst[w])  # словарь с данными о работнике
            dates_and_time = excel_work.get_grv_dates(ex_files_lst[w])  # все грв даты
            amount = excel_work.get_amount_of_dates(ex_files_lst[w])
            result = self._fill_table(worker_data)

            try:
                if 'Лицо' in result:
                    logging.info(result)
                else:
                    self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").currentCellColumn = ""
                    self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").selectedRows = "1"
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
                            self.session.findById(
                                "wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").firstVisibleColumn = f"DAY{day}"
                            try:
                                if work_s == 'МП':
                                    self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").modifyCell(
                                        1, f"DAY{day}", "МП")
                                    self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").triggerModified()
                                else:
                                    self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/"
                                                     "shell").setCurrentCell(1, f"DAY{day}")  # выбрать ячейку
                                    self.session.findById("wnd[0]/"
                                                     "usr/cntlCONTAINER0100/shellcont/shell").doubleClickCurrentCell()
                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/cntlCONTAINER2000/"
                                        "shellcont/shell").insertRows("0")  # добавить запись

                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").triggerModified()
                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").modifyCell(0, "BEGUZ", work_s)
                                    time.sleep(1)
                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").currentCellColumn = "ENDUZ"
                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").selectedRows = "0"
                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").triggerModified()

                                    if rest_s != 'Без_ОП' and rest_e != 'Без_ОП':
                                        time_s_obj = datetime.strptime(rest_s, '%H:%M')
                                        time_e_obj = datetime.strptime(rest_e, '%H:%M')
                                        work_e_obj = datetime.strptime(work_e, '%H:%M')
                                        obed_duration = time_e_obj - time_s_obj
                                        new_work_e = work_e_obj - obed_duration
                                        new_work_e_str = new_work_e.strftime('%H:%M')

                                        self.session.findById(
                                            "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                            "cntlCONTAINER2000/shellcont/shell").modifyCell(0, "ENDUZ", new_work_e_str)
                                    else:
                                        self.session.findById(
                                            "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                            "cntlCONTAINER2000/shellcont/shell").modifyCell(0, "ENDUZ", work_e)

                                    self.session.findById(
                                        "wnd[0]/usr/subREPL_AREA:ZHR_PT_PRG_SCHED_EDITOR:2000/"
                                        "cntlCONTAINER2000/shellcont/shell").pressToolbarButton("SAVE_DAY")
                            except Exception:
                                pass

                            logging.info('Добавлена информация по {0} дню'.format(day))

                            i += 1

                        self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").pressToolbarButton(
                            "SAVE")  # сохраним табель
                        self.session.findById(
                            "wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").currentCellColumn = "STATUS_NAME"
                        self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").clearSelection()
                        self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").contextMenu()
                        self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").selectContextMenuItem(
                            "CHECK_ONE")  # проверяем
                        result = self.session.findById("wnd[0]/sbar/pane[0]").text
                        if result == 'При проверке графиков были ошибки':
                            logging.info('{0} {1}'.format(worker_data['fio'], result))
                        else:
                            self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").setCurrentCell(
                                1, "STATUS_NAME")
                            self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").clearSelection()
                            self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").contextMenu()
                            self.session.findById("wnd[0]/usr/cntlCONTAINER0100/shellcont/shell").selectContextMenuItem(
                                "APPR_ONE")  # утверждаем
                            self.session.findById("wnd[1]/usr/btnBUTTON_1").press()
                            logging.info('{0} график утвержден'.format(worker_data['fio']))
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        time.sleep(1)
                    except Exception as e:
                        logging.exception(e)
            except Exception as e:
                logging.exception(e)
            w += 1

    def _fill_table(self, worker_data: dict):
        self.session.findById("wnd[0]/usr/ctxtSO_PERNR-LOW").text = ""
        self.session.findById("wnd[0]/usr/ctxtSO_PERNR-LOW").text = worker_data['tab_num']  # вводим таб номер
        self.session.findById("wnd[0]/usr/txtP_YEAR").text = ""
        self.session.findById("wnd[0]/usr/txtP_YEAR").text = worker_data['year']  # вводим расчетный год
        self.session.findById("wnd[0]/usr/txtP_M_N_B").text = ""
        self.session.findById("wnd[0]/usr/txtP_M_N_B").text = worker_data['month']  # вводим месяц с
        self.session.findById("wnd[0]/usr/txtP_M_N_E").text = ""
        self.session.findById("wnd[0]/usr/txtP_M_N_E").text = worker_data['month']  # вводим месяц по
        self.session.findById("wnd[0]/usr/txtP_M_N_E").setFocus()
        self.session.findById("wnd[0]/usr/txtP_M_N_E").caretPosition = 2

        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()  # нажимаем выполнить

        result = self.session.findById("wnd[0]/sbar/pane[0]").text

        return result