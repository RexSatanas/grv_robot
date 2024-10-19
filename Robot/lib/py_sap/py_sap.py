import json
import os
import win32com.client as win32
import time
import logging
from py_common import read_config, get_conf_path, vars


class SAP:

    def __init__(self):
        with open(f"{get_conf_path()}\SystemList.json", "r", encoding="utf-8") as f:
            self.SystemList = json.load(f)
        self.application = None
        self.connection = None

    def __get_application(self):
        """ Подключение к существующему приложению SapLogon или его запуск
        """
        if self.application is None:
            try:
                SapGuiAuto = None
                for i in range(60):
                    try:
                        # подлючаемся к запущенному sap logon 
                        try:
                            SapGuiAuto = win32.GetObject("SAPGUI")
                        # sap logon не был запущен  ранее, запускаем и подключаемся к объекту
                        except:
                            saplogon_path = conf[name]["SAPLogon"]
                            filename = f'"{saplogon_path}"'
                            logging.debug(f"SAPLogon не запущен. Запускаем {filename}")
                            os.popen(filename)
                            time.sleep(5)
                            # subprocess.run([filename], creationflags=subprocess.CREATE_NEW_CONSOLE)
                            SapGuiAuto = win32.GetObject("SAPGUI")
                        self.application = SapGuiAuto.GetScriptingEngine
                        win32.WithEvents(SapGuiAuto, ApplicationEvents)
                        if not self.application is None:
                            break
                        time.sleep(0.5)
                    except:
                        pass
            except RuntimeError as err:
                logging.critical(f"Не удалось подключиться к SAPLogon: {err}")

    def get_session(self, sysid: str):
        """ Подключение к системе в SapLogon

        Args:
            sysid (str): id системы SapLogon
        Returns:
            object | None: COM объект SAPGUI или None
        """
        found_session = None
        session = None
        try:
            if self.application is None:
                self.__get_application()
            system = [el for el in self.SystemList if el["sysid"] == sysid]
            if system:
                system = system[0]
            else:
                raise Exception(f"Система {sysid} не найдена в SystemList")
            for i in range(2):
                found_session = None
                for conn in self.application.Children:
                    found_sysid = False
                    self.connection = conn
                    for session in self.connection.Sessions:
                        if session.Info.SystemName + session.Info.Client == system["sysid"]:
                            found_sysid = True
                            if session.Info.Transaction == "SESSION_MANAGER" or session.Info.Transaction == "S000" or session.Info.Transaction == "SMEN" or session.Info.Transaction == "ZSBWP":
                                found_session = session
                                break
                    if found_sysid and found_session is None and self.connection.Sessions.Count > 0 and self.connection.Sessions.Count < 5:
                        self.connection.Sessions[0].CreateSession
                        time.sleep(1)
                        found_session = self.connection.Sessions[self.connection.Sessions.Count - 1]
                    if not found_session is None:
                        break
                if found_session is None:
                    if i == 1:
                        logging.critical(f"Не удалось соединиться с системой {system['sysid']}")
                        return False
                    else:
                        self.__create_connection(system)
            # win32.DispatchWithEvents(found_session, ApplicationEvents)
            return found_session
        except RuntimeError as err:
            logging.critical(f"Не удалось соединиться с системой {system['sysid']}: {err}")
            return None

    def __create_connection(self, system: dict):
        """ Вход в систему в SapLogon

        Args:
            system (dict): запись из файла SystemList
        """
        try:
            login = ""
            password = ""
            if "pm" in vars and not vars["pm"] is None:
                username, passw = vars["pm"].get_credential(system["sysid"])
                if username != "":
                    login = username
                if passw != "":
                    password = passw
            self.connection = self.application.OpenConnectionByConnectionString(
                f"{system['connectionString']} /SAP_CODEPAGE=1504 /UPDOWNLOAD_CP=2", True)
            session = self.connection.Sessions[0]
            session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = system["sysid"][3:]
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = login
            session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "RU"
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
            session.findById("wnd[0]").sendVKey(0)
            if session.ActiveWindow.Type == "GuiModalWindow":
                if session.ActiveWindow.Text == "Информация по лицензии при многократной регистрации":
                    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Selected = True
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
            time.sleep(2)
            if self.connection.DisabledByServer:
                logging.critical(
                    f"В системе {system['sysid']} у пользователя {login} отсутствуют полномочия на SAPscript")
        except RuntimeError as err:
            logging.critical(f"Не удалось войти в систему {system['sysid']}: {err}")

    def close_session(self, session):
        """ Подключение к существующему приложению SapLogon или его запуск
        """
        self.connection.CloseSession(session.Id)


class ApplicationEvents(object):
    def EndRequest(self):
        # Сообщения в строке статуса
        res = self.ActiveWindow.findById("wnd[0]/sbar", False)
        if res:
            if res.Text != "":
                logging.debug(f"{res.messagetype} {res.Text}")
        # Popup сообщения
        if self.ActiveWindow.Type == "GuiModalWindow":
            res = self.ActiveWindow.findById("usr/txtMESSTXT1", False)
            if res:
                if res.Text != "":
                    logging.debug(f"{self.ActiveWindow.findById('usr/txtIK1').IconName} {res.Text}")
            res = self.ActiveWindow.findById("usr/txtPOPUP01", False)
            if res:
                if res.Text != "":
                    logging.debug(f"{self.ActiveWindow.findById('usr/txtPOPUP01').IconName} {res.Text}")
        # Сообщения в таблице
        res = self.ActiveWindow.findById("wnd[0]/shellcont/shell", False)
        if res:
            if res.Text == "SAPGUI.GridViewCtrl.1":
                foundMsg = False
                foundIcon = False
                for i in range(res.ColumnCount):
                    if res.ColumnOrder(i) == "T_MSG": foundMsg = True
                    if res.ColumnOrder(i) == "%_ICON": foundIcon = True
                if foundMsg and foundIcon:
                    for i in range(res.RowCount):
                        logging.debug(f"{res.GetCellTooltip(i, '%_ICON')} {res.GetCellValue(i, 'T_MSG')}")
        # Закрываем системные сообщения
        if self.ActiveWindow.Type == "GuiModalWindow" and (
                self.ActiveWindow.Text == "Системные сообщения" or self.ActiveWindow.Text == "SAPoffice - экспресс-информация" or self.ActiveWindow.Text == "Информация по вход. документам"):
            res = self.findById("wnd[1]/tbar[0]/btn[0]", False)
            if res:
                self.findById("wnd[1]/tbar[0]/btn[0]").press
            else:
                res = self.findById("wnd[1]/tbar[0]/btn[12]", False)
                if res:
                    self.findById("wnd[1]/tbar[0]/btn[12]").press
            logging.debug("Закрыто системное окно")


name = os.path.basename(__file__).split('.')[0]
conf = read_config()
