import logging
import keyring
from typing import Tuple
from py_common import vars

class py_keyring:

    @staticmethod
    def get_credential(sysid: str) -> Tuple[str, str]:
        """ Получение логина и пароля для указанной системы

        Args:
            sysid (string): Идентификатор системы

        Returns:
            string, string: логин и пароль
        """
        cred = keyring.get_credential(sysid, None)
        if cred is None:
            raise Exception(f"Пароль для системы {sysid} не задан")
        return cred.username, cred.password

    @staticmethod
    def set_credential(sysid: str, login:str, password: str):
        """ Сохранение логина и пароля для указанной системы

        Args:
            sysid (str): идентификатор системы
            login (str): логин
            password (str): пароль
        """
        keyring.set_password(sysid, login, password)

if logging.getLogger().getEffectiveLevel() <= logging.DEBUG:
    logging.getLogger('keyring').setLevel(logging.INFO)
    logging.getLogger('win32ctypes.core.cffi').setLevel(logging.INFO)

vars["pm"] = py_keyring()