import requests
import socket
from datetime import datetime
import logging
import os
from py_common import read_config

name = os.path.basename(__file__).split('.')[0]
conf = read_config()
is_blocked = False

def __check_blocked() -> bool:
    """Проверка блокировки ЭК робота в системе мониторинга

    Returns:
        bool: блокирован ли ЭК робота
    """    
    headers = {}
    conf_el =  conf[name]["conf_el"]
    url = "http://rpa.surw.oao.rzd/rpa70/service/serviceblock.php?conf_el={}&ip_vrm={}".format(conf_el,  conf[name]["ip_vrm"])
    response = requests.get(url, headers=headers)
    if "true" in response.text:
        is_blocked = True
        logging.error("Робот заблокирован")
    else:
        is_blocked = False
    return is_blocked

def start_monitoring(conf_el: str = conf[name]["conf_el"]):
    """ Старт мониторинга

    Args:
        conf_el (str, optional): ЭК робота. Defaults to conf[name]["conf_el"].
    """    
    headers = {
        "content-type": 'application/json;charset="utf-8"',
        "accept": "application/json"
    }

    if not "time_start" in conf[name] or conf[name]["time_start"] == "":
        conf[name]["time_start"] = __format_date(datetime.now())
    
    if not "ip_vrm" in conf[name] or conf[name]["ip_vrm"] == "":
        conf[name]["ip_vrm"] = __get_ip()
                
    if __check_blocked():
        return

    query = {
        "conf_el":  conf_el,
        "time_work":  conf[name]["time_work"],
        "step": "start",
        "time":  conf[name]["time_start"],
        "ip_vrm":  conf[name]["ip_vrm"]
    }

    url = "http://rpa.surw.oao.rzd/rpa70/service/servicestart.php"
    res = requests.post(url, json=query, headers=headers)
    if len(res.text.strip()) <=1 :
        logging.info(f"Старт мониторинга {query['conf_el']} выполнен успешно")
    else:
        logging.error(f"Ошибка: {res.text}")

def stop_monitoring(conf_el: str = conf[name]["conf_el"]):
    """ Окончание мониторинга

    Args:
        conf_el (str, optional): ЭК робота. Defaults to conf[name]["conf_el"].
    """    
    if is_blocked:
        return

    headers = {
        "content-type": 'application/json;charset="utf-8"',
        "accept": "application/json"
    }

    if not "time_end" in conf[name] or conf[name]["time_end"] == "":
        conf[name]["time_end"] = __format_date(datetime.now())

    query = {
        "conf_el":  conf_el,
        "time_work":  conf[name]["time_work"],
        "step": "stop",
        "time_start":  conf[name]["time_start"],
        "ip_vrm":  conf[name]["ip_vrm"],
        "time":  conf[name]["time_end"],
        "log": "Лог работы робота"
    }

    url = "http://rpa.surw.oao.rzd/rpa70/service/servicestart.php"
    res = requests.post(url, json=query, headers=headers)
    if len(res.text.strip()) <= 1:
        logging.info(f"Завершение мониторинга {query['conf_el']} выполнено успешно")
    else:
        logging.error(f"Ошибка: {res.text}")

def __format_date(curr_date: datetime) -> str:
    """ Форматирование даты в формат ГГГГ-мм-дд ЧЧ:ММ:СС

    Args:
        curr_date (datetime): текущее время и дата

    Returns:
        str: отформатированная дата
    """    
    return curr_date.strftime("%Y-%m-%d %H:%M:%S")

def __get_ip() -> str:
    """ Определение IP-адреса

    Returns:
        str: IP-адрес
    """    
    return socket.gethostbyname(socket.gethostname())