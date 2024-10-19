import jsonc
import os
import inspect
import logging
import sys
import datetime
import getpass
from pprint import pprint

class json_formatter(logging.Formatter):
    """ класс для вывода лога в формате json """
    def format(self, record: logging.LogRecord) -> str:
        super().format(record)
        # pprint(record.__dict__)
        js = {
            "time": record.asctime,
            "level": record.levelname,
            "message": record.message,
        }
        if 'log_labels' in vars:
            for label in vars['log_labels']:
                js[label] = vars['log_labels'][label]
        return jsonc.dumps(js, ensure_ascii = False)

def read_config():
    """ Чтение файла конфигурации. 
    Если в файле конфигурации встречаются разделы с конфигурацией библиотек, 
    то выполняется обновление конфигурации библиотек и повторная инициализация 
    (вызов метода initialize(), если он существует в библиотеке)
    
    Returns:
        dict: объект json файла конфига в виде словаря
    """    
    path, module = __calling()
    with open(f"{path}/config.json", "r", encoding="utf-8") as f:
        curr_conf = jsonc.load(f)
        res_conf = {}
        for lib,config in curr_conf.items():
            if lib == module:
                res_conf[lib] = config
            else:
                if lib in sys.modules and hasattr(sys.modules[lib], 'conf'):
                    dict_merge(sys.modules[lib].conf, {lib: config})
                    if hasattr(sys.modules[lib], 'initialize'):
                        sys.modules[lib].initialize()
        return res_conf

def write_config(config: dict, key: str,value: str):
    """ Запись в файл конфига

    Args:
        config (dict): конфиг файла в виде словаря
        key (str): ключ в котором надо поменять
        value (str): значение, на которое надо заменить
    """
    path, _ = __calling()
    config[key] = value
    with open(f"{path}/config.json", "w", encoding="utf-8") as f:
        f.write(jsonc.dumps(config, indent=3, ensure_ascii=False))

def initialize():
    """ Инициализация логгера робота

    Returns:
        dict: объект json файла конфига в виде словаря
    """
    path = os.path.dirname(os.path.abspath(sys.modules['__main__'].__file__))
    filename = os.path.basename(sys.modules['__main__'].__file__)
    filename = filename.split('.')[0]
    logfile = f"{path}\\{filename}.log"
    if conf[__name__]["LogFile"]:
        logfile = __get_log_filename(conf[__name__]["LogFile"])

    handlers = [
        logging.FileHandler(logfile),
        logging.StreamHandler(),
    ]

    if conf[__name__]["LogFileJson"]:
        logfilejson = __get_log_filename(conf[__name__]["LogFileJson"])
        json_handler = logging.FileHandler(logfilejson)
        jsf = json_formatter()
        json_handler.setFormatter(jsf)
        handlers.append(json_handler)
        if not 'log_labels' in vars: vars['log_labels'] = {}

    logging.basicConfig(level=logging.DEBUG,
        format = "%(asctime)s %(levelname)-8s %(message)s",
        handlers = handlers,
        force = True
    )

    #  Для JSON
    # logging.basicConfig(level=logging.DEBUG, 
    #     format=json.dumps({"time":"%(asctime)s", "level": "%(levelname)s", "message":"%(message)s", "task": "scs_web_server", "pid":"%(process)s"}),
    #     handlers=[
    #         logging.FileHandler(f"{path}/{filename}.json.log"),
    #         logging.StreamHandler()
    #     ])

def get_conf_name() -> str:
    """ Возвращает имя модуля вызвавшего функцию
    Returns:
        str: Имя модуля
    """
    _, filename = __calling()
    return filename

def get_conf_path() -> str:
    """ Возвращает путь до модуля вызвавшего функцию
    Returns:
        str: путь до модуля
    """
    path, _ = __calling()
    return path

def __calling():
    """ Получение родительского объекта, который вызвал метод 
    """        
    frame = inspect.currentframe()
    caller = inspect.getouterframes(frame, 2)[2].filename
    path = os.path.dirname(os.path.abspath(caller))
    filename = os.path.split(os.path.abspath(caller))[1].split('.')[0]
    return path, filename      

def dict_merge(d1: dict, d2: dict):
    """ Рекурсивное объединение двух словарей
    """
    for key, value in d2.items():
        if key in d1 and isinstance(d1[key], dict) and isinstance(value, dict):
            dict_merge(d1[key], value)
        else:
            d1[key] = value

def __get_log_filename(filename: str)->str:
    """ формирование имени файла из шаблона с переменными
        {path}     - каталог скрипта
        {name}     - имя скрипта, без расширения
        {date}     - дата в формате ГГГГММДД
        {datetime} -  дата и время в формате ГГГГММДДччммсс

    Args:
        filename (str): шаблон имени файла, может включать переменные

    Returns:
        str: имя файла
    """
    path = os.path.dirname(os.path.abspath(sys.modules['__main__'].__file__))
    name = os.path.basename(sys.modules['__main__'].__file__)
    name = name.split('.')[0]
    log_vars = {
        'path':     path,                                             # каталог скрипта
        'name':     name,                                             # имя скрипта, без расширения
        'date':     datetime.datetime.now().strftime('%Y%m%d'),       # дата в формате ГГГГММДД
        'datetime': datetime.datetime.now().strftime('%Y%m%d%H%M%S')  # дата и время в формате ГГГГММДДччммсс
    }
    file = filename
    for lv in log_vars:
        if file.find('{' + lv + '}') >= 0: file = file.replace('{' + lv + '}', log_vars[lv])
    
    return file

def set_log_labels(label: str = None, value: str = None):
    """ Устанавливает переменные для json журнала
    Если имя перменной не задано, то устанавливаются переменные по умолчанию: 
    task - имя выполняющегося скрипта
    pid - идентификатор процесса выполняющегося скрипта

    Args:
        label (str, optional): имя переменной.  По умолчанию None.
        value (str, optional): значение переменной. По умолчанию None.
    """
    if not 'log_labels' in vars: return
    if label is None:
        _, task = __calling()
        vars['log_labels']['task'] = task
        vars['log_labels']['pid'] = os.getpid()
    else:
        vars[label] = value

def set_credentials(systems: list):
    """ Ввод и сохранение учетных данных (логин, пароль) для работы робота

    Args:
        systems (list): список систем
    """
    if not 'pm' in vars:
        print("В роботе не предусмотрено сохранение учетных данных")
        return
    
    for system in systems:
        login, password = vars['pm'].get_credential(system)
        print(f'Введите учетные данные для {system} (Enter не редактировать):')
        newlogin = input(f'  Логин [{login}]:')
        pstr = '******' if password else ''
        newpassword = getpass.getpass(f'  Пароль [{pstr}]:')
        if not newlogin: newlogin = login
        if not newpassword: newpassword = password
        if not newlogin:
            print("Логин не может быть пустым")
            return
        if not newpassword:
            print("Пароль не может быть пустым")
            return
        if newlogin != login or newpassword != password:
            vars['pm'].set_credential(system, newlogin, newpassword)

conf = read_config()
initialize()
vars = {}
