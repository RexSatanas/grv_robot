# py_sap

Функции по работе с SAP

Файл **py_sap.py** содержит следующие процедуры и функции:

* **get_session** - Поиск открытой сессии и свободной сессии или открытие нового соединения с системой 
* **close_session** - Завершить сессию SAPGUI

В файле **config.json** содержатся настройки работы модуля:

* **SAPLogon** - Путь к SAP Logon

В файле **SystemList.json** содержится список SAP систем, с указанием строки соединения и версии системы

### Необходимые библиотеки с GitLab

* [py_common](http://gitlab.dvgd.oao.rzd/erp_rpa/Python/py_common)

### Установка библиотек Python
```
python -m pip install -r requirements.txt
```

### Использование библиотеки
```
import py_sap

sap = py_sap.SAP()
system = "QHR196"
session = sap.get_session(system)
sys_process(session) # основная функция робота по работе с SAP
sap.SAP.close_session(session)
```
