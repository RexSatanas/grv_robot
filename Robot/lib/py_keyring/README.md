# py_keyring

Функции для работы с Диспетчером учетных записей windows (безопасное хранение логинов и паролей)

Файл **py_keyring.py** содержит следующие процедуры и функции:

* **get_credential** - Получение логина и пароля
* **set_credential** - Сохранение логина и пароля для указанной системы

### Необходимые библиотеки с GitLab

* [py_common](http://gitlab.dvgd.oao.rzd/erp_rpa/Python/py_common)

### Установка библиотек Python
```
python -m pip install -r requirements.txt
```

### Использование библиотеки
```
import py_keyring
from py_common import vars

username, password = vars['pm'].get_credential("ASDOG")
vars['pm'].set_credential(system, newlogin, newpassword)

```