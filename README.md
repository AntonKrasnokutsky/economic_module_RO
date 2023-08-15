# economic_module_RO

## Автор
Краснокутский Антон

## Описание
Экономические отчеты за оказанную медицинскую помощь для лечебно-профилактических учреждений Ростовской области по страховым медицинским организация на основании данных полученных из территориального фонда обязательного медицинского страхования

## Корректировка под операционную систему
Для работы в операционной системе Windows необходимо в файле parse.py переменную delimiter указать '\\', для операционных систем Linux необходимо указать '/'

## Технологии
- Python 3.8
- openpyxl
- et-xmlfile

### Запуск проекта:
Клонировать репозиторий:
```
git clone https://github.com/AntonKrasnokutsky/economic_module_RO.git
```

в файле `settings\settings.xml` указать данные свой МО.

создать виртуальное окруждение и установить пакеты
```
python3.8 -m venv venv
. venv\bin\activate
pip install -U pip
pip install -r requirements.txt
```

запустить
```
python main.py
```