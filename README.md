# Инструмент анализа контента соцсетей

Скрипт `analysis.py` читает данные из файла `data.xlsx`, строит несколько графиков и сохраняет их в папку `результаты`.

## Что делает скрипт

- считает распределение по колонкам `Promise_Type`, `Argumentation`, `Trigger`, `CTA`
- строит столбчатые диаграммы по этим категориям
- строит тепловую карту связи `Promise_Type` и `Trigger`
- сохраняет все графики в папку `результаты`

## Структура проекта

- `analysis.py` — основной скрипт анализа
- `data.xlsx` — входные данные
- `результаты/` — папка с сохранёнными графиками

## Что нужно для работы

- Python 3
- установленные библиотеки `pandas`, `matplotlib`, `openpyxl`

## Установка зависимостей

Если используешь виртуальное окружение `venv`, выполни:

```powershell
.\venv\Scripts\python.exe -m pip install pandas matplotlib openpyxl
```

Если запускаешь обычным Python:

```powershell
python -m pip install pandas matplotlib openpyxl
```

## Как пользоваться

1. Помести файл `data.xlsx` в корень проекта.
2. Убедись, что в Excel-файле есть колонки:
   `Promise_Type`, `Argumentation`, `Trigger`, `CTA`
3. Запусти скрипт.

Через виртуальное окружение:

```powershell
.\venv\Scripts\python.exe analysis.py
```

Или обычным Python:

```powershell
python analysis.py
```

## Результат работы

После запуска скрипт:

- создаст папку `результаты`, если её ещё нет
- сохранит туда файлы:

`promise_types.png`  
`argumentation_types.png`  
`triggers.png`  
`calls_to_action.png`  
`promise_type_vs_trigger.png`

По умолчанию скрипт сохраняет графики без открытия окон `matplotlib`, чтобы не зависать в терминале.

## Возможные ошибки

Если появляется ошибка вида `No module named ...`, значит не установлена одна из библиотек. Установи зависимости из раздела выше и запусти скрипт снова.

Если появляется ошибка про `openpyxl`, установи пакет:

```powershell
.\venv\Scripts\python.exe -m pip install openpyxl
```
