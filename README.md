# Сверка УИК: Excel и Google Таблицы

Локальная настольная программа на русском языке для Windows: от одной до трёх пар источников, история отправок, ручные подтверждения, объединённая сводка и экспорт XLSX. Сервер и установленный Microsoft Excel не требуются.

## Запуск из исходников

Нужен **Python 3.12 x64**. В Windows откройте PowerShell в папке проекта:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install --require-hashes -r requirements.lock
.venv\Scripts\python.exe -m pip install --no-build-isolation --no-deps -e .
.venv\Scripts\python.exe -m uik_reconcile
```

На Linux для разработки:

```bash
python3.12 -m venv .venv
.venv/bin/python -m pip install --require-hashes -r requirements.lock
.venv/bin/python -m pip install --no-build-isolation --no-deps -e .
.venv/bin/python -m uik_reconcile
```

В облачной среде без экрана проверка окна выполняется так:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m uik_reconcile --data-dir /tmp/uik-smoke --smoke-test
.venv/bin/python -m pytest -q
```

Это проверяет создание окна и локальную БД; не подтверждает доступ к реальной Google Таблице или работу Windows-сборки.

## Переносимая сборка Windows

На Windows выполните:

```powershell
powershell -ExecutionPolicy Bypass -File packaging/build-windows.ps1
```

Скрипт останавливается при ошибке установки, тестов или сборки. Результат — **вся папка** `dist/UIK-Reconcile`, запускается `UIK-Reconcile.exe`. Не переносите только EXE: зависимости лежат рядом. На компьютере пользователя Python не нужен. Рабочая папка `data/` создаётся рядом с EXE при первом запуске. Распакуйте приложение в папку с правом записи, например `D:\Сверка`, а не в `Program Files`.

Также подготовлен workflow `.github/workflows/windows.yml`: он собирает, тестирует и запускает Windows-приложение, затем сохраняет артефакт. Его выполнение не запускалось из текущей Linux-среды.

## Документация

- [Инструкция пользователя](docs/USER_GUIDE.md)
- [Архитектура и правила данных](docs/ARCHITECTURE.md)
- [Проверка перед выпуском](docs/VALIDATION.md)

OAuth требует Desktop app client JSON и включённого Google Sheets API. Программа запрашивает только чтение. Токены сохраняются через системное защищённое хранилище, пароли Google программа не получает. Для реальной интеграционной проверки нужны пользовательская авторизация и доступная таблица; секреты не следует добавлять в Git.
