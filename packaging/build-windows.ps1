$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)
if ($env:OS -ne "Windows_NT") { throw "Сборку Windows необходимо выполнять на Windows." }
python -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 12) else 1)"
if ($LASTEXITCODE -ne 0) { throw "Для сборки нужен Python 3.12. В готовом EXE Python уже включён." }
python -m venv .venv
if ($LASTEXITCODE -ne 0) { throw "Не удалось создать окружение Python 3.12" }
& .venv\Scripts\python.exe -m pip install --require-hashes -r requirements.lock
if ($LASTEXITCODE -ne 0) { throw "Ошибка установки зависимостей" }
& .venv\Scripts\python.exe -m pip install --no-build-isolation --no-deps -e .
if ($LASTEXITCODE -ne 0) { throw "Ошибка установки приложения" }
& .venv\Scripts\python.exe -m pytest -q --junitxml=build/test-results.xml
if ($LASTEXITCODE -ne 0) { throw "Тесты не пройдены" }
& .venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed --onedir --name UIK-Reconcile --collect-data googleapiclient --hidden-import keyring.backends.Windows packaging/launcher.py
if ($LASTEXITCODE -ne 0) { throw "Ошибка сборки" }
Copy-Item docs\USER_GUIDE.md dist\UIK-Reconcile\Инструкция.md
Copy-Item docs\CHANGELOG.md dist\UIK-Reconcile\Изменения.md
Copy-Item examples dist\UIK-Reconcile\examples -Recurse -Force
Write-Host "Создана папка dist\UIK-Reconcile. Проверьте запуск UIK-Reconcile.exe на Windows."
