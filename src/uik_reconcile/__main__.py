import argparse
import os
from pathlib import Path
import sys


def default_data_dir():
    base = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[2]
    return base / "data"


def main():
    parser = argparse.ArgumentParser(description="Сверка УИК — настольная программа")
    parser.add_argument("--data-dir", type=Path, default=default_data_dir(), help="Рабочая папка, по умолчанию data рядом с приложением")
    parser.add_argument("--smoke-test", action="store_true", help="Проверить создание окна без работы с Google и закрыться")
    args = parser.parse_args()
    from PySide6.QtCore import QLockFile, QLocale, QTimer, QTranslator
    from PySide6.QtCore import QLibraryInfo
    from PySide6.QtWidgets import QApplication, QMessageBox
    from .ui import MainWindow
    QLocale.setDefault(QLocale("ru_RU"))
    app = QApplication(sys.argv[:1])
    from .theme import apply_theme
    apply_theme(app)
    translator = QTranslator(app)
    translator.load("qtbase_ru", QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath))
    app.installTranslator(translator)
    root = args.data_dir.resolve()
    try:
        root.mkdir(parents=True, exist_ok=True)
        lock = QLockFile(str(root / ".instance.lock"))
        lock.setStaleLockTime(0)
        if not lock.tryLock(100):
            QMessageBox.warning(None, "Программа уже запущена", "Эта рабочая папка уже открыта другой копией программы.")
            return 1
        window = MainWindow(root)
    except Exception as exc:
        QMessageBox.critical(None, "Ошибка запуска", f"Не удалось открыть рабочую папку. Проверьте права записи.\n{exc}")
        return 1
    window.show()
    if args.smoke_test:
        QTimer.singleShot(250, app.quit)
    result = app.exec()
    lock.unlock()
    return result


if __name__ == "__main__":
    sys.exit(main())
