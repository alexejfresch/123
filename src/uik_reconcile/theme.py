"""Единая светлая тема, независимая от тёмной палитры Windows."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QPalette


def apply_theme(app):
    app.setStyle("Fusion")
    app.styleHints().setColorScheme(Qt.ColorScheme.Light)
    palette = QPalette()
    colors = {
        QPalette.ColorRole.Window: "#f3f6fa",
        QPalette.ColorRole.WindowText: "#172b3a",
        QPalette.ColorRole.Base: "#ffffff",
        QPalette.ColorRole.AlternateBase: "#edf2f7",
        QPalette.ColorRole.Text: "#172b3a",
        QPalette.ColorRole.Button: "#e4eaf1",
        QPalette.ColorRole.ButtonText: "#172b3a",
        QPalette.ColorRole.BrightText: "#ffffff",
        QPalette.ColorRole.Highlight: "#155b82",
        QPalette.ColorRole.HighlightedText: "#ffffff",
        QPalette.ColorRole.PlaceholderText: "#566675",
        QPalette.ColorRole.ToolTipBase: "#172b3a",
        QPalette.ColorRole.ToolTipText: "#ffffff",
        QPalette.ColorRole.Link: "#155b82",
        QPalette.ColorRole.LinkVisited: "#5b3477",
        QPalette.ColorRole.Accent: "#155b82",
    }
    for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive, QPalette.ColorGroup.Disabled):
        for role, color in colors.items():
            palette.setColor(group, role, QColor(color))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor("#526273"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base, QColor("#eef2f6"))
    app.setPalette(palette)
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet("""
        QWidget { color: #172b3a; placeholder-text-color: #526273; }
        QMainWindow, QDialog, QMessageBox, QInputDialog { background-color: #f3f6fa; }
        QWidget:disabled { color: #526273; placeholder-text-color: #526273; }
        QLabel#projectHint { background-color: #e1edf7; color: #173e5a;
                             border: 1px solid #a5bfd4; padding: 10px; border-radius: 4px; }
        QTabWidget::pane { border: 1px solid #c0cedc; background-color: #ffffff; }
        QTabBar::tab { padding: 9px 8px; background-color: #e4eaf1; color: #24384a;
                       border: 1px solid #c0cedc; }
        QTabBar::tab:selected { background-color: #ffffff; color: #155b82;
                                border-bottom: 2px solid #155b82; font-weight: bold; }
        QTabBar::tab:hover { background-color: #d6e4f0; color: #172b3a; }
        QPushButton { padding: 7px 12px; background-color: #e4eaf1; color: #172b3a;
                      border: 1px solid #9aafc2; border-radius: 3px; }
        QPushButton:hover { background-color: #d6e4f0; color: #172b3a; }
        QPushButton:pressed { background-color: #155b82; color: #ffffff; }
        QPushButton:disabled { background-color: #eef2f6; color: #526273; border-color: #c0cedc; }
        QPushButton:focus { border: 2px solid #155b82; }
        QLineEdit, QComboBox, QSpinBox, QTextEdit, QListWidget, QTableWidget {
            background-color: #ffffff; color: #172b3a; border: 1px solid #a6b6c6;
            selection-background-color: #155b82; selection-color: #ffffff;
        }
        QLineEdit, QComboBox, QSpinBox { padding: 5px; }
        QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled {
            background-color: #eef2f6; color: #526273;
        }
        QAbstractItemView { background-color: #ffffff; alternate-background-color: #edf2f7;
                            color: #172b3a; selection-background-color: #155b82; selection-color: #ffffff; }
        QAbstractItemView::item:selected { background-color: #155b82; color: #ffffff; }
        QTableWidget { gridline-color: #dce3eb; }
        QHeaderView::section { background-color: #e4eaf1; color: #172b3a;
                               padding: 8px; border: 1px solid #c0cedc; }
        QMenu { background-color: #ffffff; color: #172b3a; }
        QMenu::item:selected { background-color: #155b82; color: #ffffff; }
        QToolTip { background-color: #172b3a; color: #ffffff; border: 1px solid #172b3a; }
    """)
