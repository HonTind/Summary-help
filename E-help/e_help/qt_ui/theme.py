"""Qt logical-pixel tokens; no manual DPI multiplier and no remote assets."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase, QPalette
from .assets import PROJECT


def configure_application(app):
    app.setStyle("Fusion")
    app.styleHints().setColorScheme(Qt.ColorScheme.Light)
    palette = QPalette()
    for role, value in ((QPalette.ColorRole.Window, "#F5F6F8"),
                        (QPalette.ColorRole.Base, "#FFFFFF"),
                        (QPalette.ColorRole.AlternateBase, "#FAFBFD"),
                        (QPalette.ColorRole.WindowText, "#1D1D1F"),
                        (QPalette.ColorRole.Text, "#1D1D1F"),
                        (QPalette.ColorRole.Button, "#FFFFFF"),
                        (QPalette.ColorRole.ButtonText, "#1D1D1F"),
                        (QPalette.ColorRole.Highlight, "#E8F1FF"),
                        (QPalette.ColorRole.HighlightedText, "#0071E3")):
        palette.setColor(role, QColor(value))
    app.setPalette(palette)
    family = next((name for name in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI")
                   if name in QFontDatabase.families()), app.font().family())
    font = QFont(family, 12)
    font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias | QFont.StyleStrategy.NoSubpixelAntialias)
    font.setHintingPreference(QFont.HintingPreference.PreferNoHinting)
    app.setFont(font)
    from pathlib import Path
    app.setStyleSheet(STYLE.replace("FONT_FAMILY", family)
                     .replace("ASSET_PATH", (PROJECT / "assets/ui").as_posix())
                     .replace("QT_RESOURCE_PATH", Path(__file__).resolve().parent.as_posix()))


STYLE = """
QWidget { font-family: 'FONT_FAMILY'; font-size: 16px; color: #1D1D1F; }
QWidget#Root, QWidget#Main { background: #F5F6F8; }
QFrame#Sidebar { background: white; border-right: 1px solid #E2E5EC; }
QFrame#Card { background: white; border: 1px solid #E2E5EC; border-radius: 4px; }
QLabel { background: transparent; }
QWidget#CardBody { background: white; }
QLabel#Brand { font-size: 20px; font-weight: 600; }
QLabel#Title { font-size: 32px; font-weight: 700; background: transparent; }
QLabel#CardTitle { font-size: 22px; font-weight: 600; }
QLabel#FieldLabel { font-size: 17px; font-weight: 600; }
QLabel#Muted { color: #7C8492; }
QLabel#Small { font-size: 14px; color: #7C8492; }
QPushButton { background: white; border: 1px solid #D5D9E0; border-radius: 4px; padding: 6px 16px; }
QPushButton:hover { background: #F3F7FC; border-color: #98B7E6; }
QPushButton:focus { border-color: #0071E3; }
QPushButton:pressed { background: #E8F1FF; }
QPushButton:disabled { color: #A0A7B3; border-color: #E6E9EF; background: #F7F8FA; }
QPushButton#Primary { background: #0071E3; color: white; border-color: #0071E3; }
QPushButton#Primary:hover { background: #0066CC; }
QPushButton#Primary:pressed { background: #005AB4; }
QPushButton#Primary:disabled { background: #C5D8F2; border-color: #C5D8F2; }
QPushButton#Link { border: 1px solid #D5DFEC; border-radius: 7px; background: #F8FAFD; color: #1260B0; padding: 7px 12px; text-align: left; }
QPushButton#Link:hover { background: #EBF3FF; border-color: #8AB4E8; }
QPushButton#Link:pressed { background: #DCEAFF; border-color: #0071E3; }
QPushButton#Link:focus { border: 2px solid #0071E3; padding: 6px 11px; }
QPushButton#Link:disabled { color: #8B95A5; background: #F3F4F6; border-color: #E2E5EB; }
QPushButton#ChooseFiles { color: #0071E3; padding: 4px 16px; }
QPushButton#ChooseFiles:disabled { color: #A0A7B3; background: #F7F8FA; }
QPushButton#Nav { border: none; border-left: 3px solid transparent; border-radius: 0; padding: 10px 14px; text-align: left; background: transparent; color: #505969; }
QPushButton#Nav:hover { background: #F4F7FB; }
QPushButton#Nav:checked { background: #E8F1FF; border-left-color: #0071E3; color: #0071E3; font-weight: 600; }
QPushButton#Mode, QPushButton#ResultTab { background: transparent; border: none; border-bottom: 3px solid transparent; border-radius: 0; padding: 8px 14px; }
QPushButton#Mode:checked, QPushButton#ResultTab:checked { border-bottom-color: #0071E3; color: #0071E3; font-weight: 600; }
QPushButton#Mode:hover, QPushButton#ResultTab:hover { background: #EEF3FB; }
QPushButton#Nav:focus, QPushButton#Mode:focus, QPushButton#ResultTab:focus { border-color: #0071E3; background: #EFF5FF; }
QPushButton#Nav:pressed, QPushButton#Mode:pressed, QPushButton#ResultTab:pressed { background: #DCEAFF; }
QPushButton#Nav:disabled, QPushButton#Mode:disabled, QPushButton#ResultTab:disabled { color: #98A2B3; background: #F7F8FA; }
QFrame#TabBar { background: transparent; border-bottom: 1px solid #E1E6EF; }
QFrame#Info { background: #EAF3FF; border: none; border-radius: 4px; }
QFrame#Info QLabel { color: #0071E3; background: transparent; font-size: 15px; }
QFrame#Files { background: #F8F9FB; border: 1px solid #E7EBF2; border-radius: 5px; }
QLineEdit, QComboBox { background: white; border: 1px solid #D5D9E0; border-radius: 4px; padding: 6px 9px; min-height: 20px; selection-background-color: #DCEBFF; selection-color: #1D1D1F; }
QLineEdit:focus, QComboBox:focus { border-color: #0071E3; }
QLineEdit:disabled, QComboBox:disabled { color: #98A2B3; background: #F8FAFC; }
QComboBox::drop-down { border: none; width: 28px; background: transparent; }
QComboBox::down-arrow { image: url('ASSET_PATH/chevron-down.svg'); width: 18px; height: 18px; }
QComboBox QAbstractItemView { background: white; color: #1D1D1F; border: 1px solid #D8DEE8; selection-background-color: #E8F1FF; selection-color: #0071E3; padding: 5px; }
QCheckBox { spacing: 10px; background: transparent; }
QCheckBox::indicator { width: 17px; height: 17px; border: 1px solid #A4ADB9; border-radius: 2px; background: white; }
QCheckBox::indicator:checked { background: #0071E3; border-color: #0071E3; image: url('QT_RESOURCE_PATH/check.svg'); }
QScrollArea#Settings { border: none; background: transparent; }
QScrollBar:vertical { border: none; background: transparent; width: 9px; margin: 0; }
QScrollBar::handle:vertical { background: #CDD3DD; min-height: 28px; border-radius: 4px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal { border: none; background: #F7F8FB; height: 9px; margin: 0; }
QScrollBar::handle:horizontal { background: #CDD3DD; min-width: 28px; border-radius: 4px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }
QTableWidget { background: white; alternate-background-color: #FAFBFD; border: 1px solid #E1E6EF; border-radius: 2px; gridline-color: #E3E7EF; selection-background-color: #EAF3FF; selection-color: #1D1D1F; }
QTableWidget::item { padding: 0 12px; }
QTableCornerButton::section { background: #EEF0F5; border: none; }
QHeaderView::section { background: #EEF0F5; border: none; border-bottom: 1px solid #E0E5ED; border-right: 1px solid #E0E5ED; padding: 10px 12px; font-weight: 600; }
QFrame#Footer { background: white; border-top: 1px solid #E2E5EC; }
QFrame#Separator { background: #E7EBF2; min-height: 1px; max-height: 1px; border: none; }
QToolButton { border: 1px solid #D5DFEC; border-radius: 6px; background: #F8FAFD; padding: 6px 10px; }
QToolButton:hover { background: #EBF3FF; border-color: #8AB4E8; }
QToolButton:pressed, QToolButton:checked { background: #E8F1FF; border-color: #8AB4E8; }
QToolButton:focus { border-color: #0071E3; }
QToolButton:disabled { color: #8B95A5; background: #F3F4F6; border-color: #E2E5EB; }
QToolButton#Disclosure { text-align: left; color: #505969; font-weight: 600; }
QToolButton#Disclosure:hover { color: #0071E3; }
QListWidget, QTextEdit { background: white; border: 1px solid #E4E8EF; border-radius: 4px; padding: 10px; }
QListWidget::item { padding: 8px; }
QListWidget::item:hover { background: #F2F6FC; }
QListWidget::item:selected { background: #E8F1FF; color: #0071E3; }
QDialog { background: white; }
QLabel#PageLead { font-size: 17px; color: #344054; padding: 4px 0 8px 0; }
QLabel#PageBody { font-size: 15px; color: #596579; padding: 4px 0; }
"""
