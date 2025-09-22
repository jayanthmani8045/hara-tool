"""
UI Styling for HARA Tool
"""

# Professional black and light blue theme
DARK_THEME = """
QMainWindow {
    background-color: #0a0a0a;
    color: #e8e8e8;
}
QWidget {
    background-color: #0a0a0a;
    color: #e8e8e8;
    font-family: 'Segoe UI', 'Arial', sans-serif;
}
QGroupBox {
    font-weight: 600;
    font-size: 11pt;
    border: 1px solid #2c4f7c;
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 15px;
    background-color: #0f0f0f;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 15px;
    padding: 0 10px 0 10px;
    color: #5eb3f6;
    background-color: #0f0f0f;
}
QPushButton {
    background-color: #1e3a5f;
    color: #ffffff;
    border: 1px solid #2c4f7c;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: 500;
    font-size: 10pt;
    min-width: 100px;
}
QPushButton:hover {
    background-color: #2c4f7c;
    border: 1px solid #5eb3f6;
}
QPushButton:pressed {
    background-color: #1a2940;
}
QPushButton:disabled {
    background-color: #141414;
    color: #4a4a4a;
    border: 1px solid #2a2a2a;
}
QLabel {
    color: #c8d4e0;
    font-size: 10pt;
}
QComboBox {
    background-color: #141414;
    border: 1px solid #2c4f7c;
    border-radius: 4px;
    padding: 6px;
    min-width: 150px;
    color: #e8e8e8;
    font-size: 10pt;
}
QComboBox:hover {
    border: 1px solid #5eb3f6;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox::down-arrow {
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 4px solid #5eb3f6;
    margin-right: 5px;
}
QComboBox QAbstractItemView {
    background-color: #141414;
    border: 1px solid #2c4f7c;
    selection-background-color: #2c4f7c;
    selection-color: #ffffff;
}
QCheckBox {
    color: #e8e8e8;
    font-size: 10pt;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border: 1px solid #2c4f7c;
    border-radius: 3px;
    background-color: #141414;
}
QCheckBox::indicator:checked {
    background-color: #5eb3f6;
    border: 1px solid #5eb3f6;
}
QSlider::groove:horizontal {
    height: 6px;
    background: #2c4f7c;
    border-radius: 3px;
}
QSlider::handle:horizontal {
    background: #5eb3f6;
    border: 1px solid #5eb3f6;
    width: 18px;
    height: 18px;
    margin: -6px 0;
    border-radius: 9px;
}
QSlider::handle:horizontal:hover {
    background: #7fc4f8;
}
QTableWidget {
    background-color: #0f0f0f;
    alternate-background-color: #141414;
    gridline-color: #1e3a5f;
    border: 1px solid #2c4f7c;
    border-radius: 4px;
    color: #e8e8e8;
    font-size: 9pt;
}
QTableWidget::item {
    padding: 5px;
    border: none;
}
QTableWidget::item:selected {
    background-color: #2c4f7c;
    color: #ffffff;
}
QHeaderView::section {
    background-color: #1e3a5f;
    color: #ffffff;
    font-weight: 600;
    padding: 8px;
    border: none;
    border-right: 1px solid #0a0a0a;
    border-bottom: 1px solid #2c4f7c;
}
QTextEdit {
    background-color: #050505;
    color: #5eb3f6;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 9pt;
    border: 1px solid #2c4f7c;
    border-radius: 4px;
    padding: 8px;
}
QProgressBar {
    border: 1px solid #2c4f7c;
    border-radius: 4px;
    text-align: center;
    background-color: #141414;
    color: #e8e8e8;
    font-weight: 500;
    height: 22px;
}
QProgressBar::chunk {
    background-color: #5eb3f6;
    border-radius: 3px;
}
QTabWidget::pane {
    border: 1px solid #2c4f7c;
    background-color: #0f0f0f;
    border-radius: 4px;
}
QTabBar::tab {
    background-color: #141414;
    color: #808080;
    padding: 8px 16px;
    margin-right: 2px;
    border: 1px solid #1e3a5f;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}
QTabBar::tab:selected {
    background-color: #0f0f0f;
    color: #5eb3f6;
    font-weight: 500;
    border-color: #2c4f7c;
}
QTabBar::tab:hover:!selected {
    background-color: #1a1a1a;
    color: #c8d4e0;
}
QMessageBox {
    background-color: #141414;
    color: #e8e8e8;
}
QMessageBox QPushButton {
    min-width: 70px;
}
QFrame#separator {
    background-color: #2c4f7c;
    height: 1px;
    border: none;
}
"""

# Specific styles for certain widgets
TITLE_STYLE = """
font-size: 20pt; 
font-weight: 300; 
color: #5eb3f6; 
padding: 15px;
letter-spacing: 2px;
"""

SUBTITLE_STYLE = """
font-size: 10pt; 
color: #808080; 
padding-bottom: 10px;
"""

FILE_LABEL_STYLE_DEFAULT = """
padding: 8px; 
background-color: #141414; 
border: 1px solid #2c4f7c;
border-radius: 4px;
color: #808080;
"""

FILE_LABEL_STYLE_SELECTED = """
padding: 8px; 
background-color: #141414; 
border: 1px solid #5eb3f6;
border-radius: 4px;
color: #5eb3f6;
font-weight: 500;
"""

PROCESS_BUTTON_STYLE = """
QPushButton:enabled {
    background-color: #2c4f7c;
    font-weight: 600;
}
QPushButton:enabled:hover {
    background-color: #5eb3f6;
    color: #0a0a0a;
}
"""

FUZZY_CHECKBOX_STYLE = """
color: #5eb3f6; 
font-weight: 500;
"""

THRESHOLD_LABEL_STYLE = """
color: #5eb3f6; 
font-weight: bold; 
min-width: 40px;
"""

STATUS_LABEL_STYLE = """
color: #5eb3f6; 
font-weight: 500;
"""

INFO_LABEL_STYLE = """
color: #4a5a6a; 
font-size: 9pt;
"""