"""
Main entry point for HARA Automation Tool
"""

import sys
import os
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from .gui import HARAMainWindow


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application metadata
    app.setApplicationName("HARA Automation Tool")
    app.setOrganizationName("Automotive Safety Systems")
    app.setApplicationDisplayName("HARA Automation")
    
    # Set application ID for Windows taskbar
    if sys.platform == 'win32':
        try:
            import ctypes
            myappid = 'AutomotiveSafety.HARAAutomation.Tool.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass  # Not critical if this fails
    
    # Set application icon
    icon_path = Path(__file__).parent.parent / 'resources' / 'icon.ico'
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    
    # Create and show main window
    window = HARAMainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()