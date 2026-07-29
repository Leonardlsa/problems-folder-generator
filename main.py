#!/usr/bin/env python3
"""错题集生成器 v3.1 - 入口文件"""

import sys

from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main():
    """Application entry point."""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
