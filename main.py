#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Word文档数据提取器 - 主入口文件
"""

import os
import sys

from PyQt6.QtCore import QLocale
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication

from utils.config_manager import ConfigManager
from views.main_window import MainWindow


class Application(QApplication):
    """自定义应用程序类，处理全局设置和资源"""

    def __init__(self, argv):
        super().__init__(argv)
        self.setOrganizationName("WordExtractor")
        self.setApplicationName("Word文档数据提取器")
        self.setApplicationVersion("1.0.0")

        # 设置中文区域
        QLocale.setDefault(QLocale(QLocale.Language.Chinese, QLocale.Country.China))

        # 加载字体
        self._load_fonts()

        # 加载配置
        self.config_manager = ConfigManager()
        self._apply_theme()

    def _load_fonts(self):
        """加载应用程序字体"""
        default_font = QFont("微软雅黑", 10)
        self.setFont(default_font)

    def _apply_theme(self):
        """应用主题设置"""
        theme = self.config_manager.get_value("appearance/theme", "light")
        if theme == "dark":
            self._set_dark_theme()
        else:
            self._set_light_theme()

    def _set_light_theme(self):
        """设置浅色主题"""
        self.setProperty("theme", "light")
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                color: #212121;
            }
            QMenuBar, QToolBar {
                background-color: #e0e0e0;
            }
            QPushButton {
                background-color: #eeeeee;
                border: 1px solid #bdbdbd;
                border-radius: 4px;
                padding: 6px 12px;
                color: #212121;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #bdbdbd;
            }
            QLineEdit, QTextEdit {
                background-color: #ffffff;
                border: 1px solid #bdbdbd;
                border-radius: 4px;
                padding: 4px;
            }
            QTableView, QListView, QTreeView {
                background-color: #ffffff;
                alternate-background-color: #f5f5f5;
            }
        """)

    def _set_dark_theme(self):
        """设置深色主题"""
        self.setProperty("theme", "dark")
        self.setStyleSheet("""
            QWidget {
                background-color: #212121;
                color: #f5f5f5;
            }
            QMenuBar, QToolBar {
                background-color: #333333;
            }
            QPushButton {
                background-color: #424242;
                border: 1px solid #616161;
                border-radius: 4px;
                padding: 6px 12px;
                color: #f5f5f5;
            }
            QPushButton:hover {
                background-color: #616161;
            }
            QPushButton:pressed {
                background-color: #757575;
            }
            QLineEdit, QTextEdit {
                background-color: #333333;
                border: 1px solid #616161;
                border-radius: 4px;
                padding: 4px;
                color: #f5f5f5;
            }
            QTableView, QListView, QTreeView {
                background-color: #333333;
                alternate-background-color: #424242;
                color: #f5f5f5;
            }
        """)

    def toggle_theme(self):
        """切换主题"""
        current_theme = self.property("theme")
        new_theme = "dark" if current_theme == "light" else "light"
        self.config_manager.set_value("appearance/theme", new_theme)

        if new_theme == "dark":
            self._set_dark_theme()
        else:
            self._set_light_theme()


if __name__ == "__main__":
    try:
        # 检查是否在Windows系统中运行
        if os.name == 'nt':
            # 设置Windows应用程序ID
            import ctypes

            app_id = 'WordExtractor.App.1.0.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

        # 创建应用实例
        app = Application(sys.argv)

        # 启用高DPI支持
        # app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
        # app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

        # 创建并显示主窗口
        window = MainWindow(app)
        window.resize(1280, 800)
        window.show()

        # 使用标准的Qt事件循环而不是async
        sys.exit(app.exec())

    except Exception as e:
        print(f"程序启动出错: {e}")
        sys.exit(1)
