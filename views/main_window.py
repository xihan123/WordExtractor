#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
主窗口 - 应用程序的主界面
"""

from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QSplitter, QTabWidget, QStatusBar, QMessageBox,
                             QApplication)

from controllers.document_controller import DocumentController
from controllers.file_controller import FileController
from controllers.rule_controller import RuleController
from controllers.task_controller import TaskController
from views.document_viewer import DocumentViewer
from views.file_list_widget import FileListWidget
from views.rule_list_widget import RuleListWidget
from views.task_panel import TaskPanel


class MainWindow(QMainWindow):
    """应用程序主窗口"""

    def __init__(self, app, parent=None):
        super().__init__(parent)
        self.app = app

        # 设置窗口属性
        self.setWindowTitle("Word文档数据提取器")
        self.setMinimumSize(1000, 700)

        # 第一步：创建所有UI组件，但不建立连接
        self._init_ui()

        # 第二步：创建所有控制器
        self._init_controllers()

        # 第三步：建立所有连接
        self._connect_signals()

    def _init_ui(self):
        """初始化界面"""
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(2, 2, 2, 2)

        # 创建分割器
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(self.main_splitter)

        # 左侧：文件列表
        self.file_list_widget = FileListWidget()
        self.main_splitter.addWidget(self.file_list_widget)

        # 右侧：分割为上下两部分
        self.right_splitter = QSplitter(Qt.Orientation.Vertical)
        self.main_splitter.addWidget(self.right_splitter)

        # 右上：文档预览和规则管理的选项卡
        self.right_tab_widget = QTabWidget()
        self.right_splitter.addWidget(self.right_tab_widget)

        # 创建文档预览组件
        self.document_viewer = DocumentViewer()
        self.right_tab_widget.addTab(self.document_viewer, "文档预览")

        # 创建规则管理组件
        self.rule_list_widget = RuleListWidget()
        self.right_tab_widget.addTab(self.rule_list_widget, "提取规则")

        # 右下：任务面板
        self.task_panel = TaskPanel()
        self.right_splitter.addWidget(self.task_panel)

        # 设置分割器初始大小
        self.main_splitter.setSizes([300, 700])
        self.right_splitter.setSizes([500, 200])

        # 创建菜单栏
        self._create_menus()

        # 创建工具栏
        self._create_toolbars()

        # 创建状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def _init_controllers(self):
        """初始化所有控制器"""
        # 创建控制器
        self.file_controller = FileController(self)
        self.document_controller = DocumentController(self)
        self.rule_controller = RuleController(self)
        self.task_controller = TaskController(self)

    def _create_menus(self):
        """创建菜单但不连接操作"""
        # 主菜单栏
        menu_bar = self.menuBar()

        # 文件菜单
        file_menu = menu_bar.addMenu("文件(&F)")

        self.open_action = QAction("打开文件(&O)...", self)
        self.open_action.setShortcut(QKeySequence.StandardKey.Open)
        file_menu.addAction(self.open_action)

        self.open_dir_action = QAction("打开目录(&D)...", self)
        file_menu.addAction(self.open_dir_action)

        file_menu.addSeparator()

        self.clear_files_action = QAction("清空文件列表(&C)", self)
        file_menu.addAction(self.clear_files_action)

        file_menu.addSeparator()

        self.exit_action = QAction("退出(&X)", self)
        self.exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        file_menu.addAction(self.exit_action)

        # 规则菜单
        rule_menu = menu_bar.addMenu("规则(&R)")

        self.new_rule_action = QAction("新建规则(&N)...", self)
        rule_menu.addAction(self.new_rule_action)

        rule_menu.addSeparator()

        self.import_rules_action = QAction("导入规则(&I)...", self)
        rule_menu.addAction(self.import_rules_action)

        self.export_rules_action = QAction("导出规则(&E)...", self)
        rule_menu.addAction(self.export_rules_action)

        rule_menu.addSeparator()

        self.clear_rules_action = QAction("清空规则(&C)", self)
        rule_menu.addAction(self.clear_rules_action)

        # 任务菜单
        task_menu = menu_bar.addMenu("任务(&T)")

        self.start_task_action = QAction("开始处理(&S)", self)
        task_menu.addAction(self.start_task_action)

        self.stop_task_action = QAction("停止处理(&P)", self)
        task_menu.addAction(self.stop_task_action)

        # 视图菜单
        view_menu = menu_bar.addMenu("视图(&V)")

        self.theme_action = QAction("切换主题(&T)", self)
        view_menu.addAction(self.theme_action)

        # 帮助菜单
        help_menu = menu_bar.addMenu("帮助(&H)")

        self.about_action = QAction("关于(&A)", self)
        help_menu.addAction(self.about_action)

    def _create_toolbars(self):
        """创建工具栏但不连接操作"""
        # 文件工具栏
        file_toolbar = self.addToolBar("文件")
        file_toolbar.setIconSize(QSize(24, 24))
        file_toolbar.setMovable(False)

        self.open_toolbar_action = QAction("打开文件", self)
        file_toolbar.addAction(self.open_toolbar_action)

        self.open_dir_toolbar_action = QAction("打开目录", self)
        file_toolbar.addAction(self.open_dir_toolbar_action)

        file_toolbar.addSeparator()

        # 规则工具栏
        rule_toolbar = self.addToolBar("规则")
        rule_toolbar.setIconSize(QSize(24, 24))
        rule_toolbar.setMovable(False)

        self.new_rule_toolbar_action = QAction("新建规则", self)
        rule_toolbar.addAction(self.new_rule_toolbar_action)

        self.import_rules_toolbar_action = QAction("导入规则", self)
        rule_toolbar.addAction(self.import_rules_toolbar_action)

        self.export_rules_toolbar_action = QAction("导出规则", self)
        rule_toolbar.addAction(self.export_rules_toolbar_action)

        rule_toolbar.addSeparator()

        # 任务工具栏
        task_toolbar = self.addToolBar("任务")
        task_toolbar.setIconSize(QSize(24, 24))
        task_toolbar.setMovable(False)

        self.start_task_toolbar_action = QAction("开始处理", self)
        task_toolbar.addAction(self.start_task_toolbar_action)

        self.stop_task_toolbar_action = QAction("停止处理", self)
        task_toolbar.addAction(self.stop_task_toolbar_action)

    def _connect_signals(self):
        """连接所有信号和槽"""
        # 文件菜单和工具栏操作
        self.open_action.triggered.connect(self.file_controller.open_files)
        self.open_toolbar_action.triggered.connect(self.file_controller.open_files)
        self.open_dir_action.triggered.connect(self.file_controller.open_directory)
        self.open_dir_toolbar_action.triggered.connect(self.file_controller.open_directory)
        self.clear_files_action.triggered.connect(self.file_list_widget.clear_files)
        self.exit_action.triggered.connect(QApplication.instance().quit)

        # 规则菜单和工具栏操作
        self.new_rule_action.triggered.connect(self.rule_controller.create_new_rule)
        self.new_rule_toolbar_action.triggered.connect(self.rule_controller.create_new_rule)
        self.import_rules_action.triggered.connect(self.rule_controller.import_rules)
        self.import_rules_toolbar_action.triggered.connect(self.rule_controller.import_rules)
        self.export_rules_action.triggered.connect(self.rule_controller.export_rules)
        self.export_rules_toolbar_action.triggered.connect(self.rule_controller.export_rules)
        self.clear_rules_action.triggered.connect(self.rule_controller.clear_rules)

        # 规则测试连接
        self.rule_list_widget.ruleSelected.connect(self.document_controller.test_rule_on_document)

        # 任务菜单和工具栏操作
        self.start_task_action.triggered.connect(self.task_panel.start_processing)
        self.start_task_toolbar_action.triggered.connect(self.task_panel.start_processing)
        self.stop_task_action.triggered.connect(self.task_panel.stop_processing)
        self.stop_task_toolbar_action.triggered.connect(self.task_panel.stop_processing)

        # 其他菜单操作
        self.theme_action.triggered.connect(self._toggle_theme)
        self.about_action.triggered.connect(self._show_about_dialog)

        # 文件列表和文档预览的连接
        self.file_list_widget.fileSelected.connect(self.document_controller.load_document)

        # 文档预览和规则管理的连接
        self.document_viewer.textSelected.connect(self.rule_controller.create_rule_from_selection)
        self.document_viewer.tableSelected.connect(self.rule_controller.create_rule_from_table)

        # 规则管理和任务面板的连接
        self.rule_list_widget.ruleListChanged.connect(self.task_controller.update_rules)

        # 任务控制器的连接
        self.task_controller.processingStarted.connect(lambda: self.status_bar.showMessage("正在处理文件..."))
        self.task_controller.processingFinished.connect(lambda: self.status_bar.showMessage("处理完成"))
        self.task_controller.progressUpdated.connect(self._update_progress)

        # 任务面板与任务控制器的连接
        self.task_panel.startProcessing.connect(self.task_controller.start_processing)
        self.task_panel.stopProcessing.connect(self.task_controller.stop_processing)

    def _update_progress(self, current, total):
        """更新进度信息"""
        self.status_bar.showMessage(f"处理中... {current}/{total}")

    def _toggle_theme(self):
        """切换主题"""
        self.app.toggle_theme()

    def _show_about_dialog(self):
        """显示关于对话框"""
        # 创建自定义的关于对话框
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout

        dialog = QDialog(self)
        dialog.setWindowTitle("关于 Word文档数据提取器")
        dialog.setMinimumWidth(400)

        layout = QVBoxLayout(dialog)

        # 使用QLabel显示关于信息
        about_label = QLabel(
            "<h3>Word文档数据提取器 v1.0</h3>"
            "<p>一个现代化的桌面应用，用于批量处理Word文档并提取结构化数据到Excel。</p>"
            "<p>基于PyQt6和Python开发。</p>"
            "<p><b>开发者:</b> <a href='https://github.com/xihan123' style='color: #0066cc;'>xihan123</a></p>"
            "<p><b>开源项目地址:</b> <a href='https://github.com/xihan123/WordExtractor' style='color: #0066cc;'>https://github.com/xihan123/WordExtractor</a></p>",
            dialog
        )
        about_label.setOpenExternalLinks(True)  # 允许打开外部链接
        about_label.setTextFormat(Qt.TextFormat.RichText)
        about_label.setWordWrap(True)
        about_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(about_label)

        # 添加确定按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_button = QPushButton("确定", dialog)
        ok_button.setDefault(True)
        ok_button.clicked.connect(dialog.accept)
        button_layout.addWidget(ok_button)

        layout.addLayout(button_layout)

        # 显示对话框
        dialog.exec()

    def closeEvent(self, event):
        """窗口关闭事件"""
        # 检查是否有未完成的任务
        if self.task_controller.is_processing():
            reply = QMessageBox.question(
                self,
                "确认退出",
                "当前有正在进行的任务，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.task_controller.stop_processing()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
