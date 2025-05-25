#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文件列表视图 - 显示和管理文件列表
"""

from PyQt6.QtCore import Qt, pyqtSignal, QSortFilterProxyModel
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QPushButton, QHeaderView, QAbstractItemView,
                             QMenu, QLabel, QLineEdit)

from models.file_model import FileManager


class FileListWidget(QWidget):
    """文件列表视图"""

    fileSelected = pyqtSignal(str)  # 选中文件的路径

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_manager = FileManager()

        self._init_ui()
        self._connect_signals()

    def _init_ui(self):
        """初始化界面"""
        # 主布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 标题和搜索框
        header_layout = QHBoxLayout()

        title_label = QLabel("文件列表")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索文件...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMaximumWidth(200)
        header_layout.addWidget(self.search_input)

        layout.addLayout(header_layout)

        # 表格视图
        self.proxy_model = QSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.file_manager.model)
        self.proxy_model.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

        self.table_view = QTableView()
        self.table_view.setModel(self.proxy_model)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.setSortingEnabled(True)

        # 设置表头
        header = self.table_view.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)  # 文件名列自动拉伸
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # 大小列适应内容
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)  # 日期列适应内容
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)  # 状态列适应内容

        layout.addWidget(self.table_view)

        # 按钮布局
        button_layout = QHBoxLayout()

        self.add_file_btn = QPushButton("添加文件")
        button_layout.addWidget(self.add_file_btn)

        self.add_dir_btn = QPushButton("添加目录")
        button_layout.addWidget(self.add_dir_btn)

        self.remove_btn = QPushButton("删除所选")
        button_layout.addWidget(self.remove_btn)

        self.clear_btn = QPushButton("清空")
        button_layout.addWidget(self.clear_btn)

        layout.addLayout(button_layout)

        # 状态标签
        self.status_label = QLabel("0 个文件")
        layout.addWidget(self.status_label)

    def _connect_signals(self):
        """连接信号和槽"""
        # 按钮点击事件
        self.add_file_btn.clicked.connect(self.add_files)
        self.add_dir_btn.clicked.connect(self.add_directory)
        self.remove_btn.clicked.connect(self.remove_selected_files)
        self.clear_btn.clicked.connect(self.clear_files)

        # 表格事件
        self.table_view.doubleClicked.connect(self._on_file_double_clicked)
        self.table_view.clicked.connect(self._on_file_clicked)
        self.table_view.customContextMenuRequested.connect(self._show_context_menu)

        # 搜索框事件
        self.search_input.textChanged.connect(self._filter_files)

        # 文件管理器事件
        self.file_manager.filesAdded.connect(self._update_status_label)
        self.file_manager.scanComplete.connect(self._update_status_label)

    def add_files(self):
        """添加文件"""
        from PyQt6.QtWidgets import QFileDialog

        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择Word文档")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Word文档 (*.docx)")

        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            count = self.file_manager.add_files(file_paths)
            if count > 0:
                self._update_status_label()

    def add_directory(self):
        """添加目录"""
        from PyQt6.QtWidgets import QFileDialog

        dir_path = QFileDialog.getExistingDirectory(
            self,
            "选择目录",
            "",
            QFileDialog.Option.ShowDirsOnly
        )

        if dir_path:
            # 询问是否递归扫描子目录
            from PyQt6.QtWidgets import QMessageBox
            reply = QMessageBox.question(
                self,
                "递归扫描",
                "是否递归扫描所有子目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )

            recursive = reply == QMessageBox.StandardButton.Yes
            count = self.file_manager.scan_directory(dir_path, recursive)

            if count == 0:
                QMessageBox.information(
                    self,
                    "没有文件",
                    "所选目录中未找到Word文档(.docx文件)。"
                )

    def remove_selected_files(self):
        """删除选中的文件"""
        selected_rows = self.table_view.selectionModel().selectedRows()
        if not selected_rows:
            return

        # 转换为源模型索引
        source_indices = [self.proxy_model.mapToSource(index).row() for index in selected_rows]

        # 按索引从大到小排序，以便从后向前删除不会影响其他索引
        source_indices.sort(reverse=True)

        # 删除选中文件
        self.file_manager.remove_files(source_indices)
        self._update_status_label()

    def clear_files(self):
        """清空文件列表"""
        self.file_manager.model.clear()
        self._update_status_label()

    def _filter_files(self, text):
        """根据搜索文本过滤文件列表"""
        self.proxy_model.setFilterFixedString(text)

    def _update_status_label(self):
        """更新状态标签"""
        file_count = self.file_manager.model.rowCount()
        self.status_label.setText(f"{file_count} 个文件")

    def _on_file_clicked(self, index):
        """文件点击事件"""
        # 获取源模型中的索引
        source_index = self.proxy_model.mapToSource(index)
        file_item = self.file_manager.model.get_file(source_index.row())

        if file_item:
            self.fileSelected.emit(file_item.path)

    def _on_file_double_clicked(self, index):
        """文件双击事件"""
        source_index = self.proxy_model.mapToSource(index)
        file_item = self.file_manager.model.get_file(source_index.row())

        if file_item:
            # 尝试使用系统默认应用打开文件
            import os
            os.startfile(file_item.path) if os.name == 'nt' else os.system(f"xdg-open '{file_item.path}'")

    def _show_context_menu(self, position):
        """显示右键菜单"""
        selected_rows = self.table_view.selectionModel().selectedRows()
        if not selected_rows:
            return

        menu = QMenu(self)

        open_action = QAction("打开文件", self)
        open_action.triggered.connect(self._open_selected_file)
        menu.addAction(open_action)

        open_location_action = QAction("打开所在位置", self)
        open_location_action.triggered.connect(self._open_file_location)
        menu.addAction(open_location_action)

        menu.addSeparator()

        remove_action = QAction("删除", self)
        remove_action.triggered.connect(self.remove_selected_files)
        menu.addAction(remove_action)

        menu.exec(self.table_view.viewport().mapToGlobal(position))

    def _open_selected_file(self):
        """打开选中的文件"""
        selected = self.table_view.selectionModel().currentIndex()
        if selected.isValid():
            self._on_file_double_clicked(selected)

    def _open_file_location(self):
        """在文件管理器中打开文件所在位置"""
        selected = self.table_view.selectionModel().currentIndex()
        if not selected.isValid():
            return

        source_index = self.proxy_model.mapToSource(selected)
        file_item = self.file_manager.model.get_file(source_index.row())

        if file_item:
            import os
            import platform

            file_path = file_item.path
            dir_path = os.path.dirname(file_path)

            if platform.system() == "Windows":
                os.startfile(dir_path)
            elif platform.system() == "Darwin":  # macOS
                os.system(f"open '{dir_path}'")
            else:  # Linux
                os.system(f"xdg-open '{dir_path}'")

    def get_selected_files(self):
        """获取选中的文件列表"""
        selected_rows = self.table_view.selectionModel().selectedRows()
        if not selected_rows:
            return []

        # 转换为源模型索引
        source_indices = [self.proxy_model.mapToSource(index).row() for index in selected_rows]

        # 获取文件路径
        files = []
        for idx in source_indices:
            file_item = self.file_manager.model.get_file(idx)
            if file_item:
                files.append(file_item.path)

        return files

    def get_all_files(self):
        """获取所有文件路径"""
        files = []
        for i in range(self.file_manager.model.rowCount()):
            file_item = self.file_manager.model.get_file(i)
            if file_item:
                files.append(file_item.path)
        return files

    def update_file_status(self, file_path, is_processed, error=""):
        """更新文件处理状态"""
        for i in range(self.file_manager.model.rowCount()):
            file_item = self.file_manager.model.get_file(i)
            if file_item and file_item.path == file_path:
                self.file_manager.model.update_file_status(i, is_processed, error)
                break
