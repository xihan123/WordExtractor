#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文件控制器 - 处理文件操作的业务逻辑
"""

from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtWidgets import QFileDialog, QMessageBox


class FileController(QObject):
    """文件控制器"""

    fileOpened = pyqtSignal(str)  # 打开文件的路径

    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent

    def open_files(self):
        """打开文件对话框并添加文件"""
        file_dialog = QFileDialog(self.main_window)
        file_dialog.setWindowTitle("选择Word文档")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Word文档 (*.docx)")

        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            self.main_window.file_list_widget.file_manager.add_files(file_paths)

            # 如果有文件被添加，选中第一个并加载预览
            if file_paths:
                self.fileOpened.emit(file_paths[0])

    def open_directory(self):
        """打开目录对话框并扫描文件"""
        dir_path = QFileDialog.getExistingDirectory(
            self.main_window,
            "选择包含Word文档的目录",
            "",
            QFileDialog.Option.ShowDirsOnly
        )

        if dir_path:
            # 询问是否递归扫描子目录
            reply = QMessageBox.question(
                self.main_window,
                "递归扫描",
                "是否递归扫描所有子目录？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )

            recursive = reply == QMessageBox.StandardButton.Yes
            count = self.main_window.file_list_widget.file_manager.scan_directory(dir_path, recursive)

            if count == 0:
                QMessageBox.information(
                    self.main_window,
                    "没有文件",
                    "所选目录中未找到Word文档(.docx文件)。"
                )
