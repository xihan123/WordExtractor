#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文件管理模型 - 处理文件列表和文件操作
"""

import os
from datetime import datetime

from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, pyqtSignal, QObject


class FileItem:
    """表示单个文件项"""

    def __init__(self, path):
        self.path = path
        self.file_name = os.path.basename(path)
        self.file_ext = os.path.splitext(path)[1].lower()

        # 获取文件基本信息
        stat = os.stat(path)
        self.size = stat.st_size
        self.created_time = datetime.fromtimestamp(stat.st_ctime)
        self.modified_time = datetime.fromtimestamp(stat.st_mtime)

        # 状态标记
        self.is_processed = False
        self.processing_status = "未处理"
        self.error_message = ""

    @property
    def size_formatted(self):
        """格式化文件大小"""
        kb = self.size / 1024
        if kb < 1024:
            return f"{kb:.2f} KB"
        else:
            mb = kb / 1024
            return f"{mb:.2f} MB"

    @property
    def created_time_formatted(self):
        """格式化创建时间"""
        return self.created_time.strftime("%Y-%m-%d %H:%M:%S")

    @property
    def modified_time_formatted(self):
        """格式化修改时间"""
        return self.modified_time.strftime("%Y-%m-%d %H:%M:%S")

    def set_processed(self, status=True, error=""):
        """设置处理状态"""
        self.is_processed = status
        if error:
            self.processing_status = "处理失败"
            self.error_message = error
        else:
            self.processing_status = "已处理" if status else "未处理"


class FileTableModel(QAbstractTableModel):
    """文件表格模型，用于在表格视图中显示文件列表"""

    # 列定义
    COLUMNS = ["文件名", "大小", "修改日期", "状态"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.files = []  # 文件项列表

    def rowCount(self, parent=QModelIndex()):
        """返回行数"""
        return len(self.files)

    def columnCount(self, parent=QModelIndex()):
        """返回列数"""
        return len(self.COLUMNS)

    def headerData(self, section, orientation, role):
        """表头数据"""
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self.COLUMNS[section]
        return None

    def data(self, index, role):
        """单元格数据"""
        if not index.isValid() or index.row() >= len(self.files):
            return None

        file_item = self.files[index.row()]

        if role == Qt.ItemDataRole.DisplayRole:
            if index.column() == 0:  # 文件名
                return file_item.file_name
            elif index.column() == 1:  # 大小
                return file_item.size_formatted
            elif index.column() == 2:  # 修改日期
                return file_item.modified_time_formatted
            elif index.column() == 3:  # 状态
                return file_item.processing_status

        elif role == Qt.ItemDataRole.ToolTipRole:
            if index.column() == 0:
                return file_item.path
            elif index.column() == 3 and file_item.error_message:
                return file_item.error_message

        return None

    def add_file(self, path):
        """添加文件到模型"""
        # 检查文件是否已存在
        for file in self.files:
            if file.path == path:
                return False

        row = len(self.files)
        self.beginInsertRows(QModelIndex(), row, row)
        self.files.append(FileItem(path))
        self.endInsertRows()
        return True

    def add_files(self, paths):
        """批量添加文件"""
        new_files = []
        for path in paths:
            # 检查是否为Word文档
            if not path.lower().endswith('.docx'):
                continue

            # 检查文件是否已存在
            if any(file.path == path for file in self.files):
                continue

            new_files.append(FileItem(path))

        if new_files:
            row = len(self.files)
            self.beginInsertRows(QModelIndex(), row, row + len(new_files) - 1)
            self.files.extend(new_files)
            self.endInsertRows()

        return len(new_files)

    def remove_file(self, index):
        """从模型中移除文件"""
        if isinstance(index, QModelIndex):
            index = index.row()

        if 0 <= index < len(self.files):
            self.beginRemoveRows(QModelIndex(), index, index)
            self.files.pop(index)
            self.endRemoveRows()
            return True
        return False

    def remove_files(self, indices):
        """批量移除文件"""
        # 排序并反转索引，这样可以从后向前移除而不影响其它索引
        indices = sorted(indices, reverse=True)

        for index in indices:
            self.remove_file(index)

        return True

    def clear(self):
        """清空所有文件"""
        self.beginResetModel()
        self.files.clear()
        self.endResetModel()

    def get_file(self, index):
        """获取指定索引的文件项"""
        if isinstance(index, QModelIndex):
            index = index.row()

        if 0 <= index < len(self.files):
            return self.files[index]
        return None

    def update_file_status(self, index, is_processed=True, error=""):
        """更新文件处理状态"""
        file_item = self.get_file(index)
        if file_item:
            file_item.set_processed(is_processed, error)
            model_index = self.index(index if isinstance(index, int) else index.row(), 3)
            self.dataChanged.emit(model_index, model_index, [Qt.ItemDataRole.DisplayRole])
            return True
        return False


class FileManager(QObject):
    """文件管理器，处理文件操作并与模型交互"""

    filesAdded = pyqtSignal(int)  # 添加了多少文件
    scanComplete = pyqtSignal(int)  # 扫描完成，参数是添加的文件数

    def __init__(self, parent=None):
        super().__init__(parent)
        self.model = FileTableModel()

    def add_file(self, file_path):
        """添加单个文件"""
        if os.path.isfile(file_path) and file_path.lower().endswith('.docx'):
            success = self.model.add_file(file_path)
            if success:
                self.filesAdded.emit(1)
            return success
        return False

    def add_files(self, file_paths):
        """添加多个文件"""
        count = self.model.add_files(file_paths)
        if count > 0:
            self.filesAdded.emit(count)
        return count

    def scan_directory(self, directory, recursive=True):
        """扫描目录中的Word文档"""
        if not os.path.isdir(directory):
            return 0

        file_paths = []

        if recursive:
            # 递归遍历目录
            for root, _, files in os.walk(directory):
                for file in files:
                    if file.lower().endswith('.docx'):
                        file_paths.append(os.path.join(root, file))
        else:
            # 仅扫描当前目录
            for file in os.listdir(directory):
                if file.lower().endswith('.docx'):
                    file_paths.append(os.path.join(directory, file))

        # 批量添加找到的文件
        count = self.model.add_files(file_paths)
        self.scanComplete.emit(count)
        return count

    def remove_file(self, index):
        """移除单个文件"""
        return self.model.remove_file(index)

    def remove_files(self, indices):
        """移除多个文件"""
        return self.model.remove_files(indices)

    def clear_files(self):
        """清空所有文件"""
        self.model.clear()

    def update_file_status(self, index, is_processed=True, error=""):
        """更新文件处理状态"""
        if isinstance(index, int) and 0 <= index < len(self.files):
            file_item = self.files[index]
            file_item.set_processed(is_processed, error)
            model_index = self.index(index, 3)
            self.dataChanged.emit(model_index, model_index, [Qt.ItemDataRole.DisplayRole])
            return True
        return False
