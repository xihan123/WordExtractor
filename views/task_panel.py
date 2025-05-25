#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
任务面板视图 - 显示和控制任务处理
"""

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QProgressBar, QFileDialog,
                             QMessageBox, QCheckBox)


class TaskPanel(QWidget):
    """任务面板视图"""

    startProcessing = pyqtSignal(str, bool, bool)  # 输出文件路径, 追加模式, 跳过文件信息
    stopProcessing = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_processing = False

        self._init_ui()
        self._connect_signals()

    def _init_ui(self):
        """初始化界面"""
        # 主布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 标题
        header_layout = QHBoxLayout()

        title_label = QLabel("任务处理")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        layout.addLayout(header_layout)

        # 输出设置
        output_layout = QHBoxLayout()

        output_layout.addWidget(QLabel("输出到:"))

        self.output_path_label = QLabel("未设置输出文件")
        self.output_path_label.setStyleSheet("font-style: italic;")
        output_layout.addWidget(self.output_path_label, 1)

        self.browse_btn = QPushButton("浏览...")
        output_layout.addWidget(self.browse_btn)

        layout.addLayout(output_layout)

        # 追加模式选项
        append_layout = QHBoxLayout()

        self.append_checkbox = QCheckBox("追加到现有文件")
        self.append_checkbox.setChecked(False)
        append_layout.addWidget(self.append_checkbox)

        self.skip_file_info_checkbox = QCheckBox("不添加文件名和路径字段")
        self.skip_file_info_checkbox.setChecked(True)  # 默认勾选
        append_layout.addWidget(self.skip_file_info_checkbox)

        append_layout.addStretch()

        layout.addLayout(append_layout)

        # 进度条
        progress_layout = QHBoxLayout()

        progress_layout.addWidget(QLabel("处理进度:"))

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar, 1)

        layout.addLayout(progress_layout)

        # 状态信息
        self.status_label = QLabel("待处理")
        layout.addWidget(self.status_label)

        # 按钮布局
        button_layout = QHBoxLayout()

        self.start_btn = QPushButton("开始处理")
        self.start_btn.setMinimumWidth(120)
        button_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.setMinimumWidth(120)
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.stop_btn)

        button_layout.addStretch()

        self.open_output_btn = QPushButton("打开输出文件")
        self.open_output_btn.setEnabled(False)
        button_layout.addWidget(self.open_output_btn)

        layout.addLayout(button_layout)

    def _connect_signals(self):
        """连接信号和槽"""
        # 按钮点击事件
        self.browse_btn.clicked.connect(self._browse_output_file)
        self.start_btn.clicked.connect(self.start_processing)
        self.stop_btn.clicked.connect(self.stop_processing)
        self.open_output_btn.clicked.connect(self._open_output_file)

    def _browse_output_file(self):
        """浏览输出文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            "",
            "Excel文件 (*.xlsx);;所有文件 (*.*)"
        )

        if file_path:
            # 确保文件扩展名为 .xlsx
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'

            self.output_path_label.setText(file_path)
            self.output_path_label.setStyleSheet("")
            self.open_output_btn.setEnabled(True)

    def start_processing(self):
        """开始处理任务"""
        # 检查是否设置了输出文件
        output_path = self.output_path_label.text()
        if output_path == "未设置输出文件":
            QMessageBox.warning(
                self,
                "输出文件未设置",
                "请先设置输出Excel文件"
            )
            return

        # 确认开始处理
        self.is_processing = True
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.browse_btn.setEnabled(False)
        self.append_checkbox.setEnabled(False)
        self.skip_file_info_checkbox.setEnabled(False)

        # 重置进度条
        self.progress_bar.setValue(0)
        self.status_label.setText("正在处理...")

        # 发送开始处理信号
        self.startProcessing.emit(
            output_path,
            self.append_checkbox.isChecked(),
            self.skip_file_info_checkbox.isChecked()
        )

    def stop_processing(self):
        """停止处理任务"""
        if not self.is_processing:
            return

        reply = QMessageBox.question(
            self,
            "停止处理",
            "确定要停止当前处理任务吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.is_processing = False
            self.stopProcessing.emit()
            self.status_label.setText("已停止")

            # 恢复按钮状态
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.browse_btn.setEnabled(True)
            self.append_checkbox.setEnabled(True)

    def _open_output_file(self):
        """打开输出文件"""
        output_path = self.output_path_label.text()
        if output_path == "未设置输出文件":
            return

        import os
        import platform

        if os.path.exists(output_path):
            if platform.system() == "Windows":
                os.startfile(output_path)
            elif platform.system() == "Darwin":  # macOS
                os.system(f"open '{output_path}'")
            else:  # Linux
                os.system(f"xdg-open '{output_path}'")
        else:
            QMessageBox.warning(
                self,
                "文件不存在",
                "输出文件不存在，请先执行处理任务"
            )

    def update_progress(self, current, total):
        """更新进度条"""
        if total <= 0:
            self.progress_bar.setValue(0)
            return

        # 确保进度计算正确
        percent = min(int((current / total) * 100), 100)
        self.progress_bar.setValue(percent)
        self.status_label.setText(f"正在处理... {current}/{total} ({percent}%)")

        # 如果是最后一个任务，确保进度为100%
        if current >= total:
            self.progress_bar.setValue(100)
            self.status_label.setText(f"处理完成: {current}/{total} ({100}%)")

    def processing_completed(self, success_count, total_count):
        """处理完成"""
        self.progress_bar.setValue(100)
        self.status_label.setText(f"处理完成: {success_count}/{total_count} 个文件成功")

        self.is_processing = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.browse_btn.setEnabled(True)
        self.append_checkbox.setEnabled(True)
        self.skip_file_info_checkbox.setEnabled(True)  # 恢复此控件的状态

        QMessageBox.information(
            self,
            "处理完成",
            f"处理完成\n"
            f"共处理: {total_count} 个文件\n"
            f"成功: {success_count} 个\n"
            f"失败: {total_count - success_count} 个"
        )

    def processing_failed(self, error):
        """处理失败"""
        self.is_processing = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.browse_btn.setEnabled(True)
        self.append_checkbox.setEnabled(True)

        self.status_label.setText(f"处理失败: {error}")

        QMessageBox.critical(
            self,
            "处理失败",
            f"处理任务失败: {error}"
        )
