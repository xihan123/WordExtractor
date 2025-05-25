#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
任务控制器 - 处理提取任务的业务逻辑
"""

from PyQt6.QtCore import QObject, pyqtSignal

from models.task_model import TaskManager


class TaskController(QObject):
    """任务控制器"""

    processingStarted = pyqtSignal()
    processingFinished = pyqtSignal()
    progressUpdated = pyqtSignal(int, int)  # current, total

    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.task_manager = TaskManager()
        self.current_rules = []

        self._connect_signals()

    def _connect_signals(self):
        """连接信号和槽"""
        # 任务管理器信号
        self.task_manager.taskStarted.connect(self.processingStarted)
        self.task_manager.taskProgress.connect(self.progressUpdated)
        self.task_manager.taskProgress.connect(self._update_progress)
        self.task_manager.allTasksCompleted.connect(self._on_tasks_completed)
        self.task_manager.taskError.connect(self._on_task_error)

        # 任务面板信号
        self.main_window.task_panel.startProcessing.connect(self.start_processing)
        self.main_window.task_panel.stopProcessing.connect(self.stop_processing)

        # 文件列表面板信号
        self.task_manager.taskCompleted.connect(self._on_task_item_completed)
        self.task_manager.taskFailed.connect(self._on_task_item_failed)

    def _on_task_item_completed(self, index, task):
        """单个任务完成时更新文件状态"""
        self._update_file_status(task.file_path, True)

    def _on_task_item_failed(self, index, error):
        """单个任务失败时更新文件状态"""
        if index < len(self.task_manager.tasks):
            task = self.task_manager.tasks[index]
            self._update_file_status(task.file_path, False, error)

    def _update_file_status(self, file_path, is_success, error=""):
        """更新文件列表中的文件状态"""
        self.main_window.file_list_widget.update_file_status(file_path, is_success, error)

    def update_rules(self, rules):
        """更新规则列表"""
        self.current_rules = rules

    def start_processing(self, output_file, append_mode, skip_file_info=False):
        """开始处理任务"""
        # 获取文件列表
        files = self.main_window.file_list_widget.get_all_files()
        if not files:
            self.main_window.status_bar.showMessage("没有文件可处理")
            return

        # 获取规则列表
        rules = self.main_window.rule_list_widget.get_enabled_rules()
        if not rules:
            self.main_window.status_bar.showMessage("没有启用的提取规则")
            return

        # 创建任务
        tasks = self.task_manager.create_tasks(files, rules)

        # 开始处理
        self.main_window.status_bar.showMessage("开始处理任务...")
        self.task_manager.start_processing(output_file, append_mode, skip_file_info)

    def stop_processing(self):
        """停止处理任务"""
        if self.task_manager.stop_processing():
            self.main_window.status_bar.showMessage("已停止处理任务")

    def is_processing(self):
        """检查是否正在处理任务"""
        current, total = self.task_manager.get_progress()
        return current < total and total > 0

    def _update_progress(self, current, total):
        """更新进度"""
        self.main_window.task_panel.update_progress(current, total)
        self.main_window.status_bar.showMessage(f"处理中... {current}/{total}")

    def _on_tasks_completed(self):
        """任务完成回调"""
        stats = self.task_manager.get_statistics()

        self.processingFinished.emit()
        self.main_window.status_bar.showMessage("任务处理完成")

        self.main_window.task_panel.processing_completed(
            stats["completed"],
            stats["total"]
        )

    def _on_task_error(self, error):
        """任务错误回调"""
        self.processingFinished.emit()
        self.main_window.status_bar.showMessage(f"任务处理出错: {error}")

        self.main_window.task_panel.processing_failed(error)
