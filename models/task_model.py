#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
任务处理模型 - 处理批量提取任务
"""

import os
from datetime import datetime
from enum import Enum

from PyQt6.QtCore import QObject, pyqtSignal, QRunnable, QThreadPool

from models.extraction_rule import ExtractionMode
from utils.docx_parser import DocxParser
from utils.excel_exporter import ExcelExporter


class TaskStatus(Enum):
    """任务状态枚举"""
    PENDING = "待处理"
    PROCESSING = "处理中"
    COMPLETED = "已完成"
    FAILED = "失败"
    CANCELED = "已取消"


class ExtractionTask:
    """单个文件的提取任务"""

    def __init__(self, file_path, rules=None):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.rules = rules or []
        self.status = TaskStatus.PENDING
        self.start_time = None
        self.end_time = None
        self.error = ""
        self.extracted_data = {}  # 提取的数据，键为字段名，值为提取结果

    def start(self):
        """开始任务"""
        self.start_time = datetime.now()
        self.status = TaskStatus.PROCESSING

    def complete(self, data=None):
        """完成任务"""
        self.end_time = datetime.now()
        self.status = TaskStatus.COMPLETED
        if data:
            self.extracted_data = data

    def fail(self, error):
        """任务失败"""
        self.end_time = datetime.now()
        self.status = TaskStatus.FAILED
        self.error = str(error)

    def cancel(self):
        """取消任务"""
        self.end_time = datetime.now()
        self.status = TaskStatus.CANCELED

    @property
    def duration(self):
        """任务持续时间（秒）"""
        if self.start_time is None:
            return 0
        end = self.end_time or datetime.now()
        return (end - self.start_time).total_seconds()


class BatchExtractionWorker(QRunnable):
    """批量提取工作线程"""

    class Signals(QObject):
        """工作线程信号"""
        started = pyqtSignal()
        progress = pyqtSignal(int, int)  # current, total
        taskCompleted = pyqtSignal(int, ExtractionTask)  # index, task
        taskFailed = pyqtSignal(int, str)  # index, error
        completed = pyqtSignal(list)  # all tasks
        error = pyqtSignal(str)

    def __init__(self, tasks, output_file=None, append_mode=False, skip_file_info=False):
        super().__init__()
        self.tasks = tasks
        self.output_file = output_file
        self.append_mode = append_mode
        self.skip_file_info = skip_file_info
        self.should_stop = False
        self.signals = self.Signals()

    def run(self):
        """线程执行函数"""
        try:
            self.signals.started.emit()

            # 创建Excel导出器
            exporter = ExcelExporter()
            if self.output_file:
                exporter.set_output_file(self.output_file, self.append_mode)

            # 对每个任务进行处理
            total = len(self.tasks)

            # 初始进度
            self.signals.progress.emit(0, total)

            for i, task in enumerate(self.tasks):
                if self.should_stop:
                    task.cancel()
                    self.signals.taskFailed.emit(i, "任务已取消")
                    continue

                try:
                    # 处理任务
                    task.start()

                    # 解析Word文档
                    parser = DocxParser(task.file_path)
                    result = {}

                    # 应用每条规则
                    for rule in task.rules:
                        if not rule.enabled:
                            continue

                        value = self._apply_rule(parser, rule)
                        result[rule.header_name] = value

                    # 添加文件路径 (如果未设置跳过)
                    if not self.skip_file_info:
                        result["文件名"] = os.path.basename(task.file_path)
                        result["文件路径"] = task.file_path

                    # 标记任务完成
                    task.complete(result)

                    # 添加到Excel
                    if self.output_file:
                        exporter.add_row(result)

                    self.signals.taskCompleted.emit(i, task)

                except Exception as e:
                    # 任务处理失败
                    task.fail(str(e))
                    self.signals.taskFailed.emit(i, str(e))

                # 在每个任务完成后更新进度为i+1
                self.signals.progress.emit(i + 1, total)

            # 保存Excel
            if self.output_file:
                exporter.save()

            # 确保最终进度为100%
            self.signals.progress.emit(total, total)

            self.signals.completed.emit(self.tasks)

        except Exception as e:
            self.signals.error.emit(f"批量处理任务出错: {str(e)}")

    def _apply_rule(self, parser, rule):
        """应用提取规则"""
        if rule.rule_type == ExtractionMode.REGEX:
            pattern = rule.config.get("pattern", "")
            group = rule.config.get("group", 0)
            return parser.extract_with_regex(pattern, group)

        elif rule.rule_type == ExtractionMode.POSITION:
            start_index = rule.config.get("start_index", 0)
            end_index = rule.config.get("end_index")
            return parser.extract_by_position(start_index, end_index)

        elif rule.rule_type == ExtractionMode.BOOKMARK:
            bookmark_name = rule.config.get("bookmark_name", "")
            return parser.extract_by_bookmark(bookmark_name)

        elif rule.rule_type == ExtractionMode.TABLE_CELL:
            table_index = rule.config.get("table_index", 0)
            row_index = rule.config.get("row_index", 0)
            col_index = rule.config.get("column_index", 0)
            return parser.extract_table_cell(table_index, row_index, col_index)

        elif rule.rule_type == ExtractionMode.TABLE_COLUMN:
            table_index = rule.config.get("table_index", 0)
            col_index = rule.config.get("column_index", 0)
            has_header = rule.config.get("has_header", True)
            return parser.extract_table_column(table_index, col_index, has_header)

        elif rule.rule_type == ExtractionMode.TABLE_ROW:
            table_index = rule.config.get("table_index", 0)
            row_index = rule.config.get("row_index", 0)
            return parser.extract_table_row(table_index, row_index)

        elif rule.rule_type == ExtractionMode.TABLE_FULL:
            table_index = rule.config.get("table_index", 0)
            has_header = rule.config.get("has_header", True)
            return parser.extract_table(table_index, has_header)

        return None

    def stop(self):
        """停止处理"""
        self.should_stop = True


class TaskManager(QObject):
    """任务管理器"""

    taskStarted = pyqtSignal()
    taskProgress = pyqtSignal(int, int)  # current, total
    taskItemCompleted = pyqtSignal(int, bool)  # index, success
    taskCompleted = pyqtSignal(int, ExtractionTask)  # index, task
    taskFailed = pyqtSignal(int, str)  # index, error
    allTasksCompleted = pyqtSignal()
    taskError = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread_pool = QThreadPool()
        self.thread_pool.setMaxThreadCount(4)  # 限制最大线程数
        self.worker = None
        self.tasks = []

    def create_tasks(self, file_paths, rules):
        """创建批处理任务"""
        self.tasks = []
        for path in file_paths:
            task = ExtractionTask(path, rules)
            self.tasks.append(task)
        return self.tasks

    def start_processing(self, output_file=None, append_mode=False, skip_file_info=False):
        """开始处理任务"""
        if not self.tasks:
            self.taskError.emit("没有任务可处理")
            return False

        # 创建工作线程
        self.worker = BatchExtractionWorker(self.tasks, output_file, append_mode, skip_file_info)

        # 连接信号
        self.worker.signals.started.connect(self.taskStarted)
        self.worker.signals.progress.connect(self.taskProgress)
        self.worker.signals.taskCompleted.connect(self.taskCompleted)  # 新增连接
        self.worker.signals.taskFailed.connect(self.taskFailed)  # 新增连接
        self.worker.signals.completed.connect(self._on_all_completed)
        self.worker.signals.error.connect(self.taskError)

        # 启动工作线程
        self.thread_pool.start(self.worker)
        return True

    def stop_processing(self):
        """停止处理任务"""
        if self.worker:
            self.worker.stop()
            return True
        return False

    def _on_task_completed(self, index, task):
        """单个任务完成"""
        self.taskItemCompleted.emit(index, True)

    def _on_task_failed(self, index, error):
        """单个任务失败"""
        self.taskItemCompleted.emit(index, False)

    def _on_all_completed(self, tasks):
        """所有任务完成"""
        self.allTasksCompleted.emit()

    def get_progress(self):
        """获取进度信息"""
        total = len(self.tasks)
        completed = sum(1 for task in self.tasks
                        if task.status in [TaskStatus.COMPLETED, TaskStatus.FAILED, TaskStatus.CANCELED])
        return completed, total

    def get_statistics(self):
        """获取任务统计信息"""
        total = len(self.tasks)
        completed = sum(1 for task in self.tasks if task.status == TaskStatus.COMPLETED)
        failed = sum(1 for task in self.tasks if task.status == TaskStatus.FAILED)
        canceled = sum(1 for task in self.tasks if task.status == TaskStatus.CANCELED)
        pending = sum(1 for task in self.tasks if task.status == TaskStatus.PENDING)

        return {
            "total": total,
            "completed": completed,
            "failed": failed,
            "canceled": canceled,
            "pending": pending,
            "success_rate": completed / total if total > 0 else 0
        }
