#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
异步处理工具 - 处理异步任务
"""

from PyQt6.QtCore import QObject, pyqtSignal, QRunnable, QThreadPool


class AsyncWorkerSignals(QObject):
    """异步工作线程信号"""

    started = pyqtSignal()
    finished = pyqtSignal()
    result = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, int)  # current, total


class AsyncWorker(QRunnable):
    """异步工作线程"""

    def __init__(self, task_func, *args, **kwargs):
        super().__init__()

        self.signals = AsyncWorkerSignals()
        self.task_func = task_func
        self.args = args
        self.kwargs = kwargs
        self.should_stop = False

        # 如果有传递进度回调，则包装它
        if "progress_callback" in kwargs:
            self.kwargs["progress_callback"] = self._on_progress

    def run(self):
        """运行任务"""
        try:
            self.signals.started.emit()
            result = self.task_func(*self.args, **self.kwargs)
            self.signals.result.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))
        finally:
            self.signals.finished.emit()

    def stop(self):
        """停止任务"""
        self.should_stop = True

    def _on_progress(self, current, total):
        """进度回调"""
        self.signals.progress.emit(current, total)


class AsyncTaskManager:
    """异步任务管理器"""

    def __init__(self, max_threads=4):
        self.thread_pool = QThreadPool()
        self.thread_pool.setMaxThreadCount(max_threads)
        self.active_tasks = {}

    def run_task(self, task_id, task_func, *args, **kwargs):
        """运行异步任务"""
        # 如果已有同ID任务在运行，先停止它
        if task_id in self.active_tasks:
            self.active_tasks[task_id].stop()

        # 创建新任务
        worker = AsyncWorker(task_func, *args, **kwargs)
        self.active_tasks[task_id] = worker

        # 设置任务完成后的清理
        worker.signals.finished.connect(lambda: self._cleanup_task(task_id))

        # 启动任务
        self.thread_pool.start(worker)
        return worker.signals

    def stop_task(self, task_id):
        """停止指定任务"""
        if task_id in self.active_tasks:
            self.active_tasks[task_id].stop()
            return True
        return False

    def stop_all_tasks(self):
        """停止所有任务"""
        for task_id in list(self.active_tasks.keys()):
            self.stop_task(task_id)

    def is_task_running(self, task_id):
        """检查任务是否在运行"""
        return task_id in self.active_tasks

    def _cleanup_task(self, task_id):
        """清理完成的任务"""
        if task_id in self.active_tasks:
            del self.active_tasks[task_id]
