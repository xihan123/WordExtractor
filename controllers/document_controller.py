#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文档控制器 - 处理文档内容的业务逻辑
"""

from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtWidgets import QMessageBox

from models.document_model import DocumentManager


class DocumentController(QObject):
    """文档控制器"""

    documentLoaded = pyqtSignal(object)  # 加载的文档对象

    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.document_manager = DocumentManager()

        self._connect_signals()

    def _connect_signals(self):
        """连接信号和槽"""
        self.document_manager.documentLoaded.connect(self._on_document_loaded)
        self.document_manager.loadError.connect(self._on_load_error)

    def load_document(self, file_path):
        """加载文档"""
        if not file_path:
            return

        # 更新UI状态
        self.main_window.status_bar.showMessage(f"正在加载文档: {file_path}")

        # 异步加载文档
        self.document_manager.load_document(file_path)

    def _on_document_loaded(self, document):
        """文档加载完成回调"""
        if not document:
            return

        # 更新UI状态
        self.main_window.status_bar.showMessage(f"文档已加载: {document.file_path}")

        # 更新文档预览
        self.main_window.document_viewer.load_document(document)

        # 发送文档加载信号
        self.documentLoaded.emit(document)

    def _on_load_error(self, error):
        """文档加载错误回调"""
        self.main_window.status_bar.showMessage(f"加载文档失败: {error}")

        QMessageBox.warning(
            self.main_window,
            "加载失败",
            f"无法加载文档: {error}"
        )

    def get_current_document(self):
        """获取当前文档"""
        return self.document_manager.get_current_document()

    def test_rule_on_document(self, rule):
        """在当前文档上测试规则"""
        from utils.docx_parser import DocxParser
        from views.rule_test_dialog import RuleTestDialog

        document = self.get_current_document()
        if not document or not document.is_loaded:
            QMessageBox.warning(
                self.main_window,
                "无法测试",
                "请先加载文档"
            )
            return

        try:
            # 创建文档解析器
            parser = DocxParser(document.file_path)

            # 应用规则提取数据
            from models.extraction_rule import ExtractionMode

            result = None

            if rule.rule_type == ExtractionMode.REGEX:
                pattern = rule.config.get("pattern", "")
                group = rule.config.get("group", 0)
                result = parser.extract_with_regex(pattern, group)

            elif rule.rule_type == ExtractionMode.POSITION:
                start_index = rule.config.get("start_index", 0)
                end_index = rule.config.get("end_index")
                result = parser.extract_by_position(start_index, end_index)

            elif rule.rule_type == ExtractionMode.BOOKMARK:
                bookmark_name = rule.config.get("bookmark_name", "")
                result = parser.extract_by_bookmark(bookmark_name)

            elif rule.rule_type == ExtractionMode.TABLE_CELL:
                table_index = rule.config.get("table_index", 0)
                row_index = rule.config.get("row_index", 0)
                col_index = rule.config.get("column_index", 0)
                result = parser.extract_table_cell(table_index, row_index, col_index)

            elif rule.rule_type == ExtractionMode.TABLE_COLUMN:
                table_index = rule.config.get("table_index", 0)
                col_index = rule.config.get("column_index", 0)
                has_header = rule.config.get("has_header", True)
                result = parser.extract_table_column(table_index, col_index, has_header)

            elif rule.rule_type == ExtractionMode.TABLE_ROW:
                table_index = rule.config.get("table_index", 0)
                row_index = rule.config.get("row_index", 0)
                result = parser.extract_table_row(table_index, row_index)

            elif rule.rule_type == ExtractionMode.TABLE_FULL:
                table_index = rule.config.get("table_index", 0)
                has_header = rule.config.get("has_header", True)
                result = parser.extract_table(table_index, has_header)

            # 显示测试结果对话框
            dialog = RuleTestDialog(rule, result, self.main_window)
            dialog.exec()

        except Exception as e:
            QMessageBox.warning(
                self.main_window,
                "测试失败",
                f"规则测试失败: {str(e)}"
            )
