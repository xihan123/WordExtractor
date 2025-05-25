#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
规则控制器 - 处理提取规则的业务逻辑
"""

from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtWidgets import QFileDialog


class RuleController(QObject):
    """规则控制器"""

    rulesUpdated = pyqtSignal(list)  # 更新后的规则列表

    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent

    def create_new_rule(self):
        """创建新规则"""
        self.main_window.rule_list_widget._add_rule()

    def create_rule_from_selection(self, text):
        """从选中文本创建规则"""
        if not text:
            return

        # 尝试智能识别数据类型
        self.main_window.rule_list_widget.create_rule_from_selection(text)

    def create_rule_from_table(self, table_index, row_index, col_index):
        """从表格创建规则"""
        self.main_window.rule_list_widget.create_rule_from_table(table_index, row_index, col_index)

    def import_rules(self):
        """导入规则"""
        file_path, _ = QFileDialog.getOpenFileName(
            self.main_window,
            "导入规则",
            "",
            "JSON文件 (*.json);;所有文件 (*.*)"
        )

        if file_path:
            success = self.main_window.rule_list_widget.import_rules(file_path)
            if success:
                self.rulesUpdated.emit(self.main_window.rule_list_widget.get_all_rules())

    def export_rules(self):
        """导出规则"""
        file_path, _ = QFileDialog.getSaveFileName(
            self.main_window,
            "导出规则",
            "",
            "JSON文件 (*.json);;CSV文件 (*.csv);;所有文件 (*.*)"
        )

        if file_path:
            self.main_window.rule_list_widget.export_rules(file_path)

    def clear_rules(self):
        """清空规则"""
        success = self.main_window.rule_list_widget.clear_rules()
        if success:
            self.rulesUpdated.emit([])

    def update_rules(self, rules):
        """更新规则列表"""
        self.rulesUpdated.emit(rules)
