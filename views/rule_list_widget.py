#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
规则列表视图 - 管理提取规则
"""

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QListView,
                             QPushButton, QLabel, QAbstractItemView,
                             QMenu, QMessageBox, QDialog)

from models.extraction_rule import RuleManager, ExtractionRule, ExtractionMode
from views.rule_dialog import ExtractionRuleDialog


class RuleListWidget(QWidget):
    """规则列表视图"""

    ruleSelected = pyqtSignal(ExtractionRule)  # 选中的规则
    ruleListChanged = pyqtSignal(list)  # 规则列表变化

    def __init__(self, parent=None):
        super().__init__(parent)
        self.rule_manager = RuleManager()

        self._init_ui()
        self._connect_signals()

    def _init_ui(self):
        """初始化界面"""
        # 主布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 标题
        header_layout = QHBoxLayout()

        title_label = QLabel("提取规则")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        self.test_rule_btn = QPushButton("测试规则")
        self.test_rule_btn.setEnabled(False)
        header_layout.addWidget(self.test_rule_btn)

        layout.addLayout(header_layout)

        # 规则列表
        self.list_view = QListView()
        self.list_view.setModel(self.rule_manager.model)
        self.list_view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        layout.addWidget(self.list_view)

        # 按钮布局
        button_layout = QHBoxLayout()

        self.add_btn = QPushButton("添加规则")
        button_layout.addWidget(self.add_btn)

        self.edit_btn = QPushButton("编辑规则")
        self.edit_btn.setEnabled(False)
        button_layout.addWidget(self.edit_btn)

        self.remove_btn = QPushButton("删除规则")
        self.remove_btn.setEnabled(False)
        button_layout.addWidget(self.remove_btn)

        layout.addLayout(button_layout)

        # 移动按钮布局
        move_layout = QHBoxLayout()

        self.move_up_btn = QPushButton("上移")
        self.move_up_btn.setEnabled(False)
        move_layout.addWidget(self.move_up_btn)

        self.move_down_btn = QPushButton("下移")
        self.move_down_btn.setEnabled(False)
        move_layout.addWidget(self.move_down_btn)

        self.toggle_btn = QPushButton("启用/禁用")
        self.toggle_btn.setEnabled(False)
        move_layout.addWidget(self.toggle_btn)

        layout.addLayout(move_layout)

        # 状态标签
        self.status_label = QLabel("0 条规则")
        layout.addWidget(self.status_label)

    def _connect_signals(self):
        """连接信号和槽"""
        # 按钮点击事件
        self.add_btn.clicked.connect(self._add_rule)
        self.edit_btn.clicked.connect(self._edit_rule)
        self.remove_btn.clicked.connect(self._remove_rule)
        self.move_up_btn.clicked.connect(self._move_rule_up)
        self.move_down_btn.clicked.connect(self._move_rule_down)
        self.toggle_btn.clicked.connect(self._toggle_rule)
        self.test_rule_btn.clicked.connect(self._test_rule)

        # 列表事件
        self.list_view.clicked.connect(self._on_rule_clicked)
        self.list_view.doubleClicked.connect(self._edit_rule)
        self.list_view.customContextMenuRequested.connect(self._show_context_menu)
        self.list_view.selectionModel().selectionChanged.connect(self._on_selection_changed)

        # 规则管理器事件
        self.rule_manager.ruleAdded.connect(lambda: self._update_status_and_notify())
        self.rule_manager.ruleUpdated.connect(lambda: self._update_status_and_notify())
        self.rule_manager.ruleRemoved.connect(lambda: self._update_status_and_notify())
        self.rule_manager.rulesLoaded.connect(lambda: self._update_status_and_notify())

    def _add_rule(self):
        """添加规则"""
        dialog = ExtractionRuleDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            rule = dialog.get_rule()
            if rule:
                self.rule_manager.add_rule(rule)

    def _edit_rule(self):
        """编辑规则"""
        selected = self.list_view.currentIndex()
        if not selected.isValid():
            return

        rule = self.rule_manager.get_rule(selected.row())
        if not rule:
            return

        dialog = ExtractionRuleDialog(self, rule)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_rule = dialog.get_rule()
            if updated_rule:
                self.rule_manager.update_rule(selected.row(), updated_rule)

    def _remove_rule(self):
        """删除规则"""
        selected = self.list_view.currentIndex()
        if not selected.isValid():
            return

        reply = QMessageBox.question(
            self,
            "删除规则",
            "确定要删除选中的规则吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.rule_manager.remove_rule(selected.row())

    def _move_rule_up(self):
        """上移规则"""
        selected = self.list_view.currentIndex()
        if not selected.isValid():
            return

        if self.rule_manager.model.move_rule_up(selected.row()):
            self.list_view.setCurrentIndex(self.rule_manager.model.index(selected.row() - 1, 0))
            self._update_status_and_notify()

    def _move_rule_down(self):
        """下移规则"""
        selected = self.list_view.currentIndex()
        if not selected.isValid():
            return

        if self.rule_manager.model.move_rule_down(selected.row()):
            self.list_view.setCurrentIndex(self.rule_manager.model.index(selected.row() + 1, 0))
            self._update_status_and_notify()

    def _toggle_rule(self):
        """切换规则启用状态"""
        selected = self.list_view.currentIndex()
        if not selected.isValid():
            return

        rule = self.rule_manager.get_rule(selected.row())
        if rule:
            rule.enabled = not rule.enabled
            self.rule_manager.update_rule(selected.row(), rule)

    def _test_rule(self):
        """测试规则"""
        selected = self.list_view.currentIndex()
        if not selected.isValid():
            return

        rule = self.rule_manager.get_rule(selected.row())
        if not rule:
            return

        # 发送规则选中信号，以便文档控制器可以测试它
        self.ruleSelected.emit(rule)

    def _on_rule_clicked(self, index):
        """规则点击事件"""
        if not index.isValid():
            return

        rule = self.rule_manager.get_rule(index.row())
        if rule:
            self.ruleSelected.emit(rule)

    def _on_selection_changed(self, selected, deselected):
        """选择变化事件"""
        has_selection = len(selected.indexes()) > 0

        self.edit_btn.setEnabled(has_selection)
        self.remove_btn.setEnabled(has_selection)
        self.test_rule_btn.setEnabled(has_selection)
        self.toggle_btn.setEnabled(has_selection)

        # 启用/禁用上移下移按钮
        if has_selection:
            row = selected.indexes()[0].row()
            rule_count = self.rule_manager.model.rowCount()

            self.move_up_btn.setEnabled(row > 0)
            self.move_down_btn.setEnabled(row < rule_count - 1)
        else:
            self.move_up_btn.setEnabled(False)
            self.move_down_btn.setEnabled(False)

    def _show_context_menu(self, position):
        """显示右键菜单"""
        selected = self.list_view.indexAt(position)
        if not selected.isValid():
            return

        rule = self.rule_manager.get_rule(selected.row())
        if not rule:
            return

        menu = QMenu(self)

        # 编辑规则
        edit_action = QAction("编辑规则", self)
        edit_action.triggered.connect(self._edit_rule)
        menu.addAction(edit_action)

        # 删除规则
        remove_action = QAction("删除规则", self)
        remove_action.triggered.connect(self._remove_rule)
        menu.addAction(remove_action)

        menu.addSeparator()

        # 启用/禁用规则
        toggle_text = "禁用规则" if rule.enabled else "启用规则"
        toggle_action = QAction(toggle_text, self)
        toggle_action.triggered.connect(self._toggle_rule)
        menu.addAction(toggle_action)

        menu.addSeparator()

        # 上移/下移规则
        move_up_action = QAction("上移", self)
        move_up_action.setEnabled(selected.row() > 0)
        move_up_action.triggered.connect(self._move_rule_up)
        menu.addAction(move_up_action)

        move_down_action = QAction("下移", self)
        move_down_action.setEnabled(selected.row() < self.rule_manager.model.rowCount() - 1)
        move_down_action.triggered.connect(self._move_rule_down)
        menu.addAction(move_down_action)

        menu.addSeparator()

        # 测试规则
        test_action = QAction("测试规则", self)
        test_action.triggered.connect(self._test_rule)
        menu.addAction(test_action)

        # 复制规则
        duplicate_action = QAction("复制规则", self)
        duplicate_action.triggered.connect(lambda: self._duplicate_rule(selected.row()))
        menu.addAction(duplicate_action)

        menu.exec(self.list_view.viewport().mapToGlobal(position))

    def _duplicate_rule(self, index):
        """复制规则"""
        rule = self.rule_manager.get_rule(index)
        if rule:
            new_rule = rule.clone()
            new_rule.field_name = f"{rule.field_name} (副本)"
            new_rule.header_name = f"{rule.header_name} (副本)"
            self.rule_manager.add_rule(new_rule)

    def _update_status_and_notify(self):
        """更新状态并通知规则变化"""
        rule_count = self.rule_manager.model.rowCount()
        self.status_label.setText(f"{rule_count} 条规则")

        # 通知规则列表变化
        rules = self.rule_manager.model.get_all_rules()
        self.ruleListChanged.emit(rules)

    def create_rule_from_selection(self, text, rule_type=ExtractionMode.REGEX):
        """从选中文本创建规则"""
        if not text:
            return

        # 尝试识别数据类型
        field_name, detected_type, config = self.rule_manager.identify_data_type(text)

        # 创建规则
        rule = ExtractionRule(field_name=field_name, rule_type=detected_type, config=config)
        rule.header_name = field_name

        # 显示规则编辑对话框
        dialog = ExtractionRuleDialog(self, rule)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_rule = dialog.get_rule()
            if new_rule:
                self.rule_manager.add_rule(new_rule)

    def create_rule_from_table(self, table_index, row_index, col_index):
        """从表格创建规则"""
        if table_index < 0:
            return

        # 根据参数确定规则类型
        if row_index < 0 and col_index < 0:
            # 整个表格
            rule_type = ExtractionMode.TABLE_FULL
            config = {
                "table_index": table_index,
                "has_header": True
            }
            field_name = f"表格{table_index + 1}"
        elif row_index < 0:
            # 表格列
            rule_type = ExtractionMode.TABLE_COLUMN
            config = {
                "table_index": table_index,
                "column_index": col_index,
                "has_header": True
            }
            field_name = f"表格{table_index + 1}_列{col_index + 1}"
        elif col_index < 0:
            # 表格行
            rule_type = ExtractionMode.TABLE_ROW
            config = {
                "table_index": table_index,
                "row_index": row_index
            }
            field_name = f"表格{table_index + 1}_行{row_index + 1}"
        else:
            # 单元格
            rule_type = ExtractionMode.TABLE_CELL
            config = {
                "table_index": table_index,
                "row_index": row_index,
                "column_index": col_index
            }
            field_name = f"表格{table_index + 1}_单元格{row_index + 1}_{col_index + 1}"

        # 创建规则
        rule = ExtractionRule(field_name=field_name, rule_type=rule_type, config=config)
        rule.header_name = field_name

        # 显示规则编辑对话框
        dialog = ExtractionRuleDialog(self, rule)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_rule = dialog.get_rule()
            if new_rule:
                self.rule_manager.add_rule(new_rule)

    def import_rules(self, file_path):
        """导入规则"""
        if not file_path:
            return False

        success = self.rule_manager.load_rules(file_path)
        if success:
            QMessageBox.information(
                self,
                "导入成功",
                f"成功导入规则列表"
            )
        else:
            QMessageBox.warning(
                self,
                "导入失败",
                "导入规则列表时出错"
            )

        return success

    def export_rules(self, file_path):
        """导出规则"""
        if not file_path:
            return False

        success = self.rule_manager.save_rules(file_path)
        if success:
            QMessageBox.information(
                self,
                "导出成功",
                f"成功导出规则列表到：\n{file_path}"
            )
        else:
            QMessageBox.warning(
                self,
                "导出失败",
                "导出规则列表时出错"
            )

        return success

    def clear_rules(self):
        """清空规则列表"""
        if self.rule_manager.model.rowCount() == 0:
            return

        reply = QMessageBox.question(
            self,
            "清空规则",
            "确定要清空所有规则吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.rule_manager.model.clear()
            self._update_status_and_notify()
            return True

        return False

    def get_all_rules(self):
        """获取所有规则"""
        return self.rule_manager.model.get_all_rules()

    def get_enabled_rules(self):
        """获取所有启用的规则"""
        return self.rule_manager.model.get_enabled_rules()

    def update_file_status(self, file_path, is_processed, error=""):
        """更新文件处理状态"""
        for i in range(self.file_manager.model.rowCount()):
            file_item = self.file_manager.model.get_file(i)
            if file_item and file_item.path == file_path:
                self.file_manager.model.update_file_status(i, is_processed, error)
                # 在模型更新后刷新视图
                view_index = self.proxy_model.mapFromSource(
                    self.file_manager.model.index(i, 3)
                )
                if view_index.isValid():
                    self.proxy_model.dataChanged.emit(view_index, view_index)
                return True
        return False
