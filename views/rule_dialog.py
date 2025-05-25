#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
规则对话框 - 创建和编辑提取规则
"""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QFormLayout,
                             QLabel, QLineEdit, QComboBox, QSpinBox,
                             QCheckBox, QTabWidget,
                             QTextEdit, QGroupBox, QWidget,
                             QDialogButtonBox, QMessageBox)

from models.extraction_rule import ExtractionRule, ExtractionMode


class ExtractionRuleDialog(QDialog):
    """提取规则对话框"""

    def __init__(self, parent=None, rule=None):
        super().__init__(parent)

        self.rule = rule or ExtractionRule()
        self.result_rule = None

        self._init_ui()
        self._load_rule_data()

    def _init_ui(self):
        """初始化界面"""
        self.setWindowTitle("提取规则设置")
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)

        # 基本信息
        form_layout = QFormLayout()

        self.field_name_edit = QLineEdit()
        form_layout.addRow("字段名称:", self.field_name_edit)

        self.header_name_edit = QLineEdit()
        form_layout.addRow("表头名称:", self.header_name_edit)

        self.rule_type_combo = QComboBox()
        for mode in ExtractionMode:
            self.rule_type_combo.addItem(mode.value, mode)
        self.rule_type_combo.currentIndexChanged.connect(self._on_rule_type_changed)
        form_layout.addRow("提取方式:", self.rule_type_combo)

        self.enabled_checkbox = QCheckBox("启用此规则")
        form_layout.addRow("", self.enabled_checkbox)

        # 添加到主布局
        layout.addLayout(form_layout)

        # 规则配置选项卡
        self.tab_widget = QTabWidget()

        # 规则配置页
        self.config_tab = QWidget()
        self.config_layout = QVBoxLayout(self.config_tab)
        self.tab_widget.addTab(self.config_tab, "规则配置")

        # 在配置页上添加一个占位布局，后面会根据规则类型替换
        self.config_place_holder = QVBoxLayout()
        self.config_layout.addLayout(self.config_place_holder)

        # 高级选项页
        self.advanced_tab = QWidget()
        advanced_layout = QVBoxLayout(self.advanced_tab)

        self.description_edit = QTextEdit()
        self.description_edit.setPlaceholderText("输入规则描述...")
        advanced_layout.addWidget(QLabel("规则描述:"))
        advanced_layout.addWidget(self.description_edit)

        self.tab_widget.addTab(self.advanced_tab, "高级选项")

        # 添加到主布局
        layout.addWidget(self.tab_widget)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _load_rule_data(self):
        """加载规则数据到界面"""
        self.field_name_edit.setText(self.rule.field_name)
        self.header_name_edit.setText(self.rule.header_name)
        self.enabled_checkbox.setChecked(self.rule.enabled)
        self.description_edit.setText(self.rule.description)

        # 设置规则类型
        index = self.rule_type_combo.findData(self.rule.rule_type)
        if index >= 0:
            self.rule_type_combo.setCurrentIndex(index)

        # 加载配置界面
        self._on_rule_type_changed(self.rule_type_combo.currentIndex())

    def _on_rule_type_changed(self, index):
        """规则类型变化处理"""
        # 清除当前配置布局中的所有部件
        while self.config_place_holder.count():
            item = self.config_place_holder.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 获取选择的规则类型
        rule_type = self.rule_type_combo.itemData(index)

        # 根据规则类型创建不同的配置界面
        if rule_type == ExtractionMode.REGEX:
            self._create_regex_config()
        elif rule_type == ExtractionMode.POSITION:
            self._create_position_config()
        elif rule_type == ExtractionMode.BOOKMARK:
            self._create_bookmark_config()
        elif rule_type == ExtractionMode.TABLE_CELL:
            self._create_table_cell_config()
        elif rule_type == ExtractionMode.TABLE_COLUMN:
            self._create_table_column_config()
        elif rule_type == ExtractionMode.TABLE_ROW:
            self._create_table_row_config()
        elif rule_type == ExtractionMode.TABLE_FULL:
            self._create_table_full_config()

    def _create_regex_config(self):
        """创建正则表达式配置界面"""
        group = QGroupBox("正则表达式设置")
        layout = QFormLayout(group)

        self.pattern_edit = QLineEdit()
        self.pattern_edit.setText(self.rule.config.get("pattern", ""))
        self.pattern_edit.setPlaceholderText("输入正则表达式，如: (\\d+)")
        layout.addRow("匹配模式:", self.pattern_edit)

        self.group_spinbox = QSpinBox()
        self.group_spinbox.setMinimum(0)
        self.group_spinbox.setMaximum(10)
        self.group_spinbox.setValue(self.rule.config.get("group", 0))
        layout.addRow("分组索引:", self.group_spinbox)

        help_label = QLabel("提示: 使用括号()来创建捕获组，0表示整个匹配")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def _create_position_config(self):
        """创建位置索引配置界面"""
        group = QGroupBox("位置索引设置")
        layout = QFormLayout(group)

        self.start_index_spinbox = QSpinBox()
        self.start_index_spinbox.setMinimum(0)
        self.start_index_spinbox.setMaximum(9999)
        self.start_index_spinbox.setValue(self.rule.config.get("start_index", 0))
        layout.addRow("起始段落索引:", self.start_index_spinbox)

        self.end_index_spinbox = QSpinBox()
        self.end_index_spinbox.setMinimum(-1)
        self.end_index_spinbox.setMaximum(9999)
        self.end_index_spinbox.setValue(self.rule.config.get("end_index", -1))
        self.end_index_spinbox.setSpecialValueText("到末尾")
        layout.addRow("结束段落索引:", self.end_index_spinbox)

        help_label = QLabel("提示: 段落索引从0开始，-1表示到文档末尾")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def _create_bookmark_config(self):
        """创建书签配置界面"""
        group = QGroupBox("书签设置")
        layout = QFormLayout(group)

        self.bookmark_edit = QLineEdit()
        self.bookmark_edit.setText(self.rule.config.get("bookmark_name", ""))
        self.bookmark_edit.setPlaceholderText("输入书签名称")
        layout.addRow("书签名称:", self.bookmark_edit)

        help_label = QLabel("提示: 书签需要在Word文档中预先设置")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def _create_table_cell_config(self):
        """创建表格单元格配置界面"""
        group = QGroupBox("表格单元格设置")
        layout = QFormLayout(group)

        self.table_index_spinbox = QSpinBox()
        self.table_index_spinbox.setMinimum(0)
        self.table_index_spinbox.setMaximum(99)
        self.table_index_spinbox.setValue(self.rule.config.get("table_index", 0))
        layout.addRow("表格索引:", self.table_index_spinbox)

        self.row_index_spinbox = QSpinBox()
        self.row_index_spinbox.setMinimum(0)
        self.row_index_spinbox.setMaximum(999)
        self.row_index_spinbox.setValue(self.rule.config.get("row_index", 0))
        layout.addRow("行索引:", self.row_index_spinbox)

        self.column_index_spinbox = QSpinBox()
        self.column_index_spinbox.setMinimum(0)
        self.column_index_spinbox.setMaximum(999)
        self.column_index_spinbox.setValue(self.rule.config.get("column_index", 0))
        layout.addRow("列索引:", self.column_index_spinbox)

        help_label = QLabel("提示: 索引从0开始")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def _create_table_column_config(self):
        """创建表格列配置界面"""
        group = QGroupBox("表格列设置")
        layout = QFormLayout(group)

        self.table_index_spinbox = QSpinBox()
        self.table_index_spinbox.setMinimum(0)
        self.table_index_spinbox.setMaximum(99)
        self.table_index_spinbox.setValue(self.rule.config.get("table_index", 0))
        layout.addRow("表格索引:", self.table_index_spinbox)

        self.column_index_spinbox = QSpinBox()
        self.column_index_spinbox.setMinimum(0)
        self.column_index_spinbox.setMaximum(999)
        self.column_index_spinbox.setValue(self.rule.config.get("column_index", 0))
        layout.addRow("列索引:", self.column_index_spinbox)

        self.has_header_checkbox = QCheckBox("第一行为表头")
        self.has_header_checkbox.setChecked(self.rule.config.get("has_header", True))
        layout.addRow("", self.has_header_checkbox)

        help_label = QLabel("提示: 索引从0开始，启用表头会在结果中跳过第一行")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def _create_table_row_config(self):
        """创建表格行配置界面"""
        group = QGroupBox("表格行设置")
        layout = QFormLayout(group)

        self.table_index_spinbox = QSpinBox()
        self.table_index_spinbox.setMinimum(0)
        self.table_index_spinbox.setMaximum(99)
        self.table_index_spinbox.setValue(self.rule.config.get("table_index", 0))
        layout.addRow("表格索引:", self.table_index_spinbox)

        self.row_index_spinbox = QSpinBox()
        self.row_index_spinbox.setMinimum(0)
        self.row_index_spinbox.setMaximum(999)
        self.row_index_spinbox.setValue(self.rule.config.get("row_index", 0))
        layout.addRow("行索引:", self.row_index_spinbox)

        help_label = QLabel("提示: 索引从0开始")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def _create_table_full_config(self):
        """创建完整表格配置界面"""
        group = QGroupBox("完整表格设置")
        layout = QFormLayout(group)

        self.table_index_spinbox = QSpinBox()
        self.table_index_spinbox.setMinimum(0)
        self.table_index_spinbox.setMaximum(99)
        self.table_index_spinbox.setValue(self.rule.config.get("table_index", 0))
        layout.addRow("表格索引:", self.table_index_spinbox)

        self.has_header_checkbox = QCheckBox("第一行为表头")
        self.has_header_checkbox.setChecked(self.rule.config.get("has_header", True))
        layout.addRow("", self.has_header_checkbox)

        help_label = QLabel("提示: 索引从0开始，启用表头会将第一行作为列标题")
        help_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addRow("", help_label)

        self.config_place_holder.addWidget(group)

    def accept(self):
        """接受对话框"""
        # 检查字段名和表头名是否为空
        field_name = self.field_name_edit.text().strip()
        if not field_name:
            QMessageBox.warning(self, "输入错误", "字段名称不能为空")
            self.field_name_edit.setFocus()
            return

        header_name = self.header_name_edit.text().strip()
        if not header_name:
            header_name = field_name

        # 获取规则类型
        rule_type = self.rule_type_combo.currentData()

        # 根据规则类型收集配置
        config = {}

        if rule_type == ExtractionMode.REGEX:
            pattern = self.pattern_edit.text().strip()
            if not pattern:
                QMessageBox.warning(self, "输入错误", "正则表达式不能为空")
                self.pattern_edit.setFocus()
                return

            config["pattern"] = pattern
            config["group"] = self.group_spinbox.value()

        elif rule_type == ExtractionMode.POSITION:
            config["start_index"] = self.start_index_spinbox.value()
            config["end_index"] = self.end_index_spinbox.value()

        elif rule_type == ExtractionMode.BOOKMARK:
            bookmark_name = self.bookmark_edit.text().strip()
            if not bookmark_name:
                QMessageBox.warning(self, "输入错误", "书签名称不能为空")
                self.bookmark_edit.setFocus()
                return

            config["bookmark_name"] = bookmark_name

        elif rule_type == ExtractionMode.TABLE_CELL:
            config["table_index"] = self.table_index_spinbox.value()
            config["row_index"] = self.row_index_spinbox.value()
            config["column_index"] = self.column_index_spinbox.value()

        elif rule_type == ExtractionMode.TABLE_COLUMN:
            config["table_index"] = self.table_index_spinbox.value()
            config["column_index"] = self.column_index_spinbox.value()
            config["has_header"] = self.has_header_checkbox.isChecked()

        elif rule_type == ExtractionMode.TABLE_ROW:
            config["table_index"] = self.table_index_spinbox.value()
            config["row_index"] = self.row_index_spinbox.value()

        elif rule_type == ExtractionMode.TABLE_FULL:
            config["table_index"] = self.table_index_spinbox.value()
            config["has_header"] = self.has_header_checkbox.isChecked()

        # 创建规则对象
        self.result_rule = ExtractionRule(field_name, rule_type, config)
        self.result_rule.header_name = header_name
        self.result_rule.enabled = self.enabled_checkbox.isChecked()
        self.result_rule.description = self.description_edit.toPlainText()

        # 如果是编辑已有规则，保留原规则ID
        if self.rule and hasattr(self.rule, 'id'):
            self.result_rule.id = self.rule.id

        super().accept()

    def get_rule(self):
        """获取规则"""
        return self.result_rule
