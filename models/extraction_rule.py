#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
提取规则模型 - 定义和管理数据提取规则
"""

import csv
import json
import uuid
from enum import Enum

from PyQt6.QtCore import Qt, QAbstractListModel, QModelIndex, pyqtSignal, QObject
from PyQt6.QtGui import QColor, QFont, QIcon


class ExtractionMode(Enum):
    """提取模式枚举"""
    REGEX = "正则表达式"
    POSITION = "位置索引"
    BOOKMARK = "文档书签"
    TABLE_CELL = "表格单元格"
    TABLE_COLUMN = "表格列"
    TABLE_ROW = "表格行"
    TABLE_FULL = "完整表格"


class ExtractionRule:
    """提取规则类"""

    def __init__(self, field_name="", rule_type=ExtractionMode.REGEX, config=None):
        self.id = str(uuid.uuid4())
        self.field_name = field_name
        self.header_name = field_name  # 导出到Excel时的表头名称
        self.rule_type = rule_type
        self.enabled = True
        self.config = config or {}
        self.description = ""

    @property
    def type_name(self):
        """获取规则类型名称"""
        return self.rule_type.value if isinstance(self.rule_type, ExtractionMode) else str(self.rule_type)

    def to_dict(self):
        """将规则转换为字典"""
        return {
            "id": self.id,
            "field_name": self.field_name,
            "header_name": self.header_name,
            "rule_type": self.rule_type.name if isinstance(self.rule_type, ExtractionMode) else str(self.rule_type),
            "enabled": self.enabled,
            "config": self.config,
            "description": self.description
        }

    @classmethod
    def from_dict(cls, data):
        """从字典创建规则"""
        try:
            rule_type = ExtractionMode[data["rule_type"]] if "rule_type" in data else ExtractionMode.REGEX
        except (KeyError, ValueError):
            rule_type = ExtractionMode.REGEX

        rule = cls(
            field_name=data.get("field_name", ""),
            rule_type=rule_type,
            config=data.get("config", {})
        )

        rule.id = data.get("id", str(uuid.uuid4()))
        rule.header_name = data.get("header_name", rule.field_name)
        rule.enabled = data.get("enabled", True)
        rule.description = data.get("description", "")

        return rule

    def clone(self):
        """克隆规则"""
        new_rule = ExtractionRule(
            field_name=self.field_name,
            rule_type=self.rule_type,
            config=self.config.copy()
        )
        new_rule.header_name = self.header_name
        new_rule.enabled = self.enabled
        new_rule.description = self.description
        # 生成新ID
        return new_rule

    def get_config_summary(self):
        """获取配置摘要"""
        if self.rule_type == ExtractionMode.REGEX:
            return f"正则: {self.config.get('pattern', '未设置')}"
        elif self.rule_type == ExtractionMode.POSITION:
            start = self.config.get('start_index', '?')
            end = self.config.get('end_index', '?')
            return f"位置: {start}-{end}"
        elif self.rule_type == ExtractionMode.BOOKMARK:
            return f"书签: {self.config.get('bookmark_name', '未设置')}"
        elif self.rule_type == ExtractionMode.TABLE_CELL:
            table = self.config.get('table_index', 0)
            row = self.config.get('row_index', 0)
            col = self.config.get('column_index', 0)
            return f"表格 {table + 1}, 单元格: [{row + 1},{col + 1}]"
        elif self.rule_type == ExtractionMode.TABLE_COLUMN:
            table = self.config.get('table_index', 0)
            col = self.config.get('column_index', 0)
            return f"表格 {table + 1}, 列: {col + 1}"
        elif self.rule_type == ExtractionMode.TABLE_ROW:
            table = self.config.get('table_index', 0)
            row = self.config.get('row_index', 0)
            return f"表格 {table + 1}, 行: {row + 1}"
        elif self.rule_type == ExtractionMode.TABLE_FULL:
            table = self.config.get('table_index', 0)
            has_header = self.config.get('has_header', True)
            return f"表格 {table + 1}, " + ("含表头" if has_header else "无表头")
        else:
            return "未知配置"


class ExtractionRuleModel(QAbstractListModel):
    """提取规则列表模型"""

    ruleOrderChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.rules = []

    def rowCount(self, parent=QModelIndex()):
        """返回行数"""
        return len(self.rules)

    def data(self, index, role):
        """返回数据"""
        if not index.isValid() or index.row() >= len(self.rules):
            return None

        rule = self.rules[index.row()]

        if role == Qt.ItemDataRole.DisplayRole:
            return f"{rule.field_name} ({rule.type_name})"
        elif role == Qt.ItemDataRole.ToolTipRole:
            return f"{rule.field_name}: {rule.get_config_summary()}"
        elif role == Qt.ItemDataRole.UserRole:
            return rule
        # 添加文本颜色和字体样式特征表示启用/禁用状态
        elif role == Qt.ItemDataRole.ForegroundRole:
            # 禁用规则使用灰色文本
            if not rule.enabled:
                return QColor("#888888")
        elif role == Qt.ItemDataRole.FontRole:
            font = QFont()
            # 启用规则使用粗体
            if rule.enabled:
                font.setBold(True)
            else:
                font.setItalic(True)  # 禁用规则使用斜体
            return font
        elif role == Qt.ItemDataRole.DecorationRole:
            # 添加启用/禁用图标
            if rule.enabled:
                return QIcon.fromTheme("dialog-ok", QIcon())  # 启用图标
            else:
                return QIcon.fromTheme("dialog-cancel", QIcon())  # 禁用图标

        return None

    def add_rule(self, rule):
        """添加规则"""
        row = len(self.rules)
        self.beginInsertRows(QModelIndex(), row, row)
        self.rules.append(rule)
        self.endInsertRows()
        return True

    def update_rule(self, index, rule):
        """更新规则"""
        if not isinstance(index, int):
            index = index.row()

        if 0 <= index < len(self.rules):
            self.rules[index] = rule
            model_index = self.index(index, 0)
            self.dataChanged.emit(model_index, model_index)
            return True
        return False

    def remove_rule(self, index):
        """移除规则"""
        if not isinstance(index, int):
            index = index.row()

        if 0 <= index < len(self.rules):
            self.beginRemoveRows(QModelIndex(), index, index)
            self.rules.pop(index)
            self.endRemoveRows()
            return True
        return False

    def move_rule_up(self, index):
        """上移规则"""
        if not isinstance(index, int):
            index = index.row()

        if 0 < index < len(self.rules):
            self.beginMoveRows(QModelIndex(), index, index, QModelIndex(), index - 1)
            self.rules[index], self.rules[index - 1] = self.rules[index - 1], self.rules[index]
            self.endMoveRows()
            self.ruleOrderChanged.emit()
            return True
        return False

    def move_rule_down(self, index):
        """下移规则"""
        if not isinstance(index, int):
            index = index.row()

        if 0 <= index < len(self.rules) - 1:
            self.beginMoveRows(QModelIndex(), index, index, QModelIndex(), index + 2)
            self.rules[index], self.rules[index + 1] = self.rules[index + 1], self.rules[index]
            self.endMoveRows()
            self.ruleOrderChanged.emit()
            return True
        return False

    def get_rule(self, index):
        """获取规则"""
        if not isinstance(index, int):
            index = index.row()

        if 0 <= index < len(self.rules):
            return self.rules[index]
        return None

    def get_all_rules(self):
        """获取所有规则"""
        return self.rules.copy()

    def get_enabled_rules(self):
        """获取所有启用的规则"""
        return [rule for rule in self.rules if rule.enabled]

    def clear(self):
        """清空所有规则"""
        self.beginResetModel()
        self.rules.clear()
        self.endResetModel()

    def save_to_json(self, file_path):
        """将规则保存到JSON文件"""
        rules_data = [rule.to_dict() for rule in self.rules]
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(rules_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存规则失败: {e}")
            return False

    def load_from_json(self, file_path):
        """从JSON文件加载规则"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                rules_data = json.load(f)

            self.beginResetModel()
            self.rules = [ExtractionRule.from_dict(data) for data in rules_data]
            self.endResetModel()
            return True
        except Exception as e:
            print(f"加载规则失败: {e}")
            return False

    def save_to_csv(self, file_path):
        """将规则保存到CSV文件"""
        try:
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                # 写入表头
                writer.writerow(["字段名称", "表头名称", "规则类型", "规则配置", "启用状态", "描述"])

                # 写入规则数据
                for rule in self.rules:
                    writer.writerow([
                        rule.field_name,
                        rule.header_name,
                        rule.type_name,
                        json.dumps(rule.config, ensure_ascii=False),
                        "是" if rule.enabled else "否",
                        rule.description
                    ])
            return True
        except Exception as e:
            print(f"保存规则到CSV失败: {e}")
            return False


class RuleManager(QObject):
    """规则管理器"""

    ruleAdded = pyqtSignal(ExtractionRule)
    ruleUpdated = pyqtSignal(ExtractionRule)
    ruleRemoved = pyqtSignal(int)
    rulesLoaded = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.model = ExtractionRuleModel()

    def add_rule(self, rule):
        """添加规则"""
        if self.model.add_rule(rule):
            self.ruleAdded.emit(rule)
            return True
        return False

    def update_rule(self, index, rule):
        """更新规则"""
        if self.model.update_rule(index, rule):
            self.ruleUpdated.emit(rule)
            return True
        return False

    def remove_rule(self, index):
        """删除规则"""
        if self.model.remove_rule(index):
            self.ruleRemoved.emit(index if isinstance(index, int) else index.row())
            return True
        return False

    def get_rule(self, index):
        """获取规则"""
        return self.model.get_rule(index)

    def create_rule_from_selection(self, text, rule_type=ExtractionMode.REGEX, field_name="", config=None):
        """从选中文本创建规则"""
        if not text:
            return None

        # 如果没有提供字段名，尝试猜测
        if not field_name:
            # 取选中文本的前10个字符作为字段名
            field_name = text[:10]
            if len(text) > 10:
                field_name += "..."

        # 根据规则类型创建默认配置
        if config is None:
            config = {}

        if rule_type == ExtractionMode.REGEX:
            if "pattern" not in config:
                # 将选中文本作为要匹配的字面文本，转义特殊字符
                escaped_text = text.replace('\\', '\\\\').replace('.', '\\.').replace('*', '\\*')
                escaped_text = escaped_text.replace('[', '\\[').replace(']', '\\]').replace('+', '\\+')
                config["pattern"] = f"({escaped_text})"

        elif rule_type == ExtractionMode.POSITION:
            # 位置索引需要start_index和end_index
            pass

        # 创建规则
        rule = ExtractionRule(field_name=field_name, rule_type=rule_type, config=config)
        rule.header_name = field_name

        return rule

    def identify_data_type(self, text):
        """智能识别数据类型"""
        import re

        # 清理文本
        text = text.strip()

        # 身份证号码 (18位)
        if re.match(r'^\d{17}[\dXx]$', text):
            return "身份证号", ExtractionMode.REGEX, {"pattern": r"(\d{17}[\dXx])"}

        # 手机号码 (11位)
        if re.match(r'^1\d{10}$', text):
            return "手机号", ExtractionMode.REGEX, {"pattern": r"(1\d{10})"}

        # 邮箱
        if re.match(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', text):
            return "邮箱", ExtractionMode.REGEX, {"pattern": r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)"}

        # 日期 (多种格式)
        if re.match(r'^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?$', text):
            return "日期", ExtractionMode.REGEX, {"pattern": r"(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?)"}

        # 金额 (数字+可能的小数点)
        if re.match(r'^\d+(\.\d+)?元?$', text):
            return "金额", ExtractionMode.REGEX, {"pattern": r"(\d+(?:\.\d+)?元?)"}

        # 地址 (中文字符+数字+可能有单元楼栋等)
        if len(text) > 10 and re.search(r'[省市区县道路街号]', text):
            return "地址", ExtractionMode.REGEX, {"pattern": re.escape(text)}

        # 姓名 (2-4个中文字符)
        if re.match(r'^[\u4e00-\u9fa5]{2,4}$', text):
            return "姓名", ExtractionMode.REGEX, {"pattern": r"([\u4e00-\u9fa5]{2,4})"}

        # 默认
        return text[:10] + ("..." if len(text) > 10 else ""), ExtractionMode.REGEX, {"pattern": re.escape(text)}

    def save_rules(self, file_path):
        """保存规则"""
        if file_path.lower().endswith('.json'):
            return self.model.save_to_json(file_path)
        elif file_path.lower().endswith('.csv'):
            return self.model.save_to_csv(file_path)
        else:
            # 默认使用JSON格式
            return self.model.save_to_json(file_path + ".json")

    def load_rules(self, file_path):
        """加载规则"""
        if file_path.lower().endswith('.json'):
            success = self.model.load_from_json(file_path)
        else:
            success = False

        if success:
            self.rulesLoaded.emit()

        return success
