#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
规则测试对话框 - 显示规则测试结果
"""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QLabel, QPushButton,
                             QTextEdit, QHBoxLayout)


class RuleTestDialog(QDialog):
    """规则测试结果对话框"""

    def __init__(self, rule, result, parent=None):
        super().__init__(parent)
        self.rule = rule
        self.result = result

        self._init_ui()

    def _init_ui(self):
        """初始化界面"""
        self.setWindowTitle("规则测试结果")
        self.setMinimumSize(500, 400)

        layout = QVBoxLayout(self)

        # 规则信息
        rule_info_label = QLabel(f"<b>规则:</b> {self.rule.field_name} ({self.rule.type_name})")
        rule_info_label.setWordWrap(True)
        layout.addWidget(rule_info_label)

        # 规则详细配置
        config_label = QLabel(f"<b>配置:</b> {self.rule.get_config_summary()}")
        config_label.setWordWrap(True)
        layout.addWidget(config_label)

        # 分隔线
        separator = QLabel()
        separator.setFrameShape(QLabel.Shape.HLine)
        separator.setFrameShadow(QLabel.Shadow.Sunken)
        layout.addWidget(separator)

        # 结果标题
        result_title = QLabel("<b>测试结果:</b>")
        layout.addWidget(result_title)

        # 结果文本区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)

        # 根据结果类型和内容设置不同的显示方式
        if isinstance(self.result, list):
            if len(self.result) > 0:
                if isinstance(self.result[0], list):
                    # 表格数据
                    html = "<table border='1' cellspacing='0' cellpadding='3'>"
                    for row in self.result:
                        html += "<tr>"
                        for cell in row:
                            html += f"<td>{cell}</td>"
                        html += "</tr>"
                    html += "</table>"
                    self.result_text.setHtml(html)
                else:
                    # 列表数据
                    self.result_text.setPlainText("\n".join(str(item) for item in self.result))
            else:
                self.result_text.setPlainText("(空结果)")
        else:
            # 字符串或其他类型
            self.result_text.setPlainText(str(self.result))

        layout.addWidget(self.result_text, 1)  # 1是拉伸因子

        # 按钮布局
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)
