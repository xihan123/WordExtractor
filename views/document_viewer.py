#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文档预览视图 - 显示Word文档内容
"""

import os

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QTextCursor, QAction
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTextBrowser,
                             QLabel, QPushButton, QComboBox, QTabWidget, QMenu, QMessageBox)


class DocumentViewer(QWidget):
    """文档预览视图"""

    textSelected = pyqtSignal(str)  # 选择的文本
    tableSelected = pyqtSignal(int, int, int)  # 表格索引, 行索引, 列索引

    def __init__(self, parent=None):
        super().__init__(parent)
        self.document = None

        self._init_ui()
        self._connect_signals()

    def _init_ui(self):
        """初始化界面"""
        # 主布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 文档信息区域
        info_layout = QHBoxLayout()

        self.title_label = QLabel("未加载文档")
        self.title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        info_layout.addWidget(self.title_label)

        info_layout.addStretch()

        self.metadata_btn = QPushButton("查看元数据")
        self.metadata_btn.setEnabled(False)
        info_layout.addWidget(self.metadata_btn)

        layout.addLayout(info_layout)

        # 创建选项卡
        self.tab_widget = QTabWidget()

        # 文本预览选项卡
        self.text_tab = QWidget()
        text_layout = QVBoxLayout(self.text_tab)
        text_layout.setContentsMargins(0, 0, 0, 0)

        self.text_browser = QTextBrowser()
        self.text_browser.setOpenLinks(False)
        self.text_browser.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        text_layout.addWidget(self.text_browser)

        self.tab_widget.addTab(self.text_tab, "文本内容")

        # 表格预览选项卡
        self.tables_tab = QWidget()
        tables_layout = QVBoxLayout(self.tables_tab)
        tables_layout.setContentsMargins(0, 0, 0, 0)

        tables_header = QHBoxLayout()

        tables_label = QLabel("选择表格:")
        tables_header.addWidget(tables_label)

        self.table_combo = QComboBox()
        self.table_combo.setMinimumWidth(150)
        tables_header.addWidget(self.table_combo)

        tables_header.addStretch()

        tables_layout.addLayout(tables_header)

        self.table_browser = QTextBrowser()
        self.table_browser.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        tables_layout.addWidget(self.table_browser)

        self.tab_widget.addTab(self.tables_tab, "表格内容")

        layout.addWidget(self.tab_widget)

    def _connect_signals(self):
        """连接信号和槽"""
        # 元数据按钮
        self.metadata_btn.clicked.connect(self._show_metadata)

        # 右键菜单
        self.text_browser.customContextMenuRequested.connect(self._show_text_context_menu)
        self.table_browser.customContextMenuRequested.connect(self._show_table_context_menu)

        # 表格选择下拉框
        self.table_combo.currentIndexChanged.connect(self._update_table_view)

        # 选项卡切换
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

    def load_document(self, document):
        """加载文档"""
        self.document = document

        if not document or not document.is_loaded:
            self._clear_content()
            self.title_label.setText("加载文档失败")
            QMessageBox.warning(self, "错误", f"无法加载文档: {document.load_error if document else '未知错误'}")
            return

        # 更新标题
        self.title_label.setText(os.path.basename(document.file_path))
        self.metadata_btn.setEnabled(True)

        # 更新文本内容
        self._update_text_view()

        # 更新表格列表
        self._update_table_list()

    def _update_text_view(self):
        """更新文本视图"""
        if not self.document or not self.document.is_loaded:
            return

        html = self.document.get_document_html()
        self.text_browser.setHtml(html)

    def _update_table_list(self):
        """更新表格列表"""
        if not self.document or not self.document.is_loaded:
            return

        self.table_combo.clear()

        table_count = len(self.document.tables)
        if table_count == 0:
            self.table_combo.addItem("没有表格")
            self.table_combo.setEnabled(False)
            self.table_browser.setHtml("<p>当前文档中没有表格</p>")
        else:
            self.table_combo.setEnabled(True)
            for i in range(table_count):
                table = self.document.tables[i]
                self.table_combo.addItem(f"表格 {i + 1} ({table['rows']}行 x {table['cols']}列)")

            # 显示第一个表格
            self._update_table_view(0)

    def _update_table_view(self, index):
        """更新表格视图"""
        if not self.document or not self.document.is_loaded or index < 0:
            return

        if index < len(self.document.tables):
            table_html = self.document.get_table_as_html(index)
            self.table_browser.setHtml(table_html)

    def _clear_content(self):
        """清空内容"""
        self.text_browser.clear()
        self.table_browser.clear()
        self.table_combo.clear()
        self.table_combo.setEnabled(False)
        self.metadata_btn.setEnabled(False)

    def _show_metadata(self):
        """显示元数据"""
        if not self.document or not self.document.is_loaded:
            return

        metadata = (
            f"<h3>文档元数据</h3>"
            f"<p><b>标题:</b> {self.document.title}</p>"
            f"<p><b>作者:</b> {self.document.author}</p>"
            f"<p><b>创建日期:</b> {self.document.created_date}</p>"
            f"<p><b>最后修改:</b> {self.document.last_modified_date}</p>"
            f"<p><b>文件路径:</b> {self.document.file_path}</p>"
            f"<p><b>文件大小:</b> {os.path.getsize(self.document.file_path) / 1024:.2f} KB</p>"
            f"<p><b>段落数:</b> {len(self.document.paragraphs)}</p>"
            f"<p><b>表格数:</b> {len(self.document.tables)}</p>"
        )

        QMessageBox.information(self, "文档元数据", metadata)

    def _show_text_context_menu(self, position):
        """显示文本右键菜单"""
        cursor = self.text_browser.cursorForPosition(position)
        cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        selected_text = cursor.selectedText()

        # 获取选定的文本
        cursor = self.text_browser.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
        else:
            return

        if not selected_text:
            return

        menu = QMenu(self)

        # 复制操作
        copy_action = QAction("复制", self)
        copy_action.triggered.connect(self.text_browser.copy)
        menu.addAction(copy_action)

        menu.addSeparator()

        # 创建提取规则
        extract_menu = menu.addMenu("创建提取规则")

        regex_action = QAction("正则表达式", self)
        regex_action.triggered.connect(lambda: self.textSelected.emit(selected_text))
        extract_menu.addAction(regex_action)

        # 智能识别数据类型
        identify_action = QAction("智能识别数据类型", self)
        identify_action.triggered.connect(lambda: self._identify_and_extract(selected_text))
        menu.addAction(identify_action)

        menu.exec(self.text_browser.viewport().mapToGlobal(position))

    def _show_table_context_menu(self, position):
        """显示表格右键菜单"""
        # 获取当前表格索引
        table_index = self.table_combo.currentIndex()
        if table_index < 0 or not self.document or table_index >= len(self.document.tables):
            return

        # 尝试确定单元格位置 (这部分需要根据实际HTML结构调整)
        # 由于QTextBrowser不提供表格单元格定位，这里是一个简化的演示
        # 实际应用中可能需要更复杂的计算或使用QTableView替代

        menu = QMenu(self)

        # 提取整个表格
        extract_table_action = QAction(f"提取整个表格", self)
        extract_table_action.triggered.connect(lambda: self.tableSelected.emit(table_index, -1, -1))
        menu.addAction(extract_table_action)

        # 提取表格列
        if self.document and table_index < len(self.document.tables):
            table = self.document.tables[table_index]
            if table and table['cols'] > 0:
                columns_menu = menu.addMenu("提取表格列")

                for col in range(table['cols']):
                    col_action = QAction(f"提取第 {col + 1} 列", self)
                    col_action.triggered.connect(lambda checked, c=col: self.tableSelected.emit(table_index, -1, c))
                    columns_menu.addAction(col_action)

        menu.exec(self.table_browser.viewport().mapToGlobal(position))

    def _identify_and_extract(self, text):
        """智能识别并提取数据"""
        # 这将由规则控制器处理
        self.textSelected.emit(text)

    def _on_tab_changed(self, index):
        """选项卡切换事件"""
        # 当切换到表格选项卡时刷新表格视图
        if index == 1:  # 表格选项卡
            current_table = self.table_combo.currentIndex()
            if current_table >= 0:
                self._update_table_view(current_table)
