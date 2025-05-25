#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
配置管理工具 - 管理应用程序配置
"""

import json
import os

from PyQt6.QtCore import QObject, pyqtSignal


class ConfigManager(QObject):
    """配置管理器"""

    configChanged = pyqtSignal(str, object)  # 配置项, 新值

    def __init__(self, config_file=None):
        super().__init__()

        # 设置默认配置文件路径
        if config_file is None:
            app_data_dir = self._get_app_data_dir()
            self.config_file = os.path.join(app_data_dir, "config.json")
        else:
            self.config_file = config_file

        # 初始化配置
        self.config = self._load_config()

    def _get_app_data_dir(self):
        """获取应用数据目录"""
        app_name = "WordExtractor"

        if os.name == 'nt':  # Windows
            app_data = os.environ.get('APPDATA')
            if app_data:
                app_dir = os.path.join(app_data, app_name)
            else:
                app_dir = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", app_name)
        else:  # macOS/Linux
            app_dir = os.path.join(os.path.expanduser("~"), f".{app_name.lower()}")

        # 确保目录存在
        os.makedirs(app_dir, exist_ok=True)
        return app_dir

    def _load_config(self):
        """加载配置文件"""
        default_config = {
            "appearance": {
                "theme": "light",
                "font_size": 10
            },
            "files": {
                "recent_files": [],
                "recent_directories": [],
                "recursive_scan": True
            },
            "rules": {
                "recent_rule_files": [],
                "auto_save": True
            },
            "export": {
                "recent_export_files": [],
                "default_format": "xlsx"
            }
        }

        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                # 合并配置，确保所有默认项都存在
                self._merge_configs(default_config, config)
                return default_config
            else:
                return default_config
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return default_config

    def _merge_configs(self, default_config, config):
        """递归合并配置，保留默认值并添加新配置"""
        for key, value in default_config.items():
            if key in config:
                if isinstance(value, dict) and isinstance(config[key], dict):
                    self._merge_configs(value, config[key])
                else:
                    default_config[key] = config[key]

    def _save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False

    def get_value(self, key_path, default=None):
        """获取配置值"""
        keys = key_path.split('/')
        config = self.config

        for key in keys:
            if isinstance(config, dict) and key in config:
                config = config[key]
            else:
                return default

        return config

    def set_value(self, key_path, value):
        """设置配置值"""
        keys = key_path.split('/')
        config = self.config

        # 遍历到最后一个键的父对象
        for i, key in enumerate(keys[:-1]):
            if key not in config or not isinstance(config[key], dict):
                config[key] = {}
            config = config[key]

        # 设置值
        last_key = keys[-1]
        config[last_key] = value

        # 保存配置
        self._save_config()

        # 发送配置变更信号
        self.configChanged.emit(key_path, value)

        return True

    def add_recent_file(self, file_path, max_items=10):
        """添加最近打开的文件"""
        recent_files = self.get_value("files/recent_files", [])

        # 如果已存在，移除旧的
        if file_path in recent_files:
            recent_files.remove(file_path)

        # 添加到列表开头
        recent_files.insert(0, file_path)

        # 限制列表长度
        recent_files = recent_files[:max_items]

        # 更新配置
        self.set_value("files/recent_files", recent_files)

    def add_recent_directory(self, dir_path, max_items=5):
        """添加最近打开的目录"""
        recent_dirs = self.get_value("files/recent_directories", [])

        # 如果已存在，移除旧的
        if dir_path in recent_dirs:
            recent_dirs.remove(dir_path)

        # 添加到列表开头
        recent_dirs.insert(0, dir_path)

        # 限制列表长度
        recent_dirs = recent_dirs[:max_items]

        # 更新配置
        self.set_value("files/recent_directories", recent_dirs)

    def add_recent_rule_file(self, file_path, max_items=5):
        """添加最近使用的规则文件"""
        recent_rules = self.get_value("rules/recent_rule_files", [])

        # 如果已存在，移除旧的
        if file_path in recent_rules:
            recent_rules.remove(file_path)

        # 添加到列表开头
        recent_rules.insert(0, file_path)

        # 限制列表长度
        recent_rules = recent_rules[:max_items]

        # 更新配置
        self.set_value("rules/recent_rule_files", recent_rules)

    def add_recent_export_file(self, file_path, max_items=5):
        """添加最近导出的文件"""
        recent_exports = self.get_value("export/recent_export_files", [])

        # 如果已存在，移除旧的
        if file_path in recent_exports:
            recent_exports.remove(file_path)

        # 添加到列表开头
        recent_exports.insert(0, file_path)

        # 限制列表长度
        recent_exports = recent_exports[:max_items]

        # 更新配置
        self.set_value("export/recent_export_files", recent_exports)
