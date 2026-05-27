#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
检查项 & 设备类型配置 — 懒加载 read/check_items.yaml

导出（与原 check_items.py 接口一致）：
    CHECK_ITEM_NAMES      — 检查项名称元组
    DEVICE_TYPE_CONFIGS   — {类型名: {检查项: 0/1}}
    DEVICE_TYPE_PATTERNS  — ((regex_string, 类型名), ...)
    CABLE_CHECK_CONFIG    — 线路检查排除规则
"""

import os
import sys
import yaml

# ============================================================
# 路径解析（兼容 PyInstaller 打包）
# ============================================================

if getattr(sys, 'frozen', False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.join(os.path.dirname(__file__), '..')

_yaml_path = os.path.join(_base_dir, 'read', 'check_items.yaml')

# ============================================================
# 懒加载：首次访问导出名时才读 YAML
# ============================================================

_data = None
_loaded = False


def _ensure_loaded():
    global _data, _loaded
    if _loaded:
        return
    with open(_yaml_path, 'r', encoding='utf-8') as f:
        _data = yaml.safe_load(f)
    _loaded = True


def _make_check_option(check_items, *enabled_items):
    enabled = set(enabled_items)
    return {name: (1 if name in enabled else 0) for name in check_items}


# 需要懒加载的导出名 → 在 _data 中的构建逻辑
_EXPORTS = {
    'CHECK_ITEM_NAMES',
    'DEVICE_TYPE_CONFIGS',
    'DEVICE_TYPE_PATTERNS',
    'CABLE_CHECK_CONFIG',
}

# 缓存已构建的值
_cache = {}


def __getattr__(name):
    if name not in _EXPORTS:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

    if name in _cache:
        return _cache[name]

    _ensure_loaded()

    if name == 'CHECK_ITEM_NAMES':
        val = tuple(_data['check_items'])
    elif name == 'DEVICE_TYPE_CONFIGS':
        names = tuple(_data['check_items'])
        val = {
            type_name: _make_check_option(names, *items)
            for type_name, items in _data['device_types'].items()
        }
    elif name == 'DEVICE_TYPE_PATTERNS':
        val = tuple(
            (pattern, type_name)
            for pattern, type_name in _data['device_patterns']
        )
    elif name == 'CABLE_CHECK_CONFIG':
        val = _data.get('cable_check', {})
    else:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

    _cache[name] = val
    return val
