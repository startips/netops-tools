#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
log_config.py - 日志配置模块

替代原 logger.py 的 logg 类，使用 Python 标准 logging 模式。
应用启动时调用 setup_logging() 一次，后续模块直接 logging.getLogger(__name__) 使用。
"""

import os
import time
import logging


def setup_logging(log_name: str, log_dir: str = 'log'):
    """
    配置 root logger，应用启动时调用一次。

    每次调用会清除旧 handler 并创建新日志文件，
    支持同一进程中运行不同功能时切换日志文件。

    参数：
        log_name: 日志文件名前缀（如 'checkConfig', 'sendCmd'）
        log_dir:  日志文件存放目录，默认 'log'
    """
    os.makedirs(log_dir, exist_ok=True)

    root = logging.getLogger()

    # 清除旧 handler（支持多次调用切换日志文件）
    for handler in root.handlers[:]:
        root.removeHandler(handler)
        handler.close()

    root.setLevel(logging.DEBUG)

    time_now = time.strftime('%Y-%m-%d_%H%M%S', time.localtime())
    log_file = f'{log_dir}/{log_name}_{time_now}.log'

    formatter = logging.Formatter(
        '%(asctime)s - %(threadName)s - %(name)s - %(levelname)s:%(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # 文件 handler
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    root.addHandler(fh)

    # 控制台 handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.ERROR)
    ch.setFormatter(formatter)
    root.addHandler(ch)

    return root
