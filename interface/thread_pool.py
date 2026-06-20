#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
thread_pool.py - 线程池模块
提供带进度条的并发执行功能。
"""

import logging
from concurrent import futures

logger = logging.getLogger(__name__)


class autoThreadingPool():  # 线程池
    def __init__(self, worker=30, bar=None):
        self.bar = bar
        self.worker_local = worker
        self.result = []

    def __call__(self, func, datalist: list):  # 函数，迭代数据
        if not datalist:
            return []
        func_local = func  # function
        datalist_local = datalist  # data
        with futures.ThreadPoolExecutor(max_workers=self.worker_local) as executor:  # max_workers 线程池的数量
            future_list = []
            for row in datalist_local:
                future = executor.submit(func_local, row)
                future_list.append(future)
            unit = 0.8 / len(future_list)
            num = 0.1
            for future in futures.as_completed(future_list):
                try:
                    res = future.result()
                    self.result.append(res)
                except Exception as e:
                    logger.error(f'任务执行失败: {e}', exc_info=True)
                    self.result.append(None)
                num += unit
                self.bar(num)
        return self.result
