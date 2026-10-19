#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
华为网络设备 display version 信息提取脚本（Windows / Python 3.7.3 兼容版）
功能：提取设备名称、原始运行时间，并折算为天数，汇总输出至 CSV 文件。
完全基于 Python 标准库，无任何外部依赖。
"""

import os
import re
import sys
import csv
from decimal import Decimal, ROUND_HALF_UP


# ---- Windows 控制台中文输出兼容处理（Python 3.7 可直接调用 reconfigure）----
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass


def read_file_safely(file_path):
    """尝试多种常见编码读取文本文件，避免因日志编码不同导致报错"""
    encodings = ['utf-8-sig', 'utf-8', 'gb18030', 'gbk', 'latin-1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()


def parse_uptime_to_days(uptime_str):
    """
    将运行时间字符串解析并换算为总天数（保留两位小数）
    示例输入: '329 weeks, 0 day, 22 hours, 49 minutes'
    """
    years = re.search(r'(\d+)\s*year', uptime_str, re.IGNORECASE)
    weeks = re.search(r'(\d+)\s*week', uptime_str, re.IGNORECASE)
    days = re.search(r'(\d+)\s*day', uptime_str, re.IGNORECASE)
    hours = re.search(r'(\d+)\s*hour', uptime_str, re.IGNORECASE)
    minutes = re.search(r'(\d+)\s*min', uptime_str, re.IGNORECASE)
    seconds = re.search(r'(\d+)\s*sec', uptime_str, re.IGNORECASE)

    val_years = Decimal(years.group(1)) if years else Decimal('0')
    val_weeks = Decimal(weeks.group(1)) if weeks else Decimal('0')
    val_days = Decimal(days.group(1)) if days else Decimal('0')
    val_hours = Decimal(hours.group(1)) if hours else Decimal('0')
    val_minutes = Decimal(minutes.group(1)) if minutes else Decimal('0')
    val_seconds = Decimal(seconds.group(1)) if seconds else Decimal('0')

    total_days = (
        val_years * Decimal('365') +
        val_weeks * Decimal('7') +
        val_days +
        (val_hours / Decimal('24')) +
        (val_minutes / Decimal('1440')) +
        (val_seconds / Decimal('86400'))
    )

    return "{:.2f}".format(
        total_days.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    )


def extract_device_info(content):
    """从文本内容中提取设备名与整机运行时间"""
    device_name = None
    uptime_raw = None

    # 1. 提取设备名称（匹配提示符，如 <CD4F01B05U38-BMC-S57Z> 或 [CD4F01B05U38-BMC-S57Z]）
    prompts = re.findall(r'[<\[]([A-Za-z0-9_\-\.]+)[>\]]', content)
    if prompts:
        device_name = prompts[-1]

    # 2. 提取整机启动时间（取第一处匹配到的 uptime is）
    uptime_match = re.search(r'uptime is\s+([^\r\n]+)', content, re.IGNORECASE)
    if uptime_match:
        uptime_raw = uptime_match.group(1).strip()

    return device_name, uptime_raw


def process_directory(target_dir, output_csv):
    """扫描目标目录并将解析结果写入 CSV 文件"""
    target_dir = os.path.abspath(target_dir)
    if not os.path.isdir(target_dir):
        raise FileNotFoundError("指定的目录不存在: {}".format(target_dir))

    valid_exts = ('.txt', '.log', '.dat', '')

    files = [
        f for f in os.listdir(target_dir)
        if os.path.isfile(os.path.join(target_dir, f))
        and not f.startswith('.')
        and os.path.splitext(f)[1].lower() in valid_exts
    ]
    files.sort()

    if not files:
        print("在目录 [{}] 中未找到符合条件的文本文件。".format(target_dir))
        return

    headers = ["文件名", "设备名称", "运行天数 (天)", "原始运行时间", "采集状态"]
    rows = []
    success_count = 0

    for file_name in files:
        file_path = os.path.join(target_dir, file_name)
        try:
            content = read_file_safely(file_path)
            device_name, uptime_raw = extract_device_info(content)

            status = "正常"
            days = ""
            if not device_name and not uptime_raw:
                status = "未匹配到设备名与运行时间"
            elif not device_name:
                status = "未匹配到提示符设备名"
            elif not uptime_raw:
                status = "未匹配到启动时间"
            else:
                days = parse_uptime_to_days(uptime_raw)
                success_count += 1

            rows.append([
                file_name,
                device_name if device_name else "未提取到",
                days,
                uptime_raw if uptime_raw else "未提取到",
                status
            ])
        except Exception as err:
            rows.append([file_name, "解析失败", "", "", "错误: {}".format(err)])

    # utf-8-sig 编码，确保用 Excel 打开 CSV 不乱码
    with open(output_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    print("处理完成！")
    print("共扫描文件: {} 个，成功提取: {} 个".format(len(files), success_count))
    print("结果已保存至: {}".format(os.path.abspath(output_csv)))


if __name__ == '__main__':
    default_dir = "./text_logs"
    default_output = "./device_uptime.csv"

    target_directory = sys.argv[1] if len(sys.argv) > 1 else default_dir
    output_file = sys.argv[2] if len(sys.argv) > 2 else default_output

    process_directory(target_directory, output_file)
