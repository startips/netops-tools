#!/usr/bin/python3
# -*- coding: utf-8 -*-
import logging
import os
import sys
import time
from interface import deviceControl_auto, ping_check

logger = logging.getLogger(__name__)

# 基础路径（兼容 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))


def deviceSend(arg=None):  # 配置下发
    if arg is None:
        arg = []
    device_ip = arg[2]
    device_user = arg[0]
    device_pass = arg[1]
    des_local = arg[3]
    if ',' in arg[4]:  # 命令
        cmds = arg[4].split(',')
    else:
        cmds = [arg[4]]

    conn = deviceControl_auto(device_ip, device_user, device_pass)
    result = [device_ip, des_local]
    total_start = time.time()

    logger.info(f'{device_ip} 开始执行 ({len(cmds)}条命令)')

    # ping检测
    try:
        t0 = time.time()
        pingDelay = ping_check(device_ip)[0]
        logger.info(f'{device_ip} | ping: {pingDelay}ms (耗时{time.time() - t0:.1f}s)')
    except (OSError, ValueError, IndexError) as e:
        logger.warning(f'{device_ip} | ping异常: {e}')
        pingDelay = 'timeout'
    result.append(pingDelay)

    # 登录
    try:
        t0 = time.time()
        resData = conn.sendCmd_auto(cmds)
        login_time = time.time() - t0
        login_way = resData['loginWay']
        logger.info(f'{device_ip} | {login_way}连接成功 (耗时{login_time:.1f}s)')
        result.append(login_way)

        # 逐条执行命令
        cmd_result = {}
        success_count = 0
        fail_count = 0
        cmd_total = len([k for k in resData if k != 'loginWay'])

        for idx, (resKey, value) in enumerate(resData.items(), 1):
            if resKey == 'loginWay':
                continue
            check_res, err_detail = checkError(value)
            byte_len = len(value.encode('utf-8', errors='ignore'))

            if check_res == '成功':
                success_count += 1
                logger.info(f'{device_ip} | [{idx}/{cmd_total}] {resKey} → 成功 (回显{byte_len}字节)')
            else:
                fail_count += 1
                logger.warning(f'{device_ip} | [{idx}/{cmd_total}] {resKey} → 失败: {err_detail}')

            cmd_result[resKey] = check_res[0]

        # 写入日志文件
        try:
            filepath = os.path.join(_base_dir, 'data', f'{device_ip}_{des_local}.log')
            with open(filepath, 'w', encoding='utf-8') as f:
                for resKey, value in resData.items():
                    if resKey != 'loginWay':
                        f.write(f'=== {resKey} ===\n')
                        f.write(f'{value}\n\n')
            logger.info(f'{device_ip} | 日志已保存: {filepath}')
        except OSError as e:
            logger.error(f'{device_ip} | 日志文件写入失败: {e}')

        # 汇总
        total_time = time.time() - total_start
        summary = "\n".join(f"{k}:{v}" for k, v in cmd_result.items())
        result.append(summary)
        logger.info(f'{device_ip} | 结果: 成功{success_count}条, 失败{fail_count}条 (总耗时{total_time:.1f}s)')

    except RuntimeError as e:
        logger.error(f'{device_ip} | 登录失败 (SSH+Telnet均失败): {e}')
        result.append('login fail')

    logger.info(f'{device_ip} 执行完成')
    return result


def checkError(dataTxt):  # 命令报错识别
    errorCode = ['Error: Unrecognized command found']
    for error in errorCode:
        if error in dataTxt:
            # 提取错误行
            for line in dataTxt.split('\n'):
                if error in line:
                    return '失败', line.strip()[:100]
            return '失败', error
    return '成功', ''


if __name__ == '__main__':
    pass
