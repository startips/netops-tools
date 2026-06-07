#!/usr/bin/python3
# -*- coding: utf-8 -*-
import logging
import os
import sys
from interface import deviceControl_auto, ping_check

logger = logging.getLogger(__name__)

# 基础路径（兼容 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))


def deviceSend(arg=None):  # 配置检查
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
    # ping检测
    try:
        pingDelay = ping_check(device_ip)[0]
    except (OSError, ValueError, IndexError):
        logger.warning('%s ping检测异常' % device_ip)
        pingDelay = 'timeout'
    result.append(pingDelay)
    try:  # 登录检查
        resData = conn.sendCmd_auto(cmds)
        logger.info('%s 登陆成功' % (device_ip))
        result.append(resData['loginWay'])  # 登录方式
        try:  # 处理数据检查
            logger.info('%s 回显:\n' % (device_ip))
            cmd_result = {}
            # 处理数据
            for resKey, value in resData.items():  # 回显日志
                if resKey != 'loginWay':
                    logger.info('命令:%s\n回显结果:%s\n' % (resKey, value))
                    cmd_result[resKey] = checkError(value)
            # 写入日志文件（best effort）
            try:
                filepath = os.path.join(_base_dir, 'data', '%s_%s.log' % (device_ip, des_local))
                with open(filepath, 'w', encoding='utf-8') as f:
                    for resKey, value in resData.items():
                        if resKey != 'loginWay':
                            f.write('%s\n' % (value))
            except OSError as e:
                logger.error('%s 日志文件写入失败 %s' % (device_ip, e))
            cmd_result = "\n".join(f"{k}:{v}" for k, v in cmd_result.items())
            result.append(cmd_result)  # 下发结果显示
            logger.info('%s 命令下发完成,结果汇总：%s' % (device_ip, cmd_result))
        except Exception as e:
            result.append('数据处理异常 %s' % (e))
            logger.error('%s 数据处理异常 %s' % (device_ip, e))
    except RuntimeError as e:
        logger.error('%s 登陆失败 %s' % (device_ip, e))
        result.append('login fail')
    logger.info('%s 执行完成' % device_ip)
    return result


def checkError(dataTxt):  # 命令报错识别
    errorCode = ['Error: Unrecognized command found']
    for error in errorCode:
        if error in dataTxt:
            return '失败'
    return '成功'


if __name__ == '__main__':
    pass
