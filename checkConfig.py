#!/usr/bin/python3
# -*- coding: utf-8 -*-
import logging
import os
import sys
from interface import deviceControl_auto, ping_check, revData_error, readTxt
import re

logger = logging.getLogger(__name__)

# 基础路径（兼容 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))


def infoDeal(data):  # 数据处理 返回list
    logger.info('开始处理配置数据...')
    data_local = data
    result_dic = {}
    result = []
    readInfo = readTxt(os.path.join(_base_dir, 'read', 'Keywords.txt'))  # 读取匹配关键字
    logger.info('成功加载关键字 %d 条', len(readInfo))
    for i in readInfo:
        cell = i.split(',')
        result_dic.update({cell[1]: ''})
    revInfo = revData_error(data_local['dis current-configuration'])  # 判断是否有命令执行错误
    if revInfo == 'NULL':
        for info in readInfo:
            keywords = info.split(',')
            configStr = re.search(r'%s' % keywords[0], data_local['dis current-configuration'], re.IGNORECASE)
            if configStr:
                if keywords[2] == '0':
                    checkRes = '多余\'%s\':%s\n' % (keywords[1], keywords[0])
                    result_dic.update({keywords[1]: result_dic[keywords[1]] + checkRes})
                    logger.warning('发现多余配置: 分类=%s, 关键字=%s', keywords[1], keywords[0])
            else:
                if keywords[2] == '1':
                    checkRes = '缺少\'%s\':%s\n' % (keywords[1], keywords[0])
                    result_dic.update({keywords[1]: result_dic[keywords[1]] + checkRes})
                    logger.warning('发现缺少配置: 分类=%s, 关键字=%s', keywords[1], keywords[0])
        else:
            for key in result_dic:
                if result_dic[key] == '':
                    result_dic.update({key: '无不合规项'})
    else:
        logger.error('命令执行报错: %s', revInfo)
        result.append(revInfo)
    for key in result_dic:  # 转换到list
        result.append(result_dic[key])
    logger.info('配置数据处理完成，返回 %d 项结果', len(result))
    return result


def deviceCheck(arg=None):  # 配置检查
    if arg is None:
        arg = []
    device_ip = arg[2]
    device_user = arg[0]
    device_pass = arg[1]
    des_local = arg[3]
    logger.info('开始检查设备: %s (%s)', device_ip, des_local)
    conn = deviceControl_auto(device_ip, device_user, device_pass)  # 登陆
    cmd = ['dis current-configuration']  # 命令
    result = [device_ip, des_local]
    pingDelay = ping_check(device_ip)[0]  # ping检测
    result.append(pingDelay)
    try:  # 登录检查
        resData = conn.sendCmd_auto(cmd)
        logger.info('%s 登陆成功, 登录方式: %s', device_ip, resData['loginWay'])
        result.append(resData['loginWay'])  # 登录方式
        try:  # 处理数据检查
            result.extend(infoDeal(resData))
            logger.info('%s 数据处理成功', device_ip)
        except Exception as e:
            result.append('数据处理失败 %s' % (e))
            logger.error('%s 数据处理失败: %s', device_ip, e, exc_info=True)
    except Exception as e:
        logger.error('%s 登陆失败: %s', device_ip, e, exc_info=True)
        result.append('login fail')
    logger.info('设备 %s 检查完成', device_ip)
    return result


if __name__ == '__main__':
    infoDeal({'dis current-configuration': '123'})
    # c = {}
    # print(type(c.get('i')))
