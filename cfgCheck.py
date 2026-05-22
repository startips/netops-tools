#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
cfgCheck.py - 华为交换机离线配置文件检查

根据设备类型（Spine/Leaf/Slf 等）对 .log 配置文件进行逐项合规检查。
结果写入 Excel 报告。

主要函数：
    deviceCheck(arg)       - 入口，读取文件并调用 checkOptions
    checkOptions(data, checkItems) - 调度器，按检查项分发到各 _check_xxx 函数
"""

from interface import get_value, excel
import re


# ============================================================
# 版本补丁对照表（启动时从 Excel 加载）
# 读取 read/版本补丁.xlsx → 版本补丁推荐（季度）sheet →
# 筛选 使用场景=新上线 → 建立 {网元类型: (版本信息, H1推荐补丁)} 映射
# ============================================================
_VERSION_PATCH_MAP = {}  # {型号: (推荐版本, 推荐补丁)}


def _load_version_patch_map():
    """
    加载版本补丁对照表。
    从本地 Excel 文件读取推荐版本和补丁信息，用于后续与设备采集结果对比。
    读取失败时直接报错退出。
    """
    global _VERSION_PATCH_MAP
    map_result = {}
    try:
        xl = excel('read/版本补丁.xlsx')
        xl.excelReadCread()  # 初始化 workbook 对象
        data = xl.excelReadSheet(sheetnum='版本补丁推荐（季度）')
    except Exception as e:
        raise RuntimeError(f'读取版本补丁.xlsx 失败: {e}')

    if not data:
        return  # 空文件，不报错但对照表为空

    headers = data[0]
    # 找到列索引
    scene_idx = next((i for i, h in enumerate(headers) if '使用场景' in str(h)), None)
    if scene_idx is None:
        raise RuntimeError('版本补丁.xlsx 中未找到"使用场景"列')

    for row in data[1:]:
        try:
            if row[scene_idx] == '新上线':
                map_result[row[0]] = (row[1], row[2])
        except IndexError:
            continue  # 跳过空行

    _VERSION_PATCH_MAP = map_result

def _get_device_model(fileTxt):
    """
    从配置文件中提取设备型号。

    通过匹配 "HUAWEI 型号 uptime is" 获取设备型号字符串。

    参数：
        fileTxt: 配置文件完整文本

    返回：
        str or None — 设备型号，未匹配到时返回 None
    """
    m = re.search(r'HUAWEI (\S+) (?:Routing Switch)?\s*uptime is', fileTxt, re.IGNORECASE)
    return m.group(1) if m else None


def _match_model(model):
    """
    将设备型号匹配到版本补丁对照表中的网元类型。

    匹配规则：
    1. 精确匹配（优先）
    2. S5731 系列前缀匹配（如 S5731-H48T4XC → S5731-H）
    3. 其他情况返回 None

    参数：
        model: 配置文件提取的设备型号，如 'CE6885-48YS8CQ'

    返回：
        str or None — 对照表中的网元类型键值
    """
    if not model:
        return None

    # 1. 精确匹配
    if model in _VERSION_PATCH_MAP:
        return model

    # 2. S5731 前缀匹配（S5731-H48T4XC / S5731-S48T4X → S5731-H / S5731-S）
    if model.startswith('S5731-'):
        for key in _VERSION_PATCH_MAP:
            if key.startswith('S5731-') and model.startswith(key):
                return key

    return None


# 模块加载时自动加载对照表
_load_version_patch_map()


def deviceCheck(arg):  # 检查
    """
    设备检查入口函数。

    读取 read/config/ 目录下的配置文件，调用 checkOptions 执行各项检查。

    参数：
        arg: dict，包含 name（设备名）、type（设备类型）、filename（配置文件名）
             以及 checkOption（33个检查项的 0/1 配置）

    返回：
        list，[设备名, 设备类型, sysname, 管理IP, 型号, 各项检查结果...]
    """
    logger = get_value('logger')
    data_local = arg
    result = [data_local['name'], data_local['type']]
    try:  # 读取文件内容
        with open('read/config/%s' % data_local['filename'], 'r', encoding='utf-8', errors='ignore') as f:
            fileTxt = f.read()
    except Exception as e:
        logger.get_log().error('%s 读取文件失败 %s' % (data_local['name'], e))
        result.append('read file fail')
        return result
    result.extend(checkOptions(fileTxt, data_local))  # 检查内容，检查项
    logger.get_log().info('%s 所有项检查完成' % data_local['name'])
    return result


# ============================================================
# 检查项调度器
# ============================================================
def checkOptions(fileTxt, checkItems):
    """
    配置检查调度器。

    固定提取 3 项基本元信息（sysname/管理IP/型号），
    然后根据 checkItems['checkOption'] 中的 0/1 配置，
    将每项检查分发到对应的 _check_xxx 函数。

    参数：
        fileTxt:     配置文件的完整文本内容
        checkItems:  dict，包含 name（设备名）、type（设备类型）、
                     checkOption（{检查项名: 0/1}）

    返回：
        list，[sysname, 管理IP, 型号, 各项检查结果...]
    """
    logger = get_value('logger')
    checkResult = []

    # ---- 固定提取项：sysname / 管理IP / 设备型号 ----
    devSysnameMatch = re.search(r'#\s*\n\s*sysname (\S+)\s*\n\s*#', fileTxt, re.IGNORECASE)
    checkResult.append(devSysnameMatch.group(1) if devSysnameMatch else '未匹配到')

    devTypeMatch = re.search(
        r'interface MEth\S+\s*\n\s*'
        r'description \S+\s*\n\s*'
        r'ip binding vpn-instance \S+\s*\n\s*'
        r'ip address (\d+\.\d+\.\d+\.\d+) \d+\.\d+\.\d+\.\d+',
        fileTxt, re.IGNORECASE
    )
    checkResult.append(devTypeMatch.group(1) if devTypeMatch else '未匹配到')

    devIpMatch = re.search(r'HUAWEI (\S+) (?:Routing Switch)?\s*uptime is', fileTxt, re.IGNORECASE)
    checkResult.append(devIpMatch.group(1) if devIpMatch else '未匹配到')

    # ---- 遍历检查项配置表，分发到各检查函数 ----
    for checkItem, value in checkItems['checkOption'].items():
        logger.get_log().info(
            f'{checkItems["name"]}的检查项"{checkItem}"设置为"{value}",开始检查'
        )
        if value == 1:
            checker = _CHECKERS.get(checkItem)
            if checker:
                checkResult.append(checker(fileTxt, checkItems))
            else:
                checkResult.append(f'未知配置项: {checkItem}')
        else:
            checkResult.append('不涉及')
        logger.get_log().info(
            f'{checkItems["name"]}的检查项"{checkItem}"设置为"{value}",检查完成'
        )
    return checkResult


# ============================================================
# 以下为各检查项的具体实现函数
# 命名规则：_check_xxx(fileTxt, checkItems) -> str
# 每个函数只包含核心检查逻辑，返回结果字符串
# ============================================================

def _check_version(fileTxt, checkItems):
    """
    检查交换机版本号，并与本地版本补丁对照表对比。

    从配置文件中提取版本号，然后根据设备型号在对照表中查找推荐版本：
    - 匹配 → 返回版本号
    - 不匹配 → 返回 '版本号@未通过'
    - 找不到对应型号 → 返回 '版本号-未知'
    - 采集不到版本号 → 返回 '未匹配到'
    """
    matchVerinfo = re.findall(r'Version \S+ \(\S+ (\S+)\)', fileTxt)
    collected = matchVerinfo[0] if matchVerinfo else '未匹配到'

    if collected == '未匹配到':
        return collected

    model = _get_device_model(fileTxt)
    matched_key = _match_model(model)

    if matched_key is None:
        return f'{collected}-未知'

    recommended = _VERSION_PATCH_MAP[matched_key][0]  # 推荐版本
    return collected if collected == recommended else f'{collected}-未通过'


def _check_patch(fileTxt, checkItems):
    """
    检查补丁版本号，并与本地版本补丁对照表对比。

    从配置文件中提取补丁号，然后根据设备型号在对照表中查找推荐补丁：
    - 匹配 → 返回补丁号
    - 不匹配 → 返回 '补丁号@未通过'
    - 找不到对应型号 → 返回 '补丁号-未知'
    - 采集不到补丁号 → 返回 '未匹配到'
    """
    matchPatInfo = re.findall(r'Patch Package Version\s?\:(\S+)', fileTxt)
    collected = matchPatInfo[0] if matchPatInfo else '未匹配到'

    if collected == '未匹配到':
        return collected

    model = _get_device_model(fileTxt)
    matched_key = _match_model(model)

    if matched_key is None:
        return f'{collected}-未知'

    recommended = _VERSION_PATCH_MAP[matched_key][1]  # 推荐补丁
    return collected if collected == recommended else f'{collected}-未通过'


def _check_extra_files(fileTxt, checkItems):
    """
    检查 flash 中多余文件（.cc .pat .cfg）。

    在 flash 目录列表中统计各系统文件类型的数量，
    用循环遍历后缀名列表，避免三段重复的 findall 代码。
    判断标准：
        - .cc（版本文件）：应恰好 1 个
        - .pat（补丁文件）：应恰好 1 个
        - .cfg（配置文件）：应 ≤ 1 个

    参数：
        fileTxt:     配置文件完整文本
        checkItems:  设备信息（保留参数，暂未使用）

    返回：
        '通过'        — 文件数量符合预期
        '未匹配到'    — 找不到 flash 目录信息或三项均为 0
        str           — 各类型文件的实际数量（如 'cc:2,pat:1,cfg:1'）
    """
    matchDirInfo = re.search(r'Directory of flash[\s\S]*?<', fileTxt, re.IGNORECASE)
    if not matchDirInfo:
        return '未匹配到'
    dirInfo = matchDirInfo.group()

    # 参数化：对每个文件后缀类型执行相同的 findall，代码量缩减 2/3
    # 原始正则除后缀名外完全一致，用循环 + f-string 动态替换后缀
    extensions = ['cc', 'pat', 'cfg']
    counts = {}
    for ext in extensions:
        counts[ext] = len(re.findall(
            rf'\d+\s+\S+\s+\S+\s+\S+\s\d+\s\d+\s\d+\:\d+\:\d+\s+\S+\.{ext}',
            dirInfo, re.IGNORECASE
        ))

    if counts['cc'] == 1 and counts['pat'] == 1 and counts['cfg'] <= 1:
        return '通过'
    elif all(v == 0 for v in counts.values()):
        return '未匹配到'
    else:
        return 'cc:%d,pat:%d,cfg:%d' % (counts['cc'], counts['pat'], counts['cfg'])


def _check_hardware(fileTxt, checkItems):
    """检查硬件设备状态（有无 Offline/Unregistered/Abnormal）"""
    deviceInfo = re.search(r'Device status:[\s\S]*?<', fileTxt, re.IGNORECASE)
    if not deviceInfo:
        return '未匹配到'
    statusStr = ['Offline', 'Unregistered', 'Abnormal']
    for s in statusStr:
        if s in deviceInfo.group():
            return '未通过'
    return '通过'


def _check_open_ports(fileTxt, checkItems):
    """检查是否存在 down 状态但未 shutdown 的端口"""
    matchDownPortInfo = re.findall(
        r'(?:Multi|100|25|10)GE\d+\/\d+\/\d+\s+(down|down\(ed\)|down\(b\))'
        r'(?:\s+\S+){3}\s+\d+\s+\d+',
        fileTxt, re.IGNORECASE
    )
    if matchDownPortInfo:
        return str(len(matchDownPortInfo))
    return '通过'


def _check_bgp_neighbor(fileTxt, checkItems):
    """检查 BGP 邻居状态"""
    bgpInfo = re.search(r'BGP local router ID[\s\S]*?<', fileTxt, re.IGNORECASE)
    if not bgpInfo:
        return '未匹配到'
    bgpNum = re.findall(r'Total number of peers\s+\:\s+(\d+)', bgpInfo.group(), re.IGNORECASE)
    if not bgpNum:
        return '未匹配到'
    normalBgpNum = len(re.findall(
        r'\d+\.\d+\.\d+\.\d+(:?\s+\d+){5}\s+\S+\s+Established',
        bgpInfo.group(), re.IGNORECASE
    ))
    return '邻居数量:%s,正常邻居数量:%d' % (bgpNum[0], normalBgpNum)


def _check_feature_software(fileTxt, checkItems):
    """检查 feature-software 状态"""
    matchFeaInfo = re.search(r'FeatureName[\s\S]*?<', fileTxt)
    if not matchFeaInfo:
        return '未匹配到'
    featureInfo = re.findall(
        r'(:?PKG_PNF|AIFABRIC|TELEMETRY|WEAKEA)\s+\S+\.cc\s+active\s+'
        r'\S+\s+\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}',
        matchFeaInfo.group(), re.IGNORECASE
    )
    return '通过' if len(featureInfo) >= 4 else '未通过'


def _check_failed_commands(fileTxt, checkItems):
    """检查配置回滚的失败命令数量"""
    matchRecover = re.findall(r'The number of failed commands is (\d+)', fileTxt, re.IGNORECASE)
    if not matchRecover:
        return '未匹配到'
    return '通过' if matchRecover[0] == '0' else matchRecover[0]


def _check_esn(fileTxt, checkItems):
    """检查设备 ESN 序列号"""
    esnInfo = re.search(r'ESN:\s*(\w+)', fileTxt, re.IGNORECASE)
    return esnInfo.group(1) if esnInfo else '未匹配到'


def _check_mlag_status(fileTxt, checkItems):
    """检查 M-LAG 心跳/主备状态"""
    return '通过' if genCheckOption(
        r'Heart beat state\s+\:\s+OK\s*\n\s*'
        r'Node [12][\s\S]*?State\s+\:\s+(:?Backup|Master)[\s\S]*?'
        r'Node [12][\s\S]*?State\s+\:\s+(:?Backup|Master)',
        fileTxt
    ) else '未通过'


def _check_hash(fileTxt, checkItems):
    """检查负载均衡 ECMP hash 模式"""
    return '通过' if genCheckOption(
        r'#\s*\n\s*'
        r'load-balance ecmp\s*\n\s*'
        r'hashmode (underlay)? 2\s*\n\s*'
        r'#', fileTxt
    ) else '未通过'


def _check_alarm_active(fileTxt, checkItems):
    """
    检查 display alarm active 活跃告警。

    在配置文件中搜索 display alarm active 命令的回显区域，
    统计该区域中是否存在活跃告警条目。

    参数：
        fileTxt:     配置文件完整文本
        checkItems:  设备信息（含 type/name）

    返回：
        '未匹配到' — 配置文件中没有 display alarm active 回显
        '通过'      — 有回显区域，但没有告警数据行
        'N'         — 有 N 条活跃告警（N 为数字字符串）
    """
    # 定位 display alarm active 命令的回显区域
    # 回显格式：
    #   display alarm active
    #   Sequence   AlarmId    Severity Date Time  Description
    #   1          0x66f00001 Major    2026-05-16 14:30:25  Entity alarms
    # 或：
    #   display alarm active | no
    #   Sequence   AlarmId    Severity Date Time  Description
    #   （无告警数据行，只有表头）
    alarmSection = re.search(
        r'display alarm active[\s\S]*?<',  # 从 display alarm active 开始到下一个 <（命令提示符标记）
        fileTxt, re.IGNORECASE
    )
    # 如果找不到 display alarm active 段，返回未匹配到
    if not alarmSection:
        return '未匹配到'

    alarmText = alarmSection.group()
    # 在回显区域中统计告警条目数量
    # 告警行特征：序号 + 16进制AlarmId + 严重级别 + 日期
    alarmCount = len(re.findall(
        r'\d+\s+0x[0-9a-fA-F]+\s+\S+\s+\d{4}-\d{2}-\d{2}',
        alarmText
    ))

    if alarmCount == 0:
        return '通过'
    else:
        return str(alarmCount)


# ============================================================
# 检查项分发表：检查项名称 -> 对应函数
# ============================================================
_CHECKERS = {
    '版本':                _check_version,
    '补丁':                _check_patch,
    '多余文件检查':         _check_extra_files,
    '硬件状态检查':         _check_hardware,
    '未关闭端口':           _check_open_ports,
    'bgp邻居状态':          _check_bgp_neighbor,
    'feature-software状态': _check_feature_software,
    '失败命令配置检查':      _check_failed_commands,
    '设备Esn':             _check_esn,
    'mlag状态':            _check_mlag_status,
    'hash配置':            _check_hash,
    '告警检查':            _check_alarm_active,
}


def genCheckOption(pattern, fileTxt):
    """
    通用正则匹配检查工具。

    在配置文本中搜索给定正则模式，匹配即返回 match 对象，否则返回 None。
    本函数作为各 _check_xxx 函数的底层工具，统一大小写不敏感的匹配行为。

    参数：
        pattern:   正则表达式字符串（会被传入 re.search 的 r'' 原始字符串）
        fileTxt:  配置文件的完整文本内容

    返回：
        re.Match or None — 匹配成功返回 match 对象，失败返回 None

    注意：
        - 默认启用 re.IGNORECASE
        - pattern 中尽量不要用 ^ 或 $，因为 fileTxt 是整个文件全文，
          建议用 #\\s*\\n 或 \\n\\s* 来控制位置
    """
    return re.search(r'%s' % pattern, fileTxt, re.IGNORECASE)


if __name__ == '__main__':
    pass