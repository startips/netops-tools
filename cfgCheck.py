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
    """检查 flash 中多余文件（.cc .pat .cfg）"""
    matchDirInfo = re.search(r'Directory of flash[\s\S]*?<', fileTxt, re.IGNORECASE)
    if not matchDirInfo:
        return '未匹配到'
    dirInfo = matchDirInfo.group()
    verCount = len(re.findall(
        r'\d+\s+\S+\s+\S+\s+\S+\s\d+\s\d+\s\d+\:\d+\:\d+\s+\S+\.cc',
        dirInfo, re.IGNORECASE
    ))
    patCount = len(re.findall(
        r'\d+\s+\S+\s+\S+\s+\S+\s\d+\s\d+\s\d+\:\d+\:\d+\s+\S+\.pat',
        dirInfo, re.IGNORECASE
    ))
    cfgCount = len(re.findall(
        r'\d+\s+\S+\s+\S+\s+\S+\s\d+\s\d+\s\d+\:\d+\:\d+\s+\S+\.cfg',
        dirInfo, re.IGNORECASE
    ))
    if verCount == 1 and patCount == 1 and cfgCount <= 1:
        return '通过'
    elif verCount == 0 and patCount == 0 and cfgCount == 0:
        return '未匹配到'
    else:
        return 'cc:%d,pat:%d,cfg:%d' % (verCount, patCount, cfgCount)


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


def _check_ftp_disabled(fileTxt, checkItems):
    """检查是否关闭 FTP 服务"""
    return '通过' if genCheckOtion(
        r'undo ftp server source all-interface\s*\n\s*'
        r'undo ftp ipv6 server source all-interface',
        fileTxt
    ) else '未通过'


def _check_mlag_status(fileTxt, checkItems):
    """检查 M-LAG 心跳/主备状态"""
    return '通过' if genCheckOtion(
        r'Heart beat state\s+\:\s+OK\s*\n\s*'
        r'Node [12][\s\S]*?State\s+\:\s+(:?Backup|Master)[\s\S]*?'
        r'Node [12][\s\S]*?State\s+\:\s+(:?Backup|Master)',
        fileTxt
    ) else '未通过'


def _check_mlag_config(fileTxt, checkItems):
    """检查 dfs-group / M-LAG 配置"""
    return '通过' if genCheckOtion(
        r'#\s*\n\s*'
        r'dfs-group 1\s*\n\s*'
        r'authentication-mode hmac-sha256 password \S+\s*\n\s*'
        r'dual-active detection source ip \S+ vpn-instance DAD peer \S+\s*\n\s*'
        r'm-lag up-delay 90\s*\n\s*'
        r'priority 1[25]0\s*\n\s*'
        r'#',
        fileTxt
    ) else '未通过'


def _check_large_route(fileTxt, checkItems):
    """检查大路由配置（system resource large-route）"""
    if genCheckOtion(
        r'#\s*\n\s*system resource large-route\s*\n\s*#',
        fileTxt
    ):
        return '通过'
    # 排除 CE6866 和 k8s vlanif
    if (genCheckOtion(r'HUAWEI CE6866\S+ uptime is', fileTxt) or
            genCheckOtion(r'interface Vlanif\d+\s*\n\s*description\s+.*?k8s', fileTxt)):
        return '不涉及'
    return '未通过'


def _check_ntp(fileTxt, checkItems):
    """检查 NTP 源接口关闭配置"""
    return '通过' if genCheckOtion(
        r'ntp server source-interface all disable\s*\n\s*'
        r'ntp ipv6 server source-interface all disable',
        fileTxt
    ) else '未通过'


def _check_vlan_global(fileTxt, checkItems):
    """检查全局 vlan batch 配置"""
    matchVlanInfo = re.search(r'vlan batch\s+(.*)', fileTxt)
    return matchVlanInfo.group(1) if matchVlanInfo else '未匹配到'


def _check_mac_flapping(fileTxt, checkItems):
    """检查 MAC 飘移检测配置"""
    return '通过' if genCheckOtion(
        r'mac-address flapping detection security-level low\s*\n\s*'
        r'mac-address flapping periodical trap enable',
        fileTxt
    ) else '未通过'


def _check_stp(fileTxt, checkItems):
    """检查 STP/RSTP 基础配置"""
    return '通过' if genCheckOtion(
        r'stp bridge-address \d{4}-\d{4}-\d{4}\s*\n\s*'
        r'stp mode rstp\s*\n\s*'
        r'stp v-stp enable\s*\n\s*'
        r'stp instance 0 root primary\s*\n\s*'
        r'stp bpdu-protection\s*\n\s*'
        r'stp tc-protection',
        fileTxt
    ) else '未通过'


def _check_arp_conflict(fileTxt, checkItems):
    """检查 ARP 冲突检测配置"""
    return '通过' if genCheckOtion(
        r'#\s*\n\s*arp ip-conflict-detect enable\s*\n\s*#',
        fileTxt
    ) else '未通过'


def _check_telnet_disabled(fileTxt, checkItems):
    """检查是否关闭 Telnet 服务"""
    return '通过' if genCheckOtion(
        r'telnet server disable\s*\n\s*'
        r'telnet ipv6 server disable\s*\n\s*'
        r'undo telnet server-source all-interface\s*\n\s*'
        r'undo telnet ipv6 server-source all-interface',
        fileTxt
    ) else '未通过'


def _check_vpn_instance(fileTxt, checkItems):
    """检查 VPN 实例 OOB / DAD 配置"""
    matchOob = re.search(
        r'#\s*\n\s*'
        r'ip vpn-instance OOB\s*\n\s*'
        r'ipv4-family\s*\n\s*'
        r'(route-distinguisher 200:1\s*\n\s*)?'
        r'#',
        fileTxt, re.IGNORECASE
    )
    matchDad = re.search(
        r'#\s*\n\s*'
        r'ip vpn-instance DAD\s*\n\s*'
        r'ipv4-family\s*\n\s*'
        r'#',
        fileTxt, re.IGNORECASE
    )
    if checkItems['type'] in ['Leaf']:  # Leaf 必须包含 2 个实例
        return '通过' if (matchOob and matchDad) else '未通过'
    else:  # 其他只需要 OOB 实例
        return '通过' if matchOob else '未通过'


def _check_aaa(fileTxt, checkItems):
    """检查 AAA 认证配置"""
    return '通过' if genCheckOtion(
        r'aaa\s*\n\s*'
        r'authentication-scheme default\s*\n\s*'
        r'authentication-mode local\s*\n\s*'
        r'authorization-scheme default\s*\n\s*'
        r'authorization-mode local\s*\n\s*'
        r'accounting-scheme default\s*\n\s*'
        r'accounting-mode none\s*\n\s*'
        r'local-aaa-user password policy administrator\s*\n\s*'
        r'password history record number 0\s*\n\s*'
        r'password alert before-expire 0\s*\n\s*'
        r'undo password alert original\s*\n\s*'
        r'password expire 0\s*\n\s*'
        r'domain default\s*\n\s*'
        r'authentication-scheme default\s*\n\s*'
        r'accounting-scheme default\s*\n\s*'
        r'domain default_admin\s*\n\s*'
        r'authentication-scheme default\s*\n\s*'
        r'accounting-scheme default\s*\n\s*'
        r'local-aaa-user user-name complexity-check disable\s*\n\s*'
        r'(?:local-user (?:admin|nmsuser) password irreversible-cipher \S+\s*\n\s*'
        r'local-user (?:admin|nmsuser) password-force-change disable\s*\n\s*'
        r'local-user (?:admin|nmsuser) privilege level [13]\s*\n\s*'
        r'local-user (?:admin|nmsuser) service-type terminal ssh\s*\n\s*){2}',
        fileTxt
    ) else '未通过'


def _check_bgp(fileTxt, checkItems):
    """检查 BGP 配置（包含 Spine/Leaf 不同模板）"""
    # Spine IPv4 邻居配置格式
    SpineIpv4Bgp = (
        r'group Leaf-IPv4 external\s*\n\s*'
        r'peer Leaf-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group Leaf-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*'
        r'group SuperSpine-IPv4 external\s*\n\s*'
        r'peer SuperSpine-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group SuperSpine-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*#\s*\n\s*ipv4-family unicast'
        r'(:?\s*\n\s*network \d+\.\d+\.\d+\.\d+ \d+\.\d+\.\d+\.\d+)+'
        r'\s*\n\s*maximum load-balancing 32\s*\n\s*'
        r' peer Leaf-IPv4 enable\s*\n\s*'
        r'peer Leaf-IPv4 route-policy To-Server-Leaf export\s*\n\s*'
        r'peer Leaf-IPv4 advertise-community\s*\n\s*'
        r'  peer Leaf-IPv4 route-update-interval 0'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ enable\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group Leaf-IPv4)+'
        r'\s*\n\s*'
        r'peer SuperSpine-IPv4 enable\s*\n\s*'
        r'peer SuperSpine-IPv4 advertise-community\s*\n\s*'
        r'peer SuperSpine-IPv4 route-update-interval 0'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ enable\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group SuperSpine-IPv4)+'
        r'[\s\S]*?'
        r'route-policy To-Server-Leaf permit node 10\s*\n\s*'
        r'apply as-path \d+ overwrite'
    )
    # Spine VPNv4 邻居配置格式
    spineVpnv4Bgp = (
        r'#\s*\n\s*ipv4-family unicast\s*\n\s*#\s*\n\s*'
        r'ipv4-family vpn-instance CMB-PRD-CRI'
        r'(:?\s*\n\s*network \d+\.\d+\.\d+\.\d+ \d+\.\d+\.\d+\.\d+)+'
        r'\s*\n\s*maximum load-balancing 32\s*\n\s*'
        r'group Leaf-IPv4 external\s*\n\s*'
        r'peer Leaf-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group Leaf-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*'
        r'peer Leaf-IPv4 route-policy To-Server-Leaf export\s*\n\s*'
        r'peer Leaf-IPv4 advertise-community\s*\n\s*'
        r'peer Leaf-IPv4 route-update-interval 0\s*\n\s*'
        r'group SuperSpine-IPv4 external\s*\n\s*'
        r'peer SuperSpine-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group SuperSpine-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*'
        r'peer SuperSpine-IPv4 route-policy To-SuperSpine-CMB-PRD-CRI-IPv4 export\s*\n\s*'
        r'peer SuperSpine-IPv4 advertise-community\s*\n\s*'
        r'peer SuperSpine-IPv4 route-update-interval 0\s*\n\s*#\s*\n\s*'
        r'ipv4-family vpn-instance CMB-PRD-STD'
        r'\s*\n\s*maximum load-balancing 32\s*\n\s*'
        r'group Leaf-IPv4 external\s*\n\s*'
        r'peer Leaf-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group Leaf-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*'
        r'peer Leaf-IPv4 route-policy To-Server-Leaf export\s*\n\s*'
        r'peer Leaf-IPv4 advertise-community\s*\n\s*'
        r'peer Leaf-IPv4 route-update-interval 0\s*\n\s*'
        r'group SuperSpine-IPv4 external\s*\n\s*'
        r'peer SuperSpine-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group SuperSpine-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*'
        r'peer SuperSpine-IPv4 advertise-community\s*\n\s*'
        r'peer SuperSpine-IPv4 route-update-interval 0'
        r'[\s\S]*?'
        r'route-policy To-Server-Leaf permit node 10\s*\n\s*'
        r'apply as-path \d+ overwrite'
        r'[\s\S]*?'
        r'route-policy To-SuperSpine-CMB-PRD-CRI-IPv4 deny node 10\s*\n\s*'
        r'if-match ip-prefix Static-SLF-CMB-PRD-CRI-IPv4'
    )
    # Leaf 邻居配置格式
    leafBgp = (
        r'group Spine-IPv4 external\s*\n\s*'
        r'peer Spine-IPv4 as-number \d+'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ as-number \d+\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group Spine-IPv4\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ description \S+)+'
        r'\s*\n\s*#\s*\n\s*ipv4-family unicast'
        r'(:?\s*\n\s*network \d+\.\d+\.\d+\.\d+ \d+\.\d+\.\d+\.\d+)+'
        r'\s*\n\s*maximum load-balancing 32\s*\n\s*'
        r' peer Spine-IPv4 enable\s*\n\s*'
        r'peer Spine-IPv4 advertise-community\s*\n\s*'
        r'  peer Spine-IPv4 route-update-interval 0'
        r'(:?\s*\n\s*peer \d+\.\d+\.\d+\.\d+ enable\s*\n\s*'
        r'peer \d+\.\d+\.\d+\.\d+ group Spine-IPv4)+'
    )
    matchPart = re.compile(
        r'bgp \d+\s*\n\s*'
        r'router-id \d+\.\d+\.\d+\.\d+\s*\n\s*'
        r'timer keepalive 30 hold 90\s*\n\s*'
        r'advertise lowest-priority all-address-family peer-up delay 120\s*\n\s*'
        r'private-4-byte-as disable\s*\n\s*'
        rf'(?:(?={SpineIpv4Bgp})|(?={spineVpnv4Bgp})|(?={leafBgp}))',
        re.IGNORECASE
    )
    return '通过' if matchPart.search(fileTxt) else '未通过'


def _check_snmp(fileTxt, checkItems):
    """检查 SNMP Agent 配置"""
    return '通过' if genCheckOtion(
        r'snmp-agent\s*\n\s*'
        r'snmp-agent local-engineid \S+\s*\n\s*'
        r'snmp-agent community read cipher.*?mib-view iso-view.*?\s*\n\s*'
        r'#\s*\n\s*'
        r'snmp-agent sys-info location.*?\s*\n\s*'
        r'snmp-agent sys-info version v2c v3\s*\n\s*'
        r'snmp-agent community complexity-check disable\s*\n\s*'
        r'#\s*\n\s*'
        r'snmp-agent usm-user password complexity-check disable\s*\n\s*'
        r'snmp-agent mib-view included iso-view iso\s*\n\s*'
        r'#\s*\n\s*'
        r'snmp-agent blacklist ip-block disable\s*\n\s*'
        r'#\s*\n\s*'
        r'snmp-agent protocol source-status all-interface\s*\n\s*'
        r'undo snmp-agent protocol source-status ipv6 all-interface\s*\n\s*'
        r'#\s*\n\s*'
        r'undo snmp-agent proxy protocol source-status all-interface\s*\n\s*'
        r'undo snmp-agent proxy protocol source-status ipv6 all-interface',
        fileTxt
    ) else '未通过'


def _check_lldp(fileTxt, checkItems):
    """检查 LLDP 使能配置"""
    return '通过' if genCheckOtion(
        r'#\s*\n\s*lldp enable\s*\n\s*#', fileTxt
    ) else '未通过'


def _check_ssh(fileTxt, checkItems):
    """检查 SSH 服务端/客户端加密配置"""
    if genCheckOtion(
        r'stelnet server enable\s*\n\s*'
        r'ssh server rsa-key min-length 3072\s*\n\s*'
        r'ssh server-source all-interface\s*\n\s*'
        r'undo ssh ipv6 server-source all-interface\s*\n\s*'
        r'ssh authorization-type default aaa\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh server cipher aes256_gcm aes128_gcm aes256_ctr aes192_ctr aes128_ctr '
        r'aes256_cbc aes192_cbc aes128_cbc arcfour256 arcfour128 3des_cbc blowfish_cbc des_cbc\s*\n\s*'
        r'ssh server hmac sha2_512 sha2_256_96 sha2_256 sha1 sha1_96 md5 md5_96\s*\n\s*'
        r'ssh server key-exchange dh_group_exchange_sha256 dh_group_exchange_sha1 '
        r'dh_group14_sha1 dh_group1_sha1 ecdh_sha2_nistp256 ecdh_sha2_nistp384 '
        r'ecdh_sha2_nistp521 sm2_kep dh_group16_sha512 curve25519_sha256\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh server publickey dsa ecc rsa rsa_sha2_256 rsa_sha2_512\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh server dh-exchange min-len 2048\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh client publickey ecc rsa_sha2_256 rsa_sha2_512\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh client cipher aes256_gcm aes128_gcm aes256_ctr aes192_ctr aes128_ctr\s*\n\s*'
        r'ssh client hmac sha2_512 sha2_256\s*\n\s*'
        r'ssh client key-exchange dh_group_exchange_sha256 dh_group16_sha512',
        fileTxt
    ) or genCheckOtion(
        r'stelnet server enable\s*\n\s*'
        r'ssh server rsa-key min-length 3072\s*\n\s*'
        r'undo ssh server authentication-type keyboard-interactive enable\s*\n\s*'
        r'ssh server-source all-interface\s*\n\s*'
        r'undo ssh ipv6 server-source all-interface\s*\n\s*'
        r'ssh authorization-type default aaa\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh server cipher aes256_gcm aes128_gcm aes256_ctr aes192_ctr aes128_ctr '
        r'aes256_cbc aes128_cbc 3des_cbc\s*\n\s*'
        r'ssh server hmac sha2_512 sha2_256_96 sha2_256 sha1 sha1_96 md5 md5_96\s*\n\s*'
        r'ssh server key-exchange dh_group_exchange_sha256 dh_group_exchange_sha1 '
        r'dh_group14_sha1 dh_group1_sha1 ecdh_sha2_nistp256 ecdh_sha2_nistp384 '
        r'ecdh_sha2_nistp521 sm2_kep dh_group16_sha512\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh server publickey dsa ecc rsa rsa_sha2_256 rsa_sha2_512\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh server dh-exchange min-len 2048\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh client publickey ecc rsa_sha2_256 rsa_sha2_512\s*\n\s*'
        r'#\s*\n\s*'
        r'ssh client cipher aes256_gcm aes128_gcm aes256_ctr aes192_ctr aes128_ctr\s*\n\s*'
        r'ssh client hmac sha2_512 sha2_256\s*\n\s*'
        r'ssh client key-exchange dh_group_exchange_sha256 dh_group16_sha512\s*\n\s*',
        fileTxt
    ):
        return '通过'
    return '未通过'


def _check_cmd_privilege(fileTxt, checkItems):
    """检查命令权限配置"""
    return '通过' if genCheckOtion(
        r'command-privilege level 1 view shell dir\s*\n\s*'
        r'command-privilege level 1 view global display\s*\n\s*'
        r'command-privilege level 1 view shell save\s*\n\s*'
        r'command-privilege level 1 view shell screen-length',
        fileTxt
    ) else '未通过'


def _check_user_interface(fileTxt, checkItems):
    """检查 VTY / CON 口登录配置"""
    return '通过' if genCheckOtion(
        r'user-interface maximum-vty 21\s*\n\s*'
        r'#\s*\n\s*'
        r'user-interface con 0\s*\n\s*'
        r'authentication-mode password\s*\n\s*'
        r'set authentication password cipher \S+\s*\n\s*'
        r'idle-timeout 10 0\s*\n\s*'
        r'#\s*\n\s*'
        r'user-interface vty 0 20\s*\n\s*'
        r'authentication-mode aaa\s*\n\s*'
        r'user privilege level 3\s*\n\s*'
        r'protocol inbound ssh',
        fileTxt
    ) else '未通过'


def _check_hash(fileTxt, checkItems):
    """检查负载均衡 ECMP hash 模式"""
    return '通过' if genCheckOtion(
        r'#\s*\n\s*'
        r'load-balance ecmp\s*\n\s*'
        r'hashmode (underlay)? 2\s*\n\s*'
        r'#', fileTxt
    ) else '未通过'


def _check_oob_interface(fileTxt, checkItems):
    """检查带外接口（MEth）配置"""
    return '通过' if genCheckOtion(
        r'interface MEth\S+\s*\n\s*'
        r'description Out-Of-OOB\s*\n\s*'
        r'ip binding vpn-instance OOB\s*\n\s*'
        r'ip address \d+\.\d+\.\d+\.\d+ 255.255.255.0',
        fileTxt
    ) else '未通过'


def _check_loopback(fileTxt, checkItems):
    """检查 LoopBack1 接口配置"""
    return '通过' if genCheckOtion(
        r'interface LoopBack1\s*\n\s*'
        r'description \S+\s*\n\s*'
        r'(?:ip binding vpn-instance\s+\S+\s*\n\s*)?'
        r'ip address \d+\.\d+\.\d+\.\d+ 255.255.255.255',
        fileTxt
    ) else '未通过'


def _check_peerlink(fileTxt, checkItems):
    """检查 peer-link (Eth-Trunk100) 配置"""
    matchPeerlinkEth = re.search(
        r'interface Eth-Trunk100\s*\n\s*'
        r'description To_\S+_Eth-Trunk100\s*\n\s*'
        r'mode lacp-static\s*\n\s*'
        r'peer-link 1',
        fileTxt, re.IGNORECASE
    )
    matchPeerlinkPort = re.findall(
        r'interface 100GE1/0/[37]\s*\n\s*'
        r'description To_\S+_100GE1/0/[37]\s*\n\s*'
        r'eth-trunk 100',
        fileTxt, re.IGNORECASE
    )
    matchPeerlinkPortSlf = re.findall(
        r'interface 25GE1/0/\d{2}\s*\n\s*'
        r'description To_\S+_25GE1/0/\d{2}\s*\n\s*'
        r'eth-trunk 100',
        fileTxt, re.IGNORECASE
    )
    if checkItems['type'] in ['Slf']:
        return '通过' if (matchPeerlinkEth and len(matchPeerlinkPortSlf) == 8) else '未通过'
    else:
        return '通过' if (matchPeerlinkEth and len(matchPeerlinkPort) == 2) else '未通过'


def _check_dad(fileTxt, checkItems):
    """检查 DAD 检测链路 (Eth-Trunk101) 配置"""
    matchDadEth = re.search(
        r'interface Eth-Trunk101\s*\n\s*'
        r'undo portswitch\s*\n\s*'
        r'description To_\S+_Eth-Trunk101\s*\n\s*'
        r'ip binding vpn-instance DAD\s*\n\s*'
        r'ip address \d+\.\d+\.\d+\.\d+ 255.255.255.252\s*\n\s*'
        r'mode lacp-static\s*\n\s*'
        r'm-lag unpaired-port reserved',
        fileTxt, re.IGNORECASE
    )
    matchDadPort = re.findall(
        r'interface 25GE1/0/4[678]\s*\n\s*'
        r'description To_\S+_25GE1/0/4[678]\s*\n\s*'
        r'eth-trunk 101',
        fileTxt, re.IGNORECASE
    )
    return '通过' if (matchDadEth and len(matchDadPort) == 2) else '未通过'


def _check_monitor_link(fileTxt, checkItems):
    """检查 monitor-link 联动组配置（Leaf / Slf 不同模板）"""
    matchMonitorLinkInfo = genCheckOtion(
        r'monitor-link group 1\s*\n\s*'
        r'port 100GE1/0/1 uplink\s*\n\s*'
        r'port 100GE1/0/2 uplink\s*\n\s*'
        r'port 100GE1/0/5 uplink\s*\n\s*'
        r'port 100GE1/0/6 uplink\s*\n\s*'
        r'port 25GE1/0/1 downlink 1\s*\n\s*'
        r'port 25GE1/0/2 downlink 2\s*\n\s*'
        r'port 25GE1/0/3 downlink 3\s*\n\s*'
        r'port 25GE1/0/4 downlink 4\s*\n\s*'
        r'port 25GE1/0/5 downlink 5\s*\n\s*'
        r'port 25GE1/0/6 downlink 6\s*\n\s*'
        r'port 25GE1/0/7 downlink 7\s*\n\s*'
        r'port 25GE1/0/8 downlink 8\s*\n\s*'
        r'port 25GE1/0/9 downlink 9\s*\n\s*'
        r'port 25GE1/0/10 downlink 10\s*\n\s*'
        r'port 25GE1/0/11 downlink 11\s*\n\s*'
        r'port 25GE1/0/12 downlink 12\s*\n\s*'
        r'port 25GE1/0/13 downlink 13\s*\n\s*'
        r'port 25GE1/0/14 downlink 14\s*\n\s*'
        r'port 25GE1/0/15 downlink 15\s*\n\s*'
        r'port 25GE1/0/16 downlink 16\s*\n\s*'
        r'port 25GE1/0/17 downlink 17\s*\n\s*'
        r'port 25GE1/0/18 downlink 18\s*\n\s*'
        r'port 25GE1/0/19 downlink 19\s*\n\s*'
        r'port 25GE1/0/20 downlink 20\s*\n\s*'
        r'port 25GE1/0/21 downlink 21\s*\n\s*'
        r'port 25GE1/0/22 downlink 22\s*\n\s*'
        r'port 25GE1/0/23 downlink 23\s*\n\s*'
        r'port 25GE1/0/24 downlink 24\s*\n\s*'
        r'port 25GE1/0/25 downlink 25\s*\n\s*'
        r'port 25GE1/0/26 downlink 26\s*\n\s*'
        r'port 25GE1/0/27 downlink 27\s*\n\s*'
        r'port 25GE1/0/28 downlink 28\s*\n\s*'
        r'port 25GE1/0/29 downlink 29\s*\n\s*'
        r'port 25GE1/0/30 downlink 30\s*\n\s*'
        r'port 25GE1/0/31 downlink 31\s*\n\s*'
        r'port 25GE1/0/32 downlink 32\s*\n\s*'
        r'port 25GE1/0/33 downlink 33\s*\n\s*'
        r'port 25GE1/0/34 downlink 34\s*\n\s*'
        r'port 25GE1/0/35 downlink 35\s*\n\s*'
        r'port 25GE1/0/36 downlink 36\s*\n\s*'
        r'port 25GE1/0/37 downlink 37\s*\n\s*'
        r'port 25GE1/0/38 downlink 38\s*\n\s*'
        r'port 25GE1/0/39 downlink 39\s*\n\s*'
        r'port 25GE1/0/40 downlink 40\s*\n\s*'
        r'port 25GE1/0/41 downlink 41\s*\n\s*'
        r'port 25GE1/0/42 downlink 42\s*\n\s*'
        r'port 25GE1/0/43 downlink 43\s*\n\s*'
        r'port 25GE1/0/44 downlink 44\s*\n\s*'
        r'port 25GE1/0/45 downlink 45\s*\n\s*'
        r'port 25GE1/0/4[68] downlink 4[68]\s*\n\s*'
        r'timer recover-time 40',
        fileTxt
    )
    matchMonitorLinkSlfInfo = genCheckOtion(
        r'monitor-link group 1\s*\n\s*'
        r'port 100GE1/0/1 uplink\s*\n\s*'
        r'port 100GE1/0/2 uplink\s*\n\s*'
        r'port 100GE1/0/5 uplink\s*\n\s*'
        r'port 100GE1/0/6 uplink\s*\n\s*'
        r'port 25GE1/0/1 downlink 1\s*\n\s*'
        r'port 25GE1/0/2 downlink 2\s*\n\s*'
        r'port 25GE1/0/3 downlink 3\s*\n\s*'
        r'port 25GE1/0/4 downlink 4\s*\n\s*'
        r'port 25GE1/0/5 downlink 5\s*\n\s*'
        r'port 25GE1/0/6 downlink 6\s*\n\s*'
        r'port 25GE1/0/7 downlink 7\s*\n\s*'
        r'port 25GE1/0/8 downlink 8\s*\n\s*'
        r'port 25GE1/0/9 downlink 9\s*\n\s*'
        r'port 25GE1/0/10 downlink 10\s*\n\s*'
        r'port 25GE1/0/11 downlink 11\s*\n\s*'
        r'port 25GE1/0/12 downlink 12\s*\n\s*'
        r'port 25GE1/0/13 downlink 13\s*\n\s*'
        r'port 25GE1/0/14 downlink 14\s*\n\s*'
        r'port 25GE1/0/15 downlink 15\s*\n\s*'
        r'port 25GE1/0/16 downlink 16\s*\n\s*'
        r'port 25GE1/0/17 downlink 17\s*\n\s*'
        r'port 25GE1/0/18 downlink 18\s*\n\s*'
        r'port 25GE1/0/19 downlink 19\s*\n\s*'
        r'port 25GE1/0/20 downlink 20\s*\n\s*'
        r'port 25GE1/0/21 downlink 21\s*\n\s*'
        r'port 25GE1/0/22 downlink 22\s*\n\s*'
        r'port 25GE1/0/23 downlink 23\s*\n\s*'
        r'port 25GE1/0/24 downlink 24\s*\n\s*'
        r'port 25GE1/0/25 downlink 25\s*\n\s*'
        r'port 25GE1/0/26 downlink 26\s*\n\s*'
        r'port 25GE1/0/27 downlink 27\s*\n\s*'
        r'port 25GE1/0/28 downlink 28\s*\n\s*'
        r'port 25GE1/0/29 downlink 29\s*\n\s*'
        r'port 25GE1/0/30 downlink 30\s*\n\s*'
        r'port 25GE1/0/31 downlink 31\s*\n\s*'
        r'port 25GE1/0/32 downlink 32\s*\n\s*'
        r'port 25GE1/0/33 downlink 33\s*\n\s*'
        r'port 25GE1/0/34 downlink 34\s*\n\s*'
        r'port 25GE1/0/35 downlink 35\s*\n\s*'
        r'port 25GE1/0/36 downlink 36\s*\n\s*'
        r'port 25GE1/0/37 downlink 37\s*\n\s*'
        r'port 100GE1/0/3 downlink 103\s*\n\s*'
        r'port 100GE1/0/7 downlink 107\s*\n\s*'
        r'timer recover-time 40',
        fileTxt
    )
    if checkItems['type'] in ['Slf']:
        return '通过' if matchMonitorLinkSlfInfo else '未通过'
    else:
        return '通过' if matchMonitorLinkInfo else '未通过'


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
    '关闭FTP配置':          _check_ftp_disabled,
    'mlag状态':            _check_mlag_status,
    'mlag配置':            _check_mlag_config,
    '大路由配置':           _check_large_route,
    'NTP配置':             _check_ntp,
    '全局vlan配置':         _check_vlan_global,
    'mac飘移配置':         _check_mac_flapping,
    'STP配置':             _check_stp,
    'arp冲突配置':         _check_arp_conflict,
    'telnet关闭配置':       _check_telnet_disabled,
    'vpn实例配置':          _check_vpn_instance,
    'aaa配置':             _check_aaa,
    'BGP配置':             _check_bgp,
    'snmp配置':            _check_snmp,
    'LLDP配置':            _check_lldp,
    'ssh配置':             _check_ssh,
    'cmd权限配置':          _check_cmd_privilege,
    'user-interface配置':  _check_user_interface,
    'hash配置':            _check_hash,
    '带外接口配置':         _check_oob_interface,
    'loopback配置':        _check_loopback,
    'peerlink配置':        _check_peerlink,
    'DAD配置':             _check_dad,
    'monitor-link配置':    _check_monitor_link,
    '告警检查':            _check_alarm_active,
}


def genCheckOtion(reinfo, fileTxt):
    """
    通用正则匹配检查工具。
    在配置文本中搜索给定正则模式，匹配即返回 match 对象，否则返回 None。
    """
    return re.search(r'%s' % reinfo, fileTxt, re.IGNORECASE)


if __name__ == '__main__':
    pass