#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
compare_configs.py — 配置下发验证比对工具

功能：
    将预期配置（config_intended/*.cfg）与设备实际采集配置（config/*.log）
    进行差异分析，找出漏配（预期有实际没有）和多配（实际有预期没有）。
    结果输出 Markdown 报告到 data/ 目录。

使用方式：
    python compare_configs.py

依赖：
    PyYAML (pip install PyYAML)

流程：
    1. 加载规则文件 (read/compare_rules.yaml)
    2. 匹配设备（按设备名连接 .cfg 和 .log）
    3. 解析配置（按 # 分块 → 按规则拆分子块 → 分类）
    4. 清洗（噪声过滤 + 漏配/多配忽略）
    5. 对比（无顺序逐行）
    6. 输出报告
"""

import os
import re
import time
from datetime import datetime

import yaml

# ============================================================
# 路径常量
# ============================================================
INTENDED_DIR = 'read/config_intended'   # 预期配置目录（.cfg）
COLLECTED_DIR = 'read/config'           # 采集配置目录（.log）
RULES_PATH = 'read/compare_rules.yaml'   # 规则文件
OUTPUT_DIR = 'data'                      # 输出目录


# ============================================================
# 工具函数
# ============================================================

def _read_file(path):
    """
    读取文件内容，统一处理编码和换行符。

    参数：
        path: 文件路径（相对项目根目录）

    返回：
        str — 文件内容（\n 换行），读取失败返回 None
    """
    try:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception:
        return None


def _compiled_rules(rules_list):
    """
    将规则列表中的 pattern 字符串编译为正则对象。

    参数：
        rules_list: list，每个元素可以是 str 或 dict（含 'pattern' 字段）

    返回：
        list of compiled re.Pattern
    """
    compiled = []
    for rule in rules_list:
        if isinstance(rule, str):
            compiled.append(re.compile(rule, re.IGNORECASE))
        elif isinstance(rule, dict) and 'pattern' in rule:
            compiled.append(re.compile(rule['pattern'], re.IGNORECASE))
    return compiled


# ============================================================
# 1. 加载规则
# ============================================================

def load_rules(rules_path):
    """
    加载 YAML 规则文件。

    参数：
        rules_path: 规则文件路径

    返回：
        dict — 包含 split_sections / section_categories / noise /
                ignore_missing / ignore_extra 等字段

    规则文件加载失败时返回空规则字典，保证主流程不崩溃。
    """
    content = _read_file(rules_path)
    if not content:
        print(f'[警告] 规则文件 {rules_path} 读取失败，使用默认规则')
        return {
            'split_sections': [],
            'section_categories': {},
            'noise': {'blocks': [], 'lines': []},
            'ignore_missing': {'global_lines': [], 'block_lines': []},
            'ignore_extra': {'global_lines': [], 'block_lines': []},
        }

    rules = yaml.safe_load(content) or {}
    return rules


# ============================================================
# 2. 设备配对
# ============================================================

def _get_device_name_from_cfg(filename):
    """
    从 .cfg 文件名提取设备名。

    参数：
        filename: 如 'SZBL1D4FC01U29-DCN-DC1_BMC-ACC.cfg'

    返回：
        str — 设备名 'SZBL1D4FC01U29-DCN-DC1_BMC-ACC'
    """
    return filename.replace('.cfg', '')


def _get_device_name_from_log(filename):
    """
    从 .log 文件名提取设备名。

    .log 文件名格式为「管理IP_设备名.log」，
    取第一个下划线之后的部分作为设备名。

    参数：
        filename: 如 '12.255.190.252_SZBL1D4FD04U44-OBN-DCN-ACC.log'

    返回：
        str — 设备名 'SZBL1D4FD04U44-OBN-DCN-ACC'
    """
    name = filename.replace('.log', '')
    # 取第一个 _ 之后的部分
    parts = name.split('_', 1)
    return parts[1] if len(parts) > 1 else parts[0]


def match_devices(intended_dir, collected_dir):
    """
    将 config_intended/*.cfg 与 config/*.log 按设备名配对。

    返回三个列表：
        matched:  [(设备名, cfg_path, log_path), ...]  — 两端都有的
        only_intended: [cfg_path, ...]  — 只有预期没有采集的
        only_collected: [log_path, ...] — 只有采集没有预期的
    """
    # 扫描目录
    intended_files = {}
    collected_files = {}

    if os.path.isdir(intended_dir):
        for f in os.listdir(intended_dir):
            if f.endswith('.cfg') and not f.startswith('.'):
                dev_name = _get_device_name_from_cfg(f)
                intended_files[dev_name] = os.path.join(intended_dir, f)

    if os.path.isdir(collected_dir):
        for f in os.listdir(collected_dir):
            if f.endswith('.log') and not f.startswith('.'):
                dev_name = _get_device_name_from_log(f)
                collected_files[dev_name] = os.path.join(collected_dir, f)

    # 配对
    all_devices = set(intended_files.keys()) | set(collected_files.keys())
    matched = []
    only_intended = []
    only_collected = []

    for dev_name in all_devices:
        cfg_path = intended_files.get(dev_name)
        log_path = collected_files.get(dev_name)
        if cfg_path and log_path:
            matched.append((dev_name, cfg_path, log_path))
        elif cfg_path:
            only_intended.append(cfg_path)
        else:
            only_collected.append(log_path)

    return matched, only_intended, only_collected


# ============================================================
# 3. 从采集日志中提取 current-configuration 段
# ============================================================

def extract_collected_config(text):
    """
    从原始采集日志中提取 display current-configuration 的输出段。

    日志格式：
        ...（之前可能有其他命令的输出）
        display current-configuration
        !Software Version V200R019C10SPC500       ← 从这里开始
        #
        sysname XXX
        #
        ...
        #
        return
        #                              ← 配置段到此结束
        <设备名>                         ← 下一个命令提示符，截断点

    参数：
        text: 原始日志全文

    返回：
        str — 提取出的配置段，找不到时返回 None
    """
    # 定位 display current-configuration 命令（可能有 | no 等管道参数）
    match = re.search(
        r'display current-configuration([\s\S]*?return)',
        text, re.IGNORECASE
    )
    return match.group(1).strip() if match else None


def extract_model_version(log_text):
    """
    从原始采集日志中提取设备型号和版本号。

    从 display version 命令的输出中提取：
        - 型号：HUAWEI CE8860-4C-EI uptime is... → CE8860-4C-EI
        - 版本：Version 8.191 (CE8860EI V200R019C10SPC800) → V200R019C10SPC800

    参数：
        log_text: 原始日志全文

    返回：
        (model, version) — 均为 str 或 None
    """
    model = None
    version = None

    # 定位 display version 输出段
    ver_section = re.search(
        r'display version[^\n]*\n(.*?)(?=\n\s*<|\Z)',
        log_text, re.IGNORECASE | re.DOTALL
    )
    if ver_section:
        ver_text = ver_section.group(0)
        # 提取型号：HUAWEI <型号> uptime is
        model_match = re.search(
            r'HUAWEI\s+(\S+)\s+(?:Routing Switch\s+)?uptime is',
            ver_text, re.IGNORECASE
        )
        if model_match:
            model = model_match.group(1)
        # 提取版本：Version X.XX (XXX V200R...)
        ver_match = re.search(
            r'Version\s+\S+\s+\(\S+\s+(V?\d+R\d+C\d+(?:SPC\d+)?)\)',
            ver_text, re.IGNORECASE
        )
        if ver_match:
            version = ver_match.group(1)

    return model, version


# ============================================================
# 4. 按 # 分割配置为配置块
# ============================================================

def parse_blocks(text):
    """
    将配置文本按 # 行分割为配置块列表。

    分割规则：
        - 以 # 单独一行为分隔符（忽略行首尾空格）
        - 连续非空行直到下一个 # 行视为一个配置块
        - 过滤掉空块

    参数：
        text: 配置文本（预期配置全文 或 提取出的采集配置段）

    返回：
        list of str — 每项为一个配置块的完整内容（不含首尾 # 行）
    """
    # 按 \n#\n 分割，兼容 \r\n
    text = text.replace('\r\n', '\n')
    raw_blocks = re.split(r'\n\s*#\s*\n', text)

    blocks = []
    for block in raw_blocks:
        block = block.strip()
        if not block:
            continue
        # 如果块的第一行就是单独的 #，去掉这行，剩下的才是配置内容
        block_lines = block.split('\n')
        if block_lines[0].strip() == '#':
            block = '\n'.join(block_lines[1:]).strip()
            if not block:
                continue
        blocks.append(block)

    return blocks


# ============================================================
# 5. 提取段落头 并 按规则拆分子块
# ============================================================

def _get_block_header(block_text):
    """
    获取配置块的第一行作为段落头。

    参数：
        block_text: 配置块文本（多行）

    返回：
        str — 第一行（去除首尾空格）
    """
    lines = block_text.strip().split('\n')
    return lines[0].strip()


def _get_block_body(block_text):
    """
    获取配置块第一行之后的内容（子命令行）。

    参数：
        block_text: 配置块文本（多行）

    返回：
        list of str — 子命令行列表（已去除首尾空格，过滤空行）
    """
    lines = block_text.strip().split('\n')
    body = [line.strip() for line in lines[1:] if line.strip()]
    return body


def _classify_section(header, categories):
    """
    根据段落头匹配中文分类名。

    参数：
        header: 段落头字符串
        categories: dict，{正则模式: 分类名}

    返回：
        str or None — 匹配到的分类名，未匹配返回 None
    """
    for pattern, category in categories.items():
        if re.match(pattern, header, re.IGNORECASE):
            return category
    return None


def _should_split_section(header, split_sections):
    """
    判断段落是否需要进一步拆分子块。

    参数：
        header: 段落头字符串
        split_sections: list of dict [{'header_pattern': '...'}]

    返回：
        bool
    """
    for rule in split_sections:
        if isinstance(rule, dict):
            pattern = rule.get('header_pattern', '')
        else:
            pattern = rule
        if re.match(pattern, header, re.IGNORECASE):
            return True
    return False


def _split_interface_block(header, body_lines):
    """
    将 interface 配置块按接口名拆分为子段。

    一个 interface 块可能包含多个接口定义：
        interface GigabitEthernet0/0/1
         port link-type access
         port default vlan 2104
        interface GigabitEthernet0/0/2
         port link-type trunk
         ...

    参数：
        header: 第一个 interface 的段落头（如 'interface GigabitEthernet0/0/1'）
        body_lines: 子命令行列表

    返回：
        list of (子段头, 子段体) — 每个子段对应一个接口
    """
    sub_blocks = []
    current_header = header
    current_body = []

    for line in body_lines:
        if line.startswith('interface '):
            # 保存上一个子段
            if current_header is not None:
                sub_blocks.append((current_header, current_body))
            current_header = line
            current_body = []
        else:
            current_body.append(line)

    # 保存最后一个子段
    if current_header is not None:
        sub_blocks.append((current_header, current_body))

    return sub_blocks


def _split_generic_block(header, body_lines):
    """
    通用子块拆分：将整个块作为一个子段。

    参数：
        header: 段落头
        body_lines: 子命令行列表

    返回：
        list of (子段头, 子段体)
    """
    return [(header, body_lines)]


def process_blocks(blocks, rules):
    """
    解析配置块，按规则拆分子块并分类，生成结构化配置数据。

    参数：
        blocks: parse_blocks() 输出的原始配置块列表
        rules: 规则字典

    返回：
        dict，格式：
        {
            'global': ['行1', '行2', ...],
            'sections': {
                '端口配置': {
                    'GigabitEthernet0/0/1': ['行1', '行2'],
                    ...
                },
                'AAA配置': {
                    'aaa': ['行1', ...],
                },
                ...
            },
            '_model': str or None,
            '_version': str or None,
        }
    """
    split_sections = rules.get('split_sections', [])
    categories = rules.get('section_categories', {})

    result = {
        'global': [],
        'sections': {},
        '_model': None,
        '_version': None,
    }

    # 全局分类（兜底用）
    GLOBAL_CATEGORY = '全局配置'

    for block in blocks:
        header = _get_block_header(block)
        body_lines = _get_block_body(block)

        # 提取型号和版本信息
        ver_match = re.search(
            r'HUAWEI\s+(\S+)\s+(?:Routing Switch\s+)?uptime is',
            header, re.IGNORECASE
        )
        if ver_match:
            result['_model'] = ver_match.group(1)
            continue

        # 提取版本号（!Software Version 行）
        sw_match = re.search(r'Software Version\s+(\S+)', header)
        if sw_match:
            result['_version'] = sw_match.group(1)
            continue

        # 分类
        category = _classify_section(header, categories)

        if _should_split_section(header, split_sections):
            # 需要拆分的段落
            if header.startswith('interface '):
                sub_blocks = _split_interface_block(header, body_lines)
            else:
                sub_blocks = _split_generic_block(header, body_lines)

            for sub_header, sub_body in sub_blocks:
                cat = category or GLOBAL_CATEGORY
                if cat not in result['sections']:
                    result['sections'][cat] = {}
                # 如果子段体为空，也保留（如接口无子命令的情况）
                result['sections'][cat][sub_header] = sub_body
        else:
            # 不需要拆分的段落 — 放入全局或分类下
            if category and category != GLOBAL_CATEGORY:
                if category not in result['sections']:
                    result['sections'][category] = {}
                result['sections'][category][header] = body_lines
            else:
                # 全局配置 — 每行单独存放
                if header.strip():
                    result['global'].append(header.strip())
                for line in body_lines:
                    if line.strip():
                        result['global'].append(line.strip())

    return result


# ============================================================
# 6. 清洗配置
# ============================================================

def _match_any(line, patterns):
    """
    检查一行是否匹配任意一个正则模式。

    参数：
        line: 待检查的文本行
        patterns: list of compiled re.Pattern

    返回：
        bool
    """
    for pat in patterns:
        if pat.search(line):
            return True
    return False


def _line_in_ignored(header, line, ignore_block_rules):
    """
    检查某行是否匹配段落级别的忽略规则。

    参数：
        header: 段落头（如 'GigabitEthernet0/0/1'）
        line: 子命令行
        ignore_block_rules: list of dict {header_pattern, lines[[pattern]]}

    返回：
        bool
    """
    for rule in ignore_block_rules:
        header_pattern = rule.get('header_pattern', '')
        line_patterns = rule.get('lines', [])
        compiled_lines = _compiled_rules(line_patterns)
        if re.match(header_pattern, header, re.IGNORECASE):
            if _match_any(line, compiled_lines):
                return True
    return False


def _filter_by_model_version(rules_list, model, version):
    """
    按设备型号/版本过滤规则列表。

    规则元素可以是 str（纯字符串规则）或 dict（含 pattern/header_pattern 等字段）。
    - str 类型：始终保留（无条件规则，所有设备生效）
    - dict 且未指定 model/version：始终保留（无条件规则）
    - dict 且指定了 model 和/或 version：仅当设备匹配时才保留该规则

    匹配规则：子串包含（例如 rule_model='CE6881' 匹配设备型号 'CE6881-48S6CQ-EI'）。
    model 和 version 同时指定时为 AND 关系（两者都满足才保留）。

    参数：
        rules_list: list — 原始规则列表
        model: str or None — 设备型号
        version: str or None — 设备版本号

    返回：
        list — 过滤后的规则列表
    """
    filtered = []
    for rule in rules_list:
        # 非 dict 类型（如纯 str）→ 无条件保留
        if not isinstance(rule, dict):
            filtered.append(rule)
            continue

        rule_model = rule.get('model')
        rule_version = rule.get('version')

        # 未指定 model 且未指定 version → 无条件保留（全设备生效）
        if rule_model is None and rule_version is None:
            filtered.append(rule)
            continue

        # 检查 model 匹配（子串包含）
        model_ok = True
        if rule_model is not None:
            if model is None:
                model_ok = False
            else:
                model_ok = rule_model in model

        # 检查 version 匹配（子串包含）
        version_ok = True
        if rule_version is not None:
            if version is None:
                version_ok = False
            else:
                version_ok = rule_version in version

        # AND 关系：两者都满足才保留
        if model_ok and version_ok:
            filtered.append(rule)

    return filtered


def _normalize_password(line):
    """
    将配置行中的密码密文归一化为占位符，避免加密结果差异导致误报。

    Huawei 设备每次 display current-configuration 时加密结果可能不同，
    但配置功能完全相同。归一化后两侧密文变成相同的占位符，集合比较
    即可正确判定为一致。

    处理的密文类型：
        - password irreversible-cipher <密文>  → <密码密文>
        - snmp-agent community read cipher <密文> → <SNMP密文>
        - set authentication password cipher <密文> → <认证密文>

    参数：
        line: str — 单行配置文本

    返回：
        str — 归一化后的行
    """
    # password irreversible-cipher 后的密文
    line = re.sub(
        r'(password\s+irreversible-cipher\s+)\S+',
        r'\1<密码密文>', line
    )
    # snmp-agent community read cipher 后的密文
    line = re.sub(
        r'(snmp-agent\s+community\s+read\s+cipher\s+)\S+',
        r'\1<SNMP密文>', line
    )
    # set authentication password cipher 后的密文
    line = re.sub(
        r'(set\s+authentication\s+password\s+cipher\s+)\S+',
        r'\1<认证密文>', line
    )
    return line


def clean_blocks(config_data, rules, side):
    """
    对已解析的配置数据进行清洗。

    清洗顺序：
        1. noise.lines — 在所有位置删除匹配行
        2. noise.blocks — 按段落头删除整块
        3. side-specific — 按 side 应用 ignore_missing 或 ignore_extra
        4. 密码归一化 — 将密文替换为占位符，避免加密结果差异导致误报

    参数：
        config_data: process_blocks() 的输出
        rules: 规则字典
        side: 'intended' 或 'collected'

    返回：
        dict — 清洗后的配置数据
    """
    # 提取设备型号/版本，用于规则过滤
    model = config_data.get('_model')
    version = config_data.get('_version')

    noise = rules.get('noise', {})
    noise_lines = _compiled_rules(
        _filter_by_model_version(noise.get('lines') or [], model, version))
    noise_blocks = _compiled_rules(
        _filter_by_model_version(noise.get('blocks') or [], model, version))

    # 选择本侧忽略规则
    if side == 'intended':
        ignore_key = 'ignore_missing'
    else:
        ignore_key = 'ignore_extra'
    ignore_rules = rules.get(ignore_key, {})
    ignore_global = _compiled_rules(
        _filter_by_model_version(ignore_rules.get('global_lines') or [], model, version))
    ignore_block = _filter_by_model_version(
        ignore_rules.get('block_lines') or [], model, version)

    result = {
        'global': [],
        'sections': {},
        '_model': config_data.get('_model'),
        '_version': config_data.get('_version'),
    }

    # ---- 清洗全局行 ----
    for line in config_data.get('global', []):
        # noise.lines
        if _match_any(line, noise_lines):
            continue
        # noise.blocks
        if _match_any(line, noise_blocks):
            continue
        # 侧特异性忽略
        if _match_any(line, ignore_global):
            continue
        result['global'].append(_normalize_password(line))

    # ---- 清洗各段落 ----
    for category, sub_dict in config_data.get('sections', {}).items():
        result['sections'][category] = {}
        for header, body_lines in sub_dict.items():
            # noise.blocks — 段落头匹配则整个段落丢弃
            if _match_any(header, noise_blocks):
                continue

            cleaned_body = []
            for line in body_lines:
                # noise.lines
                if _match_any(line, noise_lines):
                    continue
                # 段落内忽略
                if _line_in_ignored(header, line, ignore_block):
                    continue
                cleaned_body.append(_normalize_password(line))

            if cleaned_body:
                result['sections'][category][header] = cleaned_body

    return result


# ============================================================
# 7. SSH 密码套件比较（无序集合比较）
# ============================================================

def _parse_ssh_cipher_line(line):
    """
    从 SSH cipher 配置行提取密码套件集合。

    例如：
        'ssh server cipher aes256_gcm aes128_gcm aes256_ctr'
        → {'aes256_gcm', 'aes128_gcm', 'aes256_ctr'}

    参数：
        line: SSH cipher 配置行

    返回：
        set of str 或 None（如果不是 cipher 行）
    """
    # 匹配 ssh server/client cipher / hmac / key-exchange 行
    m = re.match(
        r'(?:ssh\s+(?:server|client)\s+'
        r'(?:cipher|hmac|key-exchange|publickey)\s+)(.*)',
        line, re.IGNORECASE
    )
    if m:
        return set(m.group(1).split())
    return None


def _ssh_cipher_diff(intended_lines, actual_lines):
    """
    两行之间的 SSH 密码套件差异（无序集合比较）。

    参数：
        intended_lines: 预期配置的行列表
        actual_lines: 实际配置的行列表

    返回：
        (missing, extra) — 缺失集合 和 多余集合
    """
    # 将配置行转为 行内容→套件集合 的映射
    intend_sets = {}
    for line in intended_lines:
        s = _parse_ssh_cipher_line(line)
        if s:
            intend_sets[line] = s

    actual_sets = {}
    for line in actual_lines:
        s = _parse_ssh_cipher_line(line)
        if s:
            actual_sets[line] = s

    missing = []
    extra = []

    # 逐行比较
    for i_line, i_set in intend_sets.items():
        # 在实际中找同类命令行
        matched = False
        for a_line, a_set in actual_sets.items():
            i_cmd_match = re.match(
                r'(ssh\s+(?:server|client)\s+(?:cipher|hmac|key-exchange|publickey))',
                i_line, re.IGNORECASE)
            a_cmd_match = re.match(
                r'(ssh\s+(?:server|client)\s+(?:cipher|hmac|key-exchange|publickey))',
                a_line, re.IGNORECASE)
            if i_cmd_match and a_cmd_match:
                cmd_type_i = i_cmd_match.group(1)
                cmd_type_a = a_cmd_match.group(1)
                if cmd_type_i.lower() == cmd_type_a.lower():
                    matched = True
                    # 计算套件差集
                    missing_set = i_set - a_set
                    extra_set = a_set - i_set
                    if missing_set:
                        missing.append(f'{cmd_type_i} 缺失: {" ".join(sorted(missing_set))}')
                    if extra_set:
                        extra.append(f'{cmd_type_a} 多余: {" ".join(sorted(extra_set))}')
                    # 移除已匹配的实际行
                    del actual_sets[a_line]
                    break
        if not matched:
            missing.append(f'{i_line}  (整行缺失)')

    # 剩余未匹配的实际行
    for a_line in actual_sets:
        extra.append(f'{a_line}  (整行多余)')

    return missing, extra


# ============================================================
# 8. 对比逻辑
# ============================================================

def compare_configs(intended, actual, model=None, version=None):
    """
    对比预期配置和实际配置，找出漏配和多配。

    参数：
        intended: 清洗后的预期配置数据
        actual: 清洗后的实际配置数据
        model: 设备型号（从采集日志中提取，可选）
        version: 设备版本（从采集日志中提取，可选）

    返回：
        dict，格式：
        {
            'model': str or None,
            'version': str or None,
            'diffs': {
                '端口配置': {
                    'missing': {
                        'GigabitEthernet0/0/1': ['shutdown', ...],
                    },
                    'extra': {
                        'GigabitEthernet0/0/2': ['description xxx', ...],
                    },
                },
                '全局配置': {
                    'missing': ['行1', '行2', ...],
                    'extra': ['行3', ...],
                },
                ...
            }
        }
    """
    result = {
        'model': model or intended.get('_model') or actual.get('_model'),
        'version': version or intended.get('_version') or actual.get('_version'),
        'diffs': {},
    }

    all_categories = set(list(intended.get('sections', {}).keys())
                         + list(actual.get('sections', {}).keys()))

    for category in sorted(all_categories):
        cat_diff = {}

        intended_sub = intended.get('sections', {}).get(category, {})
        actual_sub = actual.get('sections', {}).get(category, {})

        all_headers = set(list(intended_sub.keys()) + list(actual_sub.keys()))

        missing_headers = {}
        extra_headers = {}
        missing_global = []
        extra_global = []

        if category == '全局配置':
            # 全局配置 — 无顺序集合比较（SSH 密码套件特殊处理）
            intended_lines = intended.get('global', [])
            actual_lines = actual.get('global', [])

            # SSH cipher 专用比较
            ssh_missing, ssh_extra = _ssh_cipher_diff(intended_lines, actual_lines)

            # 剩余行做普通集合比较
            # 过滤掉 SSH cipher 行
            i_remaining = [l for l in intended_lines if not _parse_ssh_cipher_line(l)]
            a_remaining = [l for l in actual_lines if not _parse_ssh_cipher_line(l)]

            i_set = set(i_remaining)
            a_set = set(a_remaining)

            missing_global = sorted(i_set - a_set) + ssh_missing
            extra_global = sorted(a_set - i_set) + ssh_extra
        else:
            # 段落比较（按子段头）
            for header in sorted(all_headers):
                i_lines = intended_sub.get(header, [])
                a_lines = actual_sub.get(header, [])

                i_set = set(i_lines)
                a_set = set(a_lines)

                miss = sorted(i_set - a_set)
                ext = sorted(a_set - i_set)

                if miss:
                    missing_headers[header] = miss
                if ext:
                    extra_headers[header] = ext

        if missing_global:
            cat_diff['missing'] = missing_global
        if extra_global:
            cat_diff['extra'] = extra_global
        if missing_headers:
            cat_diff.setdefault('missing_headers', {}).update(missing_headers)
        if extra_headers:
            cat_diff.setdefault('extra_headers', {}).update(extra_headers)

        if cat_diff:
            result['diffs'][category] = cat_diff

    return result


# ============================================================
# 9. 生成报告
# ============================================================

def _mask_sensitive(text):
    """
    对报告中的敏感信息进行脱敏处理。

    替换明文密码和加密密码为占位符，防止敏感信息写入报告文件。
    """
    # 密码 irreversible-cipher 后面的值
    text = re.sub(
        r'(password irreversible-cipher\s+)\S+',
        r'\1<密文>', text
    )
    # SNMP community cipher
    text = re.sub(
        r'(snmp-agent community read cipher\s+)\S+',
        r'\1<密文>', text
    )
    # set authentication password cipher
    text = re.sub(
        r'(set authentication password cipher\s+)\S+',
        r'\1<密文>', text
    )
    return text


def generate_report(results, timestamp):
    """
    生成 Markdown 格式的比对报告。

    参数：
        results: 列表，每项为 (设备名, 比对结果dict 或 错误信息)
                 比对结果 dict 格式见 compare_configs() 返回值
        timestamp: 时间戳字符串（用于文件名和内文）

    返回：
        str — Markdown 报告全文
    """
    lines = []
    lines.append('# 配置下发验证比对报告\n')
    lines.append(f'- 生成时间: {timestamp}')
    lines.append(f'- 预期配置目录: `{INTENDED_DIR}`')
    lines.append(f'- 采集配置目录: `{COLLECTED_DIR}`')
    lines.append(f'- 忽略规则文件: `{RULES_PATH}`\n')

    for dev_name, result in results:
        if isinstance(result, str):
            # 错误信息（如未采集、解析失败）
            lines.append(f'## {dev_name}')
            lines.append(f'\n> {result}\n')
            continue

        model = result.get('model') or '未知型号'
        version = result.get('version') or '未知版本'
        lines.append(f'## {dev_name}  [{model}]  [{version}]\n')

        diffs = result.get('diffs', {})
        if not diffs:
            lines.append('无差异\n')
            continue

        for category in sorted(diffs.keys()):
            diff = diffs[category]
            lines.append(f'### {category}\n')

            # --- 段落级缺失/多余（按子段头分组） ---
            missing_headers = diff.get('missing_headers', {})
            extra_headers = diff.get('extra_headers', {})

            for header in sorted(set(list(missing_headers.keys())
                                     + list(extra_headers.keys()))):
                lines.append(f'  - **{header}**')
                for miss_line in missing_headers.get(header, []):
                    lines.append(f'    - `[缺失] {miss_line}`')
                for ext_line in extra_headers.get(header, []):
                    lines.append(f'    - `[多余] {ext_line}`')

            # --- 全局级缺失/多余 ---
            missing_global = diff.get('missing', [])
            extra_global = diff.get('extra', [])

            # 缺失的段落（整段缺失）
            missing_headers_only = [h for h in missing_headers.keys()
                                    if h not in extra_headers]
            if missing_headers_only:
                lines.append(f'- 缺失 {len(missing_headers_only)}个:')
                for h in sorted(missing_headers_only):
                    lines.append(f'  - `{h}`')

            extra_headers_only = [h for h in extra_headers.keys()
                                  if h not in missing_headers]
            if extra_headers_only:
                lines.append(f'- 多余 {len(extra_headers_only)}个:')
                for h in sorted(extra_headers_only):
                    lines.append(f'  - `{h}`')

            if missing_global:
                lines.append(f'- 缺失 {len(missing_global)}项:')
                for item in missing_global:
                    lines.append(f'  - `{item}`')

            if extra_global:
                lines.append(f'- 多余 {len(extra_global)}项:')
                for item in extra_global:
                    lines.append(f'  - `{item}`')

            lines.append('')

    # 最后对整个报告做脱敏处理
    return _mask_sensitive('\n'.join(lines))


# ============================================================
# 10. 主流程
# ============================================================

def main():
    """
    主入口。
    1. 加载规则
    2. 匹配设备
    3. 逐个设备对比
    4. 输出报告
    """
    print('=' * 50)
    print('配置下发验证比对工具')
    print('=' * 50)

    # 1. 加载规则
    print(f'\n[1/4] 加载规则文件: {RULES_PATH}')
    rules = load_rules(RULES_PATH)
    print(f'  → 段落拆分规则: {len(rules.get("split_sections", []))} 条')
    print(f'  → 分类规则: {len(rules.get("section_categories", {}))} 条')

    # 2. 设备配对
    print(f'\n[2/4] 设备配对')
    print(f'  → 预期配置目录: {INTENDED_DIR}')
    print(f'  → 采集配置目录: {COLLECTED_DIR}')

    matched, only_intended, only_collected = match_devices(
        INTENDED_DIR, COLLECTED_DIR
    )

    print(f'  → 匹配成功: {len(matched)} 台')
    if only_intended:
        print(f'  → 仅预期有、未采集: {len(only_intended)} 台')
        for path in only_intended:
            print(f'    - {os.path.basename(path)}')
    if only_collected:
        print(f'  → 仅采集有、无预期: {len(only_collected)} 台')
        for path in only_collected:
            print(f'    - {os.path.basename(path)}')

    if not matched:
        print('\n[结果] 没有匹配的设备，无需比对。')
        return

    # 3. 逐个设备对比
    print(f'\n[3/4] 开始比对 {len(matched)} 台设备...')

    results = []
    for idx, (dev_name, cfg_path, log_path) in enumerate(matched, 1):
        print(f'  [{idx}/{len(matched)}] {dev_name}...', end=' ', flush=True)

        # 读取预期配置
        cfg_text = _read_file(cfg_path)
        if cfg_text is None:
            print('❌ 读取预期配置失败')
            results.append((dev_name, '读取预期配置文件失败'))
            continue

        # 读取采集日志
        log_text = _read_file(log_path)
        if log_text is None:
            print('❌ 读取采集日志失败')
            results.append((dev_name, '读取采集日志文件失败'))
            continue

        # 提取 current-configuration 段
        collected_config = extract_collected_config(log_text)
        if collected_config is None:
            print('❌ 未找到 current-configuration 段')
            results.append((dev_name, '未从采集日志中提取到配置段'))
            continue

        # 提取型号版本（从 display version 输出段）
        model, version = extract_model_version(log_text)

        # 解析
        intended_blocks = parse_blocks(cfg_text)
        collected_blocks = parse_blocks(collected_config)

        print(f'分块{len(intended_blocks)}/{len(collected_blocks)}...', end=' ')

        intended_parsed = process_blocks(intended_blocks, rules)
        collected_parsed = process_blocks(collected_blocks, rules)

        # 清洗
        intended_clean = clean_blocks(intended_parsed, rules, 'intended')
        collected_clean = clean_blocks(collected_parsed, rules, 'collected')

        # 对比
        diff_result = compare_configs(intended_clean, collected_clean,
                                      model=model, version=version)

        total_diffs = sum(len(v) for v in diff_result.get('diffs', {}).values())
        print(f'差异{total_diffs}项 ✅')
        results.append((dev_name, diff_result))

    # 4. 输出报告
    print(f'\n[4/4] 生成报告...')

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_text = generate_report(results, timestamp)

    # 写入文件
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    report_filename = f'配置比对报告_{timestamp}.md'
    report_path = os.path.join(OUTPUT_DIR, report_filename)

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report_text)

    print(f'  → 报告已保存: {report_path}')
    print(f'\n{"=" * 50}')
    print(f'比对完成! 共 {len(matched)} 台设备, {sum(1 for _, r in results if isinstance(r, dict) and r.get("diffs"))} 台有差异')
    print(f'{"=" * 50}')


if __name__ == '__main__':
    main()
