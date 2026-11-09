#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
华为设备与软件版本 EOS（停止服务日）后台自动化查询工具
- 运行于项目专属虚拟环境 (.venv)
- 输入与凭据放置于 read/ 目录
- 结果输出放置于 data/ 目录
- 本地增量缓存优先；型号缺失、版本缺失或缓存过期时自动联网补全
- 严格按照表格中的设备型号查询与缓存，不做型号别名映射
- 联网时随机限速（单次请求间隔 2.5~4.0 秒），防止账号被风控
"""

import argparse
import datetime
import json
import os
import random
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

import openpyxl


def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _get_base_dir()
READ_DIR = os.path.join(BASE_DIR, 'read')
DATA_DIR = os.path.join(BASE_DIR, 'data')

COOKIE_FILE = os.path.join(READ_DIR, 'huawei_cookie.txt')
CACHE_FILE = os.path.join(READ_DIR, 'eos_cache.json')
DEFAULT_INPUT_EXCEL = os.path.join(READ_DIR, '设备版本信息.xlsx')

CACHE_META_KEY = '__meta__'
CACHE_TTL_DAYS = 30
UNRELEASED_TTL_DAYS = 7
MISSING_RETRY_DAYS = 7
MODEL_RETRY_DAYS = 1

MIN_DELAY = 2.5
MAX_DELAY = 4.0

USER_AGENT = (
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
    'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
)

HEADER_NAMES = {
    'model': {'网元类型', '设备型号', '型号', '设备类型'},
    'version': {'版本信息', '版本', '软件版本'},
    'eos': {'EOS日期', 'EOS', '停止服务日期'},
}


class CookieError(Exception):
    """Cookie 文件缺失或内容为空"""


class FetchError(Exception):
    """在线接口请求或解析失败"""


class CookieExpiredError(FetchError):
    """接口返回登录页/无权限，Cookie 可能已失效"""


def get_default_output_excel():
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    return os.path.join(DATA_DIR, f'设备版本_EOS日期_{date_str}.xlsx')


def print_banner(input_path, output_path):
    print("=" * 70)
    print("       【华为设备与软件版本 EOS（停止服务日）后台自动化查询工具】")
    print("=" * 70)
    print(f"▶ Cookie 凭据文件: {COOKIE_FILE}（仅在需要联网时读取）")
    print(f"▶ 待处理输入表格 : {input_path}")
    print(f"▶ 结果输出位置   : {output_path}")
    print(f"▶ 本地增量缓存   : {CACHE_FILE}")
    print(
        f"▶ 安全防风控策略 : 缓存优先；型号缺失/版本缺失/过期时联网补全"
        f"（间隔 {MIN_DELAY:.1f}~{MAX_DELAY:.1f} 秒）"
    )
    print("-" * 70)


def load_cookie():
    if not os.path.exists(COOKIE_FILE):
        raise CookieError(
            f"未找到 Cookie 凭据文件: {COOKIE_FILE}\n"
            "请按以下步骤配置凭据：\n"
            "  1. 在浏览器登录华为技术支持网站并打开生命周期查询页面；\n"
            "  2. 按 F12 打开开发者工具，切换到【网络（Network）】标签刷新页面；\n"
            "  3. 复制请求标头中的完整 Cookie 字符串；\n"
            f"  4. 保存到文件: {COOKIE_FILE}"
        )

    with open(COOKIE_FILE, 'r', encoding='utf-8') as f:
        cookie = f.read().strip()

    if not cookie:
        raise CookieError(f"Cookie 凭据文件为空: {COOKIE_FILE}")

    if 'JSESSIONID' not in cookie:
        print("\n[警告] 当前凭据未包含 JSESSIONID，若请求被重定向请重新更新凭据。")
    return cookie


def safe_delay(min_sec=MIN_DELAY, max_sec=MAX_DELAY):
    time.sleep(random.uniform(min_sec, max_sec))


def _request_json(request, action, model_name):
    """发送请求并解析 JSON；网络错误/登录页会抛出明确异常。"""
    try:
        with urllib.request.urlopen(request, timeout=15) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise CookieExpiredError(
                f"{action} 返回 HTTP {e.code}，Cookie 已失效或无权限 ({model_name})"
            ) from e
        raise FetchError(f"{action} 返回 HTTP {e.code} ({model_name})") from e
    except Exception as e:
        raise FetchError(f"{action} 网络异常: {e} ({model_name})") from e

    try:
        text = raw.decode('utf-8').lstrip('\ufeff')
    except UnicodeDecodeError as e:
        raise CookieExpiredError(
            f"{action} 返回内容无法解码，Cookie 可能已失效 ({model_name})"
        ) from e

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        snippet = re.sub(r'\s+', ' ', text)[:100]
        if text.lstrip().lower().startswith(('<!doctype html', '<html')):
            raise CookieExpiredError(
                f"{action} 返回登录页/HTML，Cookie 可能已失效 ({model_name})"
            ) from e
        raise FetchError(f"{action} 返回非 JSON 内容: {snippet} ({model_name})") from e


def _extract_pbi_id(data):
    if not isinstance(data, dict):
        return None

    payload = data.get('data')
    if isinstance(payload, dict):
        pbi_id = payload.get('pbiId')
        if pbi_id:
            return pbi_id
    elif isinstance(payload, list):
        for item in payload:
            if isinstance(item, dict) and item.get('pbiId'):
                return item['pbiId']

    pbi_id = data.get('pbiId')
    return pbi_id or None


def fetch_model_online(model_name, cookie):
    """
    当本地缓存未命中或版本缺失时，通过华为官方后端 API 实时拉取该型号的
    全部版本生命周期。严格按照表格中的原始型号检索，不使用跨型号别名。
    """
    for search_model in (model_name,):
        print(f"    [在线抓取] 正在向华为后台接口实时拉取: {model_name}...")

        # 步骤 1：获取型号的 pbiId
        url_pbi = "https://info.support.huawei.com/info-finder/openSearch/queryPbiIdByKeyword"
        req_pbi = urllib.request.Request(
            url_pbi,
            data=json.dumps(
                {"lang": "zh", "domain": "0", "query": search_model}
            ).encode('utf-8'),
            headers={
                'User-Agent': USER_AGENT,
                'Cookie': cookie,
                'Content-Type': 'application/json;charset=UTF-8',
                'X-Requested-With': 'XMLHttpRequest'
            }
        )
        data = _request_json(req_pbi, "检索产品标识", model_name)
        pbi_id = _extract_pbi_id(data)

        if not pbi_id:
            print(f"      [!] 华为平台未检索到产品编码: {search_model}")
            continue

        safe_delay()

        # 步骤 2：获取该产品下的全部软件版本 ID 集合
        url_vids = "https://info.support.huawei.com/info-finder/eox/getDocVersionListAsy"
        data_vids = urllib.parse.urlencode({
            'domain': '0',
            'tid': pbi_id,
            'offeringName': search_model,
            'permission': 'view',
            'naviTag': 'model',
            'toolId': 'eox',
            'systemType': '2',
            'allVersion': 'false'
        }).encode('utf-8')
        req_vids = urllib.request.Request(
            url_vids,
            data=data_vids,
            headers={
                'User-Agent': USER_AGENT,
                'Cookie': cookie,
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'X-Requested-With': 'XMLHttpRequest'
            }
        )
        res_vids = _request_json(req_vids, "获取版本列表", model_name)
        v_list = (res_vids.get('result') or []) if isinstance(res_vids, dict) else []

        all_vids = []
        for item in v_list:
            if not isinstance(item, dict):
                continue
            version_id = item.get('id')
            if version_id:
                all_vids.append(str(version_id))
            for sub in item.get('subTerms') or []:
                if isinstance(sub, dict) and sub.get('id'):
                    all_vids.append(str(sub['id']))
        all_vids = list(dict.fromkeys(all_vids))

        if not all_vids:
            print(f"      [!] 未找到关联版本记录: {search_model}")
            continue

        safe_delay()

        # 步骤 3：批量提取全部软件版本生命周期
        url_eox = "https://info.support.huawei.com/info-finder/eox/getEoxList"
        data_eox = urllib.parse.urlencode({
            'domain': '0',
            'tid': ",".join(all_vids),
            'eoxType': '',
            'isEdition': 'true',
            'startTime': '',
            'endTime': '',
            'toolId': 'eox',
            'pageIndex': '0',
            'pageSize': '999',
            'offeringName': search_model,
            'systemType': '2',
            'keywords': '',
            'lang': 'zh'
        }).encode('utf-8')
        req_eox = urllib.request.Request(
            url_eox,
            data=data_eox,
            headers={
                'User-Agent': USER_AGENT,
                'Cookie': cookie,
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'X-Requested-With': 'XMLHttpRequest'
            }
        )
        res_eox = _request_json(req_eox, "提取生命周期详情", model_name)
        terms = (res_eox.get('eoxTermList') or []) if isinstance(res_eox, dict) else []

        model_result = {}
        for term in terms:
            if not isinstance(term, dict):
                continue
            raw_version = str(term.get('version') or term.get('name') or '').strip()
            if not raw_version:
                continue

            # 精准提取标准版本标识（如 V100R005、V300R024C00），消除型号前缀残留
            match = re.search(r'(V\d+R\d+[A-Za-z0-9_]*)', raw_version, re.I)
            if match:
                version_name = match.group(1)
            else:
                version_name = (
                    raw_version
                    .replace(search_model, '')
                    .replace(model_name, '')
                    .strip()
                )
            if not version_name:
                continue

            model_result[version_name] = {
                'eom': term.get('eom'),
                'eofs': term.get('eofs'),
                'eos': term.get('eos'),
                'status': term.get('lifecycleStatus')
            }

        if not model_result:
            print(f"      [!] 未解析到有效版本生命周期记录: {search_model}")
            continue

        print(f"      [✓] 在线抓取成功！已拉取 {len(model_result)} 个版本的生命周期数据。")
        return model_result

    print(f"      [!] 未获取到该型号的生命周期数据: {model_name}")
    return {}


def match_version_eos(model, ver_str, cache):
    """根据型号与版本字符串匹配 EOS 日期"""
    version_dict = cache.get(model)
    if not isinstance(version_dict, dict) or not version_dict:
        return None

    version = str(ver_str or '').strip()
    if not version:
        return None

    def get_eos(key):
        info = version_dict.get(key)
        if isinstance(info, dict):
            return info.get('eos')
        return None

    # 1. 精确匹配（兼容大小写）
    if version in version_dict:
        return get_eos(version)
    lower_map = {str(k).lower(): k for k in version_dict}
    matched_key = lower_map.get(version.lower())
    if matched_key is not None:
        return get_eos(matched_key)

    # 2. 去除补丁号匹配（SPC/SPH）
    base_version = re.sub(r'SP[CH]\d+', '', version, flags=re.I).strip()
    matched_key = lower_map.get(base_version.lower())
    if matched_key is not None:
        return get_eos(matched_key)

    # 3. 去除微版本 C 号匹配
    major_version = re.sub(r'C\d+.*', '', version, flags=re.I).strip()
    matched_key = lower_map.get(major_version.lower())
    if matched_key is not None:
        return get_eos(matched_key)

    # 4. 前缀模糊匹配（仅在版本号确实包含 VxRy 时启用，避免空串误匹配）
    if not re.match(r'^V\d+R\d+', base_version, re.I):
        return None
    candidates = [
        key for key in version_dict
        if str(key).lower().startswith(base_version.lower())
    ]
    if not candidates:
        return None
    candidates.sort(key=len)
    return get_eos(candidates[0])


def _cache_meta(cache):
    meta = cache.get(CACHE_META_KEY)
    if not isinstance(meta, dict):
        meta = {}
        cache[CACHE_META_KEY] = meta
    return meta


def _cache_models_meta(cache):
    meta = _cache_meta(cache)
    models_meta = meta.get('models')
    if not isinstance(models_meta, dict):
        models_meta = {}
        meta['models'] = models_meta
    return models_meta


def _cache_missing_versions_meta(cache):
    meta = _cache_meta(cache)
    missing_meta = meta.get('missing_versions')
    if not isinstance(missing_meta, dict):
        missing_meta = {}
        meta['missing_versions'] = missing_meta
    return missing_meta


def _cache_attempts_meta(cache):
    meta = _cache_meta(cache)
    attempts = meta.get('attempts')
    if not isinstance(attempts, dict):
        attempts = {}
        meta['attempts'] = attempts
    return attempts


def _cache_model_count(cache):
    return sum(1 for key in cache if key != CACHE_META_KEY)


def _ensure_cache_meta(cache):
    """为旧缓存补齐时间戳，便于按有效期判断是否需要刷新。"""
    models_meta = _cache_models_meta(cache)
    fallback = None
    if os.path.exists(CACHE_FILE):
        try:
            fallback = datetime.datetime.fromtimestamp(
                os.path.getmtime(CACHE_FILE)
            ).isoformat(timespec='seconds')
        except OSError:
            fallback = None
    if fallback is None:
        fallback = datetime.datetime.now().isoformat(timespec='seconds')

    for model_key, info in cache.items():
        if model_key == CACHE_META_KEY:
            continue
        if isinstance(info, dict) and info and model_key not in models_meta:
            models_meta[model_key] = fallback


def _parse_timestamp(value):
    try:
        return datetime.datetime.fromisoformat(str(value))
    except (TypeError, ValueError):
        return None


def is_model_cache_stale(model_key, cache):
    version_dict = cache.get(model_key)
    if not isinstance(version_dict, dict) or not version_dict:
        return True

    fetched_at = _parse_timestamp(_cache_models_meta(cache).get(model_key))
    if fetched_at is None:
        return False

    age_days = (datetime.datetime.now() - fetched_at).total_seconds() / 86400
    if age_days >= CACHE_TTL_DAYS:
        return True

    # “未发布”状态变化较快，使用更短的有效期
    if age_days >= UNRELEASED_TTL_DAYS:
        for info in version_dict.values():
            if isinstance(info, dict) and str(info.get('eos') or '').strip() == '未发布':
                return True
    return False


def _record_cache_fetch(model_key, cache):
    _cache_models_meta(cache)[model_key] = (
        datetime.datetime.now().isoformat(timespec='seconds')
    )


def _record_cache_attempt(model_key, cache):
    _cache_attempts_meta(cache)[model_key] = (
        datetime.datetime.now().isoformat(timespec='seconds')
    )


def model_attempt_due(model_key, cache):
    checked_at = _parse_timestamp(_cache_attempts_meta(cache).get(model_key))
    if checked_at is None:
        return True
    age_days = (datetime.datetime.now() - checked_at).total_seconds() / 86400
    return age_days >= MODEL_RETRY_DAYS


def missing_versions_due(model_key, versions, cache):
    """版本查询失败后冷却一段时间，避免每次运行都重复请求华为接口。"""
    bucket = _cache_missing_versions_meta(cache).get(model_key)
    if not isinstance(bucket, dict):
        return True

    now = datetime.datetime.now()
    for version in versions:
        checked_at = _parse_timestamp(bucket.get(version))
        if checked_at is None:
            return True
        age_days = (now - checked_at).total_seconds() / 86400
        if age_days >= MISSING_RETRY_DAYS:
            return True
    return False


def _sync_missing_versions(model, versions, cache):
    """记录仍然缺失的版本；已补齐的版本从记录中移除。"""
    model_key = model
    missing_meta = _cache_missing_versions_meta(cache)
    bucket = missing_meta.get(model_key)
    if not isinstance(bucket, dict):
        bucket = {}

    now = datetime.datetime.now().isoformat(timespec='seconds')
    changed = False
    for version in versions:
        if match_version_eos(model, version, cache) is None:
            if bucket.get(version) != now:
                changed = True
            bucket[version] = now
        elif version in bucket:
            bucket.pop(version)
            changed = True

    if bucket:
        missing_meta[model_key] = bucket
    else:
        missing_meta.pop(model_key, None)
    return changed


def load_cache():
    if not os.path.exists(CACHE_FILE):
        print("[2/4] 未发现本地缓存，将自动联网获取并建立缓存")
        return {}

    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            cache = json.load(f)
        if not isinstance(cache, dict):
            raise ValueError('缓存根节点不是 JSON 对象')
    except Exception as e:
        print(f"[2/4] 读取缓存失败 ({e})，将使用空缓存")
        return {}

    _ensure_cache_meta(cache)
    print(f"[2/4] 加载本地历史缓存 ... 包含 {_cache_model_count(cache)} 款设备")
    return cache


def save_cache(cache):
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def _cell_text(ws, row, column):
    value = ws.cell(row=row, column=column).value
    return '' if value is None else str(value).strip()


def resolve_columns(ws):
    found = {}
    for column in range(1, ws.max_column + 1):
        header = _cell_text(ws, 1, column)
        for field, names in HEADER_NAMES.items():
            if header in names:
                found[field] = column

    if len(found) == 3:
        return found['model'], found['version'], found['eos']

    print(
        "[警告] 未识别到标准表头（网元类型/版本信息/EOS日期），"
        "将按 A/B/C 列处理。"
    )
    return 1, 2, 3


def read_records(ws, model_column, version_column):
    records = []
    for row_idx in range(2, ws.max_row + 1):
        model = _cell_text(ws, row_idx, model_column)
        version = _cell_text(ws, row_idx, version_column)
        if model and version:
            records.append((row_idx, model, version))
    return records


def group_records(records):
    groups = {}
    for row_idx, model, version in records:
        groups.setdefault(model, []).append((row_idx, model, version))
    return groups


def main(input_path=None, output_path=None, refresh=False):
    os.makedirs(READ_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)

    input_path = input_path or DEFAULT_INPUT_EXCEL
    output_path = output_path or get_default_output_excel()

    if not os.path.exists(input_path):
        print(f"[错误] 待处理的 Excel 文件不存在: {input_path}")
        print(f"请将待处理的表格放置于: {DEFAULT_INPUT_EXCEL}，或通过参数传入路径。")
        return 1

    print_banner(input_path, output_path)

    wb = None
    try:
        try:
            wb = openpyxl.load_workbook(input_path)
        except Exception as e:
            print(f"[错误] 无法打开 Excel 文件: {e}")
            return 1

        ws = wb.active
        if ws is None:
            print("[错误] 无法获取工作表")
            return 1

        model_column, version_column, eos_column = resolve_columns(ws)
        records = read_records(ws, model_column, version_column)
        total_rows = len(records)
        print(f"[1/4] 读取待处理表格 ... 识别到 {total_rows} 行设备版本记录")
        if total_rows == 0:
            print("[错误] 表格中没有有效的设备版本记录（型号和版本不能为空）")
            return 1

        cache = load_cache()
        groups = group_records(records)

        # 判断哪些型号需要联网：型号缺失、版本缺失、缓存过期或强制刷新
        fetch_items = []
        for cache_key, model_records in groups.items():
            model = model_records[0][1]
            versions = [rec[2] for rec in model_records]
            missing_versions = [
                rec[2] for rec in model_records
                if match_version_eos(rec[1], rec[2], cache) is None
            ]
            cached = cache.get(cache_key)
            if refresh:
                reason = '强制刷新'
            elif not isinstance(cached, dict) or not cached:
                if not model_attempt_due(cache_key, cache):
                    continue
                reason = '型号未缓存'
            elif missing_versions and missing_versions_due(cache_key, missing_versions, cache):
                reason = '版本未命中'
            elif is_model_cache_stale(cache_key, cache):
                reason = '缓存已过期'
            else:
                continue
            fetch_items.append((cache_key, model, reason, versions))

        if fetch_items:
            model_names = '、'.join(item[1] for item in fetch_items[:5])
            if len(fetch_items) > 5:
                model_names += ' 等'
            print(
                f"[3/4] 检查缓存完整性 ... 需要联网补全 {len(fetch_items)} 个型号: "
                f"{model_names}"
            )
            try:
                cookie = load_cookie()
            except CookieError as e:
                print(f"\n[错误] {e}")
                return 1
            print(
                f"       Cookie 已加载，开始在线查询"
                f"（单次请求随机间隔 {MIN_DELAY:.1f}~{MAX_DELAY:.1f} 秒）"
            )
        else:
            cookie = None
            pending_missing = sum(
                1
                for model_records in groups.values()
                for rec in model_records
                if match_version_eos(rec[1], rec[2], cache) is None
            )
            if pending_missing:
                print(
                    f"[3/4] 检查缓存完整性 ... 有 {pending_missing} 条缺失版本"
                    f"处于查询冷却期，本次不重复联网"
                )
            else:
                print("[3/4] 检查缓存完整性 ... 本地缓存已覆盖全部型号与版本，无需联网")

        failed_keys = set()
        cache_dirty = False

        for idx, (cache_key, model, reason, versions) in enumerate(fetch_items):
            print(f"    [{idx + 1}/{len(fetch_items)}] {model}（{reason}）")
            try:
                online_data = fetch_model_online(model, cookie)
            except CookieExpiredError as e:
                print(f"      [!] {e}")
                print("      [!] 已停止后续在线请求，请更新 Cookie 后重试。")
                failed_keys.add(cache_key)
                failed_keys.update(item[0] for item in fetch_items[idx + 1:])
                break
            except FetchError as e:
                print(f"      [!] {e}")
                failed_keys.add(cache_key)
            except Exception as e:
                print(f"      [!] 未预期错误: {e}")
                failed_keys.add(cache_key)
            else:
                _record_cache_attempt(cache_key, cache)
                cache_dirty = True
                if online_data:
                    cache[cache_key] = online_data
                    _record_cache_fetch(cache_key, cache)
                    print(f"      [✓] 在线抓取成功，已更新 {len(online_data)} 个版本。")
                else:
                    print("      [!] 未获取到有效数据，保留原缓存。")

                # 记录仍未命中的版本，冷却期内不再重复请求
                _sync_missing_versions(model, versions, cache)

            if idx < len(fetch_items) - 1:
                safe_delay()

        if cache_dirty:
            try:
                save_cache(cache)
                print(f"      [✓] 在线抓取数据已自动增量写回缓存: {CACHE_FILE}")
            except Exception as e:
                print(f"      [!] 缓存写入失败: {e}")

        # 逐行匹配与回填
        filled_dates = 0
        filled_unreleased = 0
        filled_other = 0
        missing_count = 0
        failed_count = 0

        print("[4/4] 正在批量匹配与回填 EOS 日期 ...")
        for row_idx, model, version in records:
            eos_res = match_version_eos(model, version, cache)
            eos_text = '' if eos_res is None else str(eos_res).strip()
            cell_eos = ws.cell(row=row_idx, column=eos_column)

            if eos_text and eos_text != '未发布':
                try:
                    parts = [int(part) for part in eos_text.split('-')]
                    if len(parts) != 3:
                        raise ValueError('not a date')
                    cell_eos.value = datetime.date(parts[0], parts[1], parts[2])
                    cell_eos.number_format = 'yyyy-mm-dd'
                    filled_dates += 1
                except Exception:
                    cell_eos.value = eos_text
                    filled_other += 1
            elif eos_text == '未发布':
                cell_eos.value = '未发布'
                filled_unreleased += 1
            elif model in failed_keys:
                cell_eos.value = '查询失败'
                failed_count += 1
                print(
                    f"      [!] 第 {row_idx:3d} 行在线查询失败"
                    f"（已标记为'查询失败'）: {model} | {version}"
                )
            else:
                cell_eos.value = '未收录'
                missing_count += 1
                print(
                    f"      [!] 第 {row_idx:3d} 行未匹配到明确 EOS"
                    f"（已标记为'未收录'）: {model} | {version}"
                )

        output_dir = os.path.dirname(os.path.abspath(output_path))
        os.makedirs(output_dir, exist_ok=True)
        try:
            wb.save(output_path)
        except PermissionError:
            print(f"[错误] 无法保存结果文件（可能正被 Excel 占用）: {output_path}")
            return 1
        except Exception as e:
            print(f"[错误] 保存结果文件失败: {e}")
            return 1

        matched_count = filled_dates + filled_unreleased + filled_other
        coverage = matched_count / total_rows * 100

        print("\n" + "=" * 70)
        print("                     【全流程处理完成】")
        print("=" * 70)
        print(f"✔ 成功写入标准日期行数 : {filled_dates} 行")
        print(f"✔ 标注未发布新版本行数 : {filled_unreleased} 行")
        if filled_other:
            print(f"✔ 以原始文本回填行数   : {filled_other} 行")
        print(f"✔ 未能匹配行数         : {missing_count} 行")
        if failed_count:
            print(f"✘ 在线查询失败行数     : {failed_count} 行")
        if failed_keys:
            print(f"✘ 在线补全失败型号数   : {len(failed_keys)} 个")
        print(f"✔ 全量覆盖率           : {coverage:.1f}%")
        print("-" * 70)
        print(f"★ 最终结果文件已生成在 : {output_path}")
        print("=" * 70 + "\n")

        return 1 if failed_keys else 0
    finally:
        if wb is not None:
            wb.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='华为设备与软件版本 EOS 日期自动回填工具'
    )
    parser.add_argument(
        'input_path', nargs='?', default=DEFAULT_INPUT_EXCEL,
        help='输入 Excel 路径（默认 read/设备版本信息.xlsx）'
    )
    parser.add_argument(
        'output_path', nargs='?', default=None,
        help='输出 Excel 路径（默认 data/设备版本_EOS日期_YYYYMMDD.xlsx）'
    )
    parser.add_argument(
        '--refresh', action='store_true',
        help='忽略缓存有效期，强制刷新所有型号'
    )
    args = parser.parse_args()
    sys.exit(main(args.input_path, args.output_path, refresh=args.refresh))