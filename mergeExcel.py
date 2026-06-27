import os
import zipfile
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path
from interface import excel

# ============ 配置区（只改这里）============
CONFIG = {
    # 源数据目录
    "source_dir": "/Users/shadowx/Documents/招行/巡检/2026Q2/巡检报告数据",

    # 功能1+2：提取Excel
    "target_excel_name": "NetWork Healthy Check Report(Engineer).xlsx",
    "output_zip": "data/巡检数据分析汇总",  # 分片压缩输出
    "split_size": "18m",  # 分片大小

    # 功能2：汇总原始数据
    "source_file": "/Users/shadowx/Documents/招行/巡检/2026Q2/巡检设备清单汇总20260603更新.xlsx",
    "output_original": "data/巡检原始数据汇总",

    # 功能3：提取配置文件
    "output_config_dir": "/Users/shadowx/Documents/招行/巡检/设备配置文件汇总",

    # 功能4：合并巡检资产数据
    "output_assets": "data/巡检资产汇总",

    # 临时目录（统一管理）
    "temp_dir": "/Users/shadowx/Documents/招行/巡检/2026Q2/temp_extract",
}


# ==========================================


# ============ 通用函数 ============
def extract_zip_with_chinese(zip_path, target_dir):
    """
    解压ZIP文件，处理中文乱码问题（cp437 -> gbk）
    """
    os.makedirs(target_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for member in zip_ref.namelist():
            try:
                filename = member.encode('cp437').decode('gbk')
            except (UnicodeDecodeError, UnicodeEncodeError):
                filename = member

            dest_path = os.path.join(target_dir, filename)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)

            if not filename.endswith('/'):
                with zip_ref.open(member) as source, open(dest_path, "wb") as target:
                    shutil.copyfileobj(source, target)


def get_zip_files(source_dir):
    """获取目录下所有ZIP文件"""
    if not os.path.exists(source_dir):
        print(f"❌ 错误: 找不到目录 {source_dir}")
        return []
    return [f for f in os.listdir(source_dir) if f.endswith('.zip')]


def clean_temp_dir(temp_dir):
    """清理临时目录"""
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)


def ask_clean_temp(temp_dir):
    """询问是否清理临时目录"""
    if not os.path.exists(temp_dir):
        return
    ans = input(f"\n🗑️  是否删除临时目录 {temp_dir}？(y/N): ").strip().lower()
    if ans == 'y':
        shutil.rmtree(temp_dir)
        print(f"✅ 已删除临时目录")
    else:
        print(f"📁 临时目录保留: {temp_dir}")


def read_xls_file(filepath, sheet_name, header_row=0):
    """
    读取 .xls 文件（旧格式）
    使用 openpyxl 无法读取，需要 xlrd 库
    """
    try:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_name(sheet_name)
        data = []
        for rx in range(header_row, ws.nrows):
            row = []
            for cx in range(ws.ncols):
                row.append(ws.cell_value(rx, cx))
            data.append(row)
        return data
    except ImportError:
        print(f"  ⚠️ 无法读取 .xls 文件，需要安装 xlrd: pip install xlrd")
        return []
    except Exception as e:
        print(f"  ⚠️ 读取 {filepath} 失败: {e}")
        return []


# ==========================================


# ============ 公共解压（功能1/3/4共用）============
def extract_all_zips(config, force_extract=False):
    """
    统一解压所有ZIP文件到临时目录，供功能1/3/4共用
    
    Args:
        config: 配置字典
        force_extract: 是否强制重新解压（默认False，已解压则跳过）
    
    Returns:
        dict: {
            'excel_files': [(源路径, 新文件名), ...],
            'zip_count': int,
            'temp_dir': str
        }
    """
    source_dir = config["source_dir"]
    temp_dir = config["temp_dir"]
    target_name = config["target_excel_name"]

    extracted_files = []

    # 检查临时目录是否已存在且有内容
    if os.path.exists(temp_dir) and os.listdir(temp_dir) and not force_extract:
        print(f"📁 临时目录已存在，跳过解压: {temp_dir}")
        # 遍历已解压的文件
        zip_files = get_zip_files(source_dir)
        for zip_name in zip_files:
            sub_folder = zip_name.replace('.zip', '').strip()
            extract_dir = os.path.join(temp_dir, sub_folder)
            if os.path.exists(extract_dir):
                found_files = list(Path(extract_dir).rglob(target_name))
                for i, file_path in enumerate(found_files):
                    new_name = f"{sub_folder}_{i}_{target_name}"
                    extracted_files.append((file_path, new_name))
        print(f"📋 找到 {len(extracted_files)} 个Excel文件")
        return {"excel_files": extracted_files, "zip_count": len(zip_files), "temp_dir": temp_dir}

    # 需要解压
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    zip_files = get_zip_files(source_dir)
    if not zip_files:
        print("❌ 未发现源压缩包")
        return {"excel_files": [], "zip_count": 0, "temp_dir": temp_dir}

    total_zips = len(zip_files)
    for idx, zip_name in enumerate(zip_files, 1):
        print(f"\r  [{idx}/{total_zips}] 解压中...", end="", flush=True)
        zip_path = os.path.join(source_dir, zip_name)
        sub_folder = zip_name.replace('.zip', '').strip()
        extract_dir = os.path.join(temp_dir, sub_folder)

        # 跳过已存在的目录
        if os.path.exists(extract_dir):
            found_files = list(Path(extract_dir).rglob(target_name))
            for i, file_path in enumerate(found_files):
                new_name = f"{sub_folder}_{i}_{target_name}"
                extracted_files.append((file_path, new_name))
            continue

        try:
            extract_zip_with_chinese(zip_path, extract_dir)
            found_files = list(Path(extract_dir).rglob(target_name))
            for i, file_path in enumerate(found_files):
                new_name = f"{sub_folder}_{i}_{target_name}"
                extracted_files.append((file_path, new_name))
        except Exception as e:
            print(f"\n  ❌ 处理 {zip_name} 失败: {e}")

    print()
    print(f"📋 解压 {total_zips} 个压缩包, 找到 {len(extracted_files)} 个Excel文件")

    return {"excel_files": extracted_files, "zip_count": total_zips, "temp_dir": temp_dir}


# ==========================================


# ============ 功能1：提取打包 ============
def collect_and_split_zip(config, extract_result=None, do_cleanup=True):
    """提取文件并压缩分片"""
    if extract_result is None:
        extract_result = extract_all_zips(config)

    extracted_files = extract_result["excel_files"]
    zip_total = extract_result["zip_count"]

    if not extracted_files:
        print(f"❌ 统计: {zip_total} 个压缩包, 0 个Excel文件")
        return

    collect_dir = os.path.join(config["temp_dir"], "collected_reports")
    if os.path.exists(collect_dir):
        shutil.rmtree(collect_dir)
    os.makedirs(collect_dir)

    total_files = len(extracted_files)
    for i, (file_path, new_name) in enumerate(extracted_files, 1):
        print(f"\r  [{i}/{total_files}] 收集中...", end="", flush=True)
        dest = os.path.join(collect_dir, new_name)
        shutil.copy2(file_path, dest)

    print()

    split_size = config["split_size"]
    output_zip = config["output_zip"]
    today = datetime.now().strftime("%Y-%m-%d")
    output_zip = f"{output_zip}_{today}"

    output_zip_abs = os.path.abspath(output_zip)
    output_dir = os.path.dirname(output_zip)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for ext in ['.zip', '.z01', '.z02', '.z03', '.z04', '.z05']:
        old_file = f"{output_zip_abs}{ext}"
        if os.path.exists(old_file):
            os.remove(old_file)

    original_dir = os.getcwd()
    os.chdir(collect_dir)

    zip_cmd = ["zip", "-s", split_size, "-r", f"{output_zip_abs}.zip", "."]

    try:
        result = subprocess.run(zip_cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(
                f"✅ 完成: {zip_total} 个压缩包 → {len(extracted_files)} 个Excel → {output_zip_abs}.zip (分片: {split_size})")
        else:
            print(f"❌ 压缩失败 (返回码: {result.returncode})")
            print(f"   stderr: {result.stderr}")
    except Exception as e:
        print(f"❌ 调用压缩命令出错: {e}")
    finally:
        os.chdir(original_dir)

    if do_cleanup:
        ask_clean_temp(config["temp_dir"])


# ==========================================


# ============ 功能2：汇总原始数据 ============
def mergeOriginalData(config):
    """巡检前汇总原始表格数据"""
    source_file = config["source_file"]
    output = config["output_original"]

    if not os.path.exists(source_file):
        print(f"❌ 找不到文件: {source_file}")
        return

    data = []
    open_excel = excel(source_file)
    sheet_names = open_excel.excelReadCread()

    sheet_count = 0
    for sheet_name in sheet_names:
        if sheet_name == '汇总':
            continue
        sheet_count += 1

    print(f"📋 读取文件: {os.path.basename(source_file)}, {sheet_count} 个sheet")

    sheet_done = 0
    for sheet_name in sheet_names:
        if sheet_name == '汇总':
            continue
        sheet_done += 1
        print(f"\r  [{sheet_done}/{sheet_count}] 读取sheet...", end="", flush=True)
        res = open_excel.excelReadSheet(sheet=sheet_name, column=5)
        data.extend(res)

    print()
    open_excel.excelClose()

    if not data:
        print(f"❌ 统计: {sheet_count} 个sheet, 0 条数据")
        return

    excel_obj = excel(output)
    excel_obj.excel_write(
        title=['登陆IP', '登陆用户名', 'AB角色', '备注', '时间'],
        data=data,
        sheetname='巡检原始数据'
    )
    excel_obj.save_file()
    print(f"✅ 完成: {sheet_count} 个sheet → {len(data)} 条记录")


# ==========================================


# ============ 功能3：提取配置文件 ============
def merge_zip_txt_files(config, extract_result=None, do_cleanup=True):
    """巡检后汇总压缩包中的配置文件到一个文件夹"""
    source_dir = config["source_dir"]
    output_dir = config["output_config_dir"]
    temp_dir = config["temp_dir"]

    base_path = Path(source_dir).resolve()
    target_path = Path(output_dir).resolve()
    temp_extract_path = Path(temp_dir)

    def _rename_config_file(original_name):
        """重命名配置文件：去掉'设备配置_'前缀，IP部分 - 改 ."""
        name = original_name
        # 去掉 设备配置_ 前缀
        if name.startswith("设备配置_"):
            name = name[len("设备配置_"):]
        # 去掉 .txt 后缀
        if name.endswith(".txt"):
            name = name[:-4]
        # 按 _ 分割，最后一段是 IP，把 - 改成 .
        parts = name.rsplit("_", 1)  # 从右边分割一次
        if len(parts) == 2:
            device, ip = parts
            ip = ip.replace("-", ".")
            return f"{device}_{ip}.txt"
        return original_name  # 格式不匹配就返回原名

    target_path.mkdir(parents=True, exist_ok=True)

    zip_files = list(base_path.glob("*.zip"))
    if not zip_files:
        print("❌ 未找到 .zip 文件")
        return

    print(f"📋 读取 {len(zip_files)} 个压缩包, 提取目录: {target_path}")

    total_count = 0
    skip_count = 0
    err_count = 0

    # 如果未预先解压，则自行解压处理
    if extract_result is None:
        total_zips = len(zip_files)
        for idx, zip_file in enumerate(zip_files, 1):
            print(f"\r  [{idx}/{total_zips}] 提取配置...", end="", flush=True)
            try:
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        try:
                            file_info.filename = file_info.filename.encode('cp437').decode('gbk')
                        except (UnicodeDecodeError, UnicodeEncodeError):
                            pass
                        zip_ref.extract(file_info, temp_extract_path)

                top_folders = [f for f in temp_extract_path.iterdir()
                               if f.is_dir() and not f.name.startswith((".", "__"))]

                if top_folders:
                    cfg_dir = top_folders[0] / "资产报告/设备配置"
                    if cfg_dir.exists():
                        txt_files = list(cfg_dir.glob("*.txt"))
                        for txt_file in txt_files:
                            new_name = _rename_config_file(txt_file.name)
                            shutil.move(str(txt_file), str(target_path / new_name))
                        total_count += len(txt_files)
                    else:
                        skip_count += 1
            except Exception as e:
                err_count += 1
            finally:
                if temp_extract_path.exists():
                    shutil.rmtree(temp_extract_path)
                    os.makedirs(temp_extract_path)
        print()
    else:
        # 已预先解压，直接从临时目录递归查找配置文件
        cfg_dirs = [f for f in temp_extract_path.rglob("设备配置") if f.is_dir()]
        for cfg_dir in cfg_dirs:
            txt_files = list(cfg_dir.glob("*.txt"))
            for txt_file in txt_files:
                new_name = _rename_config_file(txt_file.name)
                shutil.move(str(txt_file), str(target_path / new_name))
            total_count += len(txt_files)
        if not cfg_dirs:
            skip_count = len([f for f in temp_extract_path.iterdir() if f.is_dir()])

    fail_info = f", 失败 {err_count}" if err_count else ""
    skip_info = f", 跳过 {skip_count}" if skip_count else ""
    print(f"✅ 完成: {len(zip_files)} 个压缩包 → {total_count} 个配置文件{skip_info}{fail_info}")
    print(f"   输出目录: {target_path}")

    if do_cleanup:
        ask_clean_temp(str(temp_extract_path))


# ==========================================


# ============ 功能4：合并巡检资产数据 ============
def merge_inspection_assets(config, extract_result=None, do_cleanup=True):
    """
    合并巡检资产数据：
    - 设备版本报告 → 网元版本信息
    - 设备SN清单 → 网元结果
    - 设备SN清单 → 硬件网元
    - DEVICE_CHART → 汇总表
    """
    if extract_result is None:
        extract_result = extract_all_zips(config)

    temp_dir = extract_result["temp_dir"]
    zip_total = extract_result["zip_count"]

    if not os.path.exists(temp_dir):
        print(f"❌ 临时目录不存在: {temp_dir}")
        return

    # 定义要合并的文件和sheet
    # header_row: 表头所在行（第一个文件读取此行作为标题）
    # data_start: 数据起始行（后续文件从此行开始读，跳过表头）
    merge_tasks = [
        {
            "name": "网元版本信息",
            "file_pattern": "设备版本报告",
            "sheet_name": "网元版本信息",
            "header_row": 2,  # 表头在第2行
            "data_start": 3,  # 数据从第3行开始
            "data": [],
            "first_file": True  # 标记是否是第一个文件
        },
        {
            "name": "网元结果",
            "file_pattern": "设备SN清单",
            "sheet_name": "网元结果",
            "header_row": 1,  # 表头在第1行
            "data_start": 2,  # 数据从第2行开始
            "data": [],
            "first_file": True
        },
        {
            "name": "硬件网元",
            "file_pattern": "设备SN清单",
            "sheet_name": "硬件网元",
            "header_row": 1,
            "data_start": 2,
            "data": [],
            "first_file": True
        },
        {
            "name": "硬件资产",
            "file_pattern": "DEVICE_CHART.xls",
            "sheet_name": "汇总表",
            "header_row": 1,
            "data_start": 2,
            "data": [],
            "first_file": True
        }
    ]

    # 遍历临时目录查找文件
    total_files = 0
    err_count = 0

    for root, dirs, files in os.walk(temp_dir):
        for f in files:
            filepath = os.path.join(root, f)

            for task in merge_tasks:
                # 匹配文件
                if task["file_pattern"] in f and f.endswith(('.xlsx', '.xls')):
                    # 跳过不需要的文件
                    if task["file_pattern"] == "DEVICE_CHART.xls" and ("Optical" in f or "FT_" in f):
                        continue
                    if task["file_pattern"] == "设备SN清单" and not f.endswith('.xlsx'):
                        continue

                    try:
                        # 判断读取起始行：第一个文件保留表头，后续文件跳过表头
                        if task["first_file"]:
                            row_start = task["header_row"]
                            task["first_file"] = False
                        else:
                            row_start = task["data_start"]

                        if f.endswith('.xls'):
                            # .xls 文件使用 xlrd 读取
                            data = read_xls_file(filepath, task["sheet_name"], header_row=row_start - 1)
                            task["data"].extend(data)
                        else:
                            # .xlsx 文件使用项目的 excel 模块读取
                            read_excel = excel(filepath)
                            read_excel.excelReadCread()  # 打开Excel对象
                            data = read_excel.excelReadSheet(
                                sheet=task["sheet_name"],
                                row_start=row_start
                            )
                            read_excel.excelClose()  # 关闭Excel对象
                            task["data"].extend(data)
                        total_files += 1
                    except Exception as e:
                        err_count += 1
                        print(f"  ⚠️ 读取 {f} 的 {task['sheet_name']} 失败: {e}")

    # 统计结果
    print(f"\n📊 统计:")
    for task in merge_tasks:
        print(f"  - {task['name']}: {len(task['data'])} 条记录")

    if not any(task["data"] for task in merge_tasks):
        print(f"❌ 未找到任何数据")
        return

    # 输出到Excel（使用配置中的输出路径，excel模块会自动添加时间戳）
    output = config["output_assets"]

    # 确保输出目录存在
    output_dir = os.path.dirname(output)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # 准备多个sheet的数据
    sheets_data = []
    for task in merge_tasks:
        if task["data"]:
            # 第一行是表头（标题），其余是数据
            title = task["data"][0] if task["data"] else []
            data = task["data"][1:] if len(task["data"]) > 1 else []
            sheets_data.append({
                'title': title,
                'data': data,
                'sheetname': task["name"]
            })

    # 使用 excel 模块的多sheet写入功能
    write_excel = excel(output)
    write_excel.excel_write_multi_sheet(sheets_data, highlight=False)  # 组装数据
    saved_file = write_excel.save_file()  # 统一保存

    fail_info = f", 失败 {err_count}" if err_count else ""
    print(f"\n✅ 完成: {zip_total} 个压缩包 → {total_files} 个文件读取 → {saved_file}{fail_info}")

    if do_cleanup:
        ask_clean_temp(config["temp_dir"])


# ==========================================


# ============ 主函数 ============
def print_menu():
    """打印菜单"""
    print("\n" + "=" * 50)
    print("📊 巡检数据处理工具")
    print("=" * 50)
    print("  1. 巡检后提取Excel并分片压缩")
    print("  2. 巡检前汇总原始数据")
    print("  3. 提取配置文件到文件夹")
    print("  4. 合并巡检资产数据（版本+SN+资产+登陆状态）")
    print("  0. 退出")
    print("=" * 50)


if __name__ == "__main__":
    while True:
        print_menu()
        choice = input("请选择功能 [0-4，可多选如 1234]: ").strip()

        if choice == "0":
            print("👋 再见！")
            break

        if not choice or not all(c in '1234' for c in choice):
            print("❌ 无效选择，请输入 0-4")
            input("\n按回车键继续...")
            continue

        choices = list(choice)
        zip_functions = [c for c in choices if c in ('1', '3', '4')]

        # 功能1/3/4共用解压，只解一次
        extract_result = None
        if zip_functions:
            extract_result = extract_all_zips(CONFIG)

        if extract_result and not extract_result["excel_files"] and len(zip_functions) > 1:
            # 解压无结果且选了多个功能，跳过
            pass
        else:
            func_map = {
                '1': lambda: collect_and_split_zip(CONFIG, extract_result, do_cleanup=False),
                '2': lambda: mergeOriginalData(CONFIG),
                '3': lambda: merge_zip_txt_files(CONFIG, extract_result, do_cleanup=False),
                '4': lambda: merge_inspection_assets(CONFIG, extract_result, do_cleanup=False),
            }
            for c in choices:
                print()
                func_map[c]()

        # 统一清理临时目录（只问一次）
        if zip_functions and extract_result:
            ask_clean_temp(CONFIG["temp_dir"])

        input("\n按回车键继续...")
