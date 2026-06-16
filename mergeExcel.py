import os
import zipfile
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from interface import excel

# ============ 配置区（只改这里）============
CONFIG = {
    # 源数据目录
    "source_dir": "/Users/shadowx/Documents/招行/巡检/2026Q2/巡检报告数据",

    # 功能1+2：提取Excel
    "target_excel_name": "NetWork Healthy Check Report(Engineer).xlsx",
    "output_excel": "data/巡检成功数据汇总",
    "output_zip": "data/巡检数据分析汇总",  # 分片压缩输出
    "split_size": "18m",  # 分片大小

    # 功能3：汇总原始数据
    "source_file": "/Users/shadowx/Documents/招行/巡检/2026Q2/巡检设备清单汇总20260603更新.xlsx",
    "output_original": "data/巡检原始数据汇总",

    # 功能4：提取配置文件
    "output_config_dir": "/Users/shadowx/Documents/招行/巡检/设备配置文件汇总",

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
            except:
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


# ==========================================


# ============ 公共解压（功能1/2/4共用）============
def extract_all_zips(config):
    """
    统一解压所有ZIP文件到临时目录，供功能1/2/4共用

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
    clean_temp_dir(temp_dir)

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


# ============ 功能1：汇总数据 ============
def mergeData(config, extract_result=None, do_cleanup=True):
    """巡检后汇总所有数据到一个表"""
    if extract_result is None:
        extract_result = extract_all_zips(config)

    extracted_files = extract_result["excel_files"]
    zip_total = extract_result["zip_count"]

    if not extracted_files:
        print(f"❌ 统计: {zip_total} 个压缩包, 0 个Excel文件")
        return

    data_all = []
    err_count = 0
    total_files = len(extracted_files)
    for i, (file_path, _) in enumerate(extracted_files, 1):
        print(f"\r  [{i}/{total_files}] 读取Excel...", end="", flush=True)
        try:
            read_data = excel(file_path)
            data_list = read_data.excel_read(row_start=3)
            data_all.extend(data_list)
        except Exception as e:
            err_count += 1

    print()

    if not data_all:
        print(f"❌ 统计: {zip_total} 个压缩包, {len(extracted_files)} 个Excel, 0 条数据")
        return

    output = config["output_excel"]
    data_wr = excel(output)
    data_wr.excel_write(
        title=['网元名称', '网元分组', '网元类型', '版本信息', '补丁信息', '评估场景', '网元IP', 'ESN号'],
        data=data_all,
        sheetname='巡检网元汇总'
    )
    data_wr.save_file()

    fail_info = f", 失败 {err_count}" if err_count else ""
    print(f"✅ 完成: {zip_total} 个压缩包 → {len(extracted_files)} 个Excel → {len(data_all)} 条记录{fail_info}")

    if do_cleanup:
        ask_clean_temp(config["temp_dir"])


# ==========================================


# ============ 功能2：提取打包 ============
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
            print(f"✅ 完成: {zip_total} 个压缩包 → {len(extracted_files)} 个Excel → {output_zip_abs}.zip (分片: {split_size})")
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


# ============ 功能3：汇总原始数据 ============
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
        res = open_excel.excelReadSheet(sheetnum=sheet_name, column=5)
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


# ============ 功能4：提取配置文件 ============
def merge_zip_txt_files(config, extract_result=None, do_cleanup=True):
    """巡检后汇总压缩包中的配置文件到一个文件夹"""
    source_dir = config["source_dir"]
    output_dir = config["output_config_dir"]
    temp_dir = config["temp_dir"]

    base_path = Path(source_dir).resolve()
    target_path = Path(output_dir).resolve()
    temp_extract_path = Path(temp_dir)

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
                        except:
                            pass
                        zip_ref.extract(file_info, temp_extract_path)

                top_folders = [f for f in temp_extract_path.iterdir()
                               if f.is_dir() and not f.name.startswith((".", "__"))]

                if top_folders:
                    cfg_dir = top_folders[0] / "资产报告/设备配置"
                    if cfg_dir.exists():
                        txt_files = list(cfg_dir.glob("*.txt"))
                        for txt_file in txt_files:
                            shutil.move(str(txt_file), str(target_path / txt_file.name))
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
                shutil.move(str(txt_file), str(target_path / txt_file.name))
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


# ============ 主函数 ============
def print_menu():
    """打印菜单"""
    print("\n" + "=" * 50)
    print("📊 巡检数据处理工具")
    print("=" * 50)
    print("  1. 巡检后汇总数据到Excel")
    print("  2. 巡检后提取Excel并分片压缩")
    print("  3. 巡检前汇总原始数据")
    print("  4. 提取配置文件到文件夹")
    print("  0. 退出")
    print("=" * 50)


if __name__ == "__main__":
    while True:
        print_menu()
        choice = input("请选择功能 [0-4，可多选如 124]: ").strip()

        if choice == "0":
            print("👋 再见！")
            break

        if not choice or not all(c in '1234' for c in choice):
            print("❌ 无效选择，请输入 0-4")
            input("\n按回车键继续...")
            continue

        choices = list(choice)
        zip_functions = [c for c in choices if c in ('1', '2', '4')]

        # 功能1/2/4共用解压，只解一次
        extract_result = None
        if zip_functions:
            extract_result = extract_all_zips(CONFIG)

        if extract_result and not extract_result["excel_files"] and len(zip_functions) > 1:
            # 解压无结果且选了多个功能，跳过
            pass
        else:
            func_map = {
                '1': lambda: mergeData(CONFIG, extract_result, do_cleanup=False),
                '2': lambda: collect_and_split_zip(CONFIG, extract_result, do_cleanup=False),
                '3': lambda: mergeOriginalData(CONFIG),
                '4': lambda: merge_zip_txt_files(CONFIG, extract_result, do_cleanup=False),
            }
            for c in choices:
                print()
                func_map[c]()

        # 统一清理临时目录（只问一次）
        if zip_functions and extract_result:
            ask_clean_temp(CONFIG["temp_dir"])

        input("\n按回车键继续...")
