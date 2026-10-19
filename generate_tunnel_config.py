#!/usr/bin/env python3
"""
自动化：根据隧道参数 Excel 生成 MPLS TE 隧道设备配置

流程：读取 Excel（可含多个 sheet，如 研发PE / 上海PE）→ 按"设备名称"分组
      → 每行参数套用配置模板 → 每个设备的隧道配置写入独立 .txt 文件

规则:
  - priority=5 的隧道不生成 `mpls te service-class` 行
  - 每个隧道配置块之间用 `#` 分隔（独立一行）
  - 输出文件名为: <设备名称>.txt
  - Excel 读取复用 interface/excel_handler.py 的 excel 类

用法:
  python generate_tunnel_config.py /path/to/隧道配置.xlsx
  python generate_tunnel_config.py /path/to/隧道配置.xlsx -o /path/to/output_dir
  python generate_tunnel_config.py /path/to/隧道配置.xlsx --dry-run   # 只打印分组统计，不写文件
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

from interface import excel

# 列索引（与表头一一对应，0 起）
COL_DEVICE = 0          # 设备名称
COL_TUNNEL_IF = 1       # Tunnel接口名称
COL_DST_IP = 2          # Tunnel目的IP
COL_COST = 3            # 隧道 cost
COL_PRIORITY = 4        # 隧道优先级
COL_TUNNEL_ID = 5       # tunnel-id
COL_SERVICE_CLASS = 6   # service-class

# 配置模板（按行顺序）
TEMPLATE_LINES = [
    "interface Tunnel{tunnel_if}",
    " ip address unnumbered interface LoopBack0",
    " tunnel-protocol mpls te",
    " destination {dst_ip}",
    " mpls te record-route label",
    " mpls te igp shortcut",
    " mpls te igp metric absolute {cost}",
    " mpls te priority {priority}",
    " mpls te backup hot-standby dynamic-bandwidth",
    " mpls te reoptimization frequency 1800",
    " statistic enable",
    " mpls te auto-bandwidth adjustment threshold 10 frequency 900 overflow-limit 3 underflow-limit 3",
    " mpls te tunnel-id {tunnel_id}",
    # service-class 行按规则动态追加
]

# 每个隧道块之间的分隔符
BLOCK_SEPARATOR = "#"


def read_excel(excel_path: Path) -> dict[str, list[dict]]:
    """通过 interface.excel 读取 Excel 所有 sheet，按设备名称分组返回隧道参数列表。"""
    open_excel = excel(str(excel_path))
    sheet_names = open_excel.excelReadCread()

    device_configs: dict[str, list[dict]] = defaultdict(list)
    for sheet_name in sheet_names:
        # 返回二维 list：每行从第2行起（跳过标题），列从第1列起
        # excelReadSheet 支持按 sheet 名(str)读取，接口类型标注未覆盖
        for row in open_excel.excelReadSheet(sheet=sheet_name):  # type: ignore[arg-type]
            if not row or not row[COL_DEVICE]:
                continue
            device_configs[row[COL_DEVICE]].append({
                "tunnel_if": str(row[COL_TUNNEL_IF]),
                "dst_ip": str(row[COL_DST_IP]),
                "cost": str(row[COL_COST]),
                "priority": str(row[COL_PRIORITY]),
                "tunnel_id": str(row[COL_TUNNEL_ID]),
                "service_class": str(row[COL_SERVICE_CLASS]) if len(row) > COL_SERVICE_CLASS and row[COL_SERVICE_CLASS] else "",
            })
    open_excel.excelClose()
    return dict(device_configs)


def build_block(t: dict) -> str:
    """根据模板生成单个隧道配置块；priority=5 时不含 service-class 行。"""
    lines = [line.format(**t) for line in TEMPLATE_LINES]
    if str(t["priority"]) != "5":
        lines.append(f' mpls te service-class {t["service_class"]}')
    return "\n".join(lines)


def generate_config(device_configs: dict[str, list[dict]], output_dir: Path) -> list[Path]:
    """为每个设备生成配置文本并写入 <设备名称>.txt，返回生成的文件列表。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for device, tunnels in device_configs.items():
        blocks = [build_block(t) for t in tunnels]
        content = f"\n{BLOCK_SEPARATOR}\n".join(blocks) + "\n"
        filepath = output_dir / f"{device}.txt"
        filepath.write_text(content, encoding="utf-8")
        written.append(filepath)
    return written


def main() -> int:
    parser = argparse.ArgumentParser(
        description="根据隧道参数 Excel 生成 MPLS TE 隧道设备配置"
    )
    parser.add_argument("excel", type=Path, help="隧道参数 Excel 文件路径")
    parser.add_argument(
        "-o", "--output", type=Path, default=Path("./data"),
        help="输出目录（默认: ./data）",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="只打印分组统计，不写文件",
    )
    args = parser.parse_args()

    if not args.excel.exists():
        sys.exit(f"Excel 文件不存在: {args.excel}")

    device_configs = read_excel(args.excel)
    total_tunnels = sum(len(v) for v in device_configs.values())
    print(f"读取 {args.excel.name}: {len(device_configs)} 个设备, {total_tunnels} 条隧道")

    for device, tunnels in device_configs.items():
        priorities = defaultdict(int)
        for t in tunnels:
            priorities[str(t["priority"])] += 1
        dist = ", ".join(f"p{p}:{c}" for p, c in sorted(priorities.items(), key=lambda x: int(x[0])))
        print(f"  {device}: {len(tunnels)} 条隧道 ({dist})")

    if args.dry_run:
        print("[dry-run] 未写入任何文件")
        return 0

    written = generate_config(device_configs, args.output)
    for f in written:
        block_count = len([s for s in f.read_text().split(BLOCK_SEPARATOR) if s.strip()])
        print(f"Generated: {f} ({block_count} 条隧道)")

    print(f"完成: {len(written)} 个配置文件已写入 {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
