# 网络设备配置检查 & 巡检数据处理工具

网络配置自动化合规检查、配置下发、下发验证比对，以及巡检报告数据汇总处理。

---

## 工具一：网络设备配置检查（main.py）

### 功能

| 序号 | 功能 | 说明 |
|------|------|------|
| 1 | 在线配置检查 | SSH/Telnet 登录设备，按 keyWords.txt 关键字匹配检查 |
| 2 | 离线配置检查 | 对采集配置按设备类型逐项合规审计（含线路检查） |
| 3 | 下发配置 | 批量登录设备下发命令 |
| 4 | 配置下发验证比对 | 预期配置 vs 采集配置 差异分析 |

### 运行

```bash
cd mywork_script
python main.py
```

### 各功能流程

#### 功能1：在线配置检查

```
读取 devices_ip.xlsx → ping 检测 → SSH/Telnet 登录
→ 执行 dis current-configuration → 按 keyWords.txt 匹配关键字
→ 写入 Excel 报告
```

#### 功能2：离线配置检查

```
读取 read/config/ 文件列表 → 按设备名匹配设备类型（check_items.yaml）
→ 逐项检查（版本/补丁/硬件/BGP/线路等13项）
→ 写入 Excel 报告
```

线路检查（第13项）：对比 `display interface description`（期望）与 `display lldp neighbor brief`（实际），排除规则在 `check_items.yaml` 的 `cable_check` 中配置。

#### 功能3：下发配置

```
读取 devices_ip.xlsx → ping 检测 → SSH/Telnet 登录
→ 批量下发命令 → 检查回显错误 → 写入 Excel 报告
```

#### 功能4：配置下发验证比对

```
加载 compare_rules_v2.yaml → 预期/采集配置配对
→ 正则提取段落 → 忽略过滤 → 密码归一化
→ 集合差集对比（忽略大小写）→ Markdown 报告
```

---

## 工具二：巡检数据处理（mergeExcel.py）

针对招商银行 eDesk Pro 巡检报告数据的后处理工具，支持多功能批量执行。

### 功能

| 序号 | 功能 | 说明 |
|------|------|------|
| 1 | 提取 Excel 并分片压缩 | 从巡检报告 ZIP 中提取目标 Excel，汇总到分片 ZIP |
| 2 | 汇总原始数据 | 将各设备原始数据合并到总表 |
| 3 | 提取配置文件 | 从 ZIP 中提取 .txt 配置文件到统一目录，自动重命名（去掉前缀、IP 分隔符 - 改 .） |
| 4 | 合并巡检资产数据 | 合并设备版本报告、SN 清单、资产报告、登录状态等多 sheet 数据 |

### 运行

```bash
cd mywork_script
python mergeExcel.py
```

支持多选，例如输入 `134` 会依次执行功能1、3、4，且共用一次解压。

### 各功能流程

#### 功能1：提取 Excel 并分片压缩

```
读取 source_dir 下所有 ZIP → 解压并编码修正(cp437→gbk)
→ 查找目标 Excel → 合并为总表 → 分片 ZIP 压缩输出
```

#### 功能2：汇总原始数据

```
读取 ZIP 中各设备 Excel → 按设备名提取原始数据 → 写入汇总表
```

#### 功能3：提取配置文件

```
解压 ZIP → 查找资产报告/设备配置/ 目录下的 .txt 文件
→ 复制到统一目录，自动重命名：
  原格式：设备配置_SHO3X11U35-CF6-H88A_12-255-21-1.txt
  新格式：SHO3X11U35-CF6-H88A_12.255.21.1.txt
```

#### 功能4：合并巡检资产数据

```
多线程并发读取 ZIP 中各文件：
  ├── 设备版本报告（header_row=2）→ 网元版本信息 sheet
  ├── 设备SN清单（header_row=1）→ 网元结果 sheet
  ├── DEVICE_CHART → 资产报告 sheet
  └── 设备登录状态 → 登陆状态 sheet
→ 多 sheet 写入 Excel（highlight=False，默认不高亮）
```

### 配置说明

编辑 `mergeExcel.py` 顶部的 `CONFIG` 字典：

```python
CONFIG = {
    "source_dir": "巡检报告ZIP目录",
    "target_excel_name": "NetWork Healthy Check Report(Engineer).xlsx",
    "output_zip": "data/巡检数据分析汇总",
    "split_size": "18m",           # 分片大小
    "source_file": "设备清单.xlsx",
    "output_original": "data/巡检原始数据汇总",
    "output_config_dir": "设备配置文件汇总目录",
    "output_assets": "data/巡检资产汇总",
    "temp_dir": "temp_extract",    # 临时解压目录
}
```

---

## 项目结构

```
mywork_script/
├── main.py                       # 入口：设备配置检查菜单
├── mergeExcel.py                 # 入口：巡检数据处理菜单
├── cfgCheck.py                   # 离线合规检查（13项 + 线路检查）
├── checkConfigOnline.py          # 在线配置检查
├── compare_configs.py            # 配置下发验证比对
├── sendCmd.py                    # 下发配置
│
├── interface/                    # 公共模块
│   ├── __init__.py               # 模块导出
│   ├── device_client.py          # SSH/Telnet 连接、数据清洗
│   ├── excel_handler.py          # Excel 读写（支持高亮开关、多 sheet、按序号/名字读取）
│   ├── thread_pool.py            # 线程池（按提交顺序返回结果，异常位置保留 None）
│   ├── log_config.py             # 日志配置
│   ├── file_utils.py             # 文件读取工具
│   └── bitFunctions.py           # 网络工具（ping、密码输入等）
│
├── read/                         # 输入文件
│   ├── config/                   # 采集配置（.txt）
│   ├── config_intended/          # 预期配置（.cfg）
│   ├── check_items.yaml          # 检查项 & 设备类型配置
│   ├── compare_rules_v2.yaml     # 比对规则
│   ├── keyWords.txt              # 在线检查关键字
│   ├── 版本补丁.xlsx             # 版本补丁推荐对照表
│   └── devices_ip.xlsx           # 设备 IP 清单
│
└── data/                         # 输出目录（Excel 报告、分片 ZIP）
```

---

## Excel 模块特性（interface/excel_handler.py）

- **`excel_write`** / **`excel_write_multi_sheet`** — 支持 `highlight` 参数，默认关闭高亮标记
- **多 sheet 写入** — 每个 sheet 可独立控制是否高亮
- **按序号或名字读取** — `excel_read` / `excelReadSheet` 支持 int（序号）和 str（sheet名）
- **自动时间戳** — `save_file()` 自动追加时间戳

## 线程池特性（interface/thread_pool.py）

- **按提交顺序返回结果** — 支持进度条回调
- **异常留空** — 报错的任务位置返回 `None`，不挤占后续位置
- **多次调用不串** — 每次 `__call__` 自动重置结果列表

---

## 配置文件说明

### check_items.yaml

- `check_items` — 检查项列表
- `device_types` — 各设备类型启用的检查项
- `device_patterns` — 设备名匹配规则
- `cable_check` — 线路检查排除规则

### keyWords.txt

功能1（在线配置检查）的关键字规则文件，每行一条，`#` 开头的行为注释。

格式：`正则表达式,显示名称,检查类型`

| 字段 | 说明 |
|------|------|
| 正则表达式 | 支持完整 Python 正则语法（`re.IGNORECASE`），在 `dis current-configuration` 回显中匹配 |
| 显示名称 | Excel 报告中的列名 |
| 检查类型 | `0` = 不应存在（匹配到则报"多余"），`1` = 应存在（未匹配则报"缺少"）|

示例：

```
telnet server disable,关闭Telnet,0
ntp server source-interface all disable,关闭NTP,1
```

### devices_ip.xlsx

功能1（在线检查）和功能3（下发配置）共用的设备清单，从第2行开始读取。运行时的用户名/密码自动插入每行前面。

列结构：

| 列 | 字段 | 说明 |
|----|------|------|
| A | IP | 设备管理 IP |
| B | Description | 设备描述（如机房/位置） |
| C | Command | 功能3下发的命令，多条命令用英文逗号隔开（功能1不读此列） |

功能1 会自动追加 ping 延迟、登录方式、关键字检查结果等列；功能3 会追加下发结果列。

### compare_rules_v2.yaml

- `settings` — 目录配置
- `sections` — 段落拆分规则（正则提取接口、BGP 等段落独立对比）
- `ignore` — 忽略规则（支持按型号/版本条件，models 列表内 AND，列表项之间 OR）

---

## 注意事项

- 仅支持 `.xlsx` 格式，不支持 `.xls`
- 巡检数据处理时，功能1/3/4会共用一次解压，避免重复操作
- 提取配置文件时，自动修正中文编码（cp437→gbk）并重命名
- 合并资产数据时，设备版本报告从第3行读数据，SN清单从第2行读数据
