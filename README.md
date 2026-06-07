# 网络设备配置检查工具

交换机配置自动化合规检查、配置下发、下发验证比对。

## 功能

| 序号 | 功能       | 说明                                   |
|----|----------|--------------------------------------|
| 1  | 在线配置检查   | SSH/Telnet 登录设备，按 keyWords.txt 关键字检查 |
| 2  | 离线配置检查   | 对采集配置按设备类型逐项合规审计（含线路检查）              |
| 3  | 下发配置     | 批量登录设备下发命令                           |
| 4  | 配置下发验证比对 | 预期配置 vs 采集配置 差异分析                    |

## 项目结构

```
mywork_script/
├── main.py                    # 入口，功能菜单
├── cfgCheck.py                # 功能2：离线合规检查（13项 + 线路检查）
├── compare_configs.py         # 功能4：配置下发验证比对
├── checkConfig.py             # 功能1：在线配置检查
├── sendCmd.py                 # 功能3：下发配置
├── interface/
│   ├── device_client.py       # SSH/Telnet 连接、数据清洗
│   ├── excel_handler.py       # Excel 读写
│   ├── thread_pool.py         # 线程池
│   ├── log_config.py          # 日志配置
│   ├── file_utils.py          # 文件读取工具
│   └── bitFunctions.py        # 网络工具（ping、密码输入等）
├── read/
│   ├── config/                # 采集配置
│   ├── config_intended/       # 预期配置
│   ├── check_items.yaml       # 检查项 & 设备类型配置
│   ├── compare_rules_v2.yaml  # 比对规则
│   ├── keyWords.txt           # 在线检查关键字
│   ├── 版本补丁.xlsx          # 版本补丁推荐对照表
│   └── devices_ip.xlsx        # 设备 IP 清单
└── data/                      # 输出报告
```

## 运行

```bash
cd mywork_script
python main.py
```

## 各功能流程

### 功能1：在线配置检查

```
读取 devices_ip.xlsx → ping 检测 → SSH/Telnet 登录
→ 执行 dis current-configuration → 按 keyWords.txt 匹配关键字
→ 写入 Excel 报告
```

### 功能2：离线配置检查

```
读取 read/config/ 文件列表 → 按设备名匹配设备类型（check_items.yaml）
→ 逐项检查（版本/补丁/硬件/BGP/线路等13项）
→ 写入 Excel 报告
```

线路检查（第13项）：对比 `display interface description`（期望）与 `display lldp neighbor brief`（实际），排除规则在
`check_items.yaml` 的 `cable_check` 中配置。

### 功能3：下发配置

```
读取 devices_ip.xlsx → ping 检测 → SSH/Telnet 登录
→ 批量下发命令 → 检查回显错误 → 写入 Excel 报告
```

### 功能4：配置下发验证比对

```
加载 compare_rules_v2.yaml → 预期/采集配置配对
→ 正则提取段落 → 忽略过滤 → 密码归一化
→ 集合差集对比（忽略大小写）→ Markdown 报告
```

## 配置文件说明

### check_items.yaml

- `check_items` — 检查项列表
- `device_types` — 各设备类型启用的检查项
- `device_patterns` — 设备名匹配规则
- `cable_check` — 线路检查排除规则

### keyWords.txt

功能1（在线配置检查）的关键字规则文件，每行一条，`#` 开头的行为注释。

格式：`正则表达式,显示名称,检查类型`

| 字段    | 说明                                                                    |
|-------|-----------------------------------------------------------------------|
| 正则表达式 | 支持完整 Python 正则语法（`re.IGNORECASE`），在 `dis current-configuration` 回显中匹配 |
| 显示名称  | Excel 报告中的列名                                                          |
| 检查类型  | `0` = 不应存在（匹配到则报"多余"），`1` = 应存在（未匹配则报"缺少"）                            |

示例：

```
telnet server disable,关闭Telnet,0
ntp server source-interface all disable,关闭NTP,1
```

### devices_ip.xlsx

功能1（在线检查）和功能3（下发配置）共用的设备清单，从第2行开始读取。
运行时输入的用户名/密码会自动插入每行前面。

列结构：

| 列 | 字段          | 说明                            |
|---|-------------|-------------------------------|
| A | IP          | 设备管理 IP                       |
| B | Description | 设备描述（如机房/位置）                  |
| C | Command     | 功能3下发的命令，多条命令用英文逗号隔开（功能1不读此列） |

功能1 会自动追加 ping 延迟、登录方式、关键字检查结果等列；功能3 会追加下发结果列。

### compare_rules_v2.yaml

- `settings` — 目录配置
- `sections` — 段落拆分规则（正则提取接口、BGP 等段落独立对比）
- `ignore` — 忽略规则（支持按型号/版本条件，models 列表内 AND，列表项之间 OR）
