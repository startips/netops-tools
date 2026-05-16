# 华为交换机配置检查工具

用于对华为交换机离线配置文件（.log）进行自动化合规检查，生成 Excel 报告。

## 功能

1. **配置检查** — 对 read/config/ 目录下的 .log 配置文件，按设备类型逐项检查合规项
2. **在线检查** — 通过 SSH/Telnet 登录设备，读取配置并实时检查
3. **命令下发** — 批量登录设备下发配置命令
4. **版本补丁对照** — 采集版本/补丁号，与本地 Excel 对照表比对

## 项目结构

```
mywork_script/
├── main.py                 # 入口，功能菜单
├── cfgCheck.py             # 离线配置文件检查（核心）
├── checkConfig.py          # 在线配置检查
├── sendCmd.py              # 命令下发
├── mergeExcel.py           # 巡检报告汇总
├── config/
│   ├── __init__.py
│   └── check_items.py      # 检查项名称 & 设备类型配置表
├── interface/
│   ├── connection.py       # 交换机登录、Excel 处理
│   ├── public_env.py       # 全局变量、日志、线程池
│   └── bitFunctions.py     # 网段计算
├── read/
│   ├── config/             # 离线配置文件 (.log)
│   ├── 版本补丁.xlsx       # 版本补丁推荐对照表
│   ├── keyWords.txt        # 在线检查关键字
│   └── devices_ip.xlsx     # 设备 IP 清单
└── data/                   # 输出目录（Excel 报告、日志）
    log/                    # 运行日志
```

## 配置检查（cfgCheck.py）

### 核心流程

1. 读取 `read/config/` 下的 `.log` 配置文件
2. 根据文件名匹配设备类型（Spine / Leaf / Slf 等）
3. 按设备类型的检查项配置（0/1），逐项执行检查
4. 结果写入 Excel 报告

### 新增检查项

两步：

1. **`config/check_items.py`** — `CHECK_ITEM_NAMES` 末尾加名称，在需要的设备类型配置中加一行
2. **`cfgCheck.py`** — 写一个 `_check_xxx()` 函数，在 `_CHECKERS` 字典注册

Excel 报告列头自动从 `CHECK_ITEM_NAMES` 生成，无需修改 `main.py`。

### 新增设备类型

1. **`config/check_items.py`** — `_make_check_option()` 定义检查项，`DEVICE_TYPE_CONFIGS` 注册，`DEVICE_TYPE_PATTERNS` 加匹配规则

### 版本补丁对照

启动时自动加载 `read/版本补丁.xlsx` → `版本补丁推荐（季度）` sheet，筛选 `使用场景=新上线` 建立对照表。

- 采集版本/补丁与推荐一致 → 正常返回
- 不一致 → 返回 `@未通过`
- 型号不在对照表中 → 返回 `-未知`

## 运行

```bash
cd mywork_script
.venv/bin/python main.py
```

菜单：
- 1 — 在线配置检查
- 2 — 离线配置文件检查（主要功能）
- 3 — 下发配置
