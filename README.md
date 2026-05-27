# 网络设备配置检查工具

网络设备配置自动化合规检查、配置下发、下发验证比对。

## 功能

| 序号 | 功能 | 说明 |
|------|------|------|
| 1 | 在线配置检查 | SSH/Telnet 登录设备，按 keyWords.txt 关键字检查 |
| 2 | 离线配置检查 | 对 read/config/*.log 按设备类型逐项合规审计 |
| 3 | 下发配置 | 批量登录设备下发命令 |
| 4 | 配置下发验证比对 | 预期配置(.cfg) vs 采集配置(.log) 差异分析 |

## 项目结构

```
mywork_script/
├── main.py                    # 入口，功能菜单
├── cfgCheck.py                # 离线配置合规检查（13 项检查项）
├── compare_configs.py         # 配置下发验证比对 v2
├── checkConfig.py             # 在线配置检查
├── sendCmd.py                 # 命令下发
├── interface/
│   ├── check_items.py         # 检查项配置加载器（读取 YAML）
│   ├── connection.py          # SSH/Telnet 连接、Excel 读写、线程池
│   ├── public_env.py          # 全局变量管理
│   └── bitFunctions.py        # 网段计算
├── read/
│   ├── config/                # 采集配置 (.log)
│   ├── config_intended/       # 预期配置 (.cfg)
│   ├── check_items.yaml       # 检查项 & 设备类型配置（直接编辑）
│   ├── compare_rules_v2.yaml  # 比对规则（段落定义 + 忽略规则）
│   ├── keyWords.txt           # 在线检查关键字
│   ├── 版本补丁.xlsx          # 版本补丁推荐对照表
│   └── devices_ip.xlsx        # 设备 IP 清单
└── data/                      # 输出报告
```

## 运行

```bash
cd mywork_script
.venv/bin/python main.py
```

或直接运行比对：`.venv/bin/python compare_configs.py`

## 配置比对规则 (compare_rules_v2.yaml)

- **settings** — 目录配置（预期/采集/输出）
- **sections** — 段落拆分（正则提取接口、BGP 等段落独立对比，加条即用）
- **ignore** — 忽略规则，两侧同时过滤，支持按型号/版本条件生效

型号/版本过滤语法（models 列表内 AND，列表项之间 OR）：

```yaml
    - header_pattern: '^user-interface '
      models:
        - model: 'CE6881'
          version: 'V200R024'
        - model: 'S5731'
          version: 'V200R023'
      lines:
        - 'protocol inbound ssh'
```

不写 `models` → 对所有设备生效。

## 比对流程

```
加载规则 → 设备配对 → 正则提取段落
→ 忽略过滤 → 密码归一化 → 集合差集对比 → Markdown 报告
```

## 线路检查（功能2第13项）

对比 `display interface description`（期望）与 `display lldp neighbor brief`（实际），检测线路接错。

排除规则在 `read/check_items.yaml` 的 `cable_check` 中配置：
- `exclude_interfaces` — 按接口名正则排除（逻辑口等）
- `exclude_phy` — 按 PHY 状态排除（down 端口）
- 无描述的端口不检查
