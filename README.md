# AI-trace — Perfetto Block IO 分析工具集

本项目提供一套完整的 Android/Linux Perfetto trace 解析工具，聚焦于 **Block IO 事件** 的提取、时延计算、IO 大小分析与可视化。

## 目录结构

```
AI-trace/
├── bin/windows-amd64/
│   └── trace_processor_shell.exe    # Perfetto 官方 trace 解析工具（v58.2）
├── trace/                            # 原始 .perfetto-trace / .perfetto 文件
├── out/                              # 输出目录，按 trace 名称分子目录
│   └── <trace名称>/
│       ├── <trace名称>.systrace                # 步骤1：systrace 文本
│       ├── <trace名称>_block_io.xlsx           # 步骤2：IO 事件表格（16列，XLSX 带格式）
│       ├── <trace名称>_block_io.csv            # 步骤2：IO 事件表格（16列，CSV 可快速分析）
│       ├── <trace名称>_segments/               # 步骤3：超大 trace 按时间切片输出
│       │   ├── <名称>_<start>_<end>_block_io.csv
│       │   └── <名称>_<start>_<end>_block_io.xlsx
│       ├── <trace名称>_latency.xlsx            # 步骤4+5+6：汇总表（4个Sheet）
│       │   ├── 调度耗时
│       │   ├── 器件耗时
│       │   ├── IO 大小汇总
│       │   └── 进程 IO 统计
│       ├── PID_<值>/ 或 TID_<值>/               # 步骤7：指定进程/线程分析（XLSX + 4图）
│       │   ├── <trace名称>_<值>.xlsx            # 多目标时 <值> 用下划线连接
│       │   └── <trace名称>_<值>_*.png           # 4张图表
│       ├── <trace名称>_diagnosis.xlsx          # 步骤8：IO 综合诊断（10个Sheet）
│       ├── <trace名称>_report.html             # 步骤9：可视化分析报告（图片内嵌）
│       ├── charts/                             # PNG 图表统一存放（步骤4/5/8）
│       │   ├── <trace名称>_schedule_latency.png     # 步骤4：调度耗时分布图
│       │   ├── <trace名称>_device_latency.png       # 步骤4：器件耗时分布图
│       │   ├── <trace名称>_io_nr_sector.png         # 步骤5：扇区数量分布图
│       │   ├── <trace名称>_io_size_kb.png           # 步骤5：请求大小分布图
│       │   └── <trace名称>_<suffix>_*.png           # 步骤8：诊断图表（9张）
│       └── stall_analysis/                     # AI 卡顿分析：单笔 IO 深度报告
│           └── <start_ns>_<dur_ns>/
│               ├── io_stall_report.md          # 四维度拆解报告
│               └── io_stall_data.txt           # 原始事件明细（数据来源可追溯）
├── scripts/
│   ├── config_loader.py              # 共享模块：配置文件加载（default.json/ai_default.json）
│   ├── convert_trace.py              # 步骤1：格式转换
│   ├── parse_block_io.py             # 步骤2：IO 事件提取
│   ├── split_trace.py                # 步骤3：超大 trace 时间切片拆分
│   ├── io_reader.py                  # 共享模块：CSV/XLSX 统一读取
│   ├── io_latency_chart.py           # 步骤4：IO 时延计算与绘图
│   ├── io_size_chart.py              # 步骤5：IO 大小分布与绘图
│   ├── io_process_stats.py           # 步骤6：进程 IO 统计
│   ├── pid_io_analysis.py            # 步骤7：指定 PID/TID 全维度分析
│   ├── io_diagnosis.py               # 步骤8：IO 综合诊断
│   ├── generate_report.py            # 步骤9：生成可视化分析报告（HTML）
│   ├── expert_notes.py               # 共享模块：存储专家级分析指引（每步/每图表/每表格）
│   └── AI/                           # AI 分析脚本（与 IO 解析脚本解耦）
│       ├── io_stall_analysis.py      # AI1: 输入 Start/Duration/Process 一键出报告
│       ├── io_chat.py                # AI2: 围绕分析报告多轮对话（支持读取本地文件/目录）
│       ├── llm_client.py             # 统一大模型客户端（模型信息读 config/ai_model.json）
│       └── convert_to_csv.py         # block_io.xlsx → CSV 缓存（加速读取）
├── config/
│   ├── default.json                  # 主配置文件（步骤 1-9 参数默认值）
│   ├── ai_default.json               # AI 分析选项配置文件（AI1/AI2 参数默认值）
│   └── ai_model.json                 # 大模型接入配置（base_url / 模型名 / API Key 环境变量）
├── skills/
│   └── io-stall-analysis/
│       └── SKILL.md                  # 存储专家 IO 卡顿根因分析经验文档
├── ai_trace.py                       # 主控脚本（交互式统一入口）
├── .last_params.json                 # 参数记忆文件（自动生成）
└── README.md                         # 本文件
```

---

## 配置文件: `config/default.json`

所有分析脚本（步骤 1-9）的默认参数集中于此，`ai_trace.py` 交互菜单中直接回车即用配置文件默认值。

### 配置结构

```jsonc
{
  "_global_comment": "所有时间字段支持秒数或纳秒数 (>= 1e10 自动视为纳秒)",
  "_note_start_end": "以下 start/end 为秒数, start_ns/end_ns 为纳秒, 均留空表示全量",

  "parse_block_io": {           // 步骤2: IO 事件提取
    "start_ns": "",               // 起始时间 (纳秒), 例: 110280321901428
    "end_ns": "",                 // 截止时间 (纳秒), 例: 110457720394070
    "start": "",                  // 起始时间 (秒), 例: 13229.127
    "end": ""                     // 截止时间 (秒), 例: 13235.444
  },

  "split_trace": {              // 步骤3: 超大 trace 切片拆分
    "pairs": [                    // 时间片段列表
      "110280321901428:110286256250923",
      "110290287672228:110401048612841",
      "110401048612841:110457864486570"
    ],
    "format": "both",             // "csv" | "xlsx" | "both"
    "timeout": "900"              // trace_processor_shell 超时秒数
  },

  "io_latency_chart": {         // 步骤4: IO 时延分析
    "sheet": "全部 IO",           // Sheet 名称
    "pcap": ""                    // 分位数截断, 如 "99.9", 空=不截断
  },

  "io_size_chart": {            // 步骤5: IO 大小分布
    "sheet": "全部 IO"
  },

  "io_process_stats": {         // 步骤6: 进程 IO 统计
    "sheet": "全部 IO"
  },

  "pid_io_analysis": {          // 步骤7: 指定 PID/TID 全维度分析
    "pid": "",                    // 目标 PID, 与 tid 二选一, 多个用逗号分隔
    "tid": "",                    // 目标 TID, 与 pid 二选一, 多个用逗号分隔
    "sheet": "全部 IO",
    "start": "", "end": "",       // 时间范围 (秒)
    "pcap": ""
  },

  "io_diagnosis": {             // 步骤8: IO 综合诊断
    "sheet": "全部 IO",
    "start": "", "end": "",       // 时间范围 (秒)
    "start_ns": "", "end_ns": "", // 时间范围 (纳秒)
    "bucket": "0.1",              // IOPS 分桶间隔 (秒)
    "top_n": "20"                 // 慢 IO Top-N 数量
  }
}
```

> 所有参数均可通过交互菜单实时修改，修改结果保存到 `.last_params.json`。下次运行时若与配置不同则作为 `(上次: xxx)` 提示，直接回车恢复配置默认值。

## 配置文件: `config/ai_default.json`

AI 分析选项（AI1 / AI2）的默认参数集中于此，与主配置分离，交互菜单中直接回车即用、也可手动输入覆盖。

### 配置结构

```jsonc
{
  "ai_stall_analysis": {         // AI1 — IO 卡顿快速分析 (阻塞 IO 根因拆解)
    "trace": "com.ss.android.ugc.aweme",
    "start_ns": 126139322335469,   // Start time (纳秒)
    "duration_ns": 32701962,       // Duration (纳秒)
    "process": "com.antutu.ABenchMark",
    "pid": "23620",
    "tid": "22288",
    "pre_ms": 50,                  // 卡顿点前窗口 (ms)
    "post_ms": 200                 // 卡顿点后窗口 (ms)
  },

  "ai_chat": {                   // AI2 — IO 报告 AI 问答 (围绕 trace 分析结果持续对话)
    "trace": "com.ss.android.ugc.aweme",
    "model": "fast"                // fast=deepseek-chat / deep=deepseek-reasoner
  }
}
```

## 配置文件: `config/ai_model.json`

大模型接入统一配置，**更换模型 / 服务商只需修改本文件，无需改代码**：

```jsonc
{
  "base_url": "https://api.deepseek.com",   // OpenAI 兼容服务端点（本地 Ollama 可设为 http://localhost:11434/v1）
  "api_key_env": "DEEPSEEK_API_KEY",        // API Key 环境变量名（实际密钥写入 .env，勿提交版本库）
  "fast_model": "deepseek-chat",            // 快速模型：结构化总结 / 指标提取
  "deep_model": "deepseek-reasoner",        // 推理模型：根因研判（输出含 reasoning_content）
  "temperature": 0.2,
  "max_tokens": 8192,
  "timeout": 600
}
```

> 服务端点优先级：环境变量 `DEEPSEEK_BASE_URL` > 配置文件 > 内置默认值。
> API Key 读取优先级：显式传参 > 环境变量（`api_key_env` 指定名称）> 项目根目录 `.env` 文件。

## 环境依赖

| 组件 | 说明 |
|---|---|
| Python 3.9+ | 脚本运行环境 |
| `trace_processor_shell.exe` | Perfetto 官方工具, 放置于 `bin/windows-amd64/` |
| `openpyxl` | `pip install openpyxl` |
| `matplotlib`, `numpy` | `pip install matplotlib numpy` |

```powershell
pip install openpyxl matplotlib numpy
```

## 数据流全景

```
原始 trace                 Block IO 表格                  分析输出
+--------------+  +----------+  +----------------------+  +----------------------+
| .perfetto-   |->| .systrace|  | block_io.xlsx        |  | latency.xlsx (4子表) |
| trace        |  | (文本)    |  | Sheet: 全部 IO       |->| png (4张分布图)      |
+--------------+  +----------+  | Sheet: 时间范围 IO    |  | PID_x/ TID_x 子目录  |
      |                          +----------------------+  +----------------------+
      |                                    |                          |
  步骤1: convert_trace.py                  |           +--------------+--------------+
                                      步骤2: parse     步骤4: io_      步骤6: io_      步骤7: pid_io_     步骤8: io_
                                     _block_io.py     latency_chart  process_stats  analysis       diagnosis
                                                      步骤5: io_size_chart 追加Sheet

                              步骤3: split_trace.py
                              超大 trace → 按时间切片拆分（多个独立 block_io 文件）
```

---

## 脚本1: `convert_trace.py` — 格式转换

### 功能

将 `trace/` 目录下的 `.perfetto-trace` 或 `.perfetto` 转换为 **systrace 文本格式**。

### 使用方式

```powershell
python scripts\convert_trace.py
```

### 输出

```
out/<trace名称>/<trace名称>.systrace
```

---

## 脚本2: `parse_block_io.py` — Block IO 事件提取

### 功能

从原始 trace 提取 **block_rq_insert / block_rq_issue / block_rq_complete** 三类事件, 输出 16 列 Excel。

### 使用方式

```powershell
python scripts\parse_block_io.py                                    # 全量
python scripts\parse_block_io.py --start 13229.127 --end 13235.444  # 时间范围
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `--start` | float(秒) | 起始时间 |
| `--end` | float(秒) | 结束时间 |

> 必须同时指定, 且 `start < end`。

### 输出

```
out/<trace名称>/<trace名称>_block_io.xlsx
out/<trace名称>/<trace名称>_block_io.csv   （同步输出，16 列数据完全一致）
```

| Sheet | 内容 | 条件 |
|---|---|---|
| 全部 IO | 全部事件 | 总是生成 |
| 时间范围 IO | 时间段内事件 | 指定 `--start --end` |

> **CSV vs XLSX**: 数据内容完全一致。CSV 生成瞬间完成（纯文本 I/O），适合快速校验；
> XLSX 含蓝色表头/Consolas 字体/自动筛选/冻结首行等格式，生成较慢但适合正式报告。
> 所有后续分析脚本（3-7）均已支持 CSV 和 XLSX 两种输入格式。

### 表格列（16列）

| # | 列名 | 来源 | 示例 |
|---|---|---|---|
| 1 | IO 进程名称 | `process.name` | `com.antutu.ABenchMark` |
| 2 | 进程 PID | `process.pid` | `5736` |
| 3 | 线程 TID | `thread.tid` | `17679` |
| 4 | 线程名称 | `thread.name` | `DefaultDispatch` |
| 5 | 时间 (秒) | `ts / 1e9` | `13229.133853383` |
| 6 | 时间 (纳秒) | `ts` | `13229133853383` |
| 7 | 运行 CPU | `cpu` | `3` |
| 8 | Trace 字段名称 | `name` | `block_rq_insert` |
| 9 | 设备号 | `dev` - MAJOR:MINOR | `8:0` |
| 10 | 请求扇区地址 | `sector` | `17686104` |
| 11 | 请求扇区数量 | `nr_sector` | `8` |
| 12 | 请求字节数 | `bytes` | `4096` |
| 13 | 读写类型 | `rwbs` decode | `WS (Write\|Sync)` |
| 14 | 进程 Comm | `comm` | - |
| 15 | 命令 | `cmd` | - |
| 16 | 错误码 | `error` | `0` |

### 设计原理

**为什么用 SQL 直查原始 trace？**

| | SQL 查询 | systrace 正则匹配 |
|---|---|---|
| 时间精度 | **纳秒（1ns）** | 微秒（1us） |
| 进程识别 | `ftrace_event - thread - process` 三级 JOIN | 仅 `pid/tgid` |
| IO 字段 | 同一行完整聚合 | 分散多行 |
| 灵活性 | 标准 SQL | 正则复杂 |

**核心 SQL**:

```sql
SELECT
    COALESCE(p.name, '<kernel>') AS process_name,
    COALESCE(CAST(p.pid AS TEXT), '') AS pid,
    t.tid, t.name AS thread_name,
    e.ts, e.cpu, e.name AS trace_event,
    MAX(CASE WHEN a.key='dev'    THEN a.int_value END) AS device,
    MAX(CASE WHEN a.key='sector' THEN a.int_value END) AS sector,
    ...
FROM ftrace_event e
JOIN thread t ON e.utid = t.id
LEFT JOIN process p ON t.upid = p.upid
LEFT JOIN args a ON e.arg_set_id = a.arg_set_id
WHERE e.name IN ('block_rq_insert','block_rq_issue','block_rq_complete')
GROUP BY e.id ORDER BY e.ts
```

**关键解码**:
- 设备号: `MKDEV(major, minor) = (major << 8) | minor`
- rwbs: `R`=Read, `W`=Write, `S`=Sync, `A`=RAhead, `M`=Meta

---

## 脚本3: `split_trace.py` — 超大 Trace 时间切片拆分

### 功能

针对超大 trace 文件（> 1GB），按指定时间片段将 Block IO 事件拆分为多个独立的 CSV/XLSX 文件，
每个片段的列结构与脚本2输出完全一致，可直接用于后续分析（脚本3-7）。

### 原理

使用 `trace_processor_shell` 加载一次 trace，通过一条 `UNION ALL` SQL 同时查询所有时间片段，
按 `seg` 标识列拆分后输出为独立文件。避免了多次重复加载 trace 的巨大开销。

### 使用方式

```powershell
# 切分为 3 个时间片段（nanoseconds）
python scripts\split_trace.py "安兔兔跑分" --pairs 110280321901428:110286256250923 110290287672228:110401048612841 110401048612841:110457864486570

# 仅输出 CSV（速度最快）
python scripts\split_trace.py "安兔兔跑分" --pairs ... --format csv

# 仅输出 XLSX（带格式）
python scripts\split_trace.py "安兔兔跑分" --pairs ... --format xlsx

# 同时输出 CSV + XLSX（默认）
python scripts\split_trace.py "安兔兔跑分" --pairs ... --format both
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `trace_name` | str (必填) | trace 文件名前缀（不含 `.perfetto-trace` / `.perfetto` 后缀） |
| `--pairs` | str 多值 (必填) | 时间片段列表，格式 `start_ns:end_ns` |
| `--format` | str | 输出格式: `csv` / `xlsx` / `both`（默认 `both`） |
| `--timeout` | int | trace_processor_shell 超时秒数（默认 `900` = 15 分钟） |

### 注意事项

- trace_processor_shell 加载 1.9GB trace 约需 **3-5 分钟**，命令行会显示 `[1/3]` 阶段耐心等待
- CSV 几乎瞬间完成，XLSX 因 openpyxl 逐单元格设置格式较慢（大量数据时明显）
- 时间戳使用**纳秒**，可从 Perfetto UI 中直接复制

### 输出

```
out/<trace名称>_segments/
├── <名称>_110280321901428_110286256250923_block_io.csv
├── <名称>_110280321901428_110286256250923_block_io.xlsx
├── <名称>_110290287672228_110401048612841_block_io.csv
├── <名称>_110290287672228_110401048612841_block_io.xlsx
├── <名称>_110401048612841_110457864486570_block_io.csv
└── <名称>_110401048612841_110457864486570_block_io.xlsx
```

每个文件 Sheet 名固定为 `全部 IO`，16 列结构与脚本2完全一致。

---

## io_reader.py — CSV/XLSX 统一读取模块

所有后续分析脚本（3-7）共享此模块，自动识别 `.csv` / `.xlsx` 格式：

| 函数 | 作用 |
|---|---|
| `load_io_events(path, sheet_name)` | 读取文件，返回 `(headers, events)` |
| `find_block_io_file(dir, name)` | 在目录中查找 `_block_io.xlsx`（优先）或 `.csv` |

> 数据列完全一致，CSV 和 XLSX 可互换使用。

---

## 脚本4: `io_latency_chart.py` — IO 时延计算与分布图

### 功能

从 block IO 表格（CSV/XLSX）读取事件, 配对计算调度/器件时延, 绘制分布图, 输出明细到 `latency.xlsx`。

### 使用方式

```powershell
python scripts\io_latency_chart.py --sheet "时间范围 IO" --pcap 99.9
python scripts\io_latency_chart.py --file "out/安兔兔跑分/安兔兔跑分_p1_block_io.csv" --sheet "全部 IO"
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `--file` | str | 手动指定 block_io 文件路径（支持任意命名，不指定则自动扫描 out/） |
| `--sheet` | str | Sheet 名称, 默认 `全部 IO` |
| `--pcap` | float | 横坐标截断分位数（如 `99.9`）, 不指定则不截断 |

### 输出

```
out/<trace名称>/charts/<trace名称>_schedule_latency.png
out/<trace名称>/charts/<trace名称>_device_latency.png
out/<trace名称>/<trace名称>_latency.xlsx  （追加 Sheet: 调度耗时 20列, 器件耗时 20列）
```

### 时延定义

```
block_rq_insert ----[ 调度耗时 ]---> block_rq_issue ----[ 器件耗时 ]---> block_rq_complete
```

| 时延类型 | 起始事件 | 终止事件 | 含义 |
|---|---|---|---|
| 调度耗时 | `block_rq_insert` | `block_rq_issue` | IO 在调度器队列中的等待时间 |
| 器件耗时 | `block_rq_issue` | `block_rq_complete` | 块设备硬件处理时间 |

### 图表结构

2 张图, 各 3 子图（读请求 / 写请求 / 其他请求）:
- X 轴: 时延 ms（对数刻度）, Y 轴: IO 数量
- 柱上标注计数, 右上角 P50/P99/均值/最大值
- `--pcap` 截断时左下角红色 `+N IO 超过 P{xx}`

### 设计原理

**从 XLSX 匹配而非 SQL**: 按 `(设备, 扇区)` 分组 - 时序排序 - 遍历配对, 逻辑直观, 与"时间范围 IO" Sheet 天然一致。

**配对算法**:
```
1. 事件按 (设备号, 扇区地址) 分组
2. 每组按时间排序后遍历:
   - insert - 待匹配 insert
   - issue + 待匹配 insert - 调度耗时 + 待匹配 issue
   - complete + 待匹配 issue - 器件耗时
3. 未配对事件自动丢弃（trace 边界截断）
```

---

## 脚本5: `io_size_chart.py` — IO 大小分布

### 功能

统计读/写/其他请求的 IO 大小分布（扇区数和 KB）, 绘制两张分布图, 汇总追加到 `latency.xlsx`。

### 使用方式

```powershell
python scripts\io_size_chart.py --sheet "时间范围 IO"
python scripts\io_size_chart.py --file "out/安兔兔跑分/安兔兔跑分_p1_block_io.csv" --sheet "全部 IO"
```

### 输出

```
out/<trace名称>/charts/<trace名称>_io_nr_sector.png
out/<trace名称>/charts/<trace名称>_io_size_kb.png
```

追加到 `latency.xlsx` 的 `IO 大小汇总` Sheet:

| 列 | 说明 |
|---|---|
| 请求类型 | 读/写/其他（分类汇总行蓝底加粗） |
| 请求扇区数量 | `nr_sector` |
| 请求大小 (KB) | `nr_sector * 512 / 1024` |
| IO 数量 | 出现次数 |

---

## 脚本6: `io_process_stats.py` — 进程 IO 统计

### 功能

按进程/线程统计 `block_rq_insert` 事件的读写请求数和数据量, 追加到 `latency.xlsx`。

> 仅统计 insert, 避免 issue/complete 阶段同一 IO 被重复计算。

### 使用方式

```powershell
python scripts\io_process_stats.py --sheet "时间范围 IO"
python scripts\io_process_stats.py --file "out/安兔兔跑分/安兔兔跑分_p1_block_io.csv" --sheet "全部 IO"
```

### 输出

追加到 `latency.xlsx` 的 `进程 IO 统计` Sheet（9列）:

| IO 发起进程 | 进程 PID | 线程 TID | IO 发起线程 | 读 | 写 | 其他 | 总计 | 总量(KB) |

按 **总数据量降序**排列, 末尾合计行。

---

## 脚本7: `pid_io_analysis.py` — 指定 PID/TID 全维度分析

### 功能

针对单个或多个 PID 或 TID, 生成 IO 分析: 调度时延、器件时延、IO 大小、线程统计。
支持同时分析多个目标（逗号/空格/分号分隔）, 多目标数据合并输出到同一文件夹与同一份表格。

### 使用方式

```powershell
python scripts\pid_io_analysis.py --pid 5736 --sheet "时间范围 IO" --pcap 99.9
python scripts\pid_io_analysis.py --tid 17679 --sheet "时间范围 IO" --pcap 99.9
python scripts\pid_io_analysis.py --pid 5736 --file "out/安兔兔跑分/安兔兔跑分_p1_block_io.csv" --sheet "全部 IO"
# 同时分析多个线程, 限定时间范围
python scripts\pid_io_analysis.py --tid 22306,22307,22308 --start 126.1 --end 126.2 --sheet "全部 IO"
```

### 命令行参数

| 参数 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `--file` | str | - | 手动指定 block_io 文件路径（支持任意命名） |
| `--pid` | str | 二选一 | 目标进程 PID, 多个用中/英文逗号、分号或空格分隔 |
| `--tid` | str | 二选一 | 目标线程 TID, 多个用中/英文逗号、分号或空格分隔 |
| `--sheet` | str | - | Sheet 名称, 默认 `全部 IO`; 不存在时交互列出所有 Sheet 供编号选择 (`?` 强制选择) |
| `--start` | float | - | 起始时间 (秒), 只分析该时间之后的 IO |
| `--end` | float | - | 结束时间 (秒), 只分析该时间之前的 IO |
| `--pcap` | float | - | 时延图表截断分位数 |

> `--pid` 和 `--tid` 必须指定且只能指定其中一个；多个目标用中/英文逗号、分号或空格分隔（如 `22306，22307` 或 `22306,22307`）。
> `--start`/`--end` 可选, 用于圈定任意时间段分析。

### 输出

```
out/<trace名称>/TID_<值...>/<trace名称>_<值...>.xlsx   # 4个Sheet（合并数据）
out/<trace名称>/TID_<值...>/<trace名称>_<值...>_*.png  # 4张图表
```

- 目录名为 `PID_<值...>` 或 `TID_<值...>`, 多目标时用 `_` 连接（如 `TID_22306_22307_22308`）。

XLSX Sheet:

| Sheet | 内容 |
|---|---|
| 调度耗时 | insert - issue 明细（20列） |
| 器件耗时 | issue - complete 明细（20列） |
| IO 大小汇总 | 读/写/其他分布 |
| 进程 IO 统计 | 各线程 IO 量与数据量（含所有目标线程） |

### 设计原理: insert 是 IO 的标识

```
 +----------- 该 PID/TID -----------+ +-- 可能为内核线程 --+
 | block_rq_insert (P=5736/T=17679)| | block_rq_issue    | | block_rq_complete |
 | 从这里识别"这是谁的 IO"          |->| (ksoftirqd/NVMe) |->|                  |
 +----------------------------------+ +-------------------+ +-------------------+
```

- **`--pid` 模式**: 仅 insert 限定目标 PID, 统计该进程下所有线程的 IO
- **`--tid` 模式**: 仅 insert 限定目标 TID, 精确到单个线程的 IO
- **多目标合并**: 多个 PID/TID 视为一个集合, 任一目标命中的 insert 事件都计入, 合并输出到同一文件夹/表格
- **配对用全部事件**: issue/complete 不限 PID/TID（可能由内核中断上下文执行）
- **时间范围**: `--start`/`--end` 按「时间 (纳秒)」字段过滤, 与综合诊断脚本逻辑一致
- **`pending_issue` 仅在 insert 匹配成功时赋值**: 防止其他进程/线程同扇区 IO 产生幽灵配对
- **统计用 insert 事件**: 大小/线程只统计目标 insert 事件

---

## 脚本8: `io_diagnosis.py` — IO 综合诊断

### 功能

一次性输出 **10 项核心诊断指标**，帮助快速定位 IO 性能瓶颈：

| # | 分析项 | 输出 |
|---|---|---|
| 1 | **时延百分位汇总** | P50/P95/P99/P99.9/P99.99/Avg/Max 对比表 |
| 2 | **读写分离** | 读/写 IO 数量占比、数据量、时延差异 |
| 3 | **队列深度时间曲线** | 分设备统计：调度队列深度 + 器件队列深度 时间变化（最多 6 个最活跃设备） |
| 4 | **IOPS & 吞吐量时间曲线** | 分设备统计：按时间分桶的 IOPS（柱状）+ 吞吐量（折线） |
| 5 | **顺序/随机 IO 检测** | 按线程+设备分组，判定顺序 vs 随机访问模式，按读写拆分 |
| 6 | **IO 大小 vs 时延散点图** | 读/写 IO 大小与端到端时延的关系，含趋势线 |
| 7 | **慢 IO Top-N 明细** | 按端到端时延降序，含分设备队列深度、线程信息、访问模式 |
| 8 | **同步/异步 IO 分类与对比** | 按 rwbs 标志位拆分同步写/异步写/同步读，百分位对比箱线图 |
| 9 | **时延时间线** | 三时延维度（端到端/调度/器件）散点图 + 底部 IO 密度，同步写红色高亮 |
| 10 | **LBA 时间线分布** | 每个 insert 事件的扇区地址(LBA)随时间分布散点图，一眼看出 IO 离散程度 |

### 使用方式

```powershell
python scripts\io_diagnosis.py                                          # 读取配置文件中的参数
python scripts\io_diagnosis.py --sheet "时间范围 IO"                    # 指定 Sheet
python scripts\io_diagnosis.py --start 13229.127 --end 13235.444        # 时间范围（秒）
python scripts\io_diagnosis.py --bucket 0.5                             # IOPS 分桶间隔（0.5秒）
python scripts\io_diagnosis.py --file "out/安兔兔跑分/安兔兔跑分_p1_block_io.csv" --sheet "全部 IO"
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `--file` | str | 手动指定 block_io 文件路径（支持任意命名，不指定则自动扫描 out/） |
| `--sheet` | str | Sheet 名称，默认 `全部 IO`；不存在时交互列出所有 Sheet 供编号选择（`?` 强制选择）|
| `--start` | float | 时间起始（秒） |
| `--end` | float | 时间结束（秒） |
| `--start-ns` | int | 时间起始（纳秒），可从 trace 直接复制 |
| `--end-ns` | int | 时间结束（纳秒） |
| `--bucket` | float | IOPS/吞吐量分桶间隔（秒），默认 0.1 (100ms) |
| `--top-n` | int | 慢 IO 显示数量，默认 20 |

### 输出

```
out/<trace名称>/
├── <trace名称>_diagnosis.xlsx          # 10 Sheet Excel
│   ├── 时延百分位                      # 调度/器件/端到端 P50-P99.99 汇总
│   ├── 时延汇总                        # 简洁汇总表格
│   ├── 读写分离                        # 读/写/Other 数量、数据量、时延
│   ├── 队列深度                        # 分设备（设备号列）+ 调度队列 + 器件队列 时间序列
│   ├── IOPS_吞吐量                     # 分设备（设备号列）+ 分桶 IOPS + MB/s 时间序列
│   ├── 顺序随机                        # 顺序/随机 IO 占比、数据量、时延
│   ├── 顺序IO-读写拆分                 # 顺序 IO 按 read/write 统计
│   ├── 同步IO                          # 同步写/异步写/同步读 数量、占比、时延对比
│   ├── 慢IO TopN                       # 慢 IO 明细（含队列深度、线程信息）
│   └── 时延时间线                      # 逐 IO 明细（insert/issue/complete 时间 + 三段时延，供 AI 分析）
└── charts/
    ├── <trace名称>_queue_depth.png     # 分设备队列深度时间曲线（每设备一子图，最多 6 个）
    ├── <trace名称>_iops_throughput.png # 分设备 IOPS + 吞吐量 时间曲线（每设备一子图，双 Y 轴）
    ├── <trace名称>_rw_latency.png      # 读写时延对比（Box Plot）
    ├── <trace名称>_rw_size.png         # 读写 IO 大小对比（直方图）
    ├── <trace名称>_seq_random.png      # 顺序/随机占比饼图 + 读写拆分柱状图
    ├── <trace名称>_sync_io_latency.png # 同步 vs 异步写时延箱线图 + 百分位对比
    ├── <trace名称>_latency_timeline.png# 时延时间线散点图（三时延维度 + 密度）
    ├── <trace名称>_lba_timeline.png    # LBA 扇区地址时间线分布图（按读写类型着色）
    └── <trace名称>_size_latency.png    # IO 大小 vs 时延 散点图（读蓝/写橙）
```

### 诊断分析指南

| 图表 | 看什么 | 常见问题判断 |
|---|---|---|
| **队列深度** | 分设备看调度队列 vs 器件队列哪个更高 | 调度队列高 = OS 层瓶颈；器件队列高 = 磁盘瓶颈；UFS/dm/loop/eMMC 各设备独立统计，互不污染 |
| **IOPS+吞吐** | 分设备看是否存在突刺(spike)、持续高负载 | 突刺说明 IO 突发；持续高负载需检查是否达到硬件极限；多设备场景按设备隔离，避免互相污染 |
| **时延百分位** | P99/P99.9 与 P50 的比值 | 比值 > 10x 说明尾延迟严重（长尾效应） |
| **读写分离** | 写的 P99 是否显著高于读 | 写延迟通常受 Sync/FUA 影响，高写延迟需检查文件系统或磁盘 |
| **读写大小** | 读/写的大小分布是否一致 | 不一致通常正常（读预取 > 写入通常碎片化） |
| **顺序/随机** | 顺序 IO 占比、顺序/随机时延差异 | 随机 IO 占比过高说明 IO 模式不理想，顺序 IO 时延应显著低于随机 |
| **散点图** | 大小与时延是否正相关、是否存在异常离群点 | 趋势线斜率过大说明大 IO 时延增长快；远离趋势线的点可能是畸变慢 IO |
| **慢IO TopN** | 慢 IO 集中在哪个线程、哪种类型、调度还是器件耗时 | 调度耗时高 = OS 排队等待；器件耗时高 = 磁盘响应慢；队列深度高时发生慢 IO = 拥塞 |
| **同步IO** | 同步写 vs 异步写占比，同步写 P99/P99.9 是否显著高于异步写 | 同步写占比高（>80%）表明大量 fsync；安兔兔跑分场景下同步写 P99.9 直接决定跑分；同步写调度时延高 = 排队等 fsync 完成 |
| **时延时间线** | 三个面板（端到端/调度/器件）X 轴时间 Y 轴时延，同步写红色三角标记，查看 P99.9 参考线和 Top-3 高时延标注 | 毛刺(jitter)时间窗口对应 IO 密度高峰 = IO 拥塞导致的滑动卡顿；调度面板高 = 排队在 OS；器件面板高 = 存储硬件慢 |
| **LBA 时间线** | X 轴时间 Y 轴扇区地址(LBA)，按读写类型着色（蓝圆读/绿方异步写/红三角同步写），观察 LBA 是否连续成带还是大幅跳跃 | LBA 连续成带 = 顺序访问；LBA 大幅跳跃散布 = 随机访问、离散度高（HDD 寻道惩罚大）；写 LBA 集中在窄区间 = 写放大风险 |

---

## 脚本9: `generate_report.py` — 生成可视化分析报告

### 功能

扫描 `out/<trace名称>/` 目录下步骤 1-8 的全部产物（systrace、xlsx、CSV、charts/ 下的 PNG），
生成一份自包含的 HTML 分析报告：所有 PNG 图表以 base64 内嵌，单文件即可独立打开/分发，
每步/每个文件均附存储专家级分析指引（来自 `expert_notes.py`）。

每个表格文件（XLSX）还会读取内部结构，在报告中附上「表格内容概览」：
各 Sheet 名称与数据行数；对 `_block_io` 表格额外统计**有效 IO Size 记录数**及
**read / write / other 分布**（口径与 `io_size_chart.py` 一致，仅按 `请求扇区数量 > 0` 的记录归类）。

### 使用方式

```powershell
python scripts\generate_report.py                            # 自动选择唯一 trace（或读配置文件）
python scripts\generate_report.py --trace-name "安兔兔跑分"    # 指定 trace 名称
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `--trace-name` | str | trace 名称，默认读取配置文件中的 `trace_name`；out 下仅一个 trace 时自动选择 |

### 输出

```
out/<trace名称>/<trace名称>_report.html   # 自包含 HTML（图片内嵌，无外链资源）
```

### 说明

- **图片查找**: PNG 图表统一从 `charts/` 子目录扫描（步骤 4/5/8 的图表均按此规范存放）
- **指定 PID/TID 分析**: 自动探测 `PID_*/`、`TID_*/` 子目录及其中的 XLSX 与 4 张进程图表（步骤 7）
- **报告结构**: 按步骤 1-8 分区展示，每步附专家分析指引与诊断思路
- **表格内容概览**: 每个 XLSX 文件下方展示其 Sheet 名称与数据行数；`_block_io` 表格另附有效 IO Size 记录数及 read/write/other 分布（无法解析时自动跳过，不影响报告生成）

---

## AI 卡顿分析: `scripts/AI/io_stall_analysis.py` — 单笔 IO 卡顿根因拆解

### 背景与定位

针对一段 `Uninterruptible Sleep (IO)` / D 状态线程卡顿，输入 **Start time / Duration / Process**，
自动定位责任 IO，按「存储专家方法论」拆解耗时点与根因。与脚本4/8的「全量统计」不同，
本脚本聚焦「单笔慢 IO」的深度归因，输出可直接落地的四维度报告。

### 经验文档: `skills/io-stall-analysis/SKILL.md`

固化存储专家的完整分析框架（详见该文件）：

- **四维度拆解**：① 该笔 IO 完整情况（insert/issue/complete 时间与上下文）② 耗时点统计（调度 vs 器件占比）③ 调度慢根因（CPU/队列满/插队）④ 器件慢根因（complete 断档/读写交叉/扇区离散度）
- **队列满阈值**：UFS 命令队列上限典型 32；eMMC 通常只能同时 issue 4 个请求
- **插队定义**：后面 insert 的 IO 反而先 issue（insert 顺序 vs issue 顺序被打乱），而非只看「有几个线程在提交」
- **根因判定决策树**：complete 断档对比 Duration → in-flight 深度 → 读写形态
- **专家补充根因**（易漏）：cgroup 限流、tag 耗尽、请求合并失配、plug/unplug、SLC cache 悬崖、写放大、GC/读干扰/ECC、Thermal Throttling、脏页 writeback 竞争、扇区离散度过高
- **关键陷阱**：`block_rq_complete` 时间戳是「软中断执行回调时刻」而非「硬件完成时刻」，见 5.4 节

### 使用方式

```powershell
# 首次运行前先做一次 CSV 缓存（60MB xlsx 遍历慢，转 CSV 后秒级读取）
python scripts\AI\convert_to_csv.py --trace com.ss.android.ugc.aweme

# 一键分析：输入 Start time / Duration / Process
python scripts\AI\io_stall_analysis.py `
    --trace com.ss.android.ugc.aweme `
    --start 126139322335469 `
    --duration 32701962 `
    --process com.antutu.ABenchMark --tid 22288
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `--trace` | str (必填) | trace 名称（如 com.ss.android.ugc.aweme） |
| `--start` | int (必填) | Start time（纳秒） |
| `--duration` | int (必填) | Duration（纳秒） |
| `--process` | str | 进程名（可选，用于定位责任 IO） |
| `--pid` | str | 进程 PID（可选） |
| `--tid` | str | 线程 TID（可选，优先级最高） |
| `--pre-ms` | int | 卡顿点前窗口（默认 50ms） |
| `--post-ms` | int | 卡顿点后窗口（默认 200ms，complete 在窗口外时需加大） |

### 输出

```
out/<trace名称>/stall_analysis/<start_ns>_<dur_ns>/
├── io_stall_report.md    # 主报告（7节规则分析 + 可选「AI 根因研判」章节，由 --ai 开启）
└── io_stall_data.txt     # 原始明细（目标线程事件、in-flight、堆积请求，数据来源可追溯）
```

### 分析框架（报告结构）

| 章节 | 内容 |
|---|---|
| 1. 目标 IO 定位 | 卡顿点前符合条件的最后一次 insert（sector/dev） |
| 2. 三段式生命周期 | insert / issue / complete 时间戳、触发者、上下文 |
| 3. 耗时点统计 | 调度耗时（issue−insert）vs 器件耗时（complete−issue）占比 |
| 4. 调度阶段排查 | CPU 资源 / in-flight 深度（UFS 32 / eMMC 4）/ 堆积请求 / 插队 |
| 5. 器件阶段排查 | complete 断档 / 读写交叉 / 扇区离散度 / **软中断延迟甄别** |
| 6. 根因结论 | 结合断档时长与 Duration 吻合度、in-flight 深度下结论 |
| 7. 数据来源 | 原始 xlsx / 实际读取 CSV 缓存 / 配对规则 |

### 5.4 软中断延迟甄别（关键认知）

`block_rq_complete` 记录的是块层**软中断执行完成回调的时刻**，不是硬件真正完成 IO 的时刻。
硬件完成 IO 后先产生硬件中断（IRQ）→ 触发软中断（BLK_SOFTIRQ）→ 软中断里才打出 complete。
若硬件中断早已回来、但 CPU 没空运行软中断，complete 时间戳会偏晚，导致「器件耗时 = complete − issue」被高估。

脚本通过两个信号自动甄别：

- **complete 聚簇**：连续 complete 间隔 < 50us 的相邻对数（显著聚簇 = ≥5 对且占比 >20%）
- **complete 记录进程**：集中在 `ksoftirqd`（软中断溢出）/ `swapper`（中断/idle 上下文）

> 均匀分布 = 真实器件耗时；显著聚簇 = 器件耗时虚高（根因在 CPU 调度侧）。

### 数据来源与缓存

- 原始数据：`out/<trace>/<trace>_block_io.xlsx`（Sheet: 全部 IO）
- CSV 缓存：`convert_to_csv.py` 生成的 `_block_io.csv`，脚本优先读取，速度远快于 xlsx
- 配对规则：按 `(设备号, 请求扇区地址)` 分组，insert→issue→complete 三段式匹配，避免全局计数漂移

### 实战案例结论（com.ss.android.ugc.aweme）

Start 126139322335469 / Duration 32701962 / TID 22288（Thread-45，com.antutu.ABenchMark）：

- 调度耗时 39.921ms (74.1%)，器件耗时 13.918ms (25.9%)
- complete 断档 32.798ms ≈ Duration 32.702ms，in-flight=26（< UFS 上限 32）、堆积=17，纯写负载无读写交叉、扇区部分随机（跳跃占比 18.8%，对器件效率有影响）
- 根因：UFS 顺序大块同步写下写缓存耗尽，触发 FTL GC/刷盘写停滞（调度耗时是「排队等设备恢复」的投影）

---

## AI 报告问答: `scripts/AI/io_chat.py` — 围绕分析结果持续对话

### 背景与定位

AI1 输出单笔 IO 卡顿报告后，AI2 提供围绕该报告（及任意本地文件/目录）的多轮交互问答。
它**不是「纯聊天」脚本**：直接粘贴一个本地文件或目录路径，会自动读取/列出内容并交由 AI 解读，
解决「AI 无法读取本地路径」的问题。

### 核心能力

| 能力 | 说明 |
|---|---|
| **多轮对话** | 对话历史持续追加，可结合报告上下文连续追问 |
| **报告上下文** | 自动加载 `out/<trace>/stall_analysis` 下最新 `io_stall_report.md` 作为系统提示 |
| **本地路径识别** | 直接粘贴存在的文件/目录路径（绝对路径或相对项目根路径），自动读取并解读 |
| **目录分级加载** | `.md` 报告完整读取；`.txt/.csv/.log/.json/.html` 等文本预览前 2000 字符；二进制文件明确提示不支持 |
| **内置命令** | `/read` 读文件、`/ls` 列目录、`/clear` 清历史、`/help` 帮助、`/quit` 退出 |
| **步骤级超时** | 主控脚本中 AI2 单独配置 3600 秒超时，适配长时对话场景 |

### 使用方式

```powershell
# 默认加载最新 stall 报告 + fast 模型
python scripts\AI\io_chat.py --trace com.ss.android.ugc.aweme

# 指定报告 + 深度推理模型
python scripts\AI\io_chat.py `
    --trace com.ss.android.ugc.aweme `
    --report out\com.ss.android.ugc.aweme\stall_analysis\126139322335469_32701962\io_stall_report.md `
    --model deep
```

### 命令行参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `--trace` | str (必填) | trace 名称，用于定位 `out/<trace>/stall_analysis` 下的报告 |
| `--report` | str | 指定报告路径（可选，默认加载最新 stall 报告） |
| `--model` | str | 模型角色：`fast`（deepseek-chat，默认）/ `deep`（deepseek-reasoner） |

### 对话命令

| 命令 | 说明 |
|---|---|
| `/help` 或 `?` | 显示帮助 |
| `/ls [目录]` | 列出目录内容（默认当前 trace 输出目录） |
| `/read <文件>` | 读取文本文件并让 AI 解读（单文件上限 15000 字符） |
| `/clear` | 清空对话历史（保留报告上下文） |
| `/quit` / `exit` / `q` | 退出 |

### 本地路径识别规则

| 输入 | 处理 |
|---|---|
| 存在的**目录** | 列出目录项 → 分级加载内容（报告完整读、文本预览、二进制跳过）→ 交 AI 解读 |
| 存在的**文件** | 读取文本内容（二进制明确提示不支持）→ 交 AI 解读 |
| 普通文本 | 作为正常对话提问 |

> 路径支持绝对路径（如 `D:\...\126139322335469_32701962`）或相对项目根路径，自动剥离首尾引号。

### 文件类型分级加载

| 类型 | 扩展名 | 处理 |
|---|---|---|
| 报告 | `.md` | 完整读取（上限 20000 字符） |
| 文本预览 | `.txt .csv .log .json .html .xml .py ...` | 预览前 2000 字符，超长提示用 `/read` 读完整 |
| 二进制 | `.xlsx .png .perfetto .pdf .db ...` | 明确提示不支持直接读取，需先转为文本 |

---

## 主控脚本: `ai_trace.py` — 交互式统一入口

### 功能

将 9 个分析步骤与 2 个 AI 分析选项（AI1 / AI2）串成统一入口，提供交互式菜单、一键全流程、参数记忆、状态持久化。

### 使用方式

```powershell
python ai_trace.py                  # 交互模式
python ai_trace.py --all            # 一键全流程（含参数配置）
python ai_trace.py --step 3         # 直接执行指定步骤
python ai_trace.py --reset          # 重置参数记忆
```

### 交互菜单

```
============================================================
  AI-trace — Perfetto Block IO 分析工具
============================================================

  Trace 文件: 2 个
  当前选择: 安兔兔跑分
  输出目录: out/安兔兔跑分

  分析流程:
  [✓] 1. 格式转换 (perfetto → systrace)            [必选]
  [✓] 2. IO 事件提取 (16列 XLSX/CSV)               [必选]
  [ ] 3. 超大 Trace 时间切片拆分                     [可选]
  [ ] 4. IO 时延分析 (调度+器件耗时分布图)           [可选]
  [ ] 5. IO 大小分布 (扇区+KB分布图)                [可选]
  [ ] 6. 进程 IO 统计 (按线程, 数据量降序)           [可选]
  [ ] 7. 指定 PID/TID 全维度分析 (独立 XLSX + 4图)   [可选]
  [ ] 8. IO 综合诊断 (队列深度/IOPS/百分位/读写分离)  [可选]
  [ ] 9. 生成分析报告                                [可选]

  AI分析选项:
  [ ] AI1. IO 卡顿快速分析 (阻塞 IO 根因拆解)
  [ ] AI2. IO 报告 AI 问答 (围绕 trace 分析结果持续对话)

  全局操作:
  [A] 一键全流程 (步骤1→9)
  [C] 切换 / 选择 trace 文件
  [R] 重置参数记忆
  [Q] 退出
```

### 核心特性

| 特性 | 说明 |
|---|---|
| **参数记忆** | 上次输入的 `--start`、`--end`、`--pid` 等参数自动保存到 `.last_params.json` |
| **配置文件** | 步骤 1-9 参数集中在 `config/default.json`；AI 分析选项参数集中在 `config/ai_default.json`；大模型接入在 `config/ai_model.json` |
| **AI 分析选项** | 菜单集成 AI1（IO 卡顿快速分析）与 AI2（IO 报告 AI 问答），AI2 单独 3600 秒超时适配长时对话 |
| **状态追踪** | 已完成步骤标记 `✓`，重开脚本也能看到进度 |
| **容错设计** | 各步骤独立子进程执行，任一步骤失败不中断后续步骤 |
| **Trace 选择** | 多 trace 文件时交互选择，单文件自动跳过 |
| **表格选择** | 步骤4-8 自动扫描 `out/` 及 `_segments/` 目录下所有 `*_block_io.*` 文件，支持交互选择切片表格 |
| **Sheet 选择** | 读取 XLSX 时，Sheet 名不存在或传 `?` 时，自动列出所有 Sheet 名称供按编号选择 |
| **多目标分隔** | `--pid`/`--tid` 支持中/英文逗号、分号或空格分隔，中文逗号「，」也能正确解析 |

### 命令行参数

| 参数 | 说明 |
|---|---|
| `(无)` | 进入交互式菜单模式 |
| `--all` | 一键执行全部 9 步，按需配置参数 |
| `--step {1-9}` | 直接执行指定步骤（需先有已保存的参数） |
| `--reset` | 清除 `.last_params.json` 重置所有记忆 |

---

## 典型工作流

### 方式一: 主控脚本（推荐）

```powershell
python ai_trace.py
# 按菜单提示依次执行，参数自动记忆，下次直接 --all 一键全流程
```

### 方式二: 手动逐脚本执行

```powershell
# 1. 提取 IO 事件（带时间范围）
python scripts\parse_block_io.py --start 13229.127 --end 13235.444

# 2. IO 时延分析
python scripts\io_latency_chart.py --sheet "时间范围 IO" --pcap 99.9

# 3. IO 大小分布
python scripts\io_size_chart.py --sheet "时间范围 IO"

# 4. 进程 IO 统计
python scripts\io_process_stats.py --sheet "时间范围 IO"

# 5. 指定进程/线程深度分析（支持多值逗号分隔 + 时间范围）
python scripts\pid_io_analysis.py --pid 5736 --sheet "时间范围 IO" --pcap 99.9
python scripts\pid_io_analysis.py --tid 22306,22307,22308 --start 126.1 --end 126.2 --sheet "全部 IO"

# 6. IO 综合诊断
python scripts\io_diagnosis.py --sheet "时间范围 IO"

# 打开 out/<trace名称>/ 查看全部结果
```

### 方式三: 超大 Trace 切片工作流

针对 > 1GB 的 trace 文件，先拆分再逐片段分析：

```powershell
# 方法 A: 交互菜单（推荐）
python ai_trace.py
# 步骤2: 提取 IO 事件
# 步骤3: 拆分 trace（输入时间片段 st:end）
# 步骤4-8: 选切片 → 菜单自动列出所有 _block_io.* 供选择

# 方法 B: 手动拆分后在终端中用 --file 指向切片文件
python scripts\split_trace.py "安兔兔跑分" --pairs 110280321901428:110286256250923 110290287672228:110401048612841 --format csv

# 后续分析直接 --file 指向任意切片文件（支持任意命名）
python scripts\io_latency_chart.py --file "out/安兔兔跑分_segments/安兔兔跑分_110280321901428_110286256250923_block_io.csv" --sheet "全部 IO"
python scripts\io_size_chart.py --file "out/安兔兔跑分_segments/安兔兔跑分_110280321901428_110286256250923_block_io.csv" --sheet "全部 IO"
```

### 方式四: AI 卡顿分析（单笔 IO 深度归因）

针对「某线程 D 状态卡顿了一段」，从 Perfetto UI 拿到 Start time / Duration / Process 后一键定位：

```powershell
# 1. 首次先做 CSV 缓存（可选，脚本会自动回退到 xlsx）
python scripts\AI\convert_to_csv.py --trace com.ss.android.ugc.aweme

# 2. 一键分析（Start time 与 Duration 均为纳秒）
python scripts\AI\io_stall_analysis.py `
    --trace com.ss.android.ugc.aweme `
    --start 126139322335469 `
    --duration 32701962 `
    --process com.antutu.ABenchMark --tid 22288
```

### 方式五: AI 报告问答（围绕分析结果持续对话）

完成 AI1 报告后，进入 AI2 围绕报告 / 本地文件 / 目录持续问答：

```powershell
# 默认加载最新 stall 报告 + fast 模型
python scripts\AI\io_chat.py --trace com.ss.android.ugc.aweme

# 交互示例：直接粘贴本地目录路径，自动读取并解读
# 你 > D:\debug-work\block-io\AI-trace\out\com.ss.android.ugc.aweme\stall_analysis\126139322335469_32701962
# AI > （列出目录项，完整加载 io_stall_report.md，预览 io_stall_data.txt，并给出解读）

# 也可用 /read、/ls 手动操作，输入 /help 查看全部命令
```