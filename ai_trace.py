#!/usr/bin/env python3
"""
AI-trace 交互式主控脚本
统一入口，串起全部 9 个分析脚本，支持:
  - 交互式菜单
  - 一键全流程
  - 参数记忆 (.last_params.json)
  - 单步执行
"""

import os
import sys
import json
import subprocess
import glob

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(PROJECT_ROOT, "scripts")
STATE_FILE = os.path.join(PROJECT_ROOT, ".last_params.json")
CONFIG_DIR = os.path.join(PROJECT_ROOT, "config")
CONFIG_FILE = os.path.join(CONFIG_DIR, "default.json")
AI_CONFIG_FILE = os.path.join(CONFIG_DIR, "ai_default.json")

STEP_ORDER = ["1", "2", "3", "4", "5", "6", "7", "8", "9"]

STEPS = {
    "1": {"name": "格式转换 (perfetto → systrace)",         "script": "convert_trace.py",
          "args": [], "required": True},
    "2": {"name": "IO 事件提取 (16列 XLSX/CSV)",            "script": "parse_block_io.py",
          "args": ["--start", "--end", "--timeout"], "required": True},
    "3": {"name": "超大 Trace 时间切片拆分",                "script": "split_trace.py",
          "args": ["--pairs", "--format", "--timeout"], "required": False, "positional": True},
    "4": {"name": "IO 时延分析 (调度+器件耗时分布图)",       "script": "io_latency_chart.py",
          "args": ["--sheet", "--pcap"], "required": False},
    "5": {"name": "IO 大小分布 (扇区+KB分布图)",             "script": "io_size_chart.py",
          "args": ["--sheet"], "required": False},
    "6": {"name": "进程 IO 统计 (按线程, 数据量降序)",       "script": "io_process_stats.py",
          "args": ["--sheet"], "required": False},
    "7": {"name": "指定 PID/TID 全维度分析 (独立 XLSX + 4图)",
          "script": "pid_io_analysis.py",
          "args": ["--pid", "--tid", "--start", "--end", "--sheet", "--pcap"], "required": False},
    "8": {"name": "IO 综合诊断 (队列深度/IOPS/百分位/读写分离)",
          "script": "io_diagnosis.py",
          "args": ["--start", "--end", "--bucket", "--sheet", "--top-n"], "required": False},
    "9": {"name": "生成分析报告 (汇总所有产物到 HTML)",      "script": "generate_report.py",
          "args": [], "required": False},
}

# AI 分析选项（独立大项，后续可继续扩充 AI3 ...）
AI_STEP_ORDER = ["AI1", "AI2"]

AI_STEPS = {
    "AI1": {
        "name": "IO 卡顿快速分析 (阻塞 IO 根因拆解)",
        "script": "io_stall_analysis.py",
        "subdir": "AI",
        "args": ["--trace", "--start", "--duration", "--process", "--pid", "--tid",
                 "--pre-ms", "--post-ms", "--ai-model"],
        "flags": ["--ai"],
        "required": False,
    },
    "AI2": {
        "name": "IO 报告 AI 问答 (围绕 trace 分析结果持续对话)",
        "script": "io_chat.py",
        "subdir": "AI",
        "args": ["--trace", "--model"],
        "required": False,
        "timeout": 3600,
    },
}


def load_state():
    if os.path.isfile(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, Exception):
            pass
    return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def load_config():
    """加载配置文件，供交互模式默认值使用"""
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, Exception):
            pass
    return {}


def load_ai_config():
    """加载 AI 分析选项配置文件 (config/ai_default.json)，供 AI 脚本参数默认值使用"""
    if os.path.isfile(AI_CONFIG_FILE):
        try:
            with open(AI_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, Exception):
            pass
    return {}


def detect_traces():
    trace_dir = os.path.join(PROJECT_ROOT, "trace")
    if not os.path.isdir(trace_dir):
        return []
    files = glob.glob(os.path.join(trace_dir, "*.perfetto-trace"))
    files += glob.glob(os.path.join(trace_dir, "*.perfetto"))
    files = sorted(set(files))
    return [os.path.basename(f) for f in files]


def select_trace(traces, state):
    if not traces:
        print("\n[错误] trace 目录下未找到 .perfetto-trace 或 .perfetto 文件，请先将 trace 放入 trace/ 目录")
        return None, None

    if len(traces) == 1:
        name = traces[0].replace(".perfetto-trace", "")
        if name.endswith(".perfetto"):
            name = name[:-len(".perfetto")]
        print(f"\n自动选择唯一 trace: {traces[0]}")
        out_dir = os.path.join(PROJECT_ROOT, "out", name)
        state["trace_file"] = traces[0]
        state["trace_name"] = name
        state["out_dir"] = out_dir
        save_state(state)
        return traces[0], name

    print("\n可用的 trace 文件:")
    for i, t in enumerate(traces, 1):
        mark = " [上次使用]" if state.get("trace_file") == t else ""
        print(f"  {i}. {t}{mark}")

    default_idx = None
    if state.get("trace_file") in traces:
        default_idx = traces.index(state["trace_file"]) + 1

    while True:
        prompt = f"请选择 (1-{len(traces)})" + (f" [默认 {default_idx}]" if default_idx else "") + ": "
        choice = input(prompt).strip()
        if choice == "" and default_idx:
            choice = str(default_idx)
        try:
            idx = int(choice)
            if 1 <= idx <= len(traces):
                sel = traces[idx - 1]
                break
        except ValueError:
            pass
        print(f"  输入无效，请输入 1-{len(traces)}")

    name = sel.replace(".perfetto-trace", "")
    if name.endswith(".perfetto"):
        name = name[:-len(".perfetto")]
    out_dir = os.path.join(PROJECT_ROOT, "out", name)
    state["trace_file"] = sel
    state["trace_name"] = name
    state["out_dir"] = out_dir
    save_state(state)
    return sel, name


def _resolve_block_io_sheets(state):
    """定位目标 block_io 表格并返回其 Sheet 名称列表（找不到返回 None）。

    优先主表格 <trace_name>_block_io.xlsx，其次 out_dir 与 segments 目录下的任意 xlsx。
    """
    sys.path.insert(0, SCRIPTS_DIR)
    from io_reader import get_sheet_names

    trace_name = state.get("trace_name", "")
    candidates = []

    out_dir = state.get("out_dir", "")
    if out_dir and os.path.isdir(out_dir):
        candidates.extend(glob.glob(os.path.join(out_dir, "*_block_io.xlsx")))

    segments_dir = os.path.join(PROJECT_ROOT, "out", trace_name + "_segments")
    if os.path.isdir(segments_dir):
        candidates.extend(glob.glob(os.path.join(segments_dir, "*_block_io.xlsx")))

    candidates = sorted(set(candidates))

    preferred = os.path.join(out_dir, f"{trace_name}_block_io.xlsx") if out_dir else ""
    if preferred and preferred in candidates:
        candidates.remove(preferred)
        candidates.insert(0, preferred)

    for f in candidates:
        names = get_sheet_names(f)
        if names:
            return names
    return None


def prompt_params(step_key, state):
    step = STEPS[step_key]
    params = {}
    cfg = load_config()

    section_map = {
        "2": "parse_block_io",
        "3": "split_trace",
        "4": "io_latency_chart",
        "5": "io_size_chart",
        "6": "io_process_stats",
        "7": "pid_io_analysis",
        "8": "io_diagnosis",
    }
    section_cfg = cfg.get(section_map.get(step_key, ""), {})

    def _cfg(key, fallback=""):
        val = state.get(key)
        if val:
            return str(val)
        cv = section_cfg.get(key)
        if cv is not None:
            return str(cv)
        return str(fallback)

    def _time_info(key):
        """返回 (显示标签, 秒值) — 优先共享时间，其次本步骤配置，再回退其他步骤时间"""
        val = state.get(key)
        if val:
            return f"{val} 秒", val
        # 兼容旧版本：从其它数字步骤继承已保存的时间（如 2_start、8_start）
        for k, v in state.items():
            if k.endswith(f"_{key}") and k.split("_")[0].isdigit() and v:
                return f"{v} 秒", str(v)
        ns_key = f"{key}_ns"
        ns_val = section_cfg.get(ns_key)
        s_val = section_cfg.get(key)
        if ns_val is not None:
            sec = ns_val / 1_000_000_000.0
            return f"{ns_val} ns (≈{sec:.9f} 秒)", str(sec)
        if s_val is not None:
            return f"{s_val} 秒", str(s_val)
        return "无", ""

    def _parse_time(val):
        """自动识别：>= 1e10 视为纳秒，否则视为秒"""
        try:
            num = float(val)
            if num >= 1e10:
                return str(num / 1_000_000_000.0)
            return str(num)
        except ValueError:
            return val

    for arg_name in step["args"]:
        key = f"{step_key}_{arg_name.lstrip('-')}"
        saved = state.get(key, "")

        if arg_name == "--start":
            _, sec_val = _time_info("start")
            cfg_default = sec_val
            saved = state.get("start", "")
            saved_hint = f" (上次: {saved} 秒)" if saved and str(saved) != str(cfg_default) else ""
            prompt = f"  起始时间 (秒或纳秒>=1e10, 默认: {cfg_default or '无'}{saved_hint}): "
            val = input(prompt).strip()
            if val:
                parsed = _parse_time(val)
                params["--start"] = parsed
            elif cfg_default:
                params["--start"] = cfg_default
            state["start"] = str(params.get("--start", cfg_default or ""))

        elif arg_name == "--end":
            _, sec_val = _time_info("end")
            cfg_default = sec_val
            saved = state.get("end", "")
            saved_hint = f" (上次: {saved} 秒)" if saved and str(saved) != str(cfg_default) else ""
            prompt = f"  结束时间 (秒或纳秒>=1e10, 默认: {cfg_default or '无'}{saved_hint}): "
            val = input(prompt).strip()
            if val:
                parsed = _parse_time(val)
                params["--end"] = parsed
            elif cfg_default:
                params["--end"] = cfg_default
            state["end"] = str(params.get("--end", cfg_default or ""))

        elif arg_name == "--sheet":
            cfg_default = _cfg("sheet", "全部 IO")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""

            sheet_names = _resolve_block_io_sheets(state)
            if sheet_names:
                print("  可用的 Sheet:")
                for i, n in enumerate(sheet_names, 1):
                    mark = "  ← 默认" if n == cfg_default else ""
                    print(f"    {i}. {n}{mark}")
                prompt = f"  请选择 Sheet 编号 (回车默认: {cfg_default}{saved_hint}): "
                val = input(prompt).strip()
                if val.isdigit():
                    idx = int(val)
                    chosen = sheet_names[idx - 1] if 1 <= idx <= len(sheet_names) else cfg_default
                elif val:
                    chosen = val
                else:
                    chosen = cfg_default
                params["--sheet"] = chosen
            else:
                prompt = f"  Sheet 名称 (默认: {cfg_default}{saved_hint}): "
                val = input(prompt).strip()
                params["--sheet"] = val if val else cfg_default
            state[key] = params["--sheet"]

        elif arg_name == "--pcap":
            cfg_default = _cfg("pcap", "无")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            prompt = f"  分位数截断 (如 99.9, 输入 null 表示不截断, 默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip()
            if val and val.lower() in ("null", "none", "无", "clear", "n/a", "na"):
                params.pop("--pcap", None)
            elif val:
                params["--pcap"] = val
            elif cfg_default != "无":
                params["--pcap"] = cfg_default
            state[key] = str(params.get("--pcap", ""))

        elif arg_name == "--pid":
            cfg_default = _cfg("pid", "无")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            prompt = f"  进程 PID (与 --tid 二选一, 多个用逗号分隔, 输入 null 表示不指定, 默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip()
            if val and val.lower() in ("null", "none", "无", "clear", "n/a", "na"):
                params.pop("--pid", None)
            elif val:
                params["--pid"] = val
                params.pop("--tid", None)
                state[f"{step_key}_tid"] = ""
            elif cfg_default != "无":
                params["--pid"] = cfg_default
                params.pop("--tid", None)
            state[key] = str(params.get("--pid", ""))

        elif arg_name == "--tid":
            cfg_default = _cfg("tid", "无")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            prompt = f"  线程 TID (与 --pid 二选一, 多个用逗号分隔, 输入 null 表示不指定, 默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip()
            if val and val.lower() in ("null", "none", "无", "clear", "n/a", "na"):
                params.pop("--tid", None)
            elif val:
                params["--tid"] = val
                params.pop("--pid", None)
                state[f"{step_key}_pid"] = ""
            elif cfg_default != "无":
                params["--tid"] = cfg_default
                params.pop("--pid", None)
            state[key] = str(params.get("--tid", ""))

        elif arg_name == "--bucket":
            cfg_default = _cfg("bucket", "0.1")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            prompt = f"  时间分桶间隔 (秒, 默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip()
            params["--bucket"] = val if val else cfg_default
            state[key] = params["--bucket"]

        elif arg_name == "--top-n":
            cfg_default = _cfg("top_n", "20")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            prompt = f"  慢 IO Top-N 数量 (默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip()
            params["--top-n"] = val if val else cfg_default
            state[key] = params["--top-n"]

        elif arg_name == "--pairs":
            pairs_list = section_cfg.get("pairs", [])
            if pairs_list:
                print(f"  配置文件中的时间片段 ({len(pairs_list)} 个):")
                for i, p in enumerate(pairs_list):
                    print(f"    {i+1}. {p}")
                use_cfg = input(f"  使用配置文件中的片段? (Y/n): ").strip().lower()
                if use_cfg in ("", "y", "yes"):
                    params["--pairs"] = pairs_list
                    state[key] = json.dumps(pairs_list)
                else:
                    pairs_list = []
            if not pairs_list:
                print(f"  请输入时间片段 (格式: start:end, 空格分隔, 直接回车跳过):")
                print(f"  例: 237863509246000:237865500000000 237865500000000:237867500000000")
                val = input(f"  : ").strip()
                if val:
                    pairs_list = val.split()
                    params["--pairs"] = pairs_list
                    state[key] = json.dumps(pairs_list)

        elif arg_name == "--format":
            cfg_default = _cfg("format", "both")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            print(f"  输出格式: csv (最快) / xlsx (带样式) / both (默认)")
            prompt = f"  选择格式 (默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip().lower()
            if val in ("csv", "xlsx", "both"):
                params["--format"] = val
            else:
                params["--format"] = cfg_default
            state[key] = params["--format"]

        elif arg_name == "--timeout":
            cfg_default = _cfg("timeout", "900")
            saved_hint = f" (上次: {saved})" if saved and saved != cfg_default else ""
            if step_key == "2":
                prompt = f"  执行超时秒数 (含 SQL 查询与 Excel 保存, 默认: {cfg_default}{saved_hint}): "
            else:
                prompt = f"  trace_processor_shell 超时秒数 (默认: {cfg_default}{saved_hint}): "
            val = input(prompt).strip()
            params["--timeout"] = val if val else cfg_default
            state[key] = params["--timeout"]

    if "--pid" in step["args"] and "--tid" in step["args"]:
        has_pid = params.get("--pid")
        has_tid = params.get("--tid")
        if not has_pid and not has_tid:
            print("  [警告] 必须指定 --pid 或 --tid 其中之一，跳过此步骤")
            return None
        if has_pid and has_tid:
            print("  [警告] --pid 和 --tid 不能同时指定，跳过此步骤")
            return None

    save_state(state)
    return params


def _prompt_ai2_params(state):
    """AI2 参数提示：报告问答需要 trace 名称与模型角色"""
    params = {}
    cfg = load_ai_config()
    section_cfg = cfg.get("ai_chat", {})

    # --trace
    trace_default = state.get("trace_name", "") or str(section_cfg.get("trace", "") or "")
    saved = state.get("AI2_trace", "")
    hint = f" (上次: {saved})" if saved and saved != trace_default else ""
    prompt = f"  Trace 名称 (默认: {trace_default or '无'}{hint}): "
    val = input(prompt).strip()
    if val:
        params["--trace"] = val
        state["AI2_trace"] = val
    elif trace_default:
        params["--trace"] = trace_default
    else:
        print("  [警告] 必须指定 trace 名称 (如 com.ss.android.ugc.aweme)")
        return None

    # --model
    model_default = str(section_cfg.get("model", "fast"))
    saved = state.get("AI2_model", "")
    hint = f" (上次: {saved})" if saved else ""
    prompt = f"  模型角色 fast=快速对话 / deep=深度推理 (默认: {model_default}{hint}): "
    val = input(prompt).strip().lower()
    if val in ("fast", "deep"):
        params["--model"] = val
        state["AI2_model"] = val
    elif val == "":
        params["--model"] = model_default
    else:
        print("  [警告] 无效模型角色，使用默认 %s" % model_default)
        params["--model"] = model_default

    save_state(state)
    return params


def prompt_ai_params(ai_key, state):
    """AI 分析选项的参数提示（AI1: 卡顿分析 / AI2: 报告问答）"""
    if ai_key == "AI2":
        return _prompt_ai2_params(state)

    params = {}
    cfg = load_ai_config()
    section_cfg = cfg.get("ai_stall_analysis", {})

    def _saved(arg):
        return state.get(f"{ai_key}_{arg.lstrip('-')}", "")

    def _parse_ns(val):
        """纳秒直输，或带 s 后缀视为秒并转纳秒"""
        v = val.strip().lower()
        if v.endswith("s"):
            v = v[:-1].strip()
            try:
                return str(int(float(v) * 1_000_000_000))
            except ValueError:
                return None
        try:
            return str(int(v))
        except ValueError:
            return None

    # --trace
    trace_default = state.get("trace_name", "") or str(section_cfg.get("trace", "") or "")
    saved = _saved("--trace")
    hint = f" (上次: {saved})" if saved and saved != trace_default else ""
    prompt = f"  Trace 名称 (默认: {trace_default or '无'}{hint}): "
    val = input(prompt).strip()
    if val:
        params["--trace"] = val
        state[f"{ai_key}_trace"] = val
    elif trace_default:
        params["--trace"] = trace_default
    else:
        print("  [警告] 必须指定 trace 名称 (如 com.ss.android.ugc.aweme)")
        return None

    # --start
    saved = _saved("--start")
    start_default = str(section_cfg.get("start_ns", "") or "")
    hint = f" (上次: {saved})" if saved and saved != start_default else ""
    prompt = f"  Start time (纳秒, 例 237866030786798, 或 237.866s 表秒, 默认: {start_default or '无'}{hint}): "
    val = input(prompt).strip()
    if val:
        parsed = _parse_ns(val)
        if parsed is None:
            print("  [警告] 无效的 Start time")
            return None
        params["--start"] = parsed
        state[f"{ai_key}_start"] = parsed
    elif start_default:
        params["--start"] = start_default
    else:
        print("  [警告] 必须指定 Start time (纳秒)")
        return None

    # --duration
    saved = _saved("--duration")
    dur_default = str(section_cfg.get("duration_ns", "") or "")
    hint = f" (上次: {saved})" if saved and saved != dur_default else ""
    prompt = f"  Duration (纳秒, 例 3840677, 或 0.00384s 表秒, 默认: {dur_default or '无'}{hint}): "
    val = input(prompt).strip()
    if val:
        parsed = _parse_ns(val)
        if parsed is None:
            print("  [警告] 无效的 Duration")
            return None
        params["--duration"] = parsed
        state[f"{ai_key}_duration"] = parsed
    elif dur_default:
        params["--duration"] = dur_default
    else:
        print("  [警告] 必须指定 Duration (纳秒)")
        return None

    # --process
    saved = _saved("--process")
    hint = f" (上次: {saved})" if saved else ""
    prompt = f"  进程名 Process (可选, 如 com.ss.android.ugc.aweme, 直接回车跳过{hint}): "
    val = input(prompt).strip()
    if val:
        params["--process"] = val
        state[f"{ai_key}_process"] = val

    # --tid
    saved = _saved("--tid")
    hint = f" (上次: {saved})" if saved else ""
    prompt = f"  线程 TID (可选, 如 14831, 直接回车跳过{hint}): "
    val = input(prompt).strip()
    if val:
        params["--tid"] = val
        state[f"{ai_key}_tid"] = val

    # --pid
    saved = _saved("--pid")
    hint = f" (上次: {saved})" if saved else ""
    prompt = f"  进程 PID (可选, 直接回车跳过{hint}): "
    val = input(prompt).strip()
    if val:
        params["--pid"] = val
        state[f"{ai_key}_pid"] = val

    # --pre-ms
    saved = _saved("--pre-ms")
    pre_default = str(section_cfg.get("pre_ms", 50))
    hint = f" (上次: {saved})" if saved and saved != pre_default else ""
    prompt = f"  卡顿点前窗口 (ms, 默认: {pre_default}{hint}): "
    val = input(prompt).strip()
    params["--pre-ms"] = val if val else pre_default
    state[f"{ai_key}_pre-ms"] = params["--pre-ms"]

    # --post-ms
    saved = _saved("--post-ms")
    post_default = str(section_cfg.get("post_ms", 200))
    hint = f" (上次: {saved})" if saved and saved != post_default else ""
    prompt = f"  卡顿点后窗口 (ms, 默认: {post_default}{hint}): "
    val = input(prompt).strip()
    params["--post-ms"] = val if val else post_default
    state[f"{ai_key}_post-ms"] = params["--post-ms"]

    # --ai (是否调用 DeepSeek AI 根因研判)
    ai_default = bool(section_cfg.get("ai", True))
    saved = state.get(f"{ai_key}_ai", "")
    hint = f" (上次: {saved})" if saved else ""
    prompt = f"  是否调用 DeepSeek AI 根因研判 (Y/n, 默认: {'Y' if ai_default else 'n'}{hint}): "
    val = input(prompt).strip().lower()
    if val in ("y", "yes", "true", "1"):
        params["--ai"] = True
        state[f"{ai_key}_ai"] = "y"
    elif val == "":
        if ai_default:
            params["--ai"] = True
            state[f"{ai_key}_ai"] = "y"
    elif val in ("n", "no", "false", "0"):
        state[f"{ai_key}_ai"] = "n"

    # 开启 AI 研判后再选模型角色
    if params.get("--ai"):
        model_default = str(section_cfg.get("ai_model", "deep"))
        saved_m = state.get(f"{ai_key}_ai-model", "")
        hint_m = f" (上次: {saved_m})" if saved_m else ""
        prompt_m = f"  AI 模型角色 fast=快速对话 / deep=深度推理 (默认: {model_default}{hint_m}): "
        val_m = input(prompt_m).strip().lower()
        if val_m in ("fast", "deep"):
            params["--ai-model"] = val_m
            state[f"{ai_key}_ai-model"] = val_m
        elif val_m == "":
            params["--ai-model"] = model_default
        else:
            print("  [警告] 无效模型角色，使用默认 %s" % model_default)
            params["--ai-model"] = model_default

    save_state(state)
    return params


def run_script(step_key, params, state, interactive=True):
    step = STEPS.get(step_key) or AI_STEPS.get(step_key)
    if not step:
        print(f"\n[错误] 未知步骤: {step_key}")
        return False
    script_path = os.path.join(SCRIPTS_DIR, step.get("subdir", ""), step["script"])

    if not os.path.isfile(script_path):
        print(f"\n[错误] 脚本不存在: {script_path}")
        return False

    cmd = [sys.executable, script_path]

    if step.get("positional"):
        cmd.append(state.get("trace_name", ""))

    if step_key in {"1", "2"}:
        trace_file = state.get("trace_file", "")
        if trace_file:
            cmd.extend(["--trace", trace_file])

    if step_key == "9":
        trace_name = state.get("trace_name", "")
        if trace_name:
            cmd.extend(["--trace-name", trace_name])

    auto_file_steps = {"4", "5", "6", "7", "8"}
    if step_key in auto_file_steps:
        trace_name = state.get("trace_name", "")
        candidates = []

        trace_dir = state.get("out_dir", "")
        if trace_dir and os.path.isdir(trace_dir):
            for ext in (".xlsx", ".csv"):
                candidates.extend(glob.glob(os.path.join(trace_dir, f"*_block_io{ext}")))

        segments_dir = os.path.join(PROJECT_ROOT, "out", trace_name + "_segments")
        if os.path.isdir(segments_dir):
            for ext in (".xlsx", ".csv"):
                candidates.extend(glob.glob(os.path.join(segments_dir, f"*_block_io{ext}")))

        candidates = sorted(set(candidates))

        seen = set()
        deduped = []
        for f in candidates:
            base = os.path.basename(f)
            if base not in seen:
                seen.add(base)
                deduped.append(f)
        candidates = deduped

        if not candidates:
            print(f"\n[警告] 未找到任何 _block_io 文件 (trace_dir + segments_dir 均无)")
            return False

        chosen = None
        if len(candidates) == 1:
            chosen = candidates[0]
            print(f"\n  自动选择唯一表格: {os.path.basename(os.path.dirname(chosen))}\\{os.path.basename(chosen)}")
        elif interactive:
            print(f"\n  可用的 block_io 表格 ({len(candidates)} 个):")
            for i, f in enumerate(candidates, 1):
                dir_label = os.path.basename(os.path.dirname(f))
                print(f"    {i}. [{dir_label}] {os.path.basename(f)}")
            while True:
                sel = input(f"  请选择 (1-{len(candidates)}, 直接回车=第一个): ").strip()
                if sel == "":
                    chosen = candidates[0]
                    break
                try:
                    idx = int(sel)
                    if 1 <= idx <= len(candidates):
                        chosen = candidates[idx - 1]
                        break
                except ValueError:
                    pass
                print(f"  输入无效, 请输入 1-{len(candidates)}")
        else:
            for f in candidates:
                base = os.path.basename(f)
                if base in (f"{trace_name}_block_io.xlsx", f"{trace_name}_block_io.csv"):
                    chosen = f
                    break
            if not chosen:
                chosen = candidates[0]

        cmd.extend(["--file", chosen])

    for arg_name in step["args"]:
        if arg_name in params:
            val = params[arg_name]
            if arg_name == "--pairs":
                pairs = val
                if isinstance(pairs, str):
                    try:
                        pairs = json.loads(pairs)
                    except (json.JSONDecodeError, ValueError):
                        pairs = val.split()
                cmd.append("--pairs")
                cmd.extend(pairs)
            else:
                cmd.extend([arg_name, str(val)])

    # 布尔型开关（如 --ai），仅当为真时追加标志本身
    for flag in step.get("flags", []):
        if params.get(flag):
            cmd.append(flag)

    print(f"\n{'='*60}")
    print(f"[步骤 {step_key}] {step['name']}")
    print(f"  命令: {' '.join(cmd)}")
    print(f"{'='*60}")

    try:
        timeout = int(params.get("--timeout", step.get("timeout", 600)))
    except (ValueError, TypeError):
        timeout = step.get("timeout", 600)

    try:
        result = subprocess.run(cmd, cwd=PROJECT_ROOT, timeout=timeout)
        if result.returncode == 0:
            print(f"\n[步骤 {step_key}] 完成")
            state[f"{step_key}_done"] = True
            save_state(state)
            return True
        else:
            print(f"\n[步骤 {step_key}] 失败 (退出码: {result.returncode})")
            return False
    except subprocess.TimeoutExpired:
        print(f"\n[步骤 {step_key}] 超时（{timeout}秒）")
        return False
    except Exception as e:
        print(f"\n[步骤 {step_key}] 异常: {e}")
        return False


def show_menu(state):
    print("\n" + "=" * 60)
    print("  AI-trace — Perfetto Block IO 分析工具")
    print("=" * 60)

    traces = detect_traces()
    print(f"\n  Trace 文件: {len(traces)} 个")
    if state.get("trace_name"):
        print(f"  当前选择: {state['trace_name']}")
    if state.get("out_dir"):
        print(f"  输出目录: {state['out_dir']}")

    print(f"\n  分析流程:")
    for key in STEP_ORDER:
        step = STEPS[key]
        done = state.get(f"{key}_done", False)
        mark = "[ok]" if done else "[  ]"
        required = "[必选]" if step["required"] else "[可选]"
        print(f"  {mark} {key}. {step['name']} {required}")

    print(f"\n  AI分析选项:")
    for key in AI_STEP_ORDER:
        step = AI_STEPS[key]
        print(f"  [ ] {key}. {step['name']}")

    print(f"\n  全局操作:")
    print(f"  [A] 一键全流程 (步骤1→9)")
    print(f"  [C] 切换 / 选择 trace 文件")
    print(f"  [R] 重置参数记忆")
    print(f"  [Q] 退出")
    print(f"\n  配置文件: config/default.json   (分析流程参数, 修改后下次运行生效)")
    print(f"  AI 配置:   config/ai_default.json (AI 分析选项参数, 修改后下次运行生效)")


def run_full_pipeline(state):
    """一键执行全流程"""
    traces = detect_traces()
    if not traces:
        print("\n[错误] 未找到 trace 文件")
        return

    if not state.get("trace_name"):
        sel, name = select_trace(traces, state)
        if sel is None:
            return

    print("\n" + "=" * 60)
    print("  一键全流程开始")
    print("=" * 60)

    for key in STEP_ORDER:
        step = STEPS[key]
        params = state.get(f"{key}_params", {})

        if params is None:
            print(f"\n[跳过] 步骤 {key}: {step['name']} (参数未配置)")
            continue

        run_script(key, params, state, interactive=False)

    print("\n" + "=" * 60)
    print("  全流程结束")
    print("=" * 60)


def interactive():
    state = load_state()

    while True:
        show_menu(state)
        ai_keys = "/".join(AI_STEP_ORDER)
        choice = input(f"\n请输入选项 (1-{len(STEP_ORDER)} / {ai_keys} / A / C / R / Q): ").strip().upper()

        if choice == "Q":
            print("再见!")
            break

        elif choice == "A":
            run_full_pipeline(state)

        elif choice == "C":
            traces = detect_traces()
            if traces:
                select_trace(traces, state)
            else:
                print("\n未找到 trace 文件，请将 .perfetto-trace 或 .perfetto 放入 trace/ 目录")

        elif choice == "R":
            if os.path.isfile(STATE_FILE):
                os.remove(STATE_FILE)
            state = {}
            print("\n参数记忆已重置")

        elif choice in AI_STEPS:
            params = prompt_ai_params(choice, state)
            if params is None:
                continue
            run_script(choice, params, state, interactive=True)

        elif choice in STEPS:
            traces = detect_traces()
            if not traces:
                print("\n[错误] 未找到 trace 文件")
                continue

            if not state.get("trace_name"):
                sel, name = select_trace(traces, state)
                if sel is None:
                    continue

            params = prompt_params(choice, state)
            if params is None:
                continue

            state[f"{choice}_params"] = params
            save_state(state)

            run_script(choice, params, state, interactive=True)

        else:
            print(f"\n无效选项: {choice}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="AI-trace Block IO 分析工具集")
    parser.add_argument("--step", type=str, choices=list(STEPS.keys()),
                        help="直接执行指定步骤 (1-9)")
    parser.add_argument("--all", action="store_true",
                        help="一键全流程")
    parser.add_argument("--reset", action="store_true",
                        help="重置参数记忆")
    args = parser.parse_args()

    if args.reset:
        if os.path.isfile(STATE_FILE):
            os.remove(STATE_FILE)
        print("参数记忆已重置")
        return

    if args.all:
        state = load_state()
        run_full_pipeline(state)
        return

    if args.step:
        state = load_state()
        traces = detect_traces()
        if not traces:
            print("[错误] 未找到 trace 文件")
            sys.exit(1)
        if not state.get("trace_name"):
            print("[提示] 请先运行交互模式选择 trace 文件，或使用 --all")
            sys.exit(1)
        params = state.get(f"{args.step}_params", {})
        if not params:
            print(f"[提示] 步骤 {args.step} 无已保存参数，请先在交互模式中配置")
            sys.exit(1)
        success = run_script(args.step, params, state, interactive=False)
        sys.exit(0 if success else 1)
        return

    interactive()


if __name__ == "__main__":
    main()