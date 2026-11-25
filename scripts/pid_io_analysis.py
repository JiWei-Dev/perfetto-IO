import os
import sys
import glob
import re
import argparse
from collections import defaultdict, Counter

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import matplotlib.pyplot as plt
import numpy as np
from config_loader import load_config
from io_reader import load_io_events, find_block_io_file, match_io_pairs as _match_io_pairs

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def get_project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def classify_rwbs(rwbs_str):
    if not rwbs_str:
        return "other"
    rwbs = rwbs_str.strip().upper()
    raw = rwbs.split(" ")[0]
    if not raw:
        return "other"
    if "W" in raw and "R" not in raw:
        return "write"
    if "R" in raw:
        return "read"
    return "other"


def _as_list(val):
    """把单个值/列表统一成字符串列表（None 返回 None，表示不限定）。"""
    if val is None:
        return None
    if isinstance(val, (list, tuple, set)):
        return [str(x) for x in val]
    return [str(val)]


def filter_by_pid(events, pid):
    pids = _as_list(pid)
    if pids is None:
        return events
    return [e for e in events if str(e.get("进程 PID", "")).strip() in pids]


def filter_by_tid(events, tid):
    tids = _as_list(tid)
    if tids is None:
        return events
    return [e for e in events if str(e.get("线程 TID", "")).strip() in tids]


NS_PER_SEC = 1_000_000_000


def filter_by_time(events, start_s=None, end_s=None):
    """按时间范围 (秒) 过滤事件，时间字段为「时间 (纳秒)」。"""
    if start_s is None and end_s is None:
        return events
    result = []
    for ev in events:
        ts_raw = ev.get("时间 (纳秒)", 0)
        try:
            ts_sec = int(ts_raw) / NS_PER_SEC if ts_raw else 0
        except (ValueError, TypeError):
            continue
        if start_s is not None and ts_sec < start_s:
            continue
        if end_s is not None and ts_sec > end_s:
            continue
        result.append(ev)
    return result


def _split_targets(raw):
    """把多目标字符串拆成去重列表，支持逗号(中/英文)/空格/分号(中/英文)分隔。"""
    if raw is None:
        return []
    if isinstance(raw, (list, tuple)):
        items = [str(x).strip() for x in raw if str(x).strip()]
    else:
        items = [x for x in re.split(r"[,，;；\s]+", str(raw).strip()) if x]
    seen = set()
    result = []
    for it in items:
        if it not in seen:
            seen.add(it)
            result.append(it)
    return result


def _filter_insert_events(events, target_pid, target_tid):
    """仅保留符合目标 PID/TID 的 block_rq_insert 事件, 其余事件原样保留。

    注意: 只过滤 insert 事件 (insert 才携带目标进程身份), issue/complete
    事件无条件保留, 与旧配对逻辑的语义一致。target_pid/target_tid 可为列表。
    """
    if target_pid is None and target_tid is None:
        return events
    pids = _as_list(target_pid)
    tids = _as_list(target_tid)
    filtered = []
    for ev in events:
        if str(ev.get("Trace 字段名称", "")).strip() == "block_rq_insert":
            matched = True
            if pids is not None:
                if str(ev.get("进程 PID", "")).strip() not in pids:
                    matched = False
            if tids is not None:
                if str(ev.get("线程 TID", "")).strip() not in tids:
                    matched = False
            if not matched:
                continue
        filtered.append(ev)
    return filtered


def match_io_pairs(events, target_pid=None, target_tid=None):
    """三段式配对封装: 按目标 PID/TID 过滤 insert 后委托共享配对核心。"""
    filtered = _filter_insert_events(events, target_pid, target_tid)
    triples = _match_io_pairs(filtered)
    schedule_latencies = []
    device_latencies = []

    for t in triples:
        insert_ev = t["insert_ev"]
        issue_ev = t["issue_ev"]
        complete_ev = t["complete_ev"]

        rwbs_sched = str(issue_ev.get("读写类型", "") or insert_ev.get("读写类型", ""))
        nr_sector_sched = issue_ev.get("请求扇区数量", 0) or insert_ev.get("请求扇区数量", 0)
        schedule_latencies.append({
            "dev": t["dev"], "sector": t["sector"],
            "insert_ts": t["insert_ts"], "issue_ts": t["issue_ts"],
            "latency_ns": t["schedule_ns"],
            "latency_ms": t["schedule_ns"] / 1_000_000,
            "rwbs": rwbs_sched, "rwbs_type": classify_rwbs(rwbs_sched),
            "nr_sector": nr_sector_sched,
            "insert_process": insert_ev.get("IO 进程名称", ""),
            "insert_pid": insert_ev.get("进程 PID", ""),
            "insert_tid": insert_ev.get("线程 TID", ""),
            "insert_thread": insert_ev.get("线程名称", ""),
            "insert_cpu": insert_ev.get("运行 CPU", ""),
            "issue_process": issue_ev.get("IO 进程名称", ""),
            "issue_pid": issue_ev.get("进程 PID", ""),
            "issue_tid": issue_ev.get("线程 TID", ""),
            "issue_thread": issue_ev.get("线程名称", ""),
            "issue_cpu": issue_ev.get("运行 CPU", ""),
        })

        rwbs_dev = str(complete_ev.get("读写类型", "") or issue_ev.get("读写类型", ""))
        nr_sector_dev = complete_ev.get("请求扇区数量", 0) or issue_ev.get("请求扇区数量", 0)
        device_latencies.append({
            "dev": t["dev"], "sector": t["sector"],
            "issue_ts": t["issue_ts"], "complete_ts": t["complete_ts"],
            "latency_ns": t["device_ns"],
            "latency_ms": t["device_ns"] / 1_000_000,
            "rwbs": rwbs_dev, "rwbs_type": classify_rwbs(rwbs_dev),
            "nr_sector": nr_sector_dev,
            "issue_process": issue_ev.get("IO 进程名称", ""),
            "issue_pid": issue_ev.get("进程 PID", ""),
            "issue_tid": issue_ev.get("线程 TID", ""),
            "issue_thread": issue_ev.get("线程名称", ""),
            "issue_cpu": issue_ev.get("运行 CPU", ""),
            "complete_process": complete_ev.get("IO 进程名称", ""),
            "complete_pid": complete_ev.get("进程 PID", ""),
            "complete_tid": complete_ev.get("线程 TID", ""),
            "complete_thread": complete_ev.get("线程名称", ""),
            "complete_cpu": complete_ev.get("运行 CPU", ""),
        })

    return schedule_latencies, device_latencies


def extract_io_sizes(events):
    sizes = []
    for ev in events:
        nr_sector_raw = ev.get("请求扇区数量", 0)
        try:
            nr_sector = int(nr_sector_raw) if nr_sector_raw else 0
        except (ValueError, TypeError):
            continue
        if nr_sector <= 0:
            continue
        rwbs_str = str(ev.get("读写类型", ""))
        sizes.append({
            "nr_sector": nr_sector,
            "size_kb": nr_sector * 512 / 1024,
            "rwbs": rwbs_str,
            "rwbs_type": classify_rwbs(rwbs_str),
        })
    return sizes


def build_process_stats(events):
    ProcessStat = lambda: {"read": 0, "write": 0, "other": 0, "total_kb": 0}
    stats = defaultdict(ProcessStat)

    for ev in events:
        trace_field = str(ev.get("Trace 字段名称", "")).strip()
        if trace_field != "block_rq_insert":
            continue
        proc = str(ev.get("IO 进程名称", "")).strip()
        pid = str(ev.get("进程 PID", "")).strip()
        tid = str(ev.get("线程 TID", "")).strip()
        thread = str(ev.get("线程名称", "")).strip()
        if not proc and not tid:
            continue
        key = (proc, pid, tid, thread)
        cat = classify_rwbs(str(ev.get("读写类型", "")))
        stats[key][cat] += 1
        try:
            nr_sector = int(ev.get("请求扇区数量", 0) or 0)
        except (ValueError, TypeError):
            nr_sector = 0
        stats[key]["total_kb"] += nr_sector * 512 / 1024
    return stats


STYLE = {
    "header_fill": PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
    "header_font": Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11),
    "cell_font": Font(name="Consolas", size=10),
    "thin_border": Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin")),
    "total_fill": PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),
    "total_font": Font(name="Microsoft YaHei", bold=True, size=11),
    "cat_fill": PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),
    "cat_font": Font(name="Microsoft YaHei", bold=True, size=11),
}


def ts_ns_to_sec(ts_ns):
    if ts_ns is None or ts_ns == "":
        return ""
    try:
        return f"{int(ts_ns) / 1_000_000_000:.9f}"
    except (ValueError, TypeError):
        return str(ts_ns)


def write_schedule_sheet(wb, data):
    columns = [
        ("设备号", "dev"), ("请求扇区地址", "sector"), ("请求扇区数量", "nr_sector"), ("读写类型", "rwbs"),
        ("insert 时间 (秒)", "insert_ts"), ("insert 时间 (纳秒)", "insert_ts"),
        ("insert 进程", "insert_process"), ("insert PID", "insert_pid"),
        ("insert 线程 TID", "insert_tid"), ("insert 线程名称", "insert_thread"), ("insert CPU", "insert_cpu"),
        ("issue 时间 (秒)", "issue_ts"), ("issue 时间 (纳秒)", "issue_ts"),
        ("issue 进程", "issue_process"), ("issue PID", "issue_pid"),
        ("issue 线程 TID", "issue_tid"), ("issue 线程名称", "issue_thread"), ("issue CPU", "issue_cpu"),
        ("调度耗时 (ms)", "latency_ms"), ("调度耗时 (ns)", "latency_ns"),
    ]
    _write_typed_sheet(wb, "调度耗时", data, columns, ts_keys={"insert_ts", "issue_ts"})


def write_device_sheet(wb, data):
    columns = [
        ("设备号", "dev"), ("请求扇区地址", "sector"), ("请求扇区数量", "nr_sector"), ("读写类型", "rwbs"),
        ("issue 时间 (秒)", "issue_ts"), ("issue 时间 (纳秒)", "issue_ts"),
        ("issue 进程", "issue_process"), ("issue PID", "issue_pid"),
        ("issue 线程 TID", "issue_tid"), ("issue 线程名称", "issue_thread"), ("issue CPU", "issue_cpu"),
        ("complete 时间 (秒)", "complete_ts"), ("complete 时间 (纳秒)", "complete_ts"),
        ("complete 进程", "complete_process"), ("complete PID", "complete_pid"),
        ("complete 线程 TID", "complete_tid"), ("complete 线程名称", "complete_thread"), ("complete CPU", "complete_cpu"),
        ("器件耗时 (ms)", "latency_ms"), ("器件耗时 (ns)", "latency_ns"),
    ]
    _write_typed_sheet(wb, "器件耗时", data, columns, ts_keys={"issue_ts", "complete_ts"})


def _write_typed_sheet(wb, sheet_title, data, columns, ts_keys=None):
    if ts_keys is None:
        ts_keys = set()

    if sheet_title in wb.sheetnames:
        del wb[sheet_title]
    ws = wb.create_sheet(title=sheet_title)

    col_names = [c[0] for c in columns]
    col_keys = [c[1] for c in columns]

    for col_idx, name in enumerate(col_names, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = STYLE["header_fill"]
        cell.font = STYLE["header_font"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = STYLE["thin_border"]

    center_keys = {"insert_cpu", "issue_cpu", "complete_cpu", "insert_tid", "issue_tid",
                   "complete_tid", "insert_pid", "issue_pid", "complete_pid"}

    for row_idx, item in enumerate(data, 2):
        for col_idx, (col_name, key) in enumerate(zip(col_names, col_keys), 1):
            raw_val = item.get(key, "")
            if key in ts_keys:
                if "纳秒" in col_name:
                    val = raw_val
                else:
                    val = ts_ns_to_sec(raw_val)
            elif key == "latency_ms":
                val = round(raw_val, 6) if isinstance(raw_val, float) else raw_val
            else:
                val = raw_val
            cell = ws.cell(row=row_idx, column=col_idx, value=val if val != "" and val is not None else "")
            cell.font = STYLE["cell_font"]
            cell.border = STYLE["thin_border"]
            if key in center_keys:
                cell.alignment = Alignment(horizontal="center")
            elif key in ("latency_ms", "latency_ns", "nr_sector"):
                cell.alignment = Alignment(horizontal="right")

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def write_size_sheet(wb, sizes):
    grouped = defaultdict(Counter)
    for item in sizes:
        grouped[item["rwbs_type"]][item["nr_sector"]] += 1

    labels_cn = {"read": "读请求", "write": "写请求", "other": "其他请求"}
    col_headers = ["请求类型", "请求扇区数量", "请求大小 (KB)", "IO 数量"]

    sheet_name = "IO 大小汇总"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    for col_idx, name in enumerate(col_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = STYLE["header_fill"]
        cell.font = STYLE["header_font"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = STYLE["thin_border"]

    row_idx = 2
    for cat in ["read", "write", "other"]:
        data = grouped.get(cat, {})
        if not data:
            continue
        cat_cell = ws.cell(row=row_idx, column=1, value=labels_cn.get(cat, cat))
        cat_cell.font = STYLE["cat_font"]
        cat_cell.fill = STYLE["cat_fill"]
        cat_cell.border = STYLE["thin_border"]
        total_cell = ws.cell(row=row_idx, column=4, value=sum(data.values()))
        total_cell.font = STYLE["cat_font"]
        total_cell.fill = STYLE["cat_fill"]
        total_cell.border = STYLE["thin_border"]
        for c in [2, 3]:
            cell = ws.cell(row=row_idx, column=c)
            cell.fill = STYLE["cat_fill"]
            cell.border = STYLE["thin_border"]
        row_idx += 1
        for nr_sector, count in sorted(data.items()):
            for c in range(1, 5):
                ws.cell(row=row_idx, column=c).border = STYLE["thin_border"]
                ws.cell(row=row_idx, column=c).font = STYLE["cell_font"]
            ws.cell(row=row_idx, column=2, value=nr_sector).alignment = Alignment(horizontal="right")
            ws.cell(row=row_idx, column=3, value=nr_sector * 512 // 1024).alignment = Alignment(horizontal="right")
            ws.cell(row=row_idx, column=4, value=count).alignment = Alignment(horizontal="center")
            row_idx += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def write_process_stats_sheet(wb, stats):
    col_headers = [
        "IO 发起进程", "进程 PID", "线程 TID", "IO 发起线程",
        "读请求数量", "写请求数量", "其他请求数量", "总请求数量", "总数据量 (KB)",
    ]
    sheet_name = "进程 IO 统计"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    for col_idx, name in enumerate(col_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = STYLE["header_fill"]
        cell.font = STYLE["header_font"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = STYLE["thin_border"]

    sorted_items = sorted(stats.items(), key=lambda x: x[1]["total_kb"], reverse=True)

    grand_total = {"read": 0, "write": 0, "other": 0, "total_kb": 0.0}
    row_idx = 2

    for (proc, pid, tid, thread), data in sorted_items:
        total_count = data["read"] + data["write"] + data["other"]
        if total_count == 0:
            continue
        values = [
            proc if proc else f"TID:{tid}",
            pid if pid else "",
            tid if tid else "",
            thread if thread else f"TID:{tid}",
            data["read"], data["write"], data["other"],
            total_count, round(data["total_kb"], 2),
        ]
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val if val != 0 or col_idx > 5 else val)
            cell.font = STYLE["cell_font"]
            cell.border = STYLE["thin_border"]
            cell.alignment = Alignment(horizontal="right" if col_idx >= 5 else "left")
        grand_total["read"] += data["read"]
        grand_total["write"] += data["write"]
        grand_total["other"] += data["other"]
        grand_total["total_kb"] += data["total_kb"]
        row_idx += 1

    grand_total_count = grand_total["read"] + grand_total["write"] + grand_total["other"]
    total_values = [
        "合计", "", "", "",
        grand_total["read"], grand_total["write"], grand_total["other"],
        grand_total_count, round(grand_total["total_kb"], 2),
    ]
    for col_idx, val in enumerate(total_values, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.font = STYLE["total_font"]
        cell.fill = STYLE["total_fill"]
        cell.border = STYLE["thin_border"]
        cell.alignment = Alignment(horizontal="right" if col_idx >= 5 else "left")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 18
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def determine_bins(data_ms, target_bins=30, cap_ms=None):
    if not data_ms:
        return np.array([0, 1]), None
    min_val = min(data_ms)
    max_val = max(data_ms)
    if max_val <= min_val:
        return np.array([min_val, min_val + 0.1]), None
    if cap_ms is not None and max_val > cap_ms:
        plot_max = cap_ms
        capped = True
    else:
        plot_max = max_val
        capped = False
    bins = np.logspace(np.log10(max(min_val, 0.001)), np.log10(plot_max * 1.1), target_bins)
    bins = np.unique(np.round(bins, 3))
    return bins, max_val if capped else None


def plot_latency_distribution(latencies, title, output_path, pcap=None):
    color_palette = {"read": "#4CAF50", "write": "#F44336", "other": "#2196F3"}
    categories = ["read", "write", "other"]
    labels_cn = {"read": "读请求", "write": "写请求", "other": "其他请求"}

    grouped = defaultdict(list)
    for item in latencies:
        grouped[item["rwbs_type"]].append(item["latency_ms"])

    # 只保留有数据的分类, 避免画出空子图; 无数据则整图跳过
    active_cats = [cat for cat in categories if grouped.get(cat)]
    if not active_cats:
        print(f"  [跳过] {title}: 无任何读/写/其他请求数据")
        return

    fig, axes = plt.subplots(len(active_cats), 1, figsize=(14, 4.5 * len(active_cats)), squeeze=False)
    fig.suptitle(title, fontsize=16, fontweight="bold", y=0.98)

    for row, cat in enumerate(active_cats):
        ax = axes[row, 0]
        data = grouped[cat]

        cap_ms = None
        if pcap is not None and pcap < 100:
            cap_ms = np.percentile(data, pcap)

        bins, actual_max = determine_bins(data, cap_ms=cap_ms)
        counts, bin_edges = np.histogram(data, bins=bins)
        bar_widths = np.diff(bin_edges) * 0.9
        bar_centers = bin_edges[:-1] + np.diff(bin_edges) / 2

        color = color_palette.get(cat, "#888888")
        bars = ax.bar(bar_centers, counts, width=bar_widths,
                      color=color, alpha=0.85, edgecolor="white", linewidth=0.5)
        max_count = max(counts) if len(counts) > 0 else 0
        for bar, count in zip(bars, counts):
            if count > 0 and count >= max_count * 0.03:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_count * 0.02,
                        str(count), ha="center", va="bottom", fontsize=7, color="#333333")

        if actual_max is not None:
            overflow = sum(1 for d in data if d > cap_ms)
            ax.text(0.97, 0.06, f"+{overflow} IO 超过 P{pcap} 截断",
                    transform=ax.transAxes, fontsize=7, va="bottom", ha="right",
                    color="#D32F2F", bbox=dict(boxstyle="round,pad=0.3", facecolor="#FFEBEE", alpha=0.95))

        ax.set_xscale("log")
        ax.set_xlabel("时延 (ms)", fontsize=11)
        ax.set_ylabel("IO 数量", fontsize=11)
        stats_text = (f"总数: {len(data)}\nP50: {np.percentile(data, 50):.2f} ms\n"
                      f"P99: {np.percentile(data, 99):.2f} ms\n"
                      f"均值: {np.mean(data):.2f} ms\n最大: {np.max(data):.2f} ms")
        ax.text(0.97, 0.97, stats_text, transform=ax.transAxes,
                fontsize=8, va="top", ha="right",
                bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.9))
        ax.set_title(labels_cn[cat], fontsize=14)
        ax.grid(axis="y", alpha=0.3, linestyle="--")

    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  图表已保存: {output_path}")


def plot_size_distribution(sizes, value_key, xlabel, title, output_path):
    color_palette = {"read": "#4CAF50", "write": "#F44336", "other": "#2196F3"}
    categories = ["read", "write", "other"]
    labels_cn = {"read": "读请求", "write": "写请求", "other": "其他请求"}

    grouped = defaultdict(list)
    for item in sizes:
        grouped[item["rwbs_type"]].append(item[value_key])

    # 只保留有数据的分类, 避免画出空子图; 无数据则整图跳过
    active_cats = [cat for cat in categories if grouped.get(cat)]
    if not active_cats:
        print(f"  [跳过] {title}: 无任何读/写/其他请求数据")
        return

    fig, axes = plt.subplots(len(active_cats), 1, figsize=(14, 4.5 * len(active_cats)), squeeze=False)
    fig.suptitle(title, fontsize=16, fontweight="bold", y=0.98)

    for row, cat in enumerate(active_cats):
        ax = axes[row, 0]
        data = grouped[cat]

        counter = Counter(data)
        items_list = sorted(counter.items())
        labels = [str(k) for k, v in items_list]
        x_positions = list(range(len(items_list)))
        heights = [v for k, v in items_list]

        color = color_palette.get(cat, "#888888")
        bar_width = max(0.6, 0.9 - 0.01 * len(items_list))
        bars = ax.bar(x_positions, heights, width=bar_width,
                      color=color, alpha=0.85, edgecolor="white", linewidth=0.5)

        max_count = max(heights) if heights else 0
        for bar, count in zip(bars, heights):
            if count > 0 and count >= max_count * 0.02:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_count * 0.02,
                        str(count), ha="center", va="bottom", fontsize=7, color="#333333")

        step = max(1, len(x_positions) // 12)
        tick_positions = x_positions[::step]
        tick_labels = labels[::step]
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=8)

        ax.set_xlabel(xlabel, fontsize=11)
        ax.set_ylabel("IO 数量", fontsize=11)

        counter_top = counter.most_common(3)
        unit_str = xlabel.split("(")[0].strip() if "(" in xlabel else ""
        top_text = "\n".join([f"{v}{' ' + unit_str}: {c} 次" for v, c in counter_top])
        stats_text = (f"总数: {len(data)}\n不同值: {len(counter)} 种\n"
                      f"均值: {np.mean(data):.2f}\n最小: {min(data)}\n最大: {max(data)}\n\nTop:\n{top_text}")
        ax.text(0.97, 0.97, stats_text, transform=ax.transAxes, fontsize=7.5,
                va="top", ha="right", bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.9))
        ax.set_title(labels_cn[cat], fontsize=14)
        ax.grid(axis="y", alpha=0.3, linestyle="--")
        ax.set_xlim(-0.5, len(items_list) - 0.5)

    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  图表已保存: {output_path}")


def parse_args():
    parser = argparse.ArgumentParser(description="指定 PID 或 TID，分析该进程/线程的 Block IO 详情")
    parser.add_argument("--pid", type=str, default=None,
                        help="要分析的进程 PID（与 --tid 二选一，多个用逗号分隔，可配置文件提供）")
    parser.add_argument("--tid", type=str, default=None,
                        help="要分析的线程 TID（与 --pid 二选一，多个用逗号分隔，可配置文件提供）")
    parser.add_argument("--sheet", type=str, default=None,
                        help="读取 XLSX 中的哪个 Sheet（默认: 全部 IO，可配置文件覆盖）")
    parser.add_argument("--pcap", type=float, default=None,
                        help="时延图表横坐标截断分位数（如 99.9），不指定则不截断")
    parser.add_argument("--start", type=float, default=None,
                        help="起始时间 (秒)，只分析该时间之后的 IO（可配置文件提供）")
    parser.add_argument("--end", type=float, default=None,
                        help="结束时间 (秒)，只分析该时间之前的 IO（可配置文件提供）")
    parser.add_argument("--file", type=str, default=None,
                        help="手动指定输入文件路径 (.csv/.xlsx)，跳过 out/ 目录自动扫描")
    return parser.parse_args()


def main():
    args = parse_args()

    cfg = load_config()
    pid_cfg = cfg.get("pid_io_analysis", {})

    if args.pid is None and args.tid is None:
        args.pid = pid_cfg.get("pid")
        args.tid = pid_cfg.get("tid")

    if args.pid is None and args.tid is None:
        print("[错误] 必须指定 --pid 或 --tid 其中之一（CLI 参数或配置文件）")
        sys.exit(1)
    if args.pid is not None and args.tid is not None:
        print("[错误] --pid 和 --tid 不能同时指定，请只选其一")
        sys.exit(1)

    if args.sheet is None:
        args.sheet = pid_cfg.get("sheet", "全部 IO")
    if args.pcap is None:
        args.pcap = pid_cfg.get("pcap")
    if args.start is None:
        args.start = pid_cfg.get("start")
    if args.end is None:
        args.end = pid_cfg.get("end")

    if args.pid is not None:
        target_type = "PID"
        target_values = _split_targets(args.pid)
    else:
        target_type = "TID"
        target_values = _split_targets(args.tid)

    if not target_values:
        print(f"[错误] 无法解析 {target_type} 目标列表，请检查 --pid/--tid 输入")
        sys.exit(1)

    if args.file:
        file_abs = os.path.abspath(args.file)
        if not os.path.isfile(file_abs):
            print(f"[错误] 文件不存在: {file_abs}")
            sys.exit(1)
        trace_dirs = [os.path.dirname(file_abs)]
        manual_file = file_abs
        manual_name = os.path.splitext(os.path.basename(file_abs))[0]
        if manual_name.endswith("_block_io"):
            manual_name = manual_name[:-len("_block_io")]
    else:
        project_root = get_project_root()
        out_dir = os.path.join(project_root, "out")
        trace_dirs = glob.glob(os.path.join(out_dir, "*"))
        trace_dirs = [d for d in trace_dirs if os.path.isdir(d)]
        manual_file = None
        manual_name = None

    if not trace_dirs:
        print(f"[提示] out 目录下未找到 trace 输出子目录")
        sys.exit(0)

    for trace_dir in trace_dirs:
        if manual_file:
            file_path = manual_file
            trace_name = manual_name
        else:
            trace_name = os.path.basename(trace_dir)
            file_path, file_fmt = find_block_io_file(trace_dir, trace_name)
        if not file_path:
            print(f"[提示] 跳过 {trace_name}：未找到 _block_io.xlsx 或 _block_io.csv")
            continue

        print(f"[读取] {file_path}")
        print(f"  Sheet: {args.sheet}")
        print(f"  目标 {target_type}: {', '.join(target_values)}")

        headers, events = load_io_events(file_path, args.sheet)
        print(f"  读取到 {len(events)} 条 IO 记录")

        if args.start is not None or args.end is not None:
            before = len(events)
            events = filter_by_time(events, args.start, args.end)
            print(f"  按时间范围筛选: {before} -> {len(events)} 条")

        target_suffix = "_".join(target_values)

        if args.pid is not None:
            target_events = filter_by_pid(events, target_values)
            match_desc = f"PID={','.join(target_values)}"
            cur_pid, cur_tid = target_values, None
        else:
            target_events = filter_by_tid(events, target_values)
            match_desc = f"TID={','.join(target_values)}"
            cur_pid, cur_tid = None, target_values

        if not target_events:
            print(f"  [警告] 未找到 {target_type}={','.join(target_values)} 的 IO 记录，跳过")
            continue
        print(f"  {target_type}={','.join(target_values)} 的记录: {len(target_events)} 条")

        target_dir = os.path.join(trace_dir, f"{target_type}_{target_suffix}")
        os.makedirs(target_dir, exist_ok=True)

        out_xlsx = os.path.join(target_dir, f"{trace_name}_{target_suffix}.xlsx")
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        print(f"  [匹配] 使用全部 {len(events)} 条事件配对，仅 insert 限定 {match_desc}")
        schedule_latencies, device_latencies = match_io_pairs(
            events,
            target_pid=cur_pid,
            target_tid=cur_tid,
        )
        print(f"  调度耗时配对: {len(schedule_latencies)}")
        print(f"  器件耗时配对: {len(device_latencies)}")

        write_schedule_sheet(wb, schedule_latencies)
        write_device_sheet(wb, device_latencies)

        sizes = extract_io_sizes(target_events)
        write_size_sheet(wb, sizes)
        print(f"  有效 IO Size 记录: {len(sizes)}")

        stats = build_process_stats(target_events)
        write_process_stats_sheet(wb, stats)
        print(f"  线程组合数: {len(stats)}")

        wb.save(out_xlsx)
        print(f"\n[完成] 表格: {out_xlsx}")

        label = f"{target_type}={','.join(target_values)}"

        charts_dir = target_dir

        if schedule_latencies:
            chart_path = os.path.join(charts_dir, f"{trace_name}_{target_suffix}_schedule_latency.png")
            print(f"[绘图] 调度耗时分布")
            plot_latency_distribution(schedule_latencies,
                                      f"{trace_name} {label} - 调度耗时分布 (insert→issue)",
                                      chart_path, pcap=args.pcap)

        if device_latencies:
            chart_path = os.path.join(charts_dir, f"{trace_name}_{target_suffix}_device_latency.png")
            print(f"[绘图] 器件耗时分布")
            plot_latency_distribution(device_latencies,
                                      f"{trace_name} {label} - 器件耗时分布 (issue→complete)",
                                      chart_path, pcap=args.pcap)

        if sizes:
            chart_path = os.path.join(charts_dir, f"{trace_name}_{target_suffix}_io_nr_sector.png")
            print(f"[绘图] 请求扇区数量分布")
            plot_size_distribution(sizes, value_key="nr_sector", xlabel="扇区数量",
                                   title=f"{trace_name} {label} - 请求扇区数量分布",
                                   output_path=chart_path)

            chart_path = os.path.join(charts_dir, f"{trace_name}_{target_suffix}_io_size_kb.png")
            print(f"[绘图] 请求大小分布 (KB)")
            plot_size_distribution(sizes, value_key="size_kb", xlabel="请求大小 (KB)",
                                   title=f"{trace_name} {label} - 请求大小分布 (KB)",
                                   output_path=chart_path)

    print(f"\n全部完成。")


if __name__ == "__main__":
    main()