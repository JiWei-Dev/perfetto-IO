import os
import sys
import glob
import argparse
import bisect
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from config_loader import load_config
from io_reader import load_io_events, find_block_io_file, match_io_pairs

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

NS_PER_SEC = 1_000_000_000
NS_PER_MS = 1_000_000
SEQ_GAP_THRESHOLD = 8
SLOW_IO_TOP_N = 20
MAX_DEV_SUBPLOTS = 6

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
CELL_FONT = Font(name="Consolas", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def get_project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def classify_rwbs(rwbs_str):
    if not rwbs_str:
        return "other"
    raw = rwbs_str.strip().upper().split(" ")[0]
    if not raw:
        return "other"
    if "W" in raw and "R" not in raw:
        return "write"
    if "R" in raw:
        return "read"
    return "other"

def classify_sync_type(rwbs_str):
    if not rwbs_str:
        return "unknown"
    raw = rwbs_str.strip().upper().split(" ")[0]
    if not raw:
        return "unknown"
    is_write = "W" in raw and "R" not in raw
    is_read = "R" in raw
    is_sync = "S" in raw
    if is_write and is_sync:
        return "sync_write"
    if is_write:
        return "async_write"
    if is_read and is_sync:
        return "sync_read"
    if is_read:
        return "async_read"
    return "unknown"


def filter_by_time(events, start_s=None, end_s=None):
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


def match_io_pairs_all(events):
    """三段式配对封装: 委托共享配对核心 match_io_pairs, 再补齐派生字段。"""
    triples = match_io_pairs(events)
    results = []
    for t in triples:
        insert_ev = t["insert_ev"]
        complete_ev = t["complete_ev"]
        rwbs = str(complete_ev.get("读写类型", "") or insert_ev.get("读写类型", ""))
        nr_sector_raw = complete_ev.get("请求扇区数量", 0) or insert_ev.get("请求扇区数量", 0)
        try:
            nr_sector = int(nr_sector_raw) if nr_sector_raw else 0
        except (ValueError, TypeError):
            nr_sector = 0
        results.append({
            "dev": t["dev"], "sector": t["sector"],
            "insert_ts": t["insert_ts"], "issue_ts": t["issue_ts"], "complete_ts": t["complete_ts"],
            "schedule_ns": t["schedule_ns"], "device_ns": t["device_ns"],
            "end_to_end_ns": t["end_to_end_ns"],
            "rwbs": rwbs, "rwbs_type": classify_rwbs(rwbs), "sync_type": classify_sync_type(rwbs),
            "nr_sector": nr_sector,
            "insert_pid": insert_ev.get("进程 PID", ""),
            "insert_tid": insert_ev.get("线程 TID", ""),
            "insert_thread": insert_ev.get("线程名称", ""),
        })
    return results


def calc_queue_depth(events):
    # 按 request 生命周期 (insert -> issue -> complete) 统计在途 IO 数量，
    # 避免简单全局 +1/-1 在事件不匹配时导致的队列深度漂移。
    # 分设备统计：不同存储设备 (UFS/dm/loop/eMMC) 各自独立累加，避免互相污染。
    pairs = match_io_pairs_all(events)

    dev_timelines = defaultdict(list)
    for p in pairs:
        dev = p.get("dev") or "未知"
        dev_timelines[dev].append((p["insert_ts"], 1, 0))
        dev_timelines[dev].append((p["issue_ts"], -1, 1))
        dev_timelines[dev].append((p["complete_ts"], 0, -1))

    qd_by_dev = {}
    for dev, timeline in dev_timelines.items():
        timeline.sort(key=lambda x: x[0])

        dispatch_q = 0
        device_q = 0
        points = []

        for ts_ns, d_dispatch, d_device in timeline:
            dispatch_q += d_dispatch
            device_q += d_device
            if dispatch_q < 0:
                dispatch_q = 0
            if device_q < 0:
                device_q = 0

            points.append({
                "time_sec": ts_ns / NS_PER_SEC,
                "dispatch_queue": dispatch_q,
                "device_queue": device_q,
            })

        qd_by_dev[dev] = points

    return qd_by_dev


def calc_iops_throughput(events, bucket_sec=0.1):
    # 分设备统计 IOPS 与吞吐量，避免多设备混合导致的时间桶互相污染。
    items_by_dev = defaultdict(list)
    for ev in events:
        if str(ev.get("Trace 字段名称", "")) != "block_rq_insert":
            continue
        ts_ns_raw = ev.get("时间 (纳秒)", 0)
        try:
            ts_sec = int(ts_ns_raw) / NS_PER_SEC if ts_ns_raw else 0
        except (ValueError, TypeError):
            continue
        nr_sector_raw = ev.get("请求扇区数量", 0)
        try:
            nr_sector = int(nr_sector_raw) if nr_sector_raw else 0
        except (ValueError, TypeError):
            nr_sector = 0
        dev = str(ev.get("设备号", "") or "").strip() or "未知"
        items_by_dev[dev].append({"ts": ts_sec, "nr_sector": nr_sector})

    result = {}
    for dev, items in items_by_dev.items():
        if not items:
            continue
        min_ts = items[0]["ts"]
        max_ts = items[-1]["ts"]
        if max_ts <= min_ts:
            continue

        num_buckets = max(1, int((max_ts - min_ts) / bucket_sec) + 1)
        buckets = [{"time_sec": min_ts + i * bucket_sec, "count": 0, "bytes": 0}
                   for i in range(num_buckets)]

        for item in items:
            idx = int((item["ts"] - min_ts) / bucket_sec)
            if 0 <= idx < num_buckets:
                buckets[idx]["count"] += 1
                buckets[idx]["bytes"] += item["nr_sector"] * 512

        for b in buckets:
            b["iops"] = b["count"] / bucket_sec
            b["throughput_mb"] = b["bytes"] / bucket_sec / (1024 * 1024)

        result[dev] = buckets

    return result


def compute_percentiles(values):
    if not values:
        return {"P50": 0, "P95": 0, "P99": 0, "P99.9": 0, "P99.99": 0, "Avg": 0, "Max": 0, "Count": 0}
    arr = np.array(values, dtype=np.float64)
    return {
        "P50": float(np.percentile(arr, 50)),
        "P95": float(np.percentile(arr, 95)),
        "P99": float(np.percentile(arr, 99)),
        "P99.9": float(np.percentile(arr, 99.9)),
        "P99.99": float(np.percentile(arr, 99.99)),
        "Avg": float(np.mean(arr)),
        "Max": float(np.max(arr)),
        "Count": len(arr),
    }


def format_ms(ns):
    if ns is None:
        return ""
    return f"{ns / NS_PER_MS:.3f}"


def write_style_sheet(ws, headers):
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def auto_col_width(ws):
    for col_cells in ws.columns:
        max_len = 0
        col_letter = None
        for cell in col_cells:
            if col_letter is None:
                col_letter = cell.column_letter
            try:
                val = str(cell.value or "")
                max_len = max(max_len, len(val))
            except Exception:
                pass
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max_len + 4, 40)


def write_data_rows(ws, rows, start_row=2):
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = ws.cell(row=start_row + r_idx, column=c_idx + 1, value=val)
            cell.font = CELL_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left")


def save_chart(fig, output_path):
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  图表已保存: {output_path}")


def chart_queue_depth(qd_by_dev, output_path):
    if not qd_by_dev:
        return
    devs = sorted(qd_by_dev.keys(), key=lambda d: len(qd_by_dev[d]), reverse=True)[:MAX_DEV_SUBPLOTS]
    n = len(devs)

    fig, axes = plt.subplots(n, 1, figsize=(16, 3 * n), squeeze=False)
    for ax, dev in zip(axes[:, 0], devs):
        points = qd_by_dev[dev]
        times = [p["time_sec"] for p in points]
        dq = [p["dispatch_queue"] for p in points]
        dv = [p["device_queue"] for p in points]
        ax.fill_between(times, dq, alpha=0.3, color="#4472C4", step="post", label="调度队列深度 (insert - issue)")
        ax.fill_between(times, dv, alpha=0.5, color="#ED7D31", step="post", label="器件队列深度 (issue - complete)")
        ax.set_ylabel("队列深度", fontsize=10)
        ax.set_title(f"设备 {dev}", fontsize=12)
        ax.legend(loc="upper right", fontsize=9)
        ax.grid(axis="y", alpha=0.3, linestyle="--")
        ax.set_xlim(times[0], times[-1])

    axes[-1, 0].set_xlabel("时间 (秒)", fontsize=11)
    title_suffix = "" if len(qd_by_dev) <= MAX_DEV_SUBPLOTS else f"（仅显示 IO 最活跃的前 {MAX_DEV_SUBPLOTS} 个设备）"
    fig.suptitle(f"IO 队列深度时间曲线（分设备）{title_suffix}", fontsize=14)
    fig.tight_layout(rect=[0, 0, 1, 0.98])
    save_chart(fig, output_path)


def chart_iops_throughput(iops_by_dev, output_path):
    if not iops_by_dev:
        return
    devs = sorted(iops_by_dev.keys(), key=lambda d: sum(b["count"] for b in iops_by_dev[d]), reverse=True)[:MAX_DEV_SUBPLOTS]
    n = len(devs)

    fig, axes = plt.subplots(n, 1, figsize=(16, 3 * n), squeeze=False)
    for ax, dev in zip(axes[:, 0], devs):
        buckets = iops_by_dev[dev]
        times = [b["time_sec"] for b in buckets]
        iops = [b["iops"] for b in buckets]
        tp = [b["throughput_mb"] for b in buckets]

        width = (times[1] - times[0]) if len(times) > 1 else 0.1
        ax.bar(times, iops, width=width, alpha=0.6, color="#4472C4", label="IOPS")
        ax.set_ylabel("IOPS", fontsize=10, color="#4472C4")
        ax.tick_params(axis="y", labelcolor="#4472C4")

        ax2 = ax.twinx()
        ax2.plot(times, tp, color="#ED7D31", linewidth=1.0, label="吞吐量 (MB/s)")
        ax2.set_ylabel("吞吐量 (MB/s)", fontsize=10, color="#ED7D31")
        ax2.tick_params(axis="y", labelcolor="#ED7D31")

        ax.set_title(f"设备 {dev}", fontsize=12)
        lines1, labels1 = ax.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax.legend(lines1 + lines2, labels1 + labels2, loc="upper right", fontsize=9)
        ax.grid(axis="y", alpha=0.3, linestyle="--")
        ax.set_xlim(times[0], times[-1])

    axes[-1, 0].set_xlabel("时间 (秒)", fontsize=11)
    title_suffix = "" if len(iops_by_dev) <= MAX_DEV_SUBPLOTS else f"（仅显示 IO 最活跃的前 {MAX_DEV_SUBPLOTS} 个设备）"
    fig.suptitle(f"IOPS & 吞吐量时间曲线（分设备）{title_suffix}", fontsize=14)
    fig.tight_layout(rect=[0, 0, 1, 0.98])
    save_chart(fig, output_path)


def chart_rw_latency(read_lat, write_lat, output_path):
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))

    categories = ["schedule_ns", "device_ns", "end_to_end_ns"]
    titles = ["调度耗时", "器件耗时", "端到端耗时"]

    for ax, cat, title in zip(axes, categories, titles):
        positions = [1, 2]
        labels = ["Read", "Write"]

        r_vals = [p[cat] for p in read_lat] if read_lat else []
        w_vals = [p[cat] for p in write_lat] if write_lat else []

        bp = ax.boxplot([r_vals, w_vals] if r_vals and w_vals else ([r_vals] if r_vals else [w_vals]),
                        positions=positions if (r_vals and w_vals) else ([1] if r_vals else [2]),
                        widths=0.4, patch_artist=True,
                        boxprops=dict(facecolor="#4472C4", alpha=0.6),
                        medianprops=dict(color="red", linewidth=1.5))

        if r_vals and w_vals:
            ax.set_xticks([1, 2])
            ax.set_xticklabels(["Read", "Write"])
        elif r_vals:
            ax.set_xticks([1])
            ax.set_xticklabels(["Read"])
        elif w_vals:
            ax.set_xticks([2])
            ax.set_xticklabels(["Write"])

        ax.set_ylabel("耗时 (ms)")
        ax.set_title(title, fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.3f"))
        ax.grid(axis="y", alpha=0.3, linestyle="--")

        if r_vals:
            ax.annotate(f"P99={np.percentile(r_vals, 99)/NS_PER_MS:.2f}ms",
                        xy=(1, np.percentile(r_vals, 95)/NS_PER_MS),
                        fontsize=7, va="bottom", ha="center",
                        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8))
        if w_vals:
            ax.annotate(f"P99={np.percentile(w_vals, 99)/NS_PER_MS:.2f}ms",
                        xy=(2, np.percentile(w_vals, 95)/NS_PER_MS),
                        fontsize=7, va="bottom", ha="center",
                        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8))

    fig.suptitle("读写 IO 时延对比 (Box Plot)", fontsize=14)
    plt.tight_layout()
    save_chart(fig, output_path)


def chart_rw_size(read_sizes, write_sizes, output_path):
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))

    for ax, sizes, title, color in [
        (axes[0], read_sizes, "Read IO 大小分布", "#4472C4"),
        (axes[1], write_sizes, "Write IO 大小分布", "#ED7D31"),
    ]:
        if not sizes:
            ax.set_title(f"{title} (无数据)", fontsize=12)
            continue
        kb_vals = [s / 1024 for s in sizes]
        ax.hist(kb_vals, bins=50, color=color, alpha=0.7, edgecolor="white")
        ax.set_xlabel("IO 大小 (KB)", fontsize=10)
        ax.set_ylabel("IO 数量", fontsize=10)
        ax.set_title(title, fontsize=12)
        ax.grid(axis="y", alpha=0.3, linestyle="--")
        p50 = np.percentile(kb_vals, 50)
        avg = np.mean(kb_vals)
        ax.axvline(p50, color="red", linestyle="--", linewidth=1.0, label=f"P50={p50:.1f}KB")
        ax.axvline(avg, color="green", linestyle="--", linewidth=1.0, label=f"Avg={avg:.1f}KB")
        ax.legend(fontsize=8)

    fig.suptitle("读写 IO 大小对比", fontsize=14)
    plt.tight_layout()
    save_chart(fig, output_path)


def detect_seq_random(pairs):
    by_thread = defaultdict(list)
    for p in pairs:
        tid = str(p.get('insert_tid', ''))
        dev = str(p.get('dev', ''))
        by_thread[(tid, dev)].append(p)

    for key in by_thread:
        by_thread[key].sort(key=lambda x: x['insert_ts'])

    seq_count = 0
    rand_count = 0

    for key, io_list in by_thread.items():
        for i in range(len(io_list)):
            if i == 0:
                io_list[i]['access_type'] = 'random'
                io_list[i]['gap_sectors'] = 0
                rand_count += 1
                continue
            prev = io_list[i - 1]
            curr = io_list[i]
            try:
                prev_sector = int(prev['sector'])
                curr_sector = int(curr['sector'])
            except (ValueError, TypeError):
                curr['access_type'] = 'random'
                curr['gap_sectors'] = 0
                rand_count += 1
                continue
            prev_end = prev_sector + int(prev['nr_sector'])
            gap = abs(curr_sector - prev_end)
            curr['gap_sectors'] = gap
            if gap <= SEQ_GAP_THRESHOLD:
                curr['access_type'] = 'sequential'
                seq_count += 1
            else:
                curr['access_type'] = 'random'
                rand_count += 1

    return pairs, seq_count, rand_count


def lookup_queue_depth(qd_by_dev, ts_ns, dev=None):
    if dev is None:
        return 0, 0
    points = qd_by_dev.get(dev)
    if not points:
        return 0, 0
    ts_sec = ts_ns / NS_PER_SEC
    timestamps = [p['time_sec'] for p in points]
    idx = bisect.bisect_right(timestamps, ts_sec) - 1
    if idx < 0:
        idx = 0
    p = points[idx]
    return p['dispatch_queue'], p['device_queue']


def chart_seq_random_pie(pairs, seq_count, rand_count, output_path):
    if not pairs:
        return
    total = seq_count + rand_count
    if total == 0:
        return

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    labels = ['顺序 IO', '随机 IO']
    sizes = [seq_count, rand_count]
    colors = ['#4472C4', '#ED7D31']
    explode = (0.05, 0)

    axes[0].pie(sizes, explode=explode, labels=labels, colors=colors,
                autopct='%1.1f%%', startangle=90, textprops={'fontsize': 11})
    axes[0].set_title(f'顺序 vs 随机 IO 占比 (总计 {total} 次)', fontsize=13)

    rw_data = defaultdict(lambda: {'seq': 0, 'rand': 0})
    for p in pairs:
        at = p.get('access_type', 'random')
        rw = p['rwbs_type']
        if at == 'sequential':
            rw_data[rw]['seq'] += 1
        else:
            rw_data[rw]['rand'] += 1

    categories = []
    seq_vals = []
    rand_vals = []
    for rw in ['read', 'write', 'other']:
        if rw_data[rw]['seq'] + rw_data[rw]['rand'] > 0:
            categories.append({'read': 'Read', 'write': 'Write', 'other': 'Other'}[rw])
            seq_vals.append(rw_data[rw]['seq'])
            rand_vals.append(rw_data[rw]['rand'])

    x = np.arange(len(categories))
    width = 0.35
    bars_seq = axes[1].bar(x - width/2, seq_vals, width, label='顺序', color='#4472C4', alpha=0.8)
    bars_rand = axes[1].bar(x + width/2, rand_vals, width, label='随机', color='#ED7D31', alpha=0.8)

    all_vals = seq_vals + rand_vals
    max_val = max(all_vals) if all_vals else 0
    for bars in (bars_seq, bars_rand):
        for bar in bars:
            h = bar.get_height()
            if h > 0:
                axes[1].text(bar.get_x() + bar.get_width() / 2, h + max_val * 0.02,
                             str(int(h)), ha="center", va="bottom", fontsize=8, color="#333333")

    axes[1].set_xticks(x)
    axes[1].set_xticklabels(categories, fontsize=10)
    axes[1].set_ylabel('IO 数量', fontsize=10)
    axes[1].set_title('按读写类型拆分 顺序/随机', fontsize=13)
    axes[1].legend(fontsize=9)
    axes[1].grid(axis='y', alpha=0.3, linestyle='--')

    fig.suptitle('IO 访问模式分析', fontsize=14)
    plt.tight_layout()
    save_chart(fig, output_path)


def _scatter_latency_panel(ax, pairs, key, ts_field):
    """在 ax 上绘制指定时延维度的散点，返回所有时延值(ms)列表。"""
    read_ts, read_lat = [], []
    write_ts, write_lat = [], []
    sync_w_ts, sync_w_lat = [], []
    for p in pairs:
        t = p[ts_field] / NS_PER_SEC
        lat = p[key] / NS_PER_MS
        st = p.get("sync_type", "")
        if st == "sync_write":
            sync_w_ts.append(t)
            sync_w_lat.append(lat)
        elif p["rwbs_type"] == "read":
            read_ts.append(t)
            read_lat.append(lat)
        else:
            write_ts.append(t)
            write_lat.append(lat)
    if read_ts:
        ax.scatter(read_ts, read_lat, alpha=0.3, s=6, c="#4472C4", edgecolors="none", label=f"Read ({len(read_ts)})")
    if write_ts:
        ax.scatter(write_ts, write_lat, alpha=0.4, s=14, c="#00B050", edgecolors="none", marker="s", label=f"Write (async) ({len(write_ts)})")
    if sync_w_ts:
        ax.scatter(sync_w_ts, sync_w_lat, alpha=0.6, s=20, c="#C00000", edgecolors="none", marker="^", label=f"Sync Write ({len(sync_w_ts)})")
    return read_lat + write_lat + sync_w_lat


def chart_latency_timeline(pairs, output_path):
    if not pairs:
        return

    # 三个时延维度: (字段, X轴时间字段, Y轴标签)
    dims = [
        ("end_to_end_ns", "insert_ts", "端到端时延 (ms)"),
        ("schedule_ns", "insert_ts", "调度耗时\ninsert→issue (ms)"),
        ("device_ns", "issue_ts", "器件耗时\nissue→complete (ms)"),
    ]

    fig, axes = plt.subplots(4, 1, figsize=(16, 14), sharex=True,
                             gridspec_kw={"height_ratios": [2, 2, 2, 1]})

    has_data = False
    for idx, (key, ts_field, ylabel) in enumerate(dims):
        ax = axes[idx]
        all_lat = _scatter_latency_panel(ax, pairs, key, ts_field)
        if all_lat:
            has_data = True
            p999 = np.percentile(all_lat, 99.9)
            ax.axhline(y=p999, color="red", linestyle=":", linewidth=1.0,
                       label=f"P99.9 = {p999:.2f}ms")
        ax.set_ylabel(ylabel, fontsize=10)
        ax.set_yscale("log")
        ax.grid(alpha=0.3, linestyle="--")

        if idx == 0:
            ax.set_title("IO 时延时间线 (按 insert/issue 时间, 对数坐标)", fontsize=13)
            ax.legend(fontsize=8, loc="upper right", ncol=2)
            # 端到端面板标注 Top-3 高时延
            e2e_sorted = sorted(p["end_to_end_ns"] / NS_PER_MS for p in pairs)
            top_k = min(3, len(e2e_sorted))
            for i in range(top_k):
                val = e2e_sorted[-(i + 1)]
                for p in pairs:
                    if abs(p["end_to_end_ns"] / NS_PER_MS - val) < 0.001:
                        t_sec = p["insert_ts"] / NS_PER_SEC
                        ax.annotate(f"{val:.2f}ms", (t_sec, val),
                                    textcoords="offset points", xytext=(0, 8),
                                    fontsize=7, color="#C00000", ha="center")
                        break

    if not has_data:
        plt.close(fig)
        return

    axes[3].set_xlabel("时间 (秒)", fontsize=11)
    axes[3].set_ylabel("IO 数量", fontsize=10)
    axes[3].set_title("IO 密度 (每 0.1s)", fontsize=10)
    if pairs:
        t_min = pairs[0]["insert_ts"] / NS_PER_SEC
        buckets = {}
        for p in pairs:
            b = int((p["insert_ts"] / NS_PER_SEC - t_min) / 0.1)
            buckets[b] = buckets.get(b, 0) + 1
        times = [t_min + (k + 0.5) * 0.1 for k in sorted(buckets)]
        counts = [buckets[k] for k in sorted(buckets)]
        axes[3].fill_between(times, counts, alpha=0.4, color="#A0A0A0")
        axes[3].plot(times, counts, linewidth=0.5, color="#505050")
    axes[3].grid(alpha=0.3, linestyle="--")

    fig.subplots_adjust(hspace=0.08)
    save_chart(fig, output_path)

def chart_lba_timeline(events, output_path):
    """LBA (扇区地址) 时间线分布图：X=时间, Y=扇区地址, 按读写类型着色标注。"""
    read_ts, read_lba = [], []
    write_ts, write_lba = [], []
    sync_w_ts, sync_w_lba = [], []
    other_ts, other_lba = [], []

    for ev in events:
        if str(ev.get("Trace 字段名称", "")) != "block_rq_insert":
            continue
        ts_raw = ev.get("时间 (纳秒)", 0)
        try:
            ts_sec = int(ts_raw) / NS_PER_SEC if ts_raw else 0
        except (ValueError, TypeError):
            continue
        sector_raw = ev.get("请求扇区地址", 0)
        try:
            sector = int(sector_raw) if sector_raw else 0
        except (ValueError, TypeError):
            continue
        rwbs = str(ev.get("读写类型", ""))
        st = classify_sync_type(rwbs)
        rw = classify_rwbs(rwbs)
        if st == "sync_write":
            sync_w_ts.append(ts_sec)
            sync_w_lba.append(sector)
        elif rw == "read":
            read_ts.append(ts_sec)
            read_lba.append(sector)
        elif rw == "write":
            write_ts.append(ts_sec)
            write_lba.append(sector)
        else:
            other_ts.append(ts_sec)
            other_lba.append(sector)

    if not read_ts and not write_ts and not sync_w_ts and not other_ts:
        return

    fig, ax = plt.subplots(figsize=(16, 6))
    if read_ts:
        ax.scatter(read_ts, read_lba, alpha=0.3, s=6, c="#4472C4", edgecolors="none", label=f"Read ({len(read_ts)})")
    if write_ts:
        ax.scatter(write_ts, write_lba, alpha=0.4, s=14, c="#00B050", edgecolors="none", marker="s", label=f"Write (async) ({len(write_ts)})")
    if sync_w_ts:
        ax.scatter(sync_w_ts, sync_w_lba, alpha=0.6, s=20, c="#C00000", edgecolors="none", marker="^", label=f"Sync Write ({len(sync_w_ts)})")
    if other_ts:
        ax.scatter(other_ts, other_lba, alpha=0.3, s=6, c="#808080", edgecolors="none", marker="x", label=f"Other ({len(other_ts)})")

    ax.set_xlabel("时间 (秒)", fontsize=11)
    ax.set_ylabel("扇区地址 (LBA)", fontsize=11)
    ax.set_title("IO 扇区地址 (LBA) 时间线分布", fontsize=13)
    ax.legend(fontsize=8, loc="upper right", ncol=2)
    ax.grid(alpha=0.3, linestyle="--")
    ax.ticklabel_format(axis="y", style="plain", useOffset=False)

    fig.subplots_adjust(top=0.92)
    save_chart(fig, output_path)

def chart_sync_io_latency(pairs, output_path):
    if not pairs:
        return

    sync_w = [p["end_to_end_ns"] / NS_PER_MS for p in pairs if p.get("sync_type") == "sync_write"]
    async_w = [p["end_to_end_ns"] / NS_PER_MS for p in pairs if p.get("sync_type") == "async_write"]
    all_r = [p["end_to_end_ns"] / NS_PER_MS for p in pairs if p["rwbs_type"] == "read"]

    if not sync_w and not async_w:
        return

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    box_data, box_labels, colors = [], [], []
    if all_r:
        box_data.append(all_r)
        box_labels.append(f"Read\n(n={len(all_r)})")
        colors.append("#4472C4")
    if async_w:
        box_data.append(async_w)
        box_labels.append(f"Async Write\n(n={len(async_w)})")
        colors.append("#ED7D31")
    if sync_w:
        box_data.append(sync_w)
        box_labels.append(f"Sync Write\n(n={len(sync_w)})")
        colors.append("#C00000")

    bp = axes[0].boxplot(box_data, patch_artist=True,
                         showfliers=True, widths=0.5)
    axes[0].set_xticklabels(box_labels)
    for patch, c in zip(bp["boxes"], colors):
        patch.set_facecolor(c)
        patch.set_alpha(0.5)
    for flier in bp["fliers"]:
        flier.set(markersize=2, alpha=0.4)
    axes[0].set_ylabel("端到端时延 (ms)", fontsize=11)
    axes[0].set_title("读 / 异步写 / 同步写 时延分布", fontsize=13)
    axes[0].grid(axis="y", alpha=0.3, linestyle="--")

    axes[1].set_title("同步写 vs 异步写 时延百分位对比", fontsize=13)
    perc_labels = ["P50", "P90", "P95", "P99", "P99.9", "Max"]
    ps = [50, 90, 95, 99, 99.9, 100]
    async_vals = [np.percentile(async_w, p) for p in ps] if async_w else [0] * len(ps)
    sync_vals = [np.percentile(sync_w, p) for p in ps] if sync_w else [0] * len(ps)

    x = range(len(perc_labels))
    w = 0.35
    bars_async = axes[1].bar([i - w/2 for i in x], async_vals, w, label="Async Write", color="#ED7D31", alpha=0.7)
    bars_sync = axes[1].bar([i + w/2 for i in x], sync_vals, w, label="Sync Write", color="#C00000", alpha=0.7)

    for bars in (bars_async, bars_sync):
        for bar in bars:
            h = bar.get_height()
            if h > 0:
                axes[1].text(bar.get_x() + bar.get_width() / 2, h,
                             f"{h:.2f}", ha="center", va="bottom", fontsize=7, color="#333333")

    axes[1].set_xticks(x)
    axes[1].set_xticklabels(perc_labels)
    axes[1].set_ylabel("时延 (ms)", fontsize=11)
    axes[1].legend(fontsize=9)
    axes[1].grid(axis="y", alpha=0.3, linestyle="--")
    if max(max(async_vals or [0]), max(sync_vals or [0])) > 0:
        axes[1].set_yscale("log")

    fig.suptitle("同步 IO vs 异步 IO 时延分析", fontsize=14)
    plt.tight_layout()
    save_chart(fig, output_path)


def chart_size_latency_scatter(pairs, output_path):
    if not pairs:
        return

    read_x, read_y = [], []
    write_x, write_y = [], []

    for p in pairs:
        size_kb = p['nr_sector'] * 512 / 1024
        lat_ms = p['end_to_end_ns'] / NS_PER_MS
        if p['rwbs_type'] == 'read':
            read_x.append(size_kb)
            read_y.append(lat_ms)
        elif p['rwbs_type'] == 'write':
            write_x.append(size_kb)
            write_y.append(lat_ms)

    fig, axes = plt.subplots(1, 2, figsize=(14, 5))

    for ax, x_vals, y_vals, title, color in [
        (axes[0], read_x, read_y, 'Read IO: 大小 vs 端到端时延', '#4472C4'),
        (axes[1], write_x, write_y, 'Write IO: 大小 vs 端到端时延', '#ED7D31'),
    ]:
        if not x_vals:
            ax.set_title(f'{title} (无数据)', fontsize=12)
            continue
        ax.scatter(x_vals, y_vals, alpha=0.4, s=8, c=color, edgecolors='none')
        ax.set_xlabel('IO 大小 (KB)', fontsize=10)
        ax.set_ylabel('端到端时延 (ms)', fontsize=10)
        ax.set_title(title, fontsize=12)
        ax.grid(alpha=0.3, linestyle='--')

        if len(x_vals) > 1:
            z = np.polyfit(x_vals, y_vals, 1)
            p_line = np.poly1d(z)
            x_line = np.linspace(min(x_vals), max(x_vals), 100)
            ax.plot(x_line, p_line(x_line), color='red', linewidth=1.0,
                    linestyle='--', label=f'趋势 y={z[0]:.4f}x+{z[1]:.3f}')
            ax.legend(fontsize=8)

        stats_text = (
            f'P50: {np.percentile(x_vals, 50):.1f}KB / {np.percentile(y_vals, 50):.2f}ms\n'
            f'P99: {np.percentile(x_vals, 99):.1f}KB / {np.percentile(y_vals, 99):.2f}ms'
        )
        ax.text(0.97, 0.97, stats_text, transform=ax.transAxes,
                fontsize=8, va='top', ha='right',
                bbox=dict(boxstyle='round', facecolor='white', alpha=0.9))

    fig.suptitle('IO 大小 vs 时延 散点图', fontsize=14)
    plt.tight_layout()
    save_chart(fig, output_path)


def write_seq_random_sheet(wb, pairs):
    ws = wb.create_sheet('顺序随机')
    headers = ['类别', 'IO 数量', '占比', '总数据量 (KB)', '平均大小 (KB)', 'P50 时延 (ms)', 'P99 时延 (ms)']
    write_style_sheet(ws, headers)

    seq_pairs = [p for p in pairs if p.get('access_type') == 'sequential']
    rand_pairs = [p for p in pairs if p.get('access_type') == 'random']
    total = len(seq_pairs) + len(rand_pairs)

    def access_stats(plist):
        if not plist:
            return {'count': 0, 'total_kb': 0, 'avg_kb': 0, 'p50_ms': 0, 'p99_ms': 0}
        e2e_ns = [p['end_to_end_ns'] for p in plist]
        sizes_kb = [p['nr_sector'] * 512 / 1024 for p in plist]
        return {
            'count': len(plist),
            'total_kb': sum(sizes_kb),
            'avg_kb': np.mean(sizes_kb),
            'p50_ms': np.percentile(e2e_ns, 50) / NS_PER_MS,
            'p99_ms': np.percentile(e2e_ns, 99) / NS_PER_MS,
        }

    ss = access_stats(seq_pairs)
    rs = access_stats(rand_pairs)

    def ratio(n):
        return f'{n / total * 100:.1f}%' if total > 0 else '0%'

    rows = [
        ['顺序 IO', ss['count'], ratio(ss['count']), f"{ss['total_kb']:.1f}",
         f"{ss['avg_kb']:.1f}", f"{ss['p50_ms']:.3f}", f"{ss['p99_ms']:.3f}"],
        ['随机 IO', rs['count'], ratio(rs['count']), f"{rs['total_kb']:.1f}",
         f"{rs['avg_kb']:.1f}", f"{rs['p50_ms']:.3f}", f"{rs['p99_ms']:.3f}"],
    ]
    write_data_rows(ws, rows)
    auto_col_width(ws)

    if seq_pairs:
        seq_by_rw = defaultdict(list)
        for p in seq_pairs:
            seq_by_rw[p['rwbs_type']].append(p)
        ws2 = wb.create_sheet('顺序IO-读写拆分')
        rw_headers = ['读写类型', '顺序 IO 数', '平均大小 (KB)', 'P50 (ms)', 'P99 (ms)']
        write_style_sheet(ws2, rw_headers)
        rw_rows = []
        for rw in ['read', 'write', 'other']:
            plist = seq_by_rw.get(rw, [])
            if not plist:
                continue
            e2e = [p['end_to_end_ns'] for p in plist]
            sizes = [p['nr_sector'] * 512 / 1024 for p in plist]
            rw_rows.append([
                rw, len(plist), f'{np.mean(sizes):.1f}',
                f'{np.percentile(e2e, 50)/NS_PER_MS:.3f}',
                f'{np.percentile(e2e, 99)/NS_PER_MS:.3f}',
            ])
        write_data_rows(ws2, rw_rows)
        auto_col_width(ws2)


def write_slow_io_sheet(wb, pairs, qd_points):
    if not pairs:
        return

    sorted_pairs = sorted(pairs, key=lambda x: x['end_to_end_ns'], reverse=True)
    top_n = sorted_pairs[:SLOW_IO_TOP_N]

    ws = wb.create_sheet(f'慢IO Top{SLOW_IO_TOP_N}')
    headers = [
        '排名', '端到端 (ms)', '调度耗时 (ms)', '器件耗时 (ms)',
        'IO 大小 (KB)', '读写类型', '访问模式',
        'insert 时间 (秒)', '线程 TID', '线程名称',
        '扇区地址', '设备号',
        'insert时调度队列', 'issue时器件队列',
    ]
    write_style_sheet(ws, headers)

    rows = []
    for rank, p in enumerate(top_n, 1):
        insert_dq, _ = lookup_queue_depth(qd_points, p['insert_ts'], p['dev'])
        _, issue_dv = lookup_queue_depth(qd_points, p['issue_ts'], p['dev'])
        rows.append([
            rank,
            f"{p['end_to_end_ns'] / NS_PER_MS:.3f}",
            f"{p['schedule_ns'] / NS_PER_MS:.3f}",
            f"{p['device_ns'] / NS_PER_MS:.3f}",
            f"{p['nr_sector'] * 512 / 1024:.1f}",
            p['rwbs'],
            p.get('access_type', 'random'),
            f"{p['insert_ts'] / NS_PER_SEC:.6f}",
            p['insert_tid'],
            p['insert_thread'],
            p['sector'],
            p['dev'],
            insert_dq,
            issue_dv,
        ])
    write_data_rows(ws, rows)
    auto_col_width(ws)

def write_latency_timeline_sheet(wb, pairs):
    """时延时间线明细：逐 IO 的调度/器件/端到端耗时，供 AI 数据分析。"""
    if not pairs:
        return
    ws = wb.create_sheet("时延时间线")
    headers = [
        'insert 时间 (秒)', 'issue 时间 (秒)', 'complete 时间 (秒)',
        '调度耗时 (ms)', '器件耗时 (ms)', '端到端耗时 (ms)',
        '读写类型', '同步类型', 'IO 大小 (KB)',
        '进程 PID', '线程 TID', '线程名称', '设备号', '扇区地址',
    ]
    write_style_sheet(ws, headers)

    rows = []
    for p in pairs:
        rows.append([
            f"{p['insert_ts'] / NS_PER_SEC:.6f}",
            f"{p['issue_ts'] / NS_PER_SEC:.6f}",
            f"{p['complete_ts'] / NS_PER_SEC:.6f}",
            f"{p['schedule_ns'] / NS_PER_MS:.3f}",
            f"{p['device_ns'] / NS_PER_MS:.3f}",
            f"{p['end_to_end_ns'] / NS_PER_MS:.3f}",
            p['rwbs'],
            p.get('sync_type', ''),
            f"{p['nr_sector'] * 512 / 1024:.1f}",
            p['insert_pid'],
            p['insert_tid'],
            p['insert_thread'],
            p['dev'],
            p['sector'],
        ])
    write_data_rows(ws, rows)
    # 明细行数可能很大，用固定列宽避免 auto_col_width 全量遍历的性能开销
    fixed_widths = [16, 16, 16, 12, 12, 12, 10, 10, 10, 8, 8, 16, 8, 10]
    for col_idx, w in enumerate(fixed_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = w

def write_sync_io_sheet(wb, pairs):
    ws = wb.create_sheet("同步IO")
    sync_w = [p for p in pairs if p.get("sync_type") == "sync_write"]
    async_w = [p for p in pairs if p.get("sync_type") == "async_write"]
    sync_r = [p for p in pairs if p.get("sync_type") == "sync_read"]
    total_w = len(sync_w) + len(async_w)

    def _s(plist):
        if not plist:
            return dict(count=0, p50_ms=0, p95_ms=0, p99_ms=0, p999_ms=0, avg_ms=0, max_ms=0,
                        total_kb=0, avg_sched=0, avg_dev=0)
        e2e = [p["end_to_end_ns"] for p in plist]
        sch = [p["schedule_ns"] for p in plist]
        dev = [p["device_ns"] for p in plist]
        return dict(
            count=len(plist),
            p50_ms=np.percentile(e2e, 50) / NS_PER_MS,
            p95_ms=np.percentile(e2e, 95) / NS_PER_MS,
            p99_ms=np.percentile(e2e, 99) / NS_PER_MS,
            p999_ms=np.percentile(e2e, 99.9) / NS_PER_MS,
            avg_ms=np.mean(e2e) / NS_PER_MS,
            max_ms=np.max(e2e) / NS_PER_MS,
            total_kb=sum(p["nr_sector"] * 512 / 1024 for p in plist),
            avg_sched=np.mean(sch) / NS_PER_MS,
            avg_dev=np.mean(dev) / NS_PER_MS,
        )

    sws = _s(sync_w)
    aws = _s(async_w)
    srs = _s(sync_r)

    headers = ["IO 类别", "数量", "写占比", "P50 (ms)", "P95 (ms)", "P99 (ms)",
               "P99.9 (ms)", "Avg (ms)", "Max (ms)", "数据量 (KB)",
               "Avg调度 (ms)", "Avg器件 (ms)"]
    write_style_sheet(ws, headers)

    def wr(n):
        return f"{n / total_w * 100:.1f}%" if total_w > 0 else "-"

    rows = [
        ["同步写 (Sync Write)", sws["count"], wr(sws["count"]),
         f'{sws["p50_ms"]:.3f}', f'{sws["p95_ms"]:.3f}', f'{sws["p99_ms"]:.3f}',
         f'{sws["p999_ms"]:.3f}', f'{sws["avg_ms"]:.3f}', f'{sws["max_ms"]:.3f}',
         f'{sws["total_kb"]:.1f}', f'{sws["avg_sched"]:.3f}', f'{sws["avg_dev"]:.3f}'],
        ["异步写 (Async Write)", aws["count"], wr(aws["count"]),
         f'{aws["p50_ms"]:.3f}', f'{aws["p95_ms"]:.3f}', f'{aws["p99_ms"]:.3f}',
         f'{aws["p999_ms"]:.3f}', f'{aws["avg_ms"]:.3f}', f'{aws["max_ms"]:.3f}',
         f'{aws["total_kb"]:.1f}', f'{aws["avg_sched"]:.3f}', f'{aws["avg_dev"]:.3f}'],
    ]
    if srs["count"] > 0:
        rows.append(["同步读 (Sync Read)", srs["count"], "-",
                     f'{srs["p50_ms"]:.3f}', f'{srs["p95_ms"]:.3f}', f'{srs["p99_ms"]:.3f}',
                     f'{srs["p999_ms"]:.3f}', f'{srs["avg_ms"]:.3f}', f'{srs["max_ms"]:.3f}',
                     f'{srs["total_kb"]:.1f}', f'{srs["avg_sched"]:.3f}', f'{srs["avg_dev"]:.3f}'])

    write_data_rows(ws, rows)
    auto_col_width(ws)

def write_percentile_sheet(wb, pairs):
    ws = wb.create_sheet("时延百分位")

    sched = [p["schedule_ns"] for p in pairs]
    device = [p["device_ns"] for p in pairs]
    e2e = [p["end_to_end_ns"] for p in pairs]

    ps = compute_percentiles(sched)
    pd_ = compute_percentiles(device)
    pe = compute_percentiles(e2e)

    headers = ["指标", "调度耗时 (ms)", "器件耗时 (ms)", "端到端耗时 (ms)"]
    write_style_sheet(ws, headers)

    rows = []
    for label in ["Count", "P50", "P95", "P99", "P99.9", "P99.99", "Avg", "Max"]:
        rows.append([
            label,
            format_ms(ps[label]) if label != "Count" else str(ps[label]),
            format_ms(pd_[label]) if label != "Count" else str(pd_[label]),
            format_ms(pe[label]) if label != "Count" else str(pe[label]),
        ])
    write_data_rows(ws, rows)
    auto_col_width(ws)

    ws_summary = wb.create_sheet("时延汇总")
    p99_e2e = pe["P99"] / NS_PER_MS if pe["Count"] > 0 else 0
    p999_e2e = pe["P99.9"] / NS_PER_MS if pe["Count"] > 0 else 0
    summary_headers = ["指标", "值"]
    write_style_sheet(ws_summary, summary_headers)
    summary_rows = [
        ["匹配 IO 对数", pe["Count"]],
        ("P50 端到端", f"{pe['P50'] / NS_PER_MS:.3f} ms"),
        ("P95 端到端", f"{pe['P95'] / NS_PER_MS:.3f} ms"),
        ("P99 端到端", f"{p99_e2e:.3f} ms"),
        ("P99.9 端到端", f"{p999_e2e:.3f} ms"),
        ("Avg 端到端", f"{pe['Avg'] / NS_PER_MS:.3f} ms"),
        ("Max 端到端", f"{pe['Max'] / NS_PER_MS:.3f} ms"),
        ("P99 调度耗时", f"{ps['P99'] / NS_PER_MS:.3f} ms"),
        ("P99 器件耗时", f"{pd_['P99'] / NS_PER_MS:.3f} ms"),
    ]
    write_data_rows(ws_summary, summary_rows)
    auto_col_width(ws_summary)


def write_rw_sheet(wb, pairs):
    ws = wb.create_sheet("读写分离")

    read_pairs = [p for p in pairs if p["rwbs_type"] == "read"]
    write_pairs = [p for p in pairs if p["rwbs_type"] == "write"]
    other_pairs = [p for p in pairs if p["rwbs_type"] == "other"]

    def stats(latency_pairs):
        if not latency_pairs:
            return {"count": 0, "total_sectors": 0, "total_kb": 0,
                    "p50_ms": 0, "p99_ms": 0, "avg_ms": 0, "max_ms": 0}
        e2e_ns = [p["end_to_end_ns"] for p in latency_pairs]
        sectors = sum(p["nr_sector"] for p in latency_pairs)
        return {
            "count": len(latency_pairs),
            "total_sectors": sectors,
            "total_kb": sectors * 512 / 1024,
            "p50_ms": np.percentile(e2e_ns, 50) / NS_PER_MS,
            "p99_ms": np.percentile(e2e_ns, 99) / NS_PER_MS,
            "avg_ms": np.mean(e2e_ns) / NS_PER_MS,
            "max_ms": np.max(e2e_ns) / NS_PER_MS,
        }

    rs = stats(read_pairs)
    ws_ = stats(write_pairs)
    os_ = stats(other_pairs)
    total = rs["count"] + ws_["count"] + os_["count"]

    headers = ["类别", "IO 数量", "占比", "总扇区数", "总数据量 (KB)", "P50 (ms)", "P99 (ms)", "Avg (ms)", "Max (ms)"]
    write_style_sheet(ws, headers)

    def ratio(n):
        return f"{n / total * 100:.1f}%" if total > 0 else "0%"

    rows = [
        ["Read", rs["count"], ratio(rs["count"]), rs["total_sectors"], f"{rs['total_kb']:.1f}",
         f"{rs['p50_ms']:.3f}", f"{rs['p99_ms']:.3f}", f"{rs['avg_ms']:.3f}", f"{rs['max_ms']:.3f}"],
        ["Write", ws_["count"], ratio(ws_["count"]), ws_["total_sectors"], f"{ws_['total_kb']:.1f}",
         f"{ws_['p50_ms']:.3f}", f"{ws_['p99_ms']:.3f}", f"{ws_['avg_ms']:.3f}", f"{ws_['max_ms']:.3f}"],
        ["Other", os_["count"], ratio(os_["count"]), os_["total_sectors"], f"{os_['total_kb']:.1f}",
         f"{os_['p50_ms']:.3f}", f"{os_['p99_ms']:.3f}", f"{os_['avg_ms']:.3f}", f"{os_['max_ms']:.3f}"],
    ]
    write_data_rows(ws, rows)
    auto_col_width(ws)

    return read_pairs, write_pairs


def write_queue_depth_sheet(wb, qd_by_dev):
    ws = wb.create_sheet("队列深度")
    headers = ["设备号", "时间 (秒)", "调度队列深度", "器件队列深度"]
    write_style_sheet(ws, headers)
    rows = []
    for dev in sorted(qd_by_dev.keys()):
        for p in qd_by_dev[dev]:
            rows.append([dev, f"{p['time_sec']:.6f}", p["dispatch_queue"], p["device_queue"]])
    write_data_rows(ws, rows)
    auto_col_width(ws)


def write_iops_sheet(wb, iops_by_dev):
    ws = wb.create_sheet("IOPS_吞吐量")
    headers = ["设备号", "时间 (秒)", "IOPS", "吞吐量 (MB/s)", "IO 次数", "数据量 (Bytes)"]
    write_style_sheet(ws, headers)
    rows = []
    for dev in sorted(iops_by_dev.keys()):
        for b in iops_by_dev[dev]:
            rows.append([dev, f"{b['time_sec']:.6f}", f"{b['iops']:.1f}", f"{b['throughput_mb']:.2f}",
                         b["count"], b["bytes"]])
    write_data_rows(ws, rows)
    auto_col_width(ws)


def parse_args():
    parser = argparse.ArgumentParser(description="IO 综合诊断：队列深度、IOPS/吞吐量、时延百分位、读写分离、顺序/随机、散点图、慢IO TopN")
    parser.add_argument("--sheet", type=str, default=None,
                        help="读取 XLSX 中的哪个 Sheet（默认: 全部 IO，可配置文件覆盖）")
    parser.add_argument("--start-ns", type=float, default=None,
                        help="时间范围起始（纳秒）")
    parser.add_argument("--end-ns", type=float, default=None,
                        help="时间范围结束（纳秒）")
    parser.add_argument("--start", type=float, default=None,
                        help="时间范围起始（秒）")
    parser.add_argument("--end", type=float, default=None,
                        help="时间范围结束（秒）")
    parser.add_argument("--bucket", type=float, default=0.1,
                        help="IOPS/吞吐量时间分桶间隔（秒），默认 0.1")
    parser.add_argument("--top-n", type=int, default=SLOW_IO_TOP_N,
                        help="慢 IO Top-N 数量，默认 20")
    parser.add_argument("--file", type=str, default=None,
                        help="手动指定输入文件路径 (.csv/.xlsx)，跳过 out/ 目录自动扫描")
    return parser.parse_args()


def main():
    args = parse_args()

    cfg = load_config()
    diag_cfg = cfg.get("io_diagnosis", {})

    top_n = diag_cfg.get("top_n", args.top_n)
    global SLOW_IO_TOP_N
    SLOW_IO_TOP_N = top_n

    if args.sheet is None:
        args.sheet = diag_cfg.get("sheet", "全部 IO")

    start_s = None
    end_s = None

    if args.start_ns is not None:
        start_s = args.start_ns / NS_PER_SEC
    elif args.start is not None:
        start_s = args.start
    else:
        start_ns_cfg = diag_cfg.get("start_ns")
        if start_ns_cfg is not None:
            start_s = start_ns_cfg / NS_PER_SEC
        else:
            start_s = diag_cfg.get("start")

    if args.end_ns is not None:
        end_s = args.end_ns / NS_PER_SEC
    elif args.end is not None:
        end_s = args.end
    else:
        end_ns_cfg = diag_cfg.get("end_ns")
        if end_ns_cfg is not None:
            end_s = end_ns_cfg / NS_PER_SEC
        else:
            end_s = diag_cfg.get("end")

    bucket_sec = diag_cfg.get("bucket", args.bucket)

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
        print("[提示] out 目录下未找到 trace 输出子目录")
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
        if start_s is not None:
            print(f"  时间范围: {start_s:.6f} - {end_s:.6f} 秒")

        _, events = load_io_events(file_path, args.sheet)
        print(f"  读取到 {len(events)} 条 IO 记录")

        events = filter_by_time(events, start_s, end_s)
        print(f"  时间范围内: {len(events)} 条")

        insert_count = sum(1 for e in events if str(e.get("Trace 字段名称", "")) == "block_rq_insert")
        issue_count = sum(1 for e in events if str(e.get("Trace 字段名称", "")) == "block_rq_issue")
        complete_count = sum(1 for e in events if str(e.get("Trace 字段名称", "")) == "block_rq_complete")
        print(f"  Insert: {insert_count}, Issue: {issue_count}, Complete: {complete_count}")

        print("\n[1/8] 计算队列深度...")
        qd_points = calc_queue_depth(events)
        print(f"  共 {len(qd_points)} 个设备, {sum(len(v) for v in qd_points.values())} 个采样点")

        print("[2/8] 计算 IOPS & 吞吐量 (分桶={}s)...".format(bucket_sec))
        iops_buckets = calc_iops_throughput(events, bucket_sec)
        print(f"  共 {len(iops_buckets)} 个设备, {sum(len(v) for v in iops_buckets.values())} 个时间桶")

        print("[3/8] 匹配 IO 配对 & 计算时延统计...")
        pairs = match_io_pairs_all(events)
        print(f"  成功配对 {len(pairs)} 条 IO")

        print("[4/8] 检测顺序/随机 IO 访问模式...")
        pairs, seq_count, rand_count = detect_seq_random(pairs)
        if len(pairs) > 0:
            print(f"  顺序: {seq_count} ({seq_count/len(pairs)*100:.1f}%),  随机: {rand_count} ({rand_count/len(pairs)*100:.1f}%)")

        read_pairs = [p for p in pairs if p["rwbs_type"] == "read"]
        write_pairs = [p for p in pairs if p["rwbs_type"] == "write"]
        print(f"  读: {len(read_pairs)}, 写: {len(write_pairs)}")

        print("[5/8] 分类同步/异步 IO...")
        sync_w_count = sum(1 for p in pairs if p.get("sync_type") == "sync_write")
        async_w_count = sum(1 for p in pairs if p.get("sync_type") == "async_write")
        print(f"  同步写: {sync_w_count}, 异步写: {async_w_count}")

        print(f"[6/8] 识别慢 IO Top{SLOW_IO_TOP_N} & 计算时延-大小相关性...")
        print("[7/8] 生成图表...")
        print("[8/8] 写入 XLSX & PNG...")

        suffix = ""
        if start_s is not None:
            suffix = f"_{start_s:.0f}-{end_s:.0f}s"

        xlsx_out = os.path.join(trace_dir, f"{trace_name}{suffix}_diagnosis.xlsx")
        wb = Workbook()
        wb.remove(wb.active)

        write_percentile_sheet(wb, pairs)
        read_pairs2, write_pairs2 = write_rw_sheet(wb, pairs)
        write_queue_depth_sheet(wb, qd_points)
        write_iops_sheet(wb, iops_buckets)
        write_seq_random_sheet(wb, pairs)
        write_sync_io_sheet(wb, pairs)
        write_slow_io_sheet(wb, pairs, qd_points)
        write_latency_timeline_sheet(wb, pairs)

        wb.save(xlsx_out)
        print(f"  XLSX 已保存: {xlsx_out}")

        charts_dir = os.path.join(trace_dir, "charts")
        os.makedirs(charts_dir, exist_ok=True)

        chart_queue_depth(qd_points,
                          os.path.join(charts_dir, f"{trace_name}{suffix}_queue_depth.png"))
        chart_iops_throughput(iops_buckets,
                              os.path.join(charts_dir, f"{trace_name}{suffix}_iops_throughput.png"))

        r_lat = [{"schedule_ns": p["schedule_ns"], "device_ns": p["device_ns"],
                   "end_to_end_ns": p["end_to_end_ns"]} for p in read_pairs]
        w_lat = [{"schedule_ns": p["schedule_ns"], "device_ns": p["device_ns"],
                   "end_to_end_ns": p["end_to_end_ns"]} for p in write_pairs]
        chart_rw_latency(r_lat, w_lat,
                         os.path.join(charts_dir, f"{trace_name}{suffix}_rw_latency.png"))

        r_sizes = [p["nr_sector"] * 512 for p in read_pairs]
        w_sizes = [p["nr_sector"] * 512 for p in write_pairs]
        chart_rw_size(r_sizes, w_sizes,
                      os.path.join(charts_dir, f"{trace_name}{suffix}_rw_size.png"))

        chart_seq_random_pie(pairs, seq_count, rand_count,
                             os.path.join(charts_dir, f"{trace_name}{suffix}_seq_random.png"))
        chart_size_latency_scatter(pairs,
                                    os.path.join(charts_dir, f"{trace_name}{suffix}_size_latency.png"))

        chart_sync_io_latency(pairs,
                              os.path.join(charts_dir, f"{trace_name}{suffix}_sync_io_latency.png"))
        chart_latency_timeline(pairs,
                                os.path.join(charts_dir, f"{trace_name}{suffix}_latency_timeline.png"))
        chart_lba_timeline(events,
                            os.path.join(charts_dir, f"{trace_name}{suffix}_lba_timeline.png"))

        print(f"\n诊断完成: {trace_name}")


if __name__ == "__main__":
    main()