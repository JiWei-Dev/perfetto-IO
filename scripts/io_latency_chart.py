import os
import sys
import glob
import argparse
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
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


def match_io_pairs(events):
    """三段式配对封装: 委托共享配对核心, 拆分为调度/器件两份时延清单。"""
    triples = _match_io_pairs(events)
    schedule_latencies = []
    device_latencies = []

    for t in triples:
        insert_ev = t["insert_ev"]
        issue_ev = t["issue_ev"]
        complete_ev = t["complete_ev"]

        rwbs_sched = str(issue_ev.get("读写类型", "") or insert_ev.get("读写类型", ""))
        nr_sector_sched = issue_ev.get("请求扇区数量", 0) or insert_ev.get("请求扇区数量", 0)
        schedule_latencies.append({
            "dev": t["dev"],
            "sector": t["sector"],
            "insert_ts": t["insert_ts"],
            "issue_ts": t["issue_ts"],
            "latency_ns": t["schedule_ns"],
            "latency_ms": t["schedule_ns"] / 1_000_000,
            "rwbs": rwbs_sched,
            "rwbs_type": classify_rwbs(rwbs_sched),
            "nr_sector": nr_sector_sched,
            "insert_process": insert_ev.get("IO 进程名称", ""),
            "insert_tid": insert_ev.get("线程 TID", ""),
            "insert_thread": insert_ev.get("线程名称", ""),
            "insert_cpu": insert_ev.get("运行 CPU", ""),
            "issue_process": issue_ev.get("IO 进程名称", ""),
            "issue_tid": issue_ev.get("线程 TID", ""),
            "issue_thread": issue_ev.get("线程名称", ""),
            "issue_cpu": issue_ev.get("运行 CPU", ""),
        })

        rwbs_dev = str(complete_ev.get("读写类型", "") or issue_ev.get("读写类型", ""))
        nr_sector_dev = complete_ev.get("请求扇区数量", 0) or issue_ev.get("请求扇区数量", 0)
        device_latencies.append({
            "dev": t["dev"],
            "sector": t["sector"],
            "issue_ts": t["issue_ts"],
            "complete_ts": t["complete_ts"],
            "latency_ns": t["device_ns"],
            "latency_ms": t["device_ns"] / 1_000_000,
            "rwbs": rwbs_dev,
            "rwbs_type": classify_rwbs(rwbs_dev),
            "nr_sector": nr_sector_dev,
            "issue_process": issue_ev.get("IO 进程名称", ""),
            "issue_tid": issue_ev.get("线程 TID", ""),
            "issue_thread": issue_ev.get("线程名称", ""),
            "issue_cpu": issue_ev.get("运行 CPU", ""),
            "complete_process": complete_ev.get("IO 进程名称", ""),
            "complete_tid": complete_ev.get("线程 TID", ""),
            "complete_thread": complete_ev.get("线程名称", ""),
            "complete_cpu": complete_ev.get("运行 CPU", ""),
        })

    return schedule_latencies, device_latencies


def determine_bins(data_ms, target_bins=30, cap_ms=None):
    if not data_ms:
        return np.array([0, 1]), None
    min_val = min(data_ms)
    max_val = max(data_ms)
    if max_val <= min_val:
        return np.array([min_val, min_val + 0.1]), None

    capped = False
    if cap_ms is not None and max_val > cap_ms:
        plot_max = cap_ms
        capped = True
    else:
        plot_max = max_val

    bins = np.logspace(np.log10(max(min_val, 0.001)), np.log10(plot_max * 1.1), target_bins)
    bins = np.unique(np.round(bins, 3))
    return bins, max_val if capped else None


def plot_latency_distribution(latencies, title, output_path, color_palette=None, pcap=None):
    if color_palette is None:
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
                    color="#D32F2F",
                    bbox=dict(boxstyle="round,pad=0.3", facecolor="#FFEBEE", alpha=0.95))

        ax.set_xscale("log")
        ax.set_xlabel("时延 (ms)", fontsize=11)
        ax.set_ylabel("IO 数量", fontsize=11)

        stats_text = (f"总数: {len(data)}\n"
                      f"P50: {np.percentile(data, 50):.2f} ms\n"
                      f"P99: {np.percentile(data, 99):.2f} ms\n"
                      f"均值: {np.mean(data):.2f} ms\n"
                      f"最大: {np.max(data):.2f} ms")
        ax.text(0.97, 0.97, stats_text, transform=ax.transAxes,
                fontsize=8, va="top", ha="right",
                bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.9))

        ax.set_title(labels_cn[cat], fontsize=14)
        ax.grid(axis="y", alpha=0.3, linestyle="--")
        ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))

    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  图表已保存: {output_path}")


def ts_ns_to_sec(ts_ns):
    if ts_ns is None or ts_ns == "":
        return ""
    try:
        return f"{int(ts_ns) / 1_000_000_000:.9f}"
    except (ValueError, TypeError):
        return str(ts_ns)


def write_latency_xlsx(schedule_latencies, device_latencies, output_path):
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
    cell_font = Font(name="Consolas", size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    schedule_columns = [
        ("设备号", "dev"),
        ("请求扇区地址", "sector"),
        ("请求扇区数量", "nr_sector"),
        ("读写类型", "rwbs"),
        ("insert 时间 (秒)", "insert_ts"),
        ("insert 时间 (纳秒)", "insert_ts"),
        ("insert 进程", "insert_process"),
        ("insert 线程 TID", "insert_tid"),
        ("insert 线程名称", "insert_thread"),
        ("insert CPU", "insert_cpu"),
        ("issue 时间 (秒)", "issue_ts"),
        ("issue 时间 (纳秒)", "issue_ts"),
        ("issue 进程", "issue_process"),
        ("issue 线程 TID", "issue_tid"),
        ("issue 线程名称", "issue_thread"),
        ("issue CPU", "issue_cpu"),
        ("调度耗时 (ms)", "latency_ms"),
        ("调度耗时 (ns)", "latency_ns"),
    ]

    device_columns = [
        ("设备号", "dev"),
        ("请求扇区地址", "sector"),
        ("请求扇区数量", "nr_sector"),
        ("读写类型", "rwbs"),
        ("issue 时间 (秒)", "issue_ts"),
        ("issue 时间 (纳秒)", "issue_ts"),
        ("issue 进程", "issue_process"),
        ("issue 线程 TID", "issue_tid"),
        ("issue 线程名称", "issue_thread"),
        ("issue CPU", "issue_cpu"),
        ("complete 时间 (秒)", "complete_ts"),
        ("complete 时间 (纳秒)", "complete_ts"),
        ("complete 进程", "complete_process"),
        ("complete 线程 TID", "complete_tid"),
        ("complete 线程名称", "complete_thread"),
        ("complete CPU", "complete_cpu"),
        ("器件耗时 (ms)", "latency_ms"),
        ("器件耗时 (ns)", "latency_ns"),
    ]

    right_align_keys = {"insert_ts", "issue_ts", "complete_ts", "latency_ms", "latency_ns", "nr_sector"}
    center_align_keys = {"insert_cpu", "issue_cpu", "complete_cpu", "insert_tid", "issue_tid", "complete_tid"}

    wb = Workbook()

    for sheet_title, data, columns in [
        ("调度耗时", schedule_latencies, schedule_columns),
        ("器件耗时", device_latencies, device_columns),
    ]:
        if sheet_title == "调度耗时":
            ws = wb.active
            ws.title = sheet_title
        else:
            ws = wb.create_sheet(title=sheet_title)

        col_names = [c[0] for c in columns]
        col_keys = [c[1] for c in columns]

        for col_idx, name in enumerate(col_names, 1):
            cell = ws.cell(row=1, column=col_idx, value=name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for row_idx, item in enumerate(data, 2):
            for col_idx, (col_name, key) in enumerate(zip(col_names, col_keys), 1):
                raw_val = item.get(key, "")
                if key in ("insert_ts", "issue_ts", "complete_ts"):
                    if "纳秒" in col_name:
                        val = raw_val
                    else:
                        val = ts_ns_to_sec(raw_val)
                elif key == "latency_ms":
                    val = round(raw_val, 6) if isinstance(raw_val, float) else raw_val
                else:
                    val = raw_val
                cell = ws.cell(row=row_idx, column=col_idx, value=val if val != "" and val is not None else "")
                cell.font = cell_font
                cell.border = thin_border
                if key in right_align_keys:
                    cell.alignment = Alignment(horizontal="right")
                elif key in center_align_keys:
                    cell.alignment = Alignment(horizontal="center")

        col_widths = [16, 16, 14, 20, 22, 20, 28, 12, 22, 8, 22, 20, 28, 12, 22, 8, 16, 18]
        for col_idx, width in enumerate(col_widths, 1):
            col_letter = chr(64 + col_idx) if col_idx <= 26 else ""
            if col_letter:
                try:
                    ws.column_dimensions[col_letter].width = width
                except Exception:
                    pass

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

    wb.save(output_path)


def parse_args():
    parser = argparse.ArgumentParser(description="计算 Block IO 时延并绘制分布图")
    parser.add_argument(
        "--sheet", type=str, default=None,
        help="读取 XLSX 中的哪个 Sheet（默认: 全部 IO，可配置文件覆盖）"
    )
    parser.add_argument(
        "--pcap", type=float, default=None,
        help="横坐标截断分位数（如 99.9），不指定则不截断"
    )
    parser.add_argument(
        "--file", type=str, default=None,
        help="手动指定输入文件路径 (.csv/.xlsx)，跳过 out/ 目录自动扫描"
    )
    return parser.parse_args()


def main():
    args = parse_args()

    cfg = load_config()
    chart_cfg = cfg.get("io_latency_chart", {})

    if args.sheet is None:
        args.sheet = chart_cfg.get("sheet", "全部 IO")
    if args.pcap is None:
        args.pcap = chart_cfg.get("pcap")

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
        pcap_info = f"P{args.pcap}" if args.pcap is not None else "全量"
        print(f"  横坐标范围: {pcap_info}")

        headers, events = load_io_events(file_path, args.sheet)

        if not events:
            print(f"  [警告] Sheet 中无数据")
            continue

        print(f"  读取到 {len(events)} 条 IO 记录")

        schedule_latencies, device_latencies = match_io_pairs(events)

        print(f"  匹配到 {len(schedule_latencies)} 对调度耗时 (insert→issue)")
        print(f"  匹配到 {len(device_latencies)} 对器件耗时 (issue→complete)")

        latency_xlsx = os.path.join(trace_dir, f"{trace_name}_latency.xlsx")
        print(f"\n[写入] 耗时明细表格")
        write_latency_xlsx(schedule_latencies, device_latencies, latency_xlsx)
        print(f"  表格已保存: {latency_xlsx}")

        charts_dir = os.path.join(trace_dir, "charts")
        os.makedirs(charts_dir, exist_ok=True)

        if schedule_latencies:
            chart_path = os.path.join(charts_dir, f"{trace_name}_schedule_latency.png")
            print(f"\n[绘图] 调度耗时分布图")
            plot_latency_distribution(
                schedule_latencies,
                f"{trace_name} - IO 调度耗时分布 (insert → issue)",
                chart_path,
                pcap=args.pcap,
            )

        if device_latencies:
            chart_path = os.path.join(charts_dir, f"{trace_name}_device_latency.png")
            print(f"\n[绘图] 器件耗时分布图")
            plot_latency_distribution(
                device_latencies,
                f"{trace_name} - 器件耗时分布 (issue → complete)",
                chart_path,
                pcap=args.pcap,
            )

    print(f"\n全部完成。")


if __name__ == "__main__":
    main()