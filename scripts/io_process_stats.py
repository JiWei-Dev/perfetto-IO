import os
import sys
import glob
import argparse
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from config_loader import load_config
from io_reader import load_io_events, find_block_io_file


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


def build_process_stats(events):
    """从 insert 事件统计每个进程/线程的 IO 数量和总数据量。"""
    key_cols = ["IO 进程名称", "线程 TID", "线程名称"]

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

        rwbs_str = str(ev.get("读写类型", ""))
        cat = classify_rwbs(rwbs_str)

        stats[key][cat] += 1

        try:
            nr_sector = int(ev.get("请求扇区数量", 0) or 0)
        except (ValueError, TypeError):
            nr_sector = 0
        stats[key]["total_kb"] += nr_sector * 512 / 1024

    return stats


def write_process_stats(stats, output_path):
    import openpyxl

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
    cell_font = Font(name="Consolas", size=10)
    bold_font = Font(name="Microsoft YaHei", bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    total_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    total_font = Font(name="Microsoft YaHei", bold=True, size=11)

    col_headers = [
        "IO 发起进程",
        "进程 PID",
        "线程 TID",
        "IO 发起线程",
        "读请求数量",
        "写请求数量",
        "其他请求数量",
        "总请求数量",
        "总数据量 (KB)",
    ]

    if os.path.isfile(output_path):
        wb = load_workbook(output_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    sheet_name = "进程 IO 统计"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    for col_idx, name in enumerate(col_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

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
            data["read"],
            data["write"],
            data["other"],
            total_count,
            round(data["total_kb"], 2),
        ]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val if val != 0 or col_idx > 5 else val)
            cell.font = cell_font
            cell.border = thin_border
            if col_idx >= 3:
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

        grand_total["read"] += data["read"]
        grand_total["write"] += data["write"]
        grand_total["other"] += data["other"]
        grand_total["total_kb"] += data["total_kb"]

        row_idx += 1

    grand_total_count = grand_total["read"] + grand_total["write"] + grand_total["other"]
    total_values = [
        "合计",
        "",
        "",
        "",
        grand_total["read"],
        grand_total["write"],
        grand_total["other"],
        grand_total_count,
        round(grand_total["total_kb"], 2),
    ]
    for col_idx, val in enumerate(total_values, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = thin_border
        if col_idx >= 3:
            cell.alignment = Alignment(horizontal="right")
        else:
            cell.alignment = Alignment(horizontal="left")

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

    wb.save(output_path)
    print(f"  表格已保存: {output_path}")


def parse_args():
    parser = argparse.ArgumentParser(description="统计 Block IO 各进程/线程的 insert 请求量和数据量")
    parser.add_argument(
        "--sheet", type=str, default=None,
        help="读取 XLSX 中的哪个 Sheet（默认: 全部 IO，可配置文件覆盖）"
    )
    parser.add_argument(
        "--file", type=str, default=None,
        help="手动指定输入文件路径 (.csv/.xlsx)，跳过 out/ 目录自动扫描"
    )
    return parser.parse_args()


def main():
    args = parse_args()

    cfg = load_config()
    proc_cfg = cfg.get("io_process_stats", {})

    if args.sheet is None:
        args.sheet = proc_cfg.get("sheet", "全部 IO")

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

        headers, events = load_io_events(file_path, args.sheet)
        print(f"  读取到 {len(events)} 条 IO 记录")

        insert_events = [e for e in events if str(e.get("Trace 字段名称", "")).strip() == "block_rq_insert"]
        print(f"  其中 insert 事件: {len(insert_events)} 条")

        stats = build_process_stats(events)
        process_count = len(stats)
        print(f"  涉及进程/线程组合: {process_count} 个")

        total_r = sum(d["read"] for d in stats.values())
        total_w = sum(d["write"] for d in stats.values())
        total_o = sum(d["other"] for d in stats.values())
        total_kb = sum(d["total_kb"] for d in stats.values())
        print(f"  R: {total_r}  W: {total_w}  Other: {total_o}  Total: {total_r+total_w+total_o}  Data: {total_kb:.1f} KB")

        latency_xlsx = os.path.join(trace_dir, f"{trace_name}_latency.xlsx")
        print(f"\n[写入] 进程 IO 统计表 → {latency_xlsx}")
        write_process_stats(stats, latency_xlsx)

    print(f"\n全部完成。")


if __name__ == "__main__":
    main()