#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
split_trace.py -- 按时间窗口切割 Perfetto trace，输出 XLSX 文件

原理:
    使用 trace_processor_shell 加载 trace 后, 用 UNION ALL SQL 一次查询所有时间分段,
    按 seg 列拆分后输出为独立 XLSX 文件, 格式与 parse_block_io.py 完全一致。

用法:
    python scripts/split_trace.py "com.ss.android.ugc.aweme" --pairs 237863509246000:237865500000000 237865500000000:237867500000000 237867500000000:237870000743000
"""

import argparse
import csv
import io
import os
import subprocess
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Excel 样式 (与 parse_block_io.py 保持一致)
# ---------------------------------------------------------------------------

COLUMN_NAMES = [
    "IO 进程名称",
    "进程 PID",
    "线程 TID",
    "线程名称",
    "时间 (秒)",
    "时间 (纳秒)",
    "运行 CPU",
    "Trace 字段名称",
    "设备号",
    "请求扇区地址",
    "请求扇区数量",
    "请求字节数",
    "读写类型",
    "进程 Comm",
    "命令",
    "错误码",
]

COLUMN_WIDTHS = [28, 10, 10, 28, 22, 20, 10, 22, 16, 16, 16, 14, 28, 28, 30, 10]

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
CELL_FONT = Font(name="Consolas", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

MAX_EXCEL_ROWS = 1_048_576
MAX_DATA_ROWS = MAX_EXCEL_ROWS - 1


# ---------------------------------------------------------------------------
# 数据转换函数 (与 parse_block_io.py 保持一致)
# ---------------------------------------------------------------------------

def clean_null(val):
    if val is None or val == "[NULL]":
        return ""
    return val


def ts_to_seconds(ts_ns):
    if ts_ns is None or ts_ns == "":
        return ""
    try:
        ts = int(ts_ns)
    except (ValueError, TypeError):
        return str(ts_ns)
    return f"{ts / 1_000_000_000:.9f}"


def decode_device(dev_num):
    if dev_num is None or dev_num == "":
        return ""
    try:
        dev = int(dev_num)
    except (ValueError, TypeError):
        return str(dev_num)
    major = (dev >> 8) & 0xFFF
    minor = dev & 0xFF
    return f"{major}:{minor}"


def decode_rwbs(rwbs):
    if not rwbs:
        return ""
    mapping = {
        "R": "Read",
        "W": "Write",
        "D": "Discard",
        "F": "Flush",
        "S": "Sync",
        "M": "Meta",
        "A": "RAhead",
        "N": "None",
    }
    parts = []
    for ch in rwbs:
        parts.append(mapping.get(ch, ch))
    return f"{rwbs} ({'|'.join(parts)})"


def build_row_values(hdrs, row):
    row_dict = dict(zip(hdrs, row))
    return [
        clean_null(row_dict.get("process_name", "")),
        clean_null(row_dict.get("pid", "")),
        clean_null(row_dict.get("tid", "")),
        clean_null(row_dict.get("thread_name", "")),
        ts_to_seconds(row_dict.get("ts", "")),
        clean_null(row_dict.get("ts", "")),
        clean_null(row_dict.get("cpu", "")),
        clean_null(row_dict.get("trace_event", "")),
        decode_device(row_dict.get("device", "")),
        clean_null(row_dict.get("sector", "")),
        clean_null(row_dict.get("nr_sector", "")),
        clean_null(row_dict.get("bytes_val", "")),
        decode_rwbs(row_dict.get("rwbs", "")),
        clean_null(row_dict.get("comm", "")),
        clean_null(row_dict.get("cmd", "")),
        clean_null(row_dict.get("error", "")),
    ]


# ---------------------------------------------------------------------------
# Sheet 写入 (与 parse_block_io.py 保持一致)
# ---------------------------------------------------------------------------

def _write_one_sheet(ws, headers, rows):
    for col_idx, name in enumerate(COLUMN_NAMES, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    for row_idx, row in enumerate(rows, 2):
        values = build_row_values(headers, row)
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val if val != "" else "")
            cell.font = CELL_FONT
            cell.border = THIN_BORDER
            if col_idx in (5, 6, 10, 11, 12, 16):
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 7:
                cell.alignment = Alignment(horizontal="center")

    for col_idx, width in enumerate(COLUMN_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def write_sheets(wb, base_title, headers, rows):
    if len(rows) <= MAX_DATA_ROWS:
        ws = wb.active if wb.active.title == "Sheet" else wb.create_sheet(title=base_title)
        ws.title = base_title
        _write_one_sheet(ws, headers, rows)
        print(f"      Sheet「{base_title}」({len(rows)} 行)")
    else:
        print(f"      数据量 {len(rows)} 行, 超过 Sheet 上限, 自动拆分 ...")
        for part_no, start in enumerate(range(0, len(rows), MAX_DATA_ROWS), 1):
            chunk = rows[start:start + MAX_DATA_ROWS]
            title = f"{base_title} ({part_no})"
            ws = wb.create_sheet(title=title)
            _write_one_sheet(ws, headers, chunk)
            print(f"      Sheet「{title}」({len(chunk)} 行)")
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]


# ---------------------------------------------------------------------------
# 路径查找
# ---------------------------------------------------------------------------

SQL_COL_NAMES = [
    "seg", "process_name", "pid", "tid", "thread_name", "ts", "cpu",
    "trace_event", "device", "sector", "nr_sector", "bytes_val",
    "rwbs", "comm", "cmd", "error",
]

_DATA_HEADERS = SQL_COL_NAMES[1:]      # 去掉 seg 列


def get_project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def locate_trace_file(trace_name):
    project_root = get_project_root()
    trace_dir = os.path.join(project_root, "trace")
    for c in [
        os.path.join(trace_dir, f"{trace_name}.perfetto-trace"),
        os.path.join(trace_dir, f"{trace_name}.perfetto"),
        os.path.join(trace_dir, trace_name),
    ]:
        if os.path.isfile(c):
            return c
    return None


def locate_trace_processor():
    project_root = get_project_root()
    tp = os.path.join(project_root, "bin", "windows-amd64", "trace_processor_shell.exe")
    if os.path.isfile(tp):
        return tp
    tp = os.path.join(project_root, "bin", "trace_processor_shell")
    if os.path.isfile(tp):
        return tp
    return None


# ---------------------------------------------------------------------------
# UNION ALL SQL 生成
# ---------------------------------------------------------------------------

def build_union_sql(time_ranges):
    parts = []
    for i, (start_ns, end_ns) in enumerate(time_ranges):
        parts.append(f"""SELECT
    {i} AS seg,
    COALESCE(p.name, '<kernel>') AS process_name,
    COALESCE(CAST(p.pid AS TEXT), '') AS pid,
    t.tid,
    t.name AS thread_name,
    e.ts,
    e.cpu,
    e.name AS trace_event,
    MAX(CASE WHEN a.key = 'dev'       THEN a.int_value END) AS device,
    MAX(CASE WHEN a.key = 'sector'    THEN a.int_value END) AS sector,
    MAX(CASE WHEN a.key = 'nr_sector' THEN a.int_value END) AS nr_sector,
    MAX(CASE WHEN a.key = 'bytes'     THEN a.int_value END) AS bytes_val,
    MAX(CASE WHEN a.key = 'rwbs'      THEN a.string_value END) AS rwbs,
    MAX(CASE WHEN a.key = 'comm'      THEN a.string_value END) AS comm,
    MAX(CASE WHEN a.key = 'cmd'       THEN a.string_value END) AS cmd,
    MAX(CASE WHEN a.key = 'error'     THEN a.int_value END) AS error
FROM ftrace_event e
JOIN thread t ON e.utid = t.id
LEFT JOIN process p ON t.upid = p.upid
LEFT JOIN args a ON e.arg_set_id = a.arg_set_id
    AND a.key IN ('dev', 'sector', 'nr_sector', 'bytes', 'rwbs', 'comm', 'cmd', 'error')
WHERE e.name IN ('block_rq_insert', 'block_rq_issue', 'block_rq_complete')
  AND e.ts BETWEEN {start_ns} AND {end_ns}
GROUP BY e.id
""")
    col_str = ", ".join(SQL_COL_NAMES)
    return f"SELECT {col_str} FROM (\n" + "\nUNION ALL\n".join(parts) + "\n) ORDER BY seg, ts"


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser(description="按时间窗口切割 Perfetto trace")
    p.add_argument("trace", help="trace 文件名或路径 (如: com.ss.android.ugc.aweme)")
    p.add_argument(
        "--pairs", type=str, nargs="+", required=True,
        help="时间对, 格式: start1:end1 start2:end2 ...  (纳秒)"
    )
    p.add_argument(
        "-o", "--out-dir", type=str, default=None,
        help="输出目录 (默认: out/<trace_name>_segments/)"
    )
    p.add_argument(
        "--timeout", type=int, default=900,
        help="trace_processor_shell 超时秒数 (默认 900 = 15 分钟)"
    )
    p.add_argument(
        "--format", type=str, default="both", choices=["csv", "xlsx", "both"],
        help="输出格式: csv (快速), xlsx (格式化报告), both (默认, 同时输出)"
    )
    args = p.parse_args()

    # ---- 解析时间范围 ----
    time_ranges = []
    for pair in args.pairs:
        parts = pair.split(":")
        if len(parts) != 2:
            print(f"[错误] 时间对格式错误: {pair} (应为 start:end)")
            sys.exit(1)
        time_ranges.append((int(parts[0]), int(parts[1])))
    for s, e in time_ranges:
        if s >= e:
            print(f"[错误] 起始时间必须 < 结束时间: {s} >= {e}")
            sys.exit(1)

    # ---- 定位 trace 文件和工具 ----
    tp_path = locate_trace_processor()
    if not tp_path:
        print("[错误] 找不到 trace_processor_shell")
        sys.exit(1)

    trace_path = locate_trace_file(args.trace)
    if not trace_path:
        trace_path = args.trace
        if not os.path.isfile(trace_path):
            print(f"[错误] trace 文件不存在: {args.trace}")
            sys.exit(1)

    trace_name = os.path.splitext(os.path.basename(trace_path))[0]
    if trace_name.endswith(".perfetto"):
        trace_name = trace_name[:-len(".perfetto")]

    if args.out_dir:
        out_dir = args.out_dir
    else:
        out_dir = os.path.join(get_project_root(), "out", trace_name + "_segments")

    total_segments = len(time_ranges)
    trace_mb = os.path.getsize(trace_path) / 1024 / 1024

    print()
    print("=" * 60)
    print("  Perfetto Trace 分割工具")
    print("=" * 60)
    print(f"  Trace:   {os.path.basename(trace_path)} ({trace_mb:.1f} MB)")
    print(f"  分段数:  {total_segments}")
    for i, (s, e) in enumerate(time_ranges, 1):
        print(f"    {i}. {s/1e9:.6f}s ~ {e/1e9:.6f}s   ({(e-s)/1e9:.3f}s)")
    print(f"  输出目录: {out_dir}")
    print(f"  输出格式: {args.format}")
    print("=" * 60)

    # ---- 生成 SQL 并写入临时文件 ----
    sql = build_union_sql(time_ranges)
    sql_file = trace_path + ".split_query.sql"
    with open(sql_file, "w", encoding="utf-8") as f:
        f.write(sql)

    # ---- 运行 trace_processor_shell ----
    est_min = trace_mb / 400
    print(f"\n[1/3] 启动 trace_processor_shell ... (预计 {est_min:.0f}-{est_min*1.5:.0f} 分钟, 请耐心等待)")
    print(f"      SQL 文件: {sql_file} ({os.path.getsize(sql_file)} bytes)")

    try:
        result = subprocess.run(
            [tp_path, "query", "-f", sql_file, trace_path],
            capture_output=True,
            timeout=args.timeout,
        )
    except subprocess.TimeoutExpired:
        print(f"\n[错误] trace_processor_shell 超时 ({args.timeout}s)")
        try:
            os.remove(sql_file)
        except OSError:
            pass
        sys.exit(1)

    try:
        os.remove(sql_file)
    except OSError:
        pass

    stdout = result.stdout.decode("utf-8", errors="replace")
    stderr = result.stderr.decode("utf-8", errors="replace")

    if result.returncode != 0:
        print(f"\n[错误] trace_processor_shell 退出码: {result.returncode}")
        if stderr:
            print(f"[STDERR] {stderr[:2000]}")
        sys.exit(1)

    # ---- 解析 CSV 输出 ----
    print("\n[2/3] 解析查询结果 ...")
    reader = csv.reader(io.StringIO(stdout))
    all_rows = list(reader)
    if not all_rows:
        print("[警告] 查询无结果")
        sys.exit(0)

    # ---- 按 seg 列拆分并写入 ----

    do_csv = args.format in ("csv", "both")
    do_xlsx = args.format in ("xlsx", "both")

    os.makedirs(out_dir, exist_ok=True)

    def _write_csv_segment(seg_rows, out_path):
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(COLUMN_NAMES)
            for row in seg_rows:
                writer.writerow(build_row_values(_DATA_HEADERS, row))

    grand_total = 0
    for i in range(total_segments):
        seg_rows = [r[1:] for r in all_rows[1:] if r and len(r) > 1 and r[0] == str(i)]
        s, e = time_ranges[i]

        if not seg_rows:
            print(f"  分段 {i+1}: 0 条记录 (时间范围无匹配数据)")
            continue

        base_name = f"{trace_name}_{s}_{e}_block_io"

        # ---- CSV (快速) ----
        if do_csv:
            csv_path = os.path.join(out_dir, base_name + ".csv")
            _write_csv_segment(seg_rows, csv_path)
            csv_sz = os.path.getsize(csv_path)
            print(f"  分段 {i+1}: {len(seg_rows):>8} 行  CSV  {csv_sz:>12,} b  {base_name}.csv")

        # ---- XLSX (慢, 带格式) ----
        if do_xlsx:
            xlsx_path = os.path.join(out_dir, base_name + ".xlsx")
            wb = Workbook()
            write_sheets(wb, "全部 IO", _DATA_HEADERS, seg_rows)
            wb.save(xlsx_path)
            xlsx_sz = os.path.getsize(xlsx_path)
            indent = "         " if do_csv else "  "
            print(f"  {indent}{len(seg_rows):>8} 行  XLSX {xlsx_sz:>12,} b  {base_name}.xlsx")

        grand_total += len(seg_rows)

    print()
    print("=" * 60)
    print(f"  完成! {total_segments} 个分段, 共 {grand_total} 条记录")
    print(f"  {out_dir}")
    print("=" * 60)


if __name__ == "__main__":
    main()