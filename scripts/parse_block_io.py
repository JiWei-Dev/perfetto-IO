import os
import sys
import csv
import argparse
import subprocess
import glob

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config_loader import load_config


def get_project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def run_sql(trace_processor, trace_file, sql, timeout=300):
    result = subprocess.run(
        [trace_processor, "query", trace_file, sql],
        capture_output=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"SQL query failed (exit {result.returncode}): {stderr}")
    return result.stdout.decode("utf-8", errors="replace")


def parse_csv_output(text):
    reader = csv.reader(text.strip().split("\n"), delimiter=",")
    rows = list(reader)
    if not rows:
        return [], []
    return rows[0], rows[1:]


def build_sql():
    return """
SELECT
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
GROUP BY e.id
ORDER BY e.ts
"""


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


def clean_null(val):
    if val is None or val == "[NULL]":
        return ""
    return val


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


def ts_to_seconds(ts_ns):
    if ts_ns is None or ts_ns == "":
        return ""
    try:
        ts = int(ts_ns)
    except (ValueError, TypeError):
        return str(ts_ns)
    return f"{ts / 1_000_000_000:.9f}"


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


MAX_EXCEL_ROWS = 1_048_576
MAX_DATA_ROWS = MAX_EXCEL_ROWS - 1


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
        print(f"    已写入 Sheet「{base_title}」({len(rows)} 行)")
    else:
        print(f"    数据量 {len(rows)} 行，超过单个 Sheet 上限 ({MAX_DATA_ROWS} 行)，将自动拆分...")
        for part_no, start in enumerate(range(0, len(rows), MAX_DATA_ROWS), 1):
            chunk = rows[start:start + MAX_DATA_ROWS]
            title = f"{base_title} ({part_no})"
            ws = wb.create_sheet(title=title)
            _write_one_sheet(ws, headers, chunk)
            print(f"    已写入 Sheet「{title}」({len(chunk)} 行, 数据行 {start + 1}-{start + len(chunk)})")
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    sheet_count = (len(rows) + MAX_DATA_ROWS - 1) // MAX_DATA_ROWS
    return sheet_count


def parse_args():
    parser = argparse.ArgumentParser(description="解析 Perfetto trace 中的 Block IO 事件")
    parser.add_argument(
        "--start-ns", type=float, default=None,
        help="时间范围起始时间（纳秒），例如 237863509246000"
    )
    parser.add_argument(
        "--end-ns", type=float, default=None,
        help="时间范围结束时间（纳秒），例如 237870000743000"
    )
    parser.add_argument(
        "--start", type=float, default=None,
        help="时间范围起始时间（秒），例如 237863.5"
    )
    parser.add_argument(
        "--end", type=float, default=None,
        help="时间范围结束时间（秒），例如 237870.0"
    )
    parser.add_argument(
        "--trace", type=str, default=None,
        help="指定单个 trace 文件名（如 xxx.perfetto-trace 或 xxx.perfetto），不指定则解析 trace/ 目录下全部"
    )
    parser.add_argument(
        "--timeout", type=int, default=300,
        help="trace_processor_shell SQL 查询超时秒数（默认 300）"
    )
    return parser.parse_args()


def main():
    args = parse_args()

    cfg = load_config()
    block_cfg = cfg.get("parse_block_io", {})

    if args.start_ns is not None:
        args.start = args.start_ns / 1_000_000_000.0
    elif args.start is None:
        start_ns_cfg = block_cfg.get("start_ns")
        if start_ns_cfg is not None:
            args.start = start_ns_cfg / 1_000_000_000.0
        else:
            args.start = block_cfg.get("start")

    if args.end_ns is not None:
        args.end = args.end_ns / 1_000_000_000.0
    elif args.end is None:
        end_ns_cfg = block_cfg.get("end_ns")
        if end_ns_cfg is not None:
            args.end = end_ns_cfg / 1_000_000_000.0
        else:
            args.end = block_cfg.get("end")

    if (args.start is None) != (args.end is None):
        print("[错误] --start 和 --end 必须同时指定")
        sys.exit(1)

    if args.start is not None and args.start >= args.end:
        print("[错误] --start 必须小于 --end")
        sys.exit(1)

    start_ns = int(args.start * 1_000_000_000) if args.start is not None else None
    end_ns = int(args.end * 1_000_000_000) if args.end is not None else None

    project_root = get_project_root()
    trace_dir = os.path.join(project_root, "trace")
    out_dir = os.path.join(project_root, "out")
    trace_processor_shell = os.path.join(
        project_root, "bin", "windows-amd64", "trace_processor_shell.exe"
    )

    if not os.path.isfile(trace_processor_shell):
        print(f"[错误] trace_processor_shell.exe 不存在: {trace_processor_shell}")
        sys.exit(1)

    trace_files = glob.glob(os.path.join(trace_dir, "*.perfetto-trace"))
    trace_files += glob.glob(os.path.join(trace_dir, "*.perfetto"))
    trace_files = sorted(set(trace_files))
    if not trace_files:
        print(f"[提示] trace 目录下未找到 .perfetto-trace 或 .perfetto 文件: {trace_dir}")
        sys.exit(0)

    if args.trace:
        target = os.path.basename(args.trace)
        matched = [f for f in trace_files if os.path.basename(f) == target]
        if not matched:
            print(f"[错误] 未找到指定的 trace 文件: {args.trace}")
            print(f"  可用文件: {', '.join(os.path.basename(f) for f in trace_files)}")
            sys.exit(1)
        trace_files = matched
        print(f"[选择] 仅解析: {target}")

    os.makedirs(out_dir, exist_ok=True)

    for trace_file in trace_files:
        trace_name = os.path.splitext(os.path.basename(trace_file))[0]
        if trace_name.endswith(".perfetto"):
            trace_name = trace_name[:-len(".perfetto")]

        trace_out_dir = os.path.join(out_dir, trace_name)
        os.makedirs(trace_out_dir, exist_ok=True)

        xlsx_output = os.path.join(trace_out_dir, f"{trace_name}_block_io.xlsx")

        print(f"[解析] {os.path.basename(trace_file)}")
        print(f"  执行 SQL 查询...")

        sql = build_sql()
        try:
            raw_output = run_sql(trace_processor_shell, trace_file, sql, timeout=args.timeout)
        except RuntimeError as e:
            print(f"[错误] {e}")
            continue

        headers, all_rows = parse_csv_output(raw_output)
        if not all_rows:
            print(f"[警告] 未查询到 block IO 事件")
            continue

        print(f"  查询到 {len(all_rows)} 条 block IO 记录")

        wb = Workbook()

        write_sheets(wb, "全部 IO", headers, all_rows)

        if start_ns is not None:
            filtered_rows = []
            for row in all_rows:
                row_dict = dict(zip(headers, row))
                ts_str = row_dict.get("ts", "")
                try:
                    ts_val = int(ts_str)
                except (ValueError, TypeError):
                    continue
                if start_ns <= ts_val <= end_ns:
                    filtered_rows.append(row)

            if filtered_rows:
                write_sheets(wb, "时间范围 IO", headers, filtered_rows)
                print(f"  时间范围 [{args.start:.3f}s, {args.end:.3f}s] 内匹配 {len(filtered_rows)} 条记录")
            else:
                ws_empty = wb.create_sheet(title="时间范围 IO")
                ws_empty.cell(row=1, column=1, value=f"时间范围 [{args.start:.3f}s, {args.end:.3f}s] 内无匹配记录")
                print(f"  时间范围内无匹配记录")

        wb.save(xlsx_output)
        print(f"[完成] XLSX: {xlsx_output}")

    print(f"\n全部解析完成。")


if __name__ == "__main__":
    main()