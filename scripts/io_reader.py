import csv
import os
import glob
from collections import defaultdict, deque
from zipfile import BadZipFile

from openpyxl import load_workbook

# 同一 request 三段事件 (insert -> issue -> complete) 之间允许的最大间隔。
# 超过此值视为过期/异常事件, 不参与配对, 避免跨请求错配。
MAX_PAIR_GAP_NS = 1_000_000_000


def load_io_events(file_path, sheet_name="全部 IO"):
    """读取 block IO 数据文件, 自动识别 CSV / XLSX 格式。

    Args:
        file_path: 文件路径 (.csv 或 .xlsx)
        sheet_name: XLSX 中的工作表名 (CSV 忽略此参数)

    Returns:
        (headers, events): headers 是列名 list, events 是 list[dict]
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        return _load_csv(file_path)
    else:
        return _load_xlsx(file_path, sheet_name)


def _load_csv(csv_path):
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames
        events = [row for row in reader]
    return headers, events


def get_sheet_names(file_path):
    """读取 XLSX 文件的所有 Sheet 名称（仅读元数据，不加载数据行）。CSV 返回 None。"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        return None
    try:
        wb = load_workbook(file_path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception as e:
        print(f"  [警告] 读取 Sheet 列表失败: {e}")
        return []


def _choose_sheet_from_names(names, default=None):
    """交互式选择 Sheet：列出所有 Sheet 名称，用户输入编号选择。

    Args:
        names: Sheet 名称列表
        default: 默认 Sheet 名（回车直接选它，若有）

    Returns:
        选中的 Sheet 名；无法交互（无 TTY / 无 Sheet）时返回 default 或第一个。
    """
    if not names:
        return default or "全部 IO"
    if len(names) == 1:
        print(f"  仅一个 Sheet「{names[0]}」，自动选择")
        return names[0]

    print("  该文件包含以下 Sheet：")
    for i, n in enumerate(names, 1):
        hint = "  ← 默认" if (default and n == default) else ""
        print(f"    [{i}] {n}{hint}")

    default_idx = names.index(default) + 1 if default in names else 1
    try:
        raw = input(f"  请选择 Sheet 编号 (回车默认 [{default_idx}] {names[default_idx - 1]}): ").strip()
    except (EOFError, KeyboardInterrupt):
        raw = ""
    if raw == "":
        return names[default_idx - 1]
    try:
        idx = int(raw)
    except ValueError:
        print(f"  无效输入，使用默认: {names[default_idx - 1]}")
        return names[default_idx - 1]
    if 1 <= idx <= len(names):
        return names[idx - 1]
    print(f"  编号超出范围，使用默认: {names[default_idx - 1]}")
    return names[default_idx - 1]


def _load_xlsx(xlsx_path, sheet_name):
    try:
        wb = load_workbook(xlsx_path, read_only=True)
    except BadZipFile:
        print(f"  [警告] {xlsx_path} 不是有效的 XLSX 文件 (可能损坏), 跳过")
        return None, []
    except Exception as e:
        print(f"  [警告] 读取 {xlsx_path} 失败: {e}, 跳过")
        return None, []

    # 未指定或显式请求选择 (--sheet "?" / "select") 时，列出所有 Sheet 让用户选编号
    if sheet_name in (None, "", "?", "select"):
        sheet_name = _choose_sheet_from_names(wb.sheetnames, default=None)

    # 指定/选中的 Sheet 不存在时，回退到交互选择
    if sheet_name not in wb.sheetnames:
        print(f"  [提示] Sheet「{sheet_name}」不存在，请重新选择：")
        sheet_name = _choose_sheet_from_names(wb.sheetnames, default=None)

    if sheet_name not in wb.sheetnames:
        available = ", ".join(wb.sheetnames)
        wb.close()
        print(f"  [警告] Sheet「{sheet_name}」不存在, 可用: {available}")
        return None, []

    ws = wb[sheet_name]
    headers = None
    events = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            headers = [str(c) if c is not None else "" for c in row]
            continue
        row_dict = dict(zip(headers, row))
        events.append(row_dict)
    wb.close()
    return headers, events


def find_block_io_file(trace_dir, trace_name):
    """在 trace_dir 中查找 block IO 文件, 优先 XLSX 其次 CSV。

    Returns:
        (file_path, label) 或 (None, None)
    """
    for ext in (".xlsx", ".csv"):
        path = os.path.join(trace_dir, f"{trace_name}_block_io{ext}")
        if os.path.isfile(path):
            return path, ext
    return None, None


def find_block_io_segments(segments_dir, trace_name):
    """在 segments_dir 中查找所有切分后的 block IO 文件。

    Returns:
        [(file_path, label), ...]  按文件名排序
    """
    results = []
    for ext in (".xlsx", ".csv"):
        pattern = os.path.join(segments_dir, f"{trace_name}_*_block_io{ext}")
        for f in sorted(glob.glob(pattern)):
            results.append((f, ext))
    return results


def _norm_key(val):
    """归一化配对键分量: None/空值 -> 空串, 其余转字符串并去首尾空格。"""
    if val is None or val == "":
        return ""
    return str(val).strip()


def _ev_ts_ns(ev):
    raw = ev.get("时间 (纳秒)", 0)
    try:
        return int(raw) if raw not in (None, "") else 0
    except (ValueError, TypeError):
        return 0


def match_io_pairs(events, time_window_ns=MAX_PAIR_GAP_NS):
    """把 block_rq_insert / block_rq_issue / block_rq_complete 三段事件配成 request 三元组。

    配对策略:
      1. 分组键升级为 (dev, sector, nr_sector): 旧实现只用 (dev, sector), 同一扇区
         被读/写重复访问时会跨 request 串扰, 加入 nr_sector 可大幅降低碰撞概率;
      2. 组内按时间排序, 用 FIFO 队列做生命周期状态机 (insert -> issue -> complete),
         比旧实现的单变量覆盖更能应对并发交错, 且符合块层"同一设备扇区的 request
         按 insert 顺序 issue"的物理语义;
      3. 时间窗约束: 相邻两段间隔超过 time_window_ns (默认 1s) 视为过期/异常, 丢弃。

    Returns:
        list[dict], 每项:
          dev, sector, nr_sector,
          insert_ev, issue_ev, complete_ev,      # 原始事件 dict
          insert_ts, issue_ts, complete_ts,       # 纳秒
          schedule_ns, device_ns, end_to_end_ns   # 三段耗时
    """
    groups = defaultdict(list)
    for ev in events:
        dev = _norm_key(ev.get("设备号", ""))
        sector = _norm_key(ev.get("请求扇区地址", ""))
        nr_sector = _norm_key(ev.get("请求扇区数量", ""))
        if not dev or not sector:
            continue
        groups[(dev, sector, nr_sector)].append(ev)

    results = []
    for key, group in groups.items():
        group.sort(key=_ev_ts_ns)
        pending_inserts = deque()  # (insert_ev, insert_ts)
        pending_issues = deque()   # (issue_ev, issue_ts, insert_ev, insert_ts)

        for ev in group:
            name = str(ev.get("Trace 字段名称", "")).strip()
            ts = _ev_ts_ns(ev)

            if name == "block_rq_insert":
                pending_inserts.append((ev, ts))

            elif name == "block_rq_issue":
                # 取最早的、且仍在时间窗内的 insert (FIFO)
                while pending_inserts:
                    insert_ev, insert_ts = pending_inserts[0]
                    gap = ts - insert_ts
                    if gap < 0 or gap > time_window_ns:
                        pending_inserts.popleft()  # 时间倒退或过期, 丢弃
                        continue
                    pending_issues.append((ev, ts, insert_ev, insert_ts))
                    pending_inserts.popleft()
                    break
                # 无可用 insert: 该 issue 为孤儿, 丢弃

            elif name == "block_rq_complete":
                while pending_issues:
                    issue_ev, issue_ts, insert_ev, insert_ts = pending_issues[0]
                    gap = ts - issue_ts
                    if gap < 0 or gap > time_window_ns:
                        pending_issues.popleft()  # 时间倒退或过期, 丢弃
                        continue
                    results.append({
                        "dev": key[0],
                        "sector": key[1],
                        "nr_sector": key[2],
                        "insert_ev": insert_ev,
                        "issue_ev": issue_ev,
                        "complete_ev": ev,
                        "insert_ts": insert_ts,
                        "issue_ts": issue_ts,
                        "complete_ts": ts,
                        "schedule_ns": issue_ts - insert_ts,
                        "device_ns": ts - issue_ts,
                        "end_to_end_ns": ts - insert_ts,
                    })
                    pending_issues.popleft()
                    break

    return results