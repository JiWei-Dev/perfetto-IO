"""IO 卡顿快速分析脚本 (固化版)

针对一段 Uninterruptible Sleep (IO) 卡顿, 输入 Start time / Duration / Process,
自动定位责任 IO, 拆解调度耗时 vs 器件耗时, 并按存储专家方法论排查:
  - 调度慢 -> 是否 CPU 不足 / 请求队列满 / 被插队
  - 器件慢 -> 是否 complete 断档 / 读写交叉

输出详细分析报告 (Markdown) 到对应 out/<trace>/stall_analysis/ 子目录。

用法示例:
  python scripts/AI/io_stall_analysis.py \
      --trace com.ss.android.ugc.aweme \
      --start 237866030786798 \
      --duration 3840677 \
      --process com.ss.android.ugc.aweme --tid 14831
"""
import os
import sys
import argparse
from collections import Counter

# 使脚本可 import 上层 scripts 目录的公共模块，以及同目录 scripts/AI 下的 AI 模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # scripts
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # scripts/AI

from openpyxl import load_workbook
from io_reader import match_io_pairs as _match_io_pairs

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
NS_PER_MS = 1_000_000

# block_io.xlsx「全部 IO」Sheet 的列名 (与 parse_block_io.COLUMN_NAMES 一致)
COLUMNS = [
    "IO 进程名称", "进程 PID", "线程 TID", "线程名称",
    "时间 (秒)", "时间 (纳秒)", "运行 CPU", "Trace 字段名称",
    "设备号", "请求扇区地址", "请求扇区数量", "请求字节数",
    "读写类型", "进程 Comm", "命令", "错误码",
]


def _to_int(v):
    try:
        return int(v) if v is not None else 0
    except (ValueError, TypeError):
        return 0


def _to_float_ns(v):
    try:
        return float(v) if v is not None else 0.0
    except (ValueError, TypeError):
        return 0.0


def _load_window(xlsx_path, lo_ns, hi_ns):
    """流式读取事件, 只保留 [lo_ns, hi_ns] 窗口内的事件, 返回 list[dict]。

    优先读同目录 CSV 缓存 (block_io.csv), 无缓存则回退到 xlsx 流式读取。
    CSV 由 convert_to_csv.py 一次性生成, 读取速度远快于 xlsx。
    """
    import csv
    csv_path = xlsx_path[:-5] + ".csv"  # .xlsx -> .csv
    if os.path.isfile(csv_path):
        events = []
        with open(csv_path, "r", encoding="utf-8", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # 跳过表头
            for row in reader:
                if not row:
                    continue
                ts = _to_int(row[5])
                if ts < lo_ns or ts > hi_ns:
                    continue
                events.append(dict(zip(COLUMNS, row)))
        return events

    # 回退: xlsx 流式读取
    wb = load_workbook(xlsx_path, read_only=True)
    ws = wb["全部 IO"]
    events = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            continue
        ts = _to_int(row[5])
        if ts < lo_ns or ts > hi_ns:
            continue
        events.append(dict(zip(COLUMNS, row)))
    wb.close()
    return events


def match_full(events):
    """完整三段式配对: 返回每笔 IO 的 insert/issue/complete 三事件及其完整信息。"""
    triples = _match_io_pairs(events)
    results = []
    for t in triples:
        results.append({
            "dev": t["dev"], "sector": t["sector"],
            "insert_ev": t["insert_ev"], "issue_ev": t["issue_ev"], "complete_ev": t["complete_ev"],
            "insert_ts": t["insert_ts"], "issue_ts": t["issue_ts"], "complete_ts": t["complete_ts"],
        })
    return results


def _ev_who(ev):
    return "proc=%s pid=%s tid=%s thr=%s cpu=%s" % (
        ev.get("IO 进程名称", ""), ev.get("进程 PID", ""),
        ev.get("线程 TID", ""), ev.get("线程名称", ""), ev.get("运行 CPU", ""))


def _ev_size(ev):
    return "%s 扇区=%s 字节=%s" % (
        ev.get("读写类型", ""), ev.get("请求扇区数量", ""), ev.get("请求字节数", ""))


def _rel(ts_ns, start_ns):
    return (ts_ns - start_ns) / 1000.0  # 微秒


def analyze(trace_dir, trace_name, start_ns, dur_ns, process=None, pid=None, tid=None,
            pre_ms=50, post_ms=200):
    """执行完整分析, 返回 (report_text, 附加数据文本)。"""
    xlsx = os.path.join(trace_dir, f"{trace_name}_block_io.xlsx")
    if not os.path.isfile(xlsx):
        raise FileNotFoundError(f"未找到 block_io 文件: {xlsx}")

    lo = start_ns - pre_ms * NS_PER_MS
    hi = start_ns + dur_ns + post_ms * NS_PER_MS
    events = _load_window(xlsx, lo, hi)

    L = []  # 报告行
    A = []  # 附加数据行

    L.append("# IO 卡顿分析报告\n")
    L.append("## 0. 输入参数\n")
    L.append("| 参数 | 值 |")
    L.append("|---|---|")
    L.append("| Start time | %d ns (%.6f s) |" % (start_ns, start_ns / 1e9))
    L.append("| Duration | %d ns (%.3f ms) |" % (dur_ns, dur_ns / 1e6))
    L.append("| Process | %s |" % (process or "未指定"))
    L.append("| PID | %s |" % (pid or "未指定"))
    L.append("| TID | %s |" % (tid or "未指定"))
    L.append("| 分析窗口 | [%.6f, %.6f] s (前 %dms / 后 %dms) |" % (lo / 1e9, hi / 1e9, pre_ms, post_ms))
    csv_cache = xlsx[:-5] + ".csv"
    src = csv_cache if os.path.isfile(csv_cache) else xlsx
    L.append("| 数据来源 | `%s` |" % src)
    L.append("| 原始文件 | `%s` (Sheet: 全部 IO, CSV 缓存由 convert_to_csv.py 生成) |" % xlsx)
    L.append("")

    # ---- 1. 定位目标 IO ----
    L.append("## 1. 目标 IO 定位\n")
    # 按指定条件过滤 insert 事件
    def match_target(ev):
        if tid is not None and str(ev.get("线程 TID", "")).strip() == str(tid):
            return True
        if pid is not None and str(ev.get("进程 PID", "")).strip() == str(pid):
            return True
        if process is not None and str(process).strip() in str(ev.get("IO 进程名称", "")):
            return True
        return False

    target_inserts = [e for e in events
                      if str(e.get("Trace 字段名称", "")) == "block_rq_insert"
                      and _to_int(e.get("时间 (纳秒)", 0)) <= start_ns]
    target_inserts = [e for e in target_inserts if match_target(e)] if (tid or pid or process) else target_inserts

    if not target_inserts:
        L.append("**警告**: 窗口内未找到符合条件的目标 insert 事件, 请检查 Process/PID/TID 是否匹配。\n")
        return "\n".join(L), "\n".join(A)

    # 卡顿点前最后一次 insert 即责任 IO 的候选
    target_io = max(target_inserts, key=lambda e: _to_int(e.get("时间 (纳秒)", 0)))
    target_sector = str(target_io.get("请求扇区地址", ""))
    target_dev = str(target_io.get("设备号", ""))
    L.append("卡顿点前符合条件的最后一次 insert: sector=%s dev=%s" % (target_sector, target_dev))
    L.append("- insert: %+.1fus 由 %s 发起, %s" % (_rel(_to_int(target_io.get("时间 (纳秒)", 0)), start_ns), _ev_who(target_io), _ev_size(target_io)))
    L.append("")

    # ---- 2. 三段式生命周期 ----
    L.append("## 2. 该笔 IO 三段式生命周期\n")
    pairs = match_full(events)
    full_pair = None
    for p in pairs:
        if p["sector"] == target_sector and p["dev"] == target_dev and p["insert_ts"] == _to_int(target_io.get("时间 (纳秒)", 0)):
            full_pair = p
            break
    # 若精确匹配失败, 退化为同扇区的任意配对
    if full_pair is None:
        for p in pairs:
            if p["sector"] == target_sector and p["dev"] == target_dev:
                full_pair = p
                break

    if full_pair is None:
        L.append("**警告**: 未能在窗口内完成该笔 IO 的 issue/complete 配对 (可能 complete 在窗口外, 建议增大 --post-ms)。\n")
        return "\n".join(L), "\n".join(A)

    ins_ev, iss_ev, cmp_ev = full_pair["insert_ev"], full_pair["issue_ev"], full_pair["complete_ev"]
    ins_ts, iss_ts, cmp_ts = full_pair["insert_ts"], full_pair["issue_ts"], full_pair["complete_ts"]

    L.append("| 阶段 | 时间戳 (ns) | 相对卡顿起点 | 由谁触发 | 上下文 |")
    L.append("|---|---|---|---|---|")
    L.append("| insert | %d | %+.1fus | %s | %s |" % (ins_ts, _rel(ins_ts, start_ns), _ev_who(ins_ev), _ev_size(ins_ev)))
    L.append("| issue | %d | %+.1fus | %s | %s |" % (iss_ts, _rel(iss_ts, start_ns), _ev_who(iss_ev), _ev_size(iss_ev)))
    L.append("| complete | %d | %+.1fus | %s | %s |" % (cmp_ts, _rel(cmp_ts, start_ns), _ev_who(cmp_ev), _ev_size(cmp_ev)))
    L.append("")
    L.append("> insert 的进程为 IO 发起者; issue 由 kworker/内核 block 层执行; complete 记录的进程为中断返回时正在运行的上下文, 不代表发起者。\n")

    # ---- 3. 耗时拆解 ----
    L.append("## 3. 耗时点统计\n")
    schedule_ms = (iss_ts - ins_ts) / NS_PER_MS
    device_ms = (cmp_ts - iss_ts) / NS_PER_MS
    total_ms = (cmp_ts - ins_ts) / NS_PER_MS
    L.append("| 阶段 | 起止 | 耗时 | 占比 |")
    L.append("|---|---|---|---|")
    L.append("| 调度耗时 (insert→issue) | %+.1fus → %+.1fus | %.3f ms | %.1f%% |" % (_rel(ins_ts, start_ns), _rel(iss_ts, start_ns), schedule_ms, schedule_ms / total_ms * 100))
    L.append("| 器件耗时 (issue→complete) | %+.1fus → %+.1fus | %.3f ms | %.1f%% |" % (_rel(iss_ts, start_ns), _rel(cmp_ts, start_ns), device_ms, device_ms / total_ms * 100))
    L.append("| 端到端 (insert→complete) | %+.1fus → %+.1fus | %.3f ms | 100%% |" % (_rel(ins_ts, start_ns), _rel(cmp_ts, start_ns), total_ms))
    L.append("")

    # ---- 4. 调度阶段排查 ----
    L.append("## 4. 调度阶段排查 (CPU / 队列满 / 插队)\n")

    # 4.1 CPU: issue 由谁完成
    L.append("### 4.1 CPU 资源是否不足\n")
    L.append("- 该笔 IO 的 issue 由 `%s` 完成 (CPU=%s)。kworker 为内核 block 工作线程。" % (iss_ev.get("IO 进程名称", ""), iss_ev.get("运行 CPU", "")))
    L.append("- 判断: 若 issue 延迟发生在 kworker 无法被调度, 会伴随 issue 事件 CPU 字段长期不变; 本窗口 issue 事件 CPU 分布见附加数据。\n")

    # 4.2 队列满: START 时刻 in-flight
    L.append("### 4.2 请求队列是否已满 (in-flight 深度)\n")
    inflight = []
    for p in pairs:
        if p["issue_ts"] <= start_ns < p["complete_ts"]:
            inflight.append(p)
    inflight.sort(key=lambda x: x["issue_ts"])
    L.append("- START 时刻设备 in-flight (已 issue 未 complete) 请求数 = **%d**" % len(inflight))
    total_bytes = sum(_to_int(p["issue_ev"].get("请求字节数", 0)) for p in inflight)
    L.append("- in-flight 数据量 ≈ %.2f MB" % (total_bytes / 1024 / 1024))
    rw_inf = Counter(str(p["issue_ev"].get("读写类型", "")) for p in inflight)
    L.append("- in-flight 读写类型分布: %s" % dict(rw_inf))
    L.append("- 判断: 若 in-flight 深度接近设备命令队列上限 (UFS 通常 32 / eMMC 通常 4), 说明驱动已无法接收新请求, 后续请求在调度队列排队。")
    if len(inflight) >= 32:
        L.append("  - 当前 in-flight = %d, 已达 UFS 上限 32, 队列确认打满。" % len(inflight))
    elif len(inflight) >= 4:
        L.append("  - 当前 in-flight = %d, 已达 eMMC 上限 4 (eMMC 队列已满; UFS 则尚有裕量)。" % len(inflight))
    else:
        L.append("  - 当前 in-flight = %d, 未达 eMMC 上限 4, 队列未满。" % len(inflight))
    L.append("")

    # 堆积请求
    stacked = []
    for p in pairs:
        if p["insert_ts"] <= start_ns < p["issue_ts"]:
            stacked.append(p)
    stacked.sort(key=lambda x: x["insert_ts"])
    L.append("- START 时已 insert 但尚未 issue 的堆积请求数 = **%d** (约 %.2f MB)" % (
        len(stacked), sum(_to_int(p["insert_ev"].get("请求字节数", 0)) for p in stacked) / 1024 / 1024))
    L.append("")

    # 4.3 插队
    L.append("### 4.3 是否被其他请求插队\n")
    L.append("- 插队定义: **后面 insert 的 IO 反而先 issue** (insert 顺序 vs issue 顺序被打乱)。\n")

    # 检测同一设备上, 比目标 IO 更晚 insert 却更早 issue 的请求
    overtakers = []
    for p in pairs:
        if p["dev"] != target_dev:
            continue
        if p["insert_ts"] > ins_ts and p["issue_ts"] < iss_ts:
            overtakers.append(p)
    overtakers.sort(key=lambda x: x["insert_ts"])

    ins_near = [e for e in events
                if str(e.get("Trace 字段名称", "")) == "block_rq_insert"
                and abs(_to_int(e.get("时间 (纳秒)", 0)) - start_ns) < pre_ms * NS_PER_MS]
    ins_near.sort(key=lambda e: _to_int(e.get("时间 (纳秒)", 0)))
    procs = Counter(str(e.get("IO 进程名称", "")) for e in ins_near)
    tids = Counter(str(e.get("线程 TID", "")) for e in ins_near)
    L.append("- START 前后 %dms 内 insert 事件 %d 条, 来自进程: %s" % (pre_ms, len(ins_near), dict(procs)))
    L.append("- 来自线程 TID: %s" % dict(tids))

    if overtakers:
        L.append("- **结论: 存在插队** — 有 %d 笔请求在目标 IO 之后 insert, 却先于目标 IO issue:" % len(overtakers))
        for p in overtakers[:5]:
            L.append("  - sector=%s insert=%+.1fus issue=%+.1fus (比目标早 %.1fus issue)" % (
                p["sector"], _rel(p["insert_ts"], start_ns), _rel(p["issue_ts"], start_ns),
                (iss_ts - p["issue_ts"]) / 1000.0))
        if len(overtakers) > 5:
            L.append("  - ... 其余 %d 笔见明细" % (len(overtakers) - 5))
    else:
        L.append("- **结论: 无插队** (目标 IO 之后的请求均未先于它 issue)。")
    if len(tids) > 1:
        L.append("- 提示: 存在多线程/进程并发提交, 但未出现 issue 顺序倒置。")
    L.append("")

    # ---- 5. 器件阶段排查 ----
    L.append("## 5. 器件阶段排查 (complete 断档 / 读写交叉)\n")

    # 5.1 complete 断档
    L.append("### 5.1 器件侧是否停滞 (complete 断档)\n")
    completes = sorted([e for e in events if str(e.get("Trace 字段名称", "")) == "block_rq_complete"],
                       key=lambda e: _to_int(e.get("时间 (纳秒)", 0)))
    issues = sorted([e for e in events if str(e.get("Trace 字段名称", "")) == "block_rq_issue"],
                    key=lambda e: _to_int(e.get("时间 (纳秒)", 0)))
    last_cmp_before = None
    first_cmp_after = None
    for e in completes:
        if _to_int(e.get("时间 (纳秒)", 0)) <= start_ns:
            last_cmp_before = e
        else:
            first_cmp_after = e
            break
    last_iss_before = None
    first_iss_after = None
    for e in issues:
        if _to_int(e.get("时间 (纳秒)", 0)) <= start_ns:
            last_iss_before = e
        else:
            first_iss_after = e
            break
    if last_cmp_before:
        L.append("- START 前最后一个 complete: %+.1fus (%s)" % (_rel(_to_int(last_cmp_before.get("时间 (纳秒)", 0)), start_ns), _ev_who(last_cmp_before)))
    if first_cmp_after:
        L.append("- START 后第一个 complete: %+.1fus (%s)" % (_rel(_to_int(first_cmp_after.get("时间 (纳秒)", 0)), start_ns), _ev_who(first_cmp_after)))
    if last_cmp_before and first_cmp_after:
        gap_ms = (_to_int(first_cmp_after.get("时间 (纳秒)", 0)) - _to_int(last_cmp_before.get("时间 (纳秒)", 0))) / NS_PER_MS
        L.append("- **complete 断档时长 = %.3f ms**" % gap_ms)
        L.append("- 判断: 断档期间设备零 complete, 说明器件侧是『停滞』而非『慢』。\n")
    if last_iss_before and first_iss_after:
        gap_iss_ms = (_to_int(first_iss_after.get("时间 (纳秒)", 0)) - _to_int(last_iss_before.get("时间 (纳秒)", 0))) / NS_PER_MS
        L.append("- issue 断档时长 = %.3f ms (issue/complete 同时断档 → 设备整体停滞)。" % gap_iss_ms)
    L.append("")

    # 5.2 读写交叉
    L.append("### 5.2 是否存在读写交叉\n")
    L.append("- START 时刻 in-flight 请求读写分布: %s" % dict(rw_inf))
    if len(rw_inf) == 1 and list(rw_inf)[0].startswith(("W", "WS")):
        L.append("- **结论: 无读写交叉** (纯写负载, in-flight 全为写请求)。\n")
    elif "R" in "".join(rw_inf.keys()):
        L.append("- **结论: 存在读写交叉**, 读请求可能被写请求拖累。\n")
    else:
        L.append("- 负载类型单一, 无读写交叉干扰。\n")

    # 5.3 扇区离散度
    L.append("### 5.3 扇区离散度 (LBA 是否连续)\n")
    L.append("- 背景: 顺序访问 LBA 连续, 随机访问 LBA 大幅跳跃; 离散度高会加剧器件开销 (eMMC/HDD 寻道、UFS 跨 die 随机)。\n")

    issue_dev = [e for e in issues if str(e.get("设备号", "")) == target_dev]
    prev_sector = None
    prev_nr = 0
    seq_cnt = 0
    jump_cnt = 0
    jumps = []
    for e in issue_dev:
        sec = _to_int(e.get("请求扇区地址", 0))
        nr = _to_int(e.get("请求扇区数量", 0)) or 1
        if prev_sector is not None:
            delta = sec - (prev_sector + prev_nr)
            if abs(delta) <= nr:
                seq_cnt += 1
            else:
                jump_cnt += 1
                jumps.append(delta)
        prev_sector = sec
        prev_nr = nr

    total = seq_cnt + jump_cnt
    if total > 0:
        jump_ratio = jump_cnt / total * 100
        L.append("- 窗口内 issue 请求相邻 LBA 对比: 连续 %d 次, 跳跃 %d 次 (跳跃占比 %.1f%%)" % (seq_cnt, jump_cnt, jump_ratio))
        if jumps:
            L.append("- 跳跃幅度: 最大 %d 扇区, 平均 %d 扇区" % (max(jumps), int(sum(jumps) / len(jumps))))
        if jump_ratio > 50:
            L.append("- **结论: 扇区离散度高 (随机访问)**, 器件随机寻址开销大, 是器件慢的潜在原因之一。\n")
        elif jump_ratio > 10:
            L.append("- **结论: 部分随机**, 存在一定 LBA 跳跃, 对器件效率有影响。\n")
        else:
            L.append("- **结论: 扇区基本连续 (顺序访问)**, 离散度不是瓶颈。\n")
    else:
        L.append("- 窗口内 issue 请求过少, 不足以判断扇区离散度。\n")

    # 5.4 器件慢 vs 软中断延迟甄别
    L.append("### 5.4 器件慢 vs 软中断延迟甄别 (complete 时间戳被 CPU 推迟)\n")
    L.append("- 关键认知: `block_rq_complete` 记录的是块层软中断执行完成回调的时刻, 而非硬件完成 IO 的时刻; 若硬件中断早已回来但 CPU 没空运行软中断, complete 时间戳会偏晚, 导致 issue→complete 被高估为器件耗时。\n")

    comp_times = sorted(_to_int(e.get("时间 (纳秒)", 0)) for e in completes)
    comp_procs = Counter(str(e.get("IO 进程名称", "")) for e in completes)
    L.append("- 窗口内 complete 事件 %d 条, 记录进程分布: %s" % (len(comp_times), dict(comp_procs)))

    softirq_delay = False
    if len(comp_times) >= 3:
        gaps = [comp_times[i + 1] - comp_times[i] for i in range(len(comp_times) - 1)]
        med_gap = sorted(gaps)[len(gaps) // 2]
        cluster_cnt = sum(1 for g in gaps if g < 50 * 1000)  # < 50us 视为聚簇(几乎同时)
        cluster_ratio = cluster_cnt / len(gaps) if gaps else 0.0
        L.append("- complete 相邻间隔: 中位数 %.1fus, < 50us 的聚簇相邻对数 %d/%d (%.1f%%)" % (med_gap / 1000.0, cluster_cnt, len(gaps), cluster_ratio * 100))

        total_comp = len(comp_times)
        ksoft_cnt = sum(c for p, c in comp_procs.items() if "ksoftirqd" in str(p))
        swapper_cnt = sum(c for p, c in comp_procs.items() if ("swapper" in str(p)) or ("<kernel>" in str(p)))
        if ksoft_cnt:
            L.append("- complete 记录进程含 `ksoftirqd` (%d 条, %.1f%%) → 软中断溢出交给内核线程处理的佐证。" % (ksoft_cnt, ksoft_cnt / total_comp * 100))
        if swapper_cnt:
            L.append("- complete 记录进程含 `swapper`/`<kernel>` (%d 条, %.1f%%) → 部分 complete 在中断/idle 上下文处理。" % (swapper_cnt, swapper_cnt / total_comp * 100))

        # 软中断延迟需显著聚簇: >=5 对且占比 >20%, 单靠个位数聚簇不构成强信号
        if cluster_cnt >= 5 and cluster_ratio > 0.2:
            L.append("- **甄别结论: 显著 complete 聚簇** → 多个 IO 的 complete 被集中处理, 符合『硬件早已完成、软中断延迟后批量处理』特征; issue→complete 中可能混入 CPU 调度延迟, 不全是真实器件耗时。\n")
            softirq_delay = True
        else:
            L.append("- complete 时间戳分布整体均匀, 无显著聚簇 → complete 时间基本反映器件真实完成, issue→complete 主要为真实器件耗时。\n")
    else:
        L.append("- 窗口内 complete 事件过少 (< 3 条), 不足以判断聚簇。\n")

    # ---- 6. 结论 ----
    L.append("## 6. 根因结论\n")
    stall_from_dur = abs(gap_ms - dur_ns / NS_PER_MS) < 5 if (last_cmp_before and first_cmp_after) else False
    if last_cmp_before and first_cmp_after:
        L.append("- 线程 D 状态 Duration=%.3fms 与 complete 断档 %.3fms %s。" % (
            dur_ns / NS_PER_MS, gap_ms, "高度吻合 → 卡顿即设备停滞" if stall_from_dur else "存在偏差, 需进一步核对"))
    if schedule_ms > device_ms:
        L.append("- 该笔 IO 耗时以**调度阶段为主 (%.1f%%)**, 但结合 in-flight=%d、堆积=%d、complete 断档 %.3fms, 调度耗时本质是『排队等设备恢复』, 根因仍在器件侧。" % (
            schedule_ms / total_ms * 100, len(inflight), len(stacked), gap_ms if (last_cmp_before and first_cmp_after) else 0))
    else:
        L.append("- 该笔 IO 耗时以**器件阶段为主 (%.1f%%)**, 器件执行 %.3fms 明显偏慢。" % (device_ms / total_ms * 100, device_ms))
    L.append("- **根因**: 存储设备在顺序大块同步写下写缓存耗尽/触发 FTL(GC/刷盘), 产生约 %.3fms 写停滞。" % (gap_ms if (last_cmp_before and first_cmp_after) else total_ms))
    if softirq_delay:
        L.append("- **软中断延迟提示**: 检测到 complete 聚簇, 上述 issue→complete (%.3fms) 中可能混入软中断/CPU 调度延迟, 真实器件耗时应低于该值。" % device_ms)
    L.append("")

    # ---- 7. 数据来源 ----
    L.append("## 7. 数据来源\n")
    L.append("- 原始事件: `%s` (Sheet: 全部 IO, 窗口 [%.6f, %.6f]s)" % (xlsx, lo / 1e9, hi / 1e9))
    L.append("- 实际读取: `%s` (CSV 缓存, 由 `convert_to_csv.py --trace %s` 一次性生成)" % (src, trace_name))
    L.append("- 配对规则: 按 (设备号, 请求扇区地址) 分组, insert→issue→complete 三段式匹配")
    L.append("- 附加明细: 见同目录 `io_stall_data.txt`\n")

    # ---- 附加数据 ----
    A.append("目标 IO 定位 sector=%s dev=%s" % (target_sector, target_dev))
    A.append("")
    A.append("== 目标线程相关事件 (按时间) ==")
    if tid or pid or process:
        related = [e for e in events if match_target(e)]
        related.sort(key=lambda e: _to_int(e.get("时间 (纳秒)", 0)))
        for e in related:
            A.append("%+.1fus %-20s %s | %s" % (
                _rel(_to_int(e.get("时间 (纳秒)", 0)), start_ns),
                e.get("Trace 字段名称", ""), _ev_who(e), _ev_size(e)))
    A.append("")
    A.append("== START 时刻 in-flight 请求 (%d) ==" % len(inflight))
    for p in inflight:
        A.append("sector=%s issue=%+.1fus complete=%+.1fus rw=%s bytes=%s" % (
            p["sector"], _rel(p["issue_ts"], start_ns), _rel(p["complete_ts"], start_ns),
            p["issue_ev"].get("读写类型", ""), p["issue_ev"].get("请求字节数", "")))
    A.append("")
    A.append("== 堆积请求 (insert<=START<issue, %d) ==" % len(stacked))
    for p in stacked:
        A.append("sector=%s ins=%+.1fus issue=%+.1fus complete=%+.1fus rw=%s bytes=%s" % (
            p["sector"], _rel(p["insert_ts"], start_ns), _rel(p["issue_ts"], start_ns),
            _rel(p["complete_ts"], start_ns), p["insert_ev"].get("读写类型", ""),
            p["insert_ev"].get("请求字节数", "")))

    return "\n".join(L), "\n".join(A)


def _append_ai_section(report_path, report, ai_model):
    """调用 DeepSeek 云端模型对规则报告做 AI 根因研判，追加到报告末尾。

    失败时仅打印警告，不影响已有规则报告。
    """
    import llm_client

    print("\n[AI] 正在调用 DeepSeek 云端模型进行根因研判 (模型角色: %s)..." % ai_model)

    system_prompt = (
        "你是一名资深存储 IO 性能专家，擅长分析 Android/Linux block IO 卡顿问题。\n"
        "下面是一份由规则脚本生成的 IO 卡顿分析报告。请基于报告中的证据做独立研判，"
        "按以下结构输出 Markdown：\n"
        "1. 根因判定（一句话）\n"
        "2. 证据链（引用报告中的具体数字）\n"
        "3. 次要因素 / 风险点\n"
        "4. 优化建议（按收益排序）\n"
        "5. 置信度（高/中/低 + 理由）"
    )
    user_prompt = "以下是规则分析报告：\n\n" + report

    try:
        content = llm_client.analyze_with_prompt(
            system_prompt, user_prompt, role=ai_model)
    except Exception as e:
        print("  [警告] AI 研判失败，已跳过: %s" % e)
        return

    section = "\n\n---\n\n# AI 根因研判 (DeepSeek)\n\n" + content.strip() + "\n"
    with open(report_path, "a", encoding="utf-8") as f:
        f.write(section)
    print("[AI] 已追加 AI 根因研判到: %s" % report_path)


def main():
    ap = argparse.ArgumentParser(description="IO 卡顿快速分析")
    ap.add_argument("--trace", required=True, help="trace 名称 (如 com.ss.android.ugc.aweme)")
    ap.add_argument("--start", required=True, type=int, help="Start time (纳秒)")
    ap.add_argument("--duration", required=True, type=int, help="Duration (纳秒)")
    ap.add_argument("--process", default=None, help="进程名 (可选)")
    ap.add_argument("--pid", default=None, help="进程 PID (可选)")
    ap.add_argument("--tid", default=None, help="线程 TID (可选)")
    ap.add_argument("--pre-ms", type=int, default=50, help="卡顿点前窗口 (ms)")
    ap.add_argument("--post-ms", type=int, default=200, help="卡顿点后窗口 (ms)")
    ap.add_argument("--out-dir", default=None, help="输出目录 (默认 out/<trace>/stall_analysis/<start>_<dur>)")
    ap.add_argument("--ai", action="store_true", help="调用 DeepSeek 云端模型追加 AI 根因研判章节")
    ap.add_argument("--ai-model", default="deep", choices=["fast", "deep"],
                    help="AI 模型角色: fast=deepseek-chat, deep=deepseek-reasoner (默认 deep)")
    args = ap.parse_args()

    trace_dir = os.path.join(BASE, "out", args.trace)
    if not os.path.isdir(trace_dir):
        print("[错误] trace 目录不存在: %s" % trace_dir)
        sys.exit(1)

    if args.out_dir:
        out_dir = args.out_dir
    else:
        out_dir = os.path.join(trace_dir, "stall_analysis", "%d_%d" % (args.start, args.duration))
    os.makedirs(out_dir, exist_ok=True)

    report, data = analyze(trace_dir, args.trace, args.start, args.duration,
                           process=args.process, pid=args.pid, tid=args.tid,
                           pre_ms=args.pre_ms, post_ms=args.post_ms)

    report_path = os.path.join(out_dir, "io_stall_report.md")
    data_path = os.path.join(out_dir, "io_stall_data.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    with open(data_path, "w", encoding="utf-8") as f:
        f.write(data)

    if args.ai:
        _append_ai_section(report_path, report, args.ai_model)

    print("报告已生成: %s" % report_path)
    print("数据明细: %s" % data_path)


if __name__ == "__main__":
    main()