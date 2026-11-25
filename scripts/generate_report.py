import os
import sys
import glob
import base64
import argparse
from datetime import datetime

from config_loader import get_project_root, load_config
from expert_notes import get_step_note, get_file_notes

STEPS = [
    {
        "id": "1",
        "title": "格式转换 (perfetto → systrace)",
        "desc": "使用 Perfetto trace_processor_shell 将原始 trace 文件转换为可读的 systrace 文本格式。",
        "files": [{"pattern": "{name}.systrace", "desc": "systrace 文本文件"}],
    },
    {
        "id": "2",
        "title": "IO 事件提取 (16 列 XLSX/CSV)",
        "desc": "从 systrace 解析 block_rq_insert / block_rq_issue / block_rq_complete 事件，提取 16 个维度到表格文件。",
        "files": [{"pattern": "{name}_block_io.xlsx", "desc": "IO 事件表格（XLSX 优先，CSV 可用于快速分析）"}],
    },
    {
        "id": "3",
        "title": "IO 时延分析 (调度 + 器件耗时分布)",
        "desc": "计算每条 IO 的调度耗时 (insert→issue) 与器件耗时 (issue→complete)，输出百分位统计表和分布直方图。",
        "files": [
            {"pattern": "{name}_schedule_latency.png", "desc": "调度耗时分布直方图", "embed": True},
            {"pattern": "{name}_device_latency.png", "desc": "器件耗时分布直方图", "embed": True},
            {"pattern": "{name}_latency.xlsx", "desc": "时延汇总 Excel（含调度耗时 / 器件耗时 / IO 大小 / 进程统计 4 个 Sheet）"},
        ],
    },
    {
        "id": "4",
        "title": "IO 大小分布 (扇区 + KB 分布图)",
        "desc": "按扇区数量和 IO 大小 (KB) 绘制请求大小分布直方图，分析 IO 碎片化程度。",
        "files": [
            {"pattern": "{name}_io_nr_sector.png", "desc": "扇区数量分布直方图", "embed": True},
            {"pattern": "{name}_io_size_kb.png", "desc": "IO 大小 (KB) 分布直方图", "embed": True},
        ],
    },
    {
        "id": "5",
        "title": "进程 IO 统计 (按线程, 数据量降序)",
        "desc": "统计每个线程的 IO 数量与数据量，按数据量降序排列，结果写入 latency.xlsx 的「进程 IO 统计」Sheet。",
        "files": [
            {"pattern": "{name}_latency.xlsx", "desc": "时延汇总 Excel（本步骤追加「进程 IO 统计」Sheet）"},
        ],
    },
    {
        "id": "6",
        "title": "指定 PID/TID 全维度分析",
        "desc": "针对指定进程/线程做独立的 IO 分析，输出包含 4 个 Sheet 的 XLSX 和 4 张图表。",
        "detect_pid": True,
        "files": [],
    },
    {
        "id": "7",
        "title": "IO 综合诊断",
        "desc": "一次性输出 10 项核心诊断指标：时延百分位 / 读写分离 / 队列深度 / IOPS & 吞吐量 / 顺序/随机 IO 检测 / IO 大小 vs 时延散点图 / 慢 IO Top-N / 同步/异步 IO 分类 / 时延时间线 / LBA 时间线分布。",
        "files": [
            {"pattern": "{name}{suffix}_diagnosis.xlsx",
             "desc": "综合诊断 Excel（10 个 Sheet，含逐 IO 的时延时间线明细）"},
            {"pattern": "{name}{suffix}_queue_depth.png",
             "desc": "队列深度时间曲线（调度队列 + 器件队列）", "embed": True},
            {"pattern": "{name}{suffix}_iops_throughput.png",
             "desc": "IOPS + 吞吐量时间曲线（双 Y 轴）", "embed": True},
            {"pattern": "{name}{suffix}_rw_latency.png",
             "desc": "读写时延对比箱线图", "embed": True},
            {"pattern": "{name}{suffix}_rw_size.png",
             "desc": "读写 IO 大小对比直方图", "embed": True},
            {"pattern": "{name}{suffix}_seq_random.png",
             "desc": "顺序/随机 IO 占比饼图 + 读写拆分", "embed": True},
            {"pattern": "{name}{suffix}_size_latency.png",
             "desc": "IO 大小 vs 时延散点图（含趋势线）", "embed": True},
            {"pattern": "{name}{suffix}_sync_io_latency.png",
             "desc": "同步 vs 异步写时延箱线图 + 百分位对比", "embed": True},
            {"pattern": "{name}{suffix}_latency_timeline.png",
             "desc": "时延时间线散点图（端到端/调度/器件三面板 + IO 密度）", "embed": True},
            {"pattern": "{name}{suffix}_lba_timeline.png",
             "desc": "LBA 扇区地址时间线分布图（按读写类型着色，观察 IO 离散程度）", "embed": True},
        ],
    },
]


def img_to_base64(filepath):
    try:
        with open(filepath, "rb") as f:
            data = f.read()
        ext = os.path.splitext(filepath)[1].lower()
        mime = "image/png" if ext == ".png" else "image/jpeg"
        return f"data:{mime};base64," + base64.b64encode(data).decode("ascii")
    except Exception:
        return None


def fmt_size(filepath):
    try:
        sz = os.path.getsize(filepath)
        if sz < 1024:
            return f"{sz} B"
        elif sz < 1024 * 1024:
            return f"{sz / 1024:.1f} KB"
        else:
            return f"{sz / (1024 * 1024):.2f} MB"
    except Exception:
        return "—"


def fmt_time(filepath):
    try:
        ts = os.path.getmtime(filepath)
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "—"


def fmt_num(n):
    """带千分位分隔符的整数格式化。"""
    try:
        return f"{int(n):,}"
    except (ValueError, TypeError):
        return "—"


def find_files(trace_dir, name, file_spec, suffix_variants=None):
    results = []
    for spec in file_spec:
        pattern = spec["pattern"]
        base_dir = os.path.join(trace_dir, "charts") if spec.get("embed") else trace_dir
        if "{suffix}" in pattern:
            if suffix_variants:
                for sfx in suffix_variants:
                    resolved = pattern.replace("{name}", name).replace("{suffix}", sfx)
                    fpath = os.path.join(base_dir, resolved)
                    if os.path.isfile(fpath):
                        results.append({"path": fpath, "desc": spec["desc"].replace("{suffix}", sfx),
                                        "embed": spec.get("embed", False), "basename": resolved})
        else:
            resolved = pattern.replace("{name}", name)
            fpath = os.path.join(base_dir, resolved)
            if os.path.isfile(fpath):
                results.append({"path": fpath, "desc": spec["desc"],
                                "embed": spec.get("embed", False), "basename": resolved})
    return results


def detect_step6(trace_dir, name):
    entries = []
    for entry in sorted(os.listdir(trace_dir)):
        target_dir = os.path.join(trace_dir, entry)
        if not os.path.isdir(target_dir):
            continue
        if not (entry.startswith("PID_") or entry.startswith("TID_")):
            continue
        target_value = entry.split("_", 1)[1]
        xlsx_bn = f"{name}_{target_value}.xlsx"
        xlsx_path = os.path.join(target_dir, xlsx_bn)

        pngs = []
        for png_pat in [
            f"{name}_{target_value}_schedule_latency.png",
            f"{name}_{target_value}_device_latency.png",
            f"{name}_{target_value}_io_nr_sector.png",
            f"{name}_{target_value}_io_size_kb.png",
        ]:
            pp = os.path.join(target_dir, png_pat)
            if os.path.isfile(pp):
                pngs.append({"path": pp,
                             "desc": png_pat.replace(f"{name}_{target_value}_", "").replace(".png", "").replace("_", " "),
                             "embed": True, "basename": png_pat})

        entries.append({
            "pid": entry,
            "xlsx": {"path": xlsx_path, "basename": os.path.join(entry, xlsx_bn)},
            "pngs": pngs,
        })
    return entries


def detect_step7_suffix_variants(trace_dir, name):
    variants = []
    for f in glob.glob(os.path.join(trace_dir, f"{name}_*diagnosis.xlsx")):
        bn = os.path.basename(f)
        mid = bn[len(name):-len("_diagnosis.xlsx")]
        if mid == "":
            variants.append(("", bn))
        elif mid.startswith("_"):
            variants.append((mid, bn))
    if not variants:
        variants.append(("", f"{name}_diagnosis.xlsx"))
    return [sfx for sfx, _ in variants]


def find_step7_pngs(trace_dir, name, suffix_variants):
    step7_specs = STEPS[6]["files"]
    return find_files(trace_dir, name, step7_specs, suffix_variants)


def classify_rwbs(rwbs_str):
    """按 rwbs 字符串将 IO 归类为 read / write / other。"""
    if not rwbs_str:
        return "other"
    rwbs = str(rwbs_str).strip().upper()
    raw = rwbs.split(" ")[0]
    if not raw:
        return "other"
    if "W" in raw and "R" not in raw:
        return "write"
    if "R" in raw:
        return "read"
    return "other"


def summarize_table_file(file_path):
    """读取表格文件的概览信息。

    对任意 XLSX/CSV 统计每个 Sheet 的数据行数；对 block_io 文件
    (文件名含 _block_io) 额外统计有效 IO Size 记录数及 read/write/other 分布，
    与 io_size_chart 的口径保持一致。

    Returns:
        dict: {"sheets": [{"name", "rows", "valid_size", "rw"}], "error": str|None}
    """
    ext = os.path.splitext(file_path)[1].lower()
    basename = os.path.basename(file_path)
    compute_rw = "_block_io" in basename

    summary = {"sheets": [], "error": None}

    if ext == ".csv":
        try:
            import csv as _csv
            with open(file_path, "r", encoding="utf-8") as f:
                reader = _csv.DictReader(f)
                headers = list(reader.fieldnames or [])
                rw_col = headers.index("读写类型") if compute_rw and "读写类型" in headers else None
                ns_col = headers.index("请求扇区数量") if compute_rw and "请求扇区数量" in headers else None
                rows = 0
                valid_size = 0
                rw = {"read": 0, "write": 0, "other": 0}
                for row in reader:
                    rows += 1
                    if ns_col is None:
                        continue
                    try:
                        ns = int(row.get(headers[ns_col]) or 0)
                    except (ValueError, TypeError):
                        continue
                    if ns > 0:
                        valid_size += 1
                        if rw_col is not None:
                            rw[classify_rwbs(row.get(headers[rw_col], ""))] += 1
            sheet_info = {"name": "全部 IO", "rows": rows}
            if compute_rw:
                sheet_info["valid_size"] = valid_size
                sheet_info["rw"] = rw
            summary["sheets"].append(sheet_info)
        except Exception as e:
            summary["error"] = str(e)
        return summary

    try:
        from openpyxl import load_workbook
        wb = load_workbook(file_path, read_only=True)
    except Exception as e:
        summary["error"] = str(e)
        return summary

    try:
        for ws in wb.worksheets:
            sheet_info = {"name": ws.title, "rows": 0}
            rw_col = None
            ns_col = None
            valid_size = 0
            rw = {"read": 0, "write": 0, "other": 0}
            for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if row_idx == 0:
                    headers = [str(c) if c is not None else "" for c in row]
                    if compute_rw and "读写类型" in headers:
                        rw_col = headers.index("读写类型")
                    if compute_rw and "请求扇区数量" in headers:
                        ns_col = headers.index("请求扇区数量")
                    continue
                sheet_info["rows"] += 1
                if ns_col is None:
                    continue
                try:
                    ns_val = row[ns_col] if ns_col < len(row) else None
                    ns = int(ns_val) if ns_val not in (None, "") else 0
                except (ValueError, TypeError):
                    continue
                if ns > 0:
                    valid_size += 1
                    if rw_col is not None:
                        rw_val = row[rw_col] if rw_col < len(row) else ""
                        rw[classify_rwbs(rw_val)] += 1
            if compute_rw:
                sheet_info["valid_size"] = valid_size
                sheet_info["rw"] = rw
            summary["sheets"].append(sheet_info)
    except Exception as e:
        summary["error"] = str(e)
    finally:
        wb.close()

    return summary


def table_summary_html(file_path):
    """为表格文件生成一段「表格内容概览」HTML；无法解析时返回空串。"""
    summary = summarize_table_file(file_path)
    if summary.get("error") or not summary.get("sheets"):
        return ""

    n_sheets = len(summary["sheets"])
    title = "表格内容概览" if n_sheets <= 1 else f"表格内容概览（{n_sheets} 个 Sheet）"
    parts = [f'<div class="table-summary"><div class="ts-title">{title}</div>']

    parts.append('<div class="ts-sheets">')
    for s in summary["sheets"]:
        parts.append(f'<span class="ts-chip">{s["name"]} · <b>{fmt_num(s["rows"])}</b> 行</span>')
    parts.append('</div>')

    rw_lines = []
    for s in summary["sheets"]:
        rw = s.get("rw")
        if rw and any(rw.values()):
            rw_lines.append(
                f'{s["name"]}: 有效 IO Size 记录 <b>{fmt_num(s["valid_size"])}</b> · '
                f'<span class="rw-read">read {fmt_num(rw["read"])}</span> · '
                f'<span class="rw-write">write {fmt_num(rw["write"])}</span> · '
                f'<span class="rw-other">other {fmt_num(rw["other"])}</span>'
            )
    if rw_lines:
        parts.append('<div class="ts-rw">' + '<br>'.join(rw_lines) + '</div>')

    parts.append('</div>')
    return "".join(parts)


def build_html(trace_name, trace_dir, report_path):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html_parts = []
    css = """
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: "Microsoft YaHei", "PingFang SC", "Segoe UI", sans-serif;
    background: #f0f2f5; color: #1a1a2e; line-height: 1.7;
    padding-bottom: 60px;
}
.header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    color: #fff; padding: 40px 48px 32px;
}
.header h1 { font-size: 28px; font-weight: 700; margin-bottom: 8px; }
.header .sub { font-size: 14px; color: #8892b0; }
.header .meta { margin-top: 16px; display: flex; gap: 32px; font-size: 13px; color: #a8b2d1; flex-wrap: wrap; }
.container { max-width: 1100px; margin: 0 auto; padding: 0 24px; }
.step {
    background: #fff; border-radius: 10px; margin-top: 24px;
    box-shadow: 0 1px 3px rgba(0,0,0,.08); overflow: hidden;
}
.step-header {
    background: #f8f9fc; padding: 18px 28px; border-bottom: 1px solid #e8ecf1;
    display: flex; align-items: center; gap: 14px;
}
.step-num {
    width: 36px; height: 36px; border-radius: 50%;
    background: linear-gradient(135deg, #4472C4, #6C8CD9);
    color: #fff; font-size: 15px; font-weight: 700;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.step-title { font-size: 17px; font-weight: 700; color: #1a1a2e; }
.step-desc { padding: 12px 28px; font-size: 14px; color: #555; background: #fafbfc; border-bottom: 1px solid #e8ecf1; }
.step-body { padding: 20px 28px; }
.file-list { list-style: none; }
.file-item { padding: 10px 0; border-bottom: 1px solid #f0f0f0; }
.file-item:last-child { border-bottom: none; }
.file-type { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; margin-right: 10px; }
.file-type.xlsx { background: #e8f5e9; color: #2e7d32; }
.file-type.png { background: #fff3e0; color: #e65100; }
.file-type.systrace { background: #e3f2fd; color: #1565c0; }
.file-desc { font-size: 14px; color: #333; }
.file-meta { font-size: 12px; color: #888; margin-top: 2px; }
.file-link { font-size: 12px; color: #4472C4; text-decoration: none; }
.file-link:hover { text-decoration: underline; }
.img-wrap { margin-top: 14px; border: 1px solid #e8ecf1; border-radius: 8px; overflow: hidden; background: #fafbfc; }
.img-wrap img { width: 100%; height: auto; display: block; }
.img-caption { padding: 10px 14px; font-size: 13px; color: #666; border-top: 1px solid #e8ecf1; }
.expert-note {
    margin: 10px 0; padding: 12px 16px;
    background: linear-gradient(135deg, #eef2ff 0%, #f0f4ff 100%);
    border-left: 4px solid #4472C4; border-radius: 0 8px 8px 0;
    font-size: 13px; color: #374151; line-height: 1.7;
}
.expert-note b { color: #1e40af; }
.expert-note-step { margin: 8px 28px 0; }
.expert-note-file { margin: 4px 0 0 0; }
.pid-section { margin-top: 16px; padding: 16px 20px; background: #f8f9fc; border-radius: 8px; border-left: 3px solid #4472C4; }
.pid-section h4 { font-size: 15px; color: #1a1a2e; margin-bottom: 10px; }
.empty-step { padding: 20px 28px; color: #999; font-size: 14px; text-align: center; }
.footer { text-align: center; padding: 40px 24px 20px; color: #999; font-size: 12px; }
.toc { background: #fff; border-radius: 10px; margin-top: 24px;
    box-shadow: 0 1px 3px rgba(0,0,0,.08); padding: 20px 28px; }
.toc h3 { font-size: 16px; margin-bottom: 12px; color: #1a1a2e; }
.toc a { display: block; padding: 6px 0; color: #4472C4; text-decoration: none; font-size: 14px; }
.toc a:hover { color: #1a56db; }
.summary { background: #fff; border-radius: 10px; margin-top: 24px;
    box-shadow: 0 1px 3px rgba(0,0,0,.08); padding: 20px 28px; }
.summary h3 { font-size: 16px; margin-bottom: 10px; color: #1a1a2e; }
.summary table { width: 100%; border-collapse: collapse; font-size: 14px; }
.summary th { background: #f8f9fc; padding: 8px 12px; text-align: left; border-bottom: 2px solid #e8ecf1; }
.summary td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; }
.summary tr:last-child td { border-bottom: none; }
.table-summary { margin-top: 10px; padding: 10px 14px; background: #f6f8fb; border: 1px solid #e3e8f0; border-radius: 8px; }
.table-summary .ts-title { font-size: 12px; font-weight: 700; color: #4472C4; margin-bottom: 6px; }
.table-summary .ts-sheets { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 6px; }
.table-summary .ts-chip { padding: 2px 10px; border-radius: 12px; background: #fff; border: 1px solid #dbe4f0; font-size: 12px; color: #444; }
.table-summary .ts-chip b { color: #1a1a2e; }
.table-summary .ts-rw { font-size: 12px; color: #555; line-height: 1.9; }
.table-summary .ts-rw b { color: #1a1a2e; }
.table-summary .rw-read { color: #2e7d32; font-weight: 700; }
.table-summary .rw-write { color: #c62828; font-weight: 700; }
.table-summary .rw-other { color: #1565c0; font-weight: 700; }
"""
    html_parts.append(f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AI-trace 分析报告 — {trace_name}</title>
<style>{css}</style>
</head>
<body>
<div class="header">
    <h1>AI-trace 分析报告</h1>
    <div class="sub">Perfetto Block IO 全流程诊断 · 一键汇总</div>
    <div class="meta">
        <span>Trace: {trace_name}</span>
        <span>生成时间: {now}</span>
    </div>
</div>
<div class="container">
""")

    # TOC
    html_parts.append('<div class="toc"><h3>报告目录</h3>')
    for s in STEPS:
        html_parts.append(f'<a href="#step{s["id"]}">步骤 {s["id"]}: {s["title"]}</a>')
    html_parts.append('</div>')

    # Summary table
    total_files = 0
    summary_rows = []
    suffix_variants = detect_step7_suffix_variants(trace_dir, trace_name)

    for step in STEPS:
        sid = step["id"]
        if sid == "6":
            pid_entries = detect_step6(trace_dir, trace_name)
            file_count = sum(1 + len(e["pngs"]) for e in pid_entries)
        elif sid == "7":
            files = find_step7_pngs(trace_dir, trace_name, suffix_variants)
            file_count = len(files)
        else:
            files = find_files(trace_dir, trace_name, step["files"])
            file_count = len(files)
        status = "已生成" if file_count > 0 else "未生成"
        summary_rows.append((sid, step["title"], status, str(file_count)))
        total_files += file_count

    html_parts.append('<div class="summary"><h3>产出总览</h3><table>')
    html_parts.append('<tr><th>步骤</th><th>名称</th><th>状态</th><th>文件数</th></tr>')
    for row in summary_rows:
        icon = "&#x2705;" if row[2] == "已生成" else "&#x26A0;"
        html_parts.append(f'<tr><td>{row[0]}</td><td>{row[1]}</td><td>{icon}&nbsp;{row[2]}</td><td>{row[3]}</td></tr>')
    html_parts.append(f'<tr style="font-weight:700"><td colspan="3">合计</td><td>{total_files} 个文件</td></tr>')
    html_parts.append('</table></div>')

    # Step sections
    for step in STEPS:
        sid = step["id"]
        html_parts.append(f'<div class="step" id="step{sid}">')
        html_parts.append(f'<div class="step-header">'
                          f'<div class="step-num">{sid}</div>'
                          f'<div class="step-title">{step["title"]}</div>'
                          f'</div>')
        html_parts.append(f'<div class="step-desc">{step["desc"]}</div>')
        
        step_note = get_step_note(sid)
        if step_note:
            html_parts.append(f'<div class="expert-note expert-note-step">⚡ <b>专家分析指引</b><br>{step_note}</div>')
        html_parts.append('<div class="step-body">')

        if sid == "6":
            pid_entries = detect_step6(trace_dir, trace_name)
            if not pid_entries:
                html_parts.append(
                    '<div class="empty-step">未检测到此步骤的输出文件（可能尚未执行此步骤，或未指定 PID/TID）</div>')
            else:
                html_parts.append(
                    f'<p style="font-size:14px;color:#666;margin-bottom:12px">'
                    f'共检测到 <b>{len(pid_entries)}</b> 个 PID/TID 的分析结果：</p>')
                for entry in pid_entries:
                    xlsx_info = entry["xlsx"]
                    html_parts.append('<div class="pid-section">')
                    html_parts.append(f'<h4>{entry["pid"]}</h4>')
                    html_parts.append('<ul class="file-list">')
                    html_parts.append(
                        f'<li class="file-item">'
                        f'<span class="file-type xlsx">XLSX</span>'
                        f'<span class="file-desc">全维度分析 Excel（{entry["pid"]}）</span>'
                        f'<br><span class="file-meta">'
                        f'{fmt_size(xlsx_info["path"])} · {fmt_time(xlsx_info["path"])} · '
                        f'<a class="file-link" href="{xlsx_info["basename"]}" target="_blank" rel="noopener noreferrer">打开文件</a>'
                        f'</span></li>')
                    ts_html6 = table_summary_html(xlsx_info["path"])
                    if ts_html6:
                        html_parts.append(ts_html6)
                    for png in entry["pngs"]:
                        html_parts.append(
                            f'<li class="file-item">'
                            f'<span class="file-type png">PNG</span>'
                            f'<span class="file-desc">{png["desc"]}</span>'
                            f'<br><span class="file-meta">'
                            f'{fmt_size(png["path"])} · {fmt_time(png["path"])}'
                            f'</span></li>')
                        b64 = img_to_base64(png["path"])
                        if b64:
                            html_parts.append(
                                f'<div class="img-wrap">'
                                f'<img src="{b64}" alt="{png["desc"]}" loading="lazy">'
                                f'<div class="img-caption">{png["desc"]}</div>'
                                f'</div>')
                            file_note6 = get_file_notes(sid, png["basename"])
                            if file_note6:
                                html_parts.append(f'<div class="expert-note expert-note-file">🔍 <b>专家解读</b><br>{file_note6}</div>')
                    html_parts.append('</ul></div>')
        else:
            if sid == "7":
                files = find_step7_pngs(trace_dir, trace_name, suffix_variants)
            else:
                files = find_files(trace_dir, trace_name, step["files"])

            if not files:
                html_parts.append('<div class="empty-step">未检测到此步骤的输出文件（可能尚未执行此步骤）</div>')
            else:
                html_parts.append('<ul class="file-list">')
                for f in files:
                    ext = os.path.splitext(f["basename"])[1].lstrip(".").upper()
                    if ext == "SYSTRACE":
                        type_class = "systrace"
                    elif ext == "XLSX":
                        type_class = "xlsx"
                    elif ext == "PNG":
                        type_class = "png"
                    else:
                        type_class = ""

                    meta = f'{fmt_size(f["path"])} · {fmt_time(f["path"])}'
                    if ext in ("XLSX", "SYSTRACE"):
                        meta += (f' · '
                                 f'<a class="file-link" href="{f["basename"]}" target="_blank" rel="noopener noreferrer">打开文件</a>')

                    html_parts.append(
                        f'<li class="file-item">'
                        f'<span class="file-type {type_class}">{ext}</span>'
                        f'<span class="file-desc">{f["desc"]}</span>'
                        f'<br><span class="file-meta">{meta}</span>'
                        f'</li>')
                    if ext == "XLSX":
                        ts_html = table_summary_html(f["path"])
                        if ts_html:
                            html_parts.append(ts_html)
                    if not f.get("embed"):
                        fn = get_file_notes(sid, f["basename"])
                        if fn:
                            html_parts.append(f'<div class="expert-note expert-note-file">🔍 <b>专家解读</b><br>{fn}</div>')

                    if f.get("embed"):
                        b64 = img_to_base64(f["path"])
                        if b64:
                            html_parts.append(
                                f'<div class="img-wrap">'
                                f'<img src="{b64}" alt="{f["desc"]}" loading="lazy">'
                                f'<div class="img-caption">{f["desc"]}</div>'
                                f'</div>')
                        file_note = get_file_notes(sid, f["basename"])
                        if file_note:
                            html_parts.append(f'<div class="expert-note expert-note-file">🔍 <b>专家解读</b><br>{file_note}</div>')
                html_parts.append('</ul>')

        html_parts.append('</div></div>')

    html_parts.append(f"""</div>
<div class="footer">
    <p>AI-trace — Perfetto Block IO 分析工具集 · 报告生成时间: {now}</p>
</div>
</body>
</html>""")

    html = "\n".join(html_parts)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html)
    return html


def main():
    parser = argparse.ArgumentParser(description="生成所有步骤产物的汇总 HTML 报告")
    parser.add_argument("--trace-name", type=str, default=None,
                        help="trace 名称（默认读取配置文件）")
    args = parser.parse_args()

    config = load_config()
    project_root = get_project_root()
    out_dir = os.path.join(project_root, "out")

    trace_name = args.trace_name or config.get("trace_name", "")
    if not trace_name:
        if os.path.isdir(out_dir):
            candidates = [
                d for d in os.listdir(out_dir)
                if os.path.isdir(os.path.join(out_dir, d))
            ]
            if len(candidates) == 1:
                trace_name = candidates[0]
                print(f"[信息] 自动选择唯一 trace: {trace_name}")
            elif not candidates:
                print("[错误] out 目录下没有 trace 输出目录，请先执行前面的步骤")
                sys.exit(1)
            else:
                print("[错误] 未指定 trace 名称，out 目录下有多个 trace：")
                for c in sorted(candidates):
                    print(f"  - {c}")
                print("请使用 --trace-name 参数指定")
                sys.exit(1)
        else:
            print("[错误] out 目录不存在")
            sys.exit(1)

    trace_dir = os.path.join(out_dir, trace_name)
    if not os.path.isdir(trace_dir):
        print(f"[错误] trace 输出目录不存在: {trace_dir}")
        sys.exit(1)

    report_path = os.path.join(trace_dir, f"{trace_name}_report.html")
    print(f"[信息] 开始生成报告 -> {report_path}")

    build_html(trace_name, trace_dir, report_path)

    report_size = fmt_size(report_path)
    print(f"\n报告已生成: {report_path} ({report_size})")
    print(f"  可直接双击 HTML 文件在浏览器中查看")


if __name__ == "__main__":
    main()