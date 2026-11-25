import os
import sys
import glob
import argparse
from collections import defaultdict, Counter

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

from config_loader import load_config
from io_reader import load_io_events, find_block_io_file

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
        rwbs_type = classify_rwbs(rwbs_str)

        sizes.append({
            "nr_sector": nr_sector,
            "size_kb": nr_sector * 512 / 1024,
            "rwbs": rwbs_str,
            "rwbs_type": rwbs_type,
        })

    return sizes


def plot_distribution(data_list, value_key, xlabel, title, output_path,
                      color_palette=None, is_integer=False):
    if color_palette is None:
        color_palette = {"read": "#4CAF50", "write": "#F44336", "other": "#2196F3"}

    categories = ["read", "write", "other"]
    labels_cn = {"read": "读请求", "write": "写请求", "other": "其他请求"}

    grouped = defaultdict(list)
    for item in data_list:
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
        items = sorted(counter.items())

        if is_integer:
            labels = [str(k) for k, v in items]
            x_positions = list(range(len(items)))
            heights = [v for k, v in items]
        else:
            labels = [f"{k:.1f}" for k, v in items]
            x_positions = list(range(len(items)))
            heights = [v for k, v in items]

        color = color_palette.get(cat, "#888888")
        bar_width = max(0.6, 0.9 - 0.01 * len(items))
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

        total = len(data)
        most_common = counter.most_common(3)
        top_text = "\n".join([f"{val}{' ' + xlabel.split('(')[0].strip() if is_integer else ' KB'}: {cnt} 次"
                              for val, cnt in most_common])

        stats_text = (f"总数: {total}\n"
                      f"不同值: {len(counter)} 种\n"
                      f"均值: {np.mean(data):.2f}\n"
                      f"最小: {min(data)}\n"
                      f"最大: {max(data)}\n\n"
                      f"Top:\n{top_text}")
        ax.text(0.97, 0.97, stats_text, transform=ax.transAxes,
                fontsize=7.5, va="top", ha="right",
                bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.9))

        ax.set_title(labels_cn[cat], fontsize=14)
        ax.grid(axis="y", alpha=0.3, linestyle="--")
        ax.set_xlim(-0.5, len(items) - 0.5)

    plt.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  图表已保存: {output_path}")


def write_size_sheet(sizes, output_path):
    import openpyxl

    grouped = defaultdict(Counter)
    for item in sizes:
        grouped[item["rwbs_type"]][item["nr_sector"]] += 1

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=11)
    cell_font = Font(name="Consolas", size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    cat_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    cat_font = Font(name="Microsoft YaHei", bold=True, size=11)

    labels_cn = {"read": "读请求", "write": "写请求", "other": "其他请求"}
    col_headers = ["请求类型", "请求扇区数量", "请求大小 (KB)", "IO 数量"]

    if os.path.isfile(output_path):
        wb = load_workbook(output_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    sheet_name = "IO 大小汇总"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    for col_idx, name in enumerate(col_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_idx = 2
    for cat in ["read", "write", "other"]:
        data = grouped.get(cat, {})
        if not data:
            continue

        cat_cell = ws.cell(row=row_idx, column=1, value=labels_cn.get(cat, cat))
        cat_cell.font = cat_font
        cat_cell.fill = cat_fill
        cat_cell.border = thin_border
        total_cell = ws.cell(row=row_idx, column=4, value=sum(data.values()))
        total_cell.font = cat_font
        total_cell.fill = cat_fill
        total_cell.border = thin_border
        for c in [2, 3]:
            cell = ws.cell(row=row_idx, column=c)
            cell.fill = cat_fill
            cell.border = thin_border
        row_idx += 1

        for nr_sector, count in sorted(data.items()):
            ws.cell(row=row_idx, column=1).border = thin_border
            ws.cell(row=row_idx, column=1).font = cell_font
            ws.cell(row=row_idx, column=2, value=nr_sector).font = cell_font
            ws.cell(row=row_idx, column=2).border = thin_border
            ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal="right")
            ws.cell(row=row_idx, column=3, value=nr_sector * 512 // 1024).font = cell_font
            ws.cell(row=row_idx, column=3).border = thin_border
            ws.cell(row=row_idx, column=3).alignment = Alignment(horizontal="right")
            ws.cell(row=row_idx, column=4, value=count).font = cell_font
            ws.cell(row=row_idx, column=4).border = thin_border
            ws.cell(row=row_idx, column=4).alignment = Alignment(horizontal="center")
            row_idx += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"  表格已保存: {output_path}")


def parse_args():
    parser = argparse.ArgumentParser(description="分析 Block IO 请求大小分布")
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
    size_cfg = cfg.get("io_size_chart", {})

    if args.sheet is None:
        args.sheet = size_cfg.get("sheet", "全部 IO")

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

        if not events:
            print(f"  [警告] Sheet 中无数据")
            continue

        print(f"  读取到 {len(events)} 条 IO 记录")

        sizes = extract_io_sizes(events)
        print(f"  有效 IO Size 记录: {len(sizes)}")
        for cat in ["read", "write", "other"]:
            cnt = sum(1 for s in sizes if s["rwbs_type"] == cat)
            if cnt:
                print(f"    {cat}: {cnt}")

        if sizes:
            charts_dir = os.path.join(trace_dir, "charts")
            os.makedirs(charts_dir, exist_ok=True)

            chart_path = os.path.join(charts_dir, f"{trace_name}_io_nr_sector.png")
            print(f"\n[绘图] 请求扇区数量分布")
            plot_distribution(
                sizes,
                value_key="nr_sector",
                xlabel="扇区数量",
                title=f"{trace_name} - 请求扇区数量分布",
                output_path=chart_path,
                is_integer=True,
            )

            chart_path = os.path.join(charts_dir, f"{trace_name}_io_size_kb.png")
            print(f"\n[绘图] 请求大小分布 (KB)")
            plot_distribution(
                sizes,
                value_key="size_kb",
                xlabel="请求大小 (KB)",
                title=f"{trace_name} - 请求大小分布 (KB)",
                output_path=chart_path,
                is_integer=True,
            )

            latency_xlsx = os.path.join(trace_dir, f"{trace_name}_latency.xlsx")
            print(f"\n[写入] IO 大小汇总表 → {latency_xlsx}")
            write_size_sheet(sizes, latency_xlsx)

    print(f"\n全部完成。")


if __name__ == "__main__":
    main()