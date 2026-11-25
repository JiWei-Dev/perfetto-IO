"""将 block_io.xlsx 一次性转换为 block_io.csv 缓存, 加速后续窗口读取。

xlsx (openpyxl read_only) 遍历 60MB 文件较慢, 转成 CSV 后可用 csv 模块快速流式读取。
"""
import os
import sys
import csv
import time
import argparse

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from openpyxl import load_workbook

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def convert(trace_dir, trace_name, force=False):
    xlsx = os.path.join(trace_dir, f"{trace_name}_block_io.xlsx")
    csv_path = os.path.join(trace_dir, f"{trace_name}_block_io.csv")
    if os.path.isfile(csv_path) and not force:
        print(f"[跳过] 已存在缓存: {csv_path}")
        return csv_path
    if not os.path.isfile(xlsx):
        raise FileNotFoundError(f"未找到 {xlsx}")

    t0 = time.time()
    wb = load_workbook(xlsx, read_only=True)
    ws = wb["全部 IO"]
    tmp = csv_path + ".tmp"
    n = 0
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if c is None else c for c in row])
            n += 1
            if n % 100000 == 0:
                print("  ... %d 行 (%.1fs)" % (n, time.time() - t0))
    wb.close()
    os.replace(tmp, csv_path)
    print(f"[完成] {n} 行, 耗时 {time.time() - t0:.1f}s -> {csv_path}")
    return csv_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--trace", required=True)
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    trace_dir = os.path.join(BASE, "out", args.trace)
    convert(trace_dir, args.trace, force=args.force)


if __name__ == "__main__":
    main()