"""AI 交互式问答（AI2）

围绕当前 TRACE 的 IO 分析报告进行多轮对话。
除纯对话外，还支持读取本地文件 / 目录，便于直接解读报告、查看数据明细：
  - 直接粘贴一个存在的文件或目录路径，会自动读取/列出并交由 AI 解读；
  - /read <文件> 读取文件内容加载到上下文；
  - /ls <目录> 列出目录内容。

用法:
  python scripts/AI/io_chat.py --trace com.ss.android.ugc.aweme [--model fast|deep]
  python scripts/AI/io_chat.py --trace <trace> --report <报告路径>

对话命令:
  /quit         退出
  /clear        清空对话历史（保留报告上下文）
  /help         显示帮助
  /ls [目录]    列出目录内容（默认当前 trace 输出目录）
  /read <文件>  读取文本文件内容并让 AI 解读
"""
import os
import sys
import argparse

# 使脚本可 import 同目录的 llm_client
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import llm_client

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

MAX_REPORT_CHARS = 20000   # 报告完整加载上限
MAX_FILE_CHARS = 15000     # /read 单文件加载上限
PREVIEW_CHARS = 2000       # 目录内非报告文本文件的预览上限

# 报告类文件（目录内完整读取）
REPORT_EXTS = {".md"}
# 文本预览类文件（目录内读取前若干字符）
TEXT_EXTS = {".txt", ".csv", ".log", ".json", ".html", ".htm", ".xml",
             ".py", ".sh", ".c", ".h", ".cpp", ".java", ".kt"}
# 明确不支持直接读取的二进制文件
BINARY_EXTS = {".xlsx", ".xls", ".png", ".jpg", ".jpeg", ".gif", ".bmp",
               ".pdf", ".zip", ".gz", ".tar", ".bin", ".pyc", ".exe",
               ".dll", ".so", ".proto", ".perfetto", ".perfetto-trace", ".db"}


def find_reports(trace_name):
    """返回 out/<trace>/stall_analysis 下所有 io_stall_report.md 路径（按修改时间降序）。"""
    stall_dir = os.path.join(BASE, "out", trace_name, "stall_analysis")
    reports = []
    if os.path.isdir(stall_dir):
        for name in sorted(os.listdir(stall_dir)):
            d = os.path.join(stall_dir, name)
            if os.path.isdir(d):
                rp = os.path.join(d, "io_stall_report.md")
                if os.path.isfile(rp):
                    reports.append(rp)
    reports.sort(key=os.path.getmtime, reverse=True)
    return reports


def load_report(report_path, max_chars=MAX_REPORT_CHARS):
    """读取报告内容，超长则截断（防止超出模型上下文窗口）。"""
    with open(report_path, "r", encoding="utf-8") as f:
        content = f.read()
    if len(content) > max_chars:
        content = content[:max_chars] + "\n\n[报告过长，已截断]"
    return content


def build_system_prompt(report_content):
    prompt = (
        "你是一名资深存储 IO 性能分析专家，精通 Android / Linux block IO 分析。\n"
        "用户会围绕 TRACE 的 IO 分析报告 / 数据向你提问，请结合上下文中的具体数据作答。\n"
        "用户可能会提供本地文件或目录路径，其内容会由系统读取后附在消息中，请基于这些内容作答。\n"
        "回答要求：\n"
        "  1. 简洁、准确，优先引用报告中的具体数字，并注明出处章节；\n"
        "  2. 若上下文信息不足以回答，请明确指出缺失的证据，不要编造；\n"
        "  3. 对报告中的判断若有不同看法，可给出修正建议并说明理由。\n"
    )
    if report_content:
        prompt += "\n\n===== 分析报告上下文 =====\n\n" + report_content
    return prompt


def resolve_path(raw):
    """解析用户输入的路径：支持绝对路径 / 相对项目根路径，剥离首尾引号。"""
    raw = (raw or "").strip().strip('"').strip("'")
    if not raw:
        return None
    if os.path.isabs(raw):
        return os.path.normpath(raw)
    return os.path.normpath(os.path.join(BASE, raw))


def read_text_file(path, max_chars):
    """读取文本文件，返回 (content, truncated)。二进制或失败返回 (None, err)。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in BINARY_EXTS:
        return None, "二进制文件（%s）不支持直接读取，请先转换为文本" % ext
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read()
    except Exception as e:
        return None, str(e)
    truncated = len(content) > max_chars
    if truncated:
        content = content[:max_chars]
    return content, truncated


def collect_dir(path):
    """列出目录并收集其中文本文件内容。返回 (desc, report_texts, preview_texts)。"""
    try:
        entries = sorted(os.listdir(path))
    except Exception as e:
        return None, [], []

    desc_lines = ["目录 %s 共 %d 项:" % (path, len(entries))]
    report_texts = []
    preview_texts = []

    for e in entries:
        p = os.path.join(path, e)
        if os.path.isdir(p):
            desc_lines.append("  - [目录] %s" % e)
            continue
        size = os.path.getsize(p)
        desc_lines.append("  - %s (%d 字节)" % (e, size))

        ext = os.path.splitext(e)[1].lower()
        if ext in REPORT_EXTS:
            content, truncated = read_text_file(p, MAX_REPORT_CHARS)
            if content is not None:
                note = "\n\n[内容过长，已截断]" if truncated else ""
                report_texts.append("===== 文件: %s =====\n%s%s" % (e, content, note))
        elif ext in TEXT_EXTS:
            content, truncated = read_text_file(p, PREVIEW_CHARS)
            if content is not None:
                note = ("\n\n[仅预览前 %d 字符，可用 /read 读取完整]" % PREVIEW_CHARS) if truncated else ""
                preview_texts.append("===== 文件: %s (预览) =====\n%s%s" % (e, content, note))

    return "\n".join(desc_lines), report_texts, preview_texts


def make_file_message(path):
    """构造解读单个文件的消息文本。"""
    content, truncated = read_text_file(path, MAX_FILE_CHARS)
    if content is None:
        return "用户请求解读本地文件: %s\n\n[系统提示] 无法读取该文件: %s" % (path, truncated)
    note = "\n\n[内容过长，已截断]" if truncated else ""
    return ("用户请求解读本地文件: %s\n\n===== 文件内容 =====\n%s%s\n\n请解读该文件的内容。"
            % (path, content, note))


def ask_ai(messages, model):
    """发起一次模型调用并打印回复，失败时回滚未答复的提问。"""
    print("AI > ", end="", flush=True)
    try:
        content, _reasoning, _usage = llm_client.chat(messages, role=model)
        print(content)
        messages.append({"role": "assistant", "content": content})
    except Exception as e:
        print("\n[错误] %s" % e)
        if messages and messages[-1]["role"] == "user":
            messages.pop()


def print_help():
    print(
        "  命令:\n"
        "    /quit          退出\n"
        "    /clear         清空对话历史 (报告上下文保留)\n"
        "    /help          显示本帮助\n"
        "    /ls [目录]     列出目录内容 (默认当前 trace 输出目录)\n"
        "    /read <文件>   读取文本文件内容并让 AI 解读\n"
        "    直接粘贴一个存在的文件/目录路径，会自动读取并解读"
    )


def main():
    ap = argparse.ArgumentParser(description="IO 报告 AI 问答 (AI2)")
    ap.add_argument("--trace", required=True, help="trace 名称 (如 com.ss.android.ugc.aweme)")
    ap.add_argument("--model", default="fast", choices=["fast", "deep"],
                    help="模型角色: fast=deepseek-chat, deep=deepseek-reasoner (默认 fast)")
    ap.add_argument("--report", default=None, help="指定报告路径 (可选，默认最新 stall 报告)")
    args = ap.parse_args()

    # 定位报告上下文
    report_content = ""
    if args.report:
        if not os.path.isfile(args.report):
            print("[错误] 报告不存在: %s" % args.report)
            sys.exit(1)
        report_content = load_report(args.report)
        print("[上下文] 已加载报告: %s (%d 字符)" % (args.report, len(report_content)))
    else:
        reports = find_reports(args.trace)
        if reports:
            report_content = load_report(reports[0])
            rel = os.path.relpath(reports[0], BASE)
            print("[上下文] 已加载报告: %s (%d 字符)" % (rel, len(report_content)))
            if len(reports) > 1:
                print("[提示] 共 %d 份 stall 报告，默认最新一份，可用 --report 指定其他" % len(reports))
        else:
            print("[提示] 未找到 stall_analysis 报告，将以无报告上下文模式开始 (仅回答通用 IO 问题)")

    system_prompt = build_system_prompt(report_content)
    messages = [{"role": "system", "content": system_prompt}]

    model_label = (llm_client.DEFAULT_DEEP_MODEL if args.model == "deep"
                   else llm_client.DEFAULT_FAST_MODEL)
    print("\n[AI2] 已进入 IO 报告问答模式 (模型: %s)" % model_label)
    print("  输入问题即可对话；可直接粘贴文件/目录路径；输入 /help 查看全部命令\n")

    while True:
        try:
            user = input("你 > ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if not user:
            continue

        low = user.lower()

        if low in ("/quit", "/exit", "quit", "exit", "q"):
            print("再见!")
            break
        if low in ("/help", "help", "?"):
            print_help()
            continue
        if low in ("/clear", "clear"):
            messages = [{"role": "system", "content": system_prompt}]
            print("  已清空对话历史 (报告上下文保留)")
            continue

        # /ls 列出目录
        if low.startswith("/ls") or low.startswith("/dir"):
            arg = user.split(None, 1)[1].strip() if len(user.split(None, 1)) > 1 else ""
            if arg:
                path = resolve_path(arg)
            else:
                path = os.path.join(BASE, "out", args.trace)
            if not path or not os.path.isdir(path):
                print("  [错误] 目录不存在: %s" % (arg or path))
            else:
                desc, _report_texts, _preview_texts = collect_dir(path)
                print("  " + (desc or "(空目录)").replace("\n", "\n  "))
            continue

        # /read 读取文件
        if low.startswith("/read"):
            arg = user.split(None, 1)[1].strip() if len(user.split(None, 1)) > 1 else ""
            if not arg:
                print("  用法: /read <文件路径>")
                continue
            path = resolve_path(arg)
            if not path or not os.path.isfile(path):
                print("  [错误] 文件不存在: %s" % arg)
                continue
            messages.append({"role": "user", "content": make_file_message(path)})
            ask_ai(messages, args.model)
            continue

        # 自动路径检测：直接粘贴存在的文件/目录路径
        path = resolve_path(user)
        if path and os.path.exists(path):
            if os.path.isdir(path):
                desc, report_texts, preview_texts = collect_dir(path)
                msg = "用户请求解读本地目录: %s\n\n%s" % (path, desc or "(空目录)")
                for t in report_texts:
                    msg += "\n\n" + t
                for t in preview_texts:
                    msg += "\n\n" + t
                msg += "\n\n请解读该目录的内容。"
            else:
                msg = make_file_message(path)
            messages.append({"role": "user", "content": msg})
            ask_ai(messages, args.model)
            continue

        # 正常对话
        messages.append({"role": "user", "content": user})
        ask_ai(messages, args.model)


if __name__ == "__main__":
    main()