import os
import sys
import subprocess
import glob
import argparse


def get_project_root():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(script_dir)


def main():
    parser = argparse.ArgumentParser(description="格式转换: perfetto → systrace")
    parser.add_argument("--trace", type=str, default=None,
                        help="指定单个 trace 文件名（如 xxx.perfetto-trace 或 xxx.perfetto），不指定则转换 trace/ 目录下全部")
    args = parser.parse_args()

    project_root = get_project_root()
    trace_dir = os.path.join(project_root, "trace")
    out_dir = os.path.join(project_root, "out")
    trace_processor_shell = os.path.join(
        project_root, "bin", "windows-amd64", "trace_processor_shell.exe"
    )

    if not os.path.isdir(trace_dir):
        print(f"[错误] trace 目录不存在: {trace_dir}")
        sys.exit(1)

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
        print(f"[选择] 仅转换: {target}")

    os.makedirs(out_dir, exist_ok=True)

    success_count = 0
    for trace_file in trace_files:
        trace_name = os.path.splitext(os.path.basename(trace_file))[0]
        if trace_name.endswith(".perfetto"):
            trace_name = trace_name[:-len(".perfetto")]

        trace_out_dir = os.path.join(out_dir, trace_name)
        os.makedirs(trace_out_dir, exist_ok=True)

        systrace_output = os.path.join(trace_out_dir, f"{trace_name}.systrace")

        print(f"[转换] {os.path.basename(trace_file)} -> {systrace_output}")

        try:
            result = subprocess.run(
                [
                    trace_processor_shell,
                    "convert", "systrace",
                    trace_file,
                    systrace_output,
                ],
                capture_output=True,
                timeout=600,
                cwd=project_root,
            )

            if result.returncode != 0:
                print(f"[警告] trace_processor_shell 返回非零退出码: {result.returncode}")
                stderr_text = result.stderr.decode("utf-8", errors="replace").strip()
                if stderr_text:
                    print(f"  stderr: {stderr_text}")
                continue

            output_size = os.path.getsize(systrace_output)
            if output_size == 0:
                print(f"[警告] 输出文件为空，转换可能失败")
                continue

            success_count += 1
            print(f"[完成] 输出文件: {systrace_output} ({output_size:,} bytes)")

        except subprocess.TimeoutExpired:
            print(f"[错误] trace_processor_shell 执行超时（10分钟）")
        except Exception as e:
            print(f"[错误] 转换失败: {e}")

    print(f"\n所有文件转换完成，成功 {success_count}/{len(trace_files)} 个文件。")


if __name__ == "__main__":
    main()