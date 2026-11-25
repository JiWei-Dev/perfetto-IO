---
name: "io-stall-analysis"
description: "资深存储专家视角的 IO 卡顿(Uninterruptible Sleep / D 状态)根因分析方法论, 覆盖 insert→issue→complete 三段式拆解与调度/器件双路排查. 当用户提供一段 IO 卡顿的 Start time/Duration/Process 要求定位耗时根因时使用."
---

# IO 卡顿根因分析方法论（存储专家版）

> 适用对象：`Uninterruptible Sleep (IO)` / D 状态线程卡顿、单笔慢 IO、应用 IO 长尾延迟。
> 配套脚本：`scripts/AI/io_stall_analysis.py`（输入 Start time / Duration / Process 一键出报告）。

## 一、总框架：先认清「单笔 IO 生命周期」与「线程卡顿」不是一回事

一个块 IO 请求在内核 block 层有三个关键时间点，这也是 trace 里三类事件的来源：

| 事件 | 含义 | 阶段 |
|---|---|---|
| `block_rq_insert` | 请求进入 IO 调度器队列 | 起点 |
| `block_rq_issue` | 请求被下发到设备驱动 | 调度完成 / 器件开始 |
| `block_rq_complete` | 请求被设备完成 | 终点 |

由此得到两个独立阶段：

- **调度耗时 = issue − insert**（瓶颈在 Block Layer / IO 调度器 / CPU 抢占 / 队列排队）
- **器件耗时 = complete − issue**（瓶颈在存储硬件）

**关键认知**：线程在 D 状态卡了 32ms，往往**不等于**某一笔 IO 的端到端耗时是 32ms。D 状态线程等待的是「它提交的那笔请求的 complete」，而 complete 迟迟不来，是因为**器件在这段时间整体停滞**（GC/写缓存刷新），而不是这一笔 IO 本身跑得慢。所以拆解的第一步永远是：先看这段时间器件有没有 `complete`，而不是盯着单笔 IO 的 latency。

## 二、四维度标准拆解流程

### 维度 1：该笔 IO 的完整情况（基本信息）

要回答三个「谁、何时、上下文」：
- **insert**：什么时间、由哪个进程/线程（PID/TID/线程名）发起，CPU 是多少；
- **issue**：什么时候被 kworker/内核下发；
- **complete**：什么时候完成。

数据来源：`block_io.xlsx` 的「全部 IO」Sheet，按 `(设备号, 请求扇区地址)` 分组做 insert→issue→complete 三段式配对。注意 `complete` 事件里记录的进程是「中断返回时正在运行的上下文」，不一定是 IO 发起者，不要误判。

### 维度 2：耗时点统计（调度 vs 器件占比）

- 调度耗时占比高（> 器件耗时的 10%~30%）：瓶颈在 OS 侧；
- 器件耗时占比高：瓶颈在存储硬件。

**陷阱 1**：当器件停滞时，后续请求全部积压在调度队列，表现为「调度耗时虚高」。所以「调度耗时高」不必然等于「调度器有问题」，必须结合维度 4 的 complete 断档一起判断——很多场景下调度耗时只是「排队等设备恢复」的投影。

**陷阱 2（complete 时间戳 ≠ 硬件完成时刻）**：`block_rq_complete` 记录的是块层**软中断执行完成回调的时刻**，而不是硬件真正完成 IO 的时刻。硬件完成 IO 后先产生硬件中断（IRQ），再触发软中断（softirq / BLK_SOFTIRQ），软中断里才调用 `blk_complete_request` 打出 `block_rq_complete`。若硬件中断早已回来、但 CPU 没空运行软中断，complete 时间戳就会偏晚，导致「器件耗时 = complete − issue」被**高估**——此时器件耗时里混入了「中断返回 + 软中断调度」的 CPU 延迟，根因其实在 CPU 侧而非存储硬件。

### 维度 3：调度慢的根因排查

用户容易想到的三条：CPU 得不到资源 / 请求队列已满 / 被后来请求插队。逐条定位：

- **CPU 不足**：看 `issue` 事件的 `运行 CPU` 字段。若 issue 长期不出现、且 kworker 线程无 CPU 记录，说明 block 工作线程没被调度到。
- **队列满**：统计 START 时刻 `in-flight`（已 issue 未 complete）请求数。若接近设备命令队列上限（**UFS 典型 32 / eMMC 典型 4**），驱动已无法接收新请求，后续只能在调度器排队。注意 eMMC 通常只能同时 issue 4 个请求，in-flight ≥ 4 即可能已打满。
- **插队**：**「插队」的正确定义是——后面 insert 的 IO 反而先 issue 了**（insert 顺序 vs issue 顺序被打乱）。判断时不要只看「有几个线程在提交」，而要看目标 IO 的 insert→issue 之间，是否有更晚 insert 的请求抢在其前面 issue。单一线程也可能因调度器重排/请求合并而出现乱序。

**专家补充（用户常漏掉的调度慢原因）**：
- **cgroup / io.weight 限流**：IO 优先级被压低，即使 CPU 空闲也被 throttling；
- **CFQ/BFQ 调度器时间片与权重**：低权重 cgroup 的请求被长期搁置；
- **请求合并（merge/coalesce）**：insert 的请求扇区被合并后，issue 的是「合并后的大请求」，直接用 insert 的扇区去 match issue 会失配，需按时间窗口而非严格扇区来对齐；
- **plug/unplug 机制**：blk-mq 的 plug 会把请求短暂 hold 住延迟下发，产生几 ms 级调度延迟；
- **设备 tag 耗尽**：blk-mq hardware queue 的 tag 用完后，新请求无法 issue，这是「调度队列堆积但器件队列没满」的典型信号。

### 维度 4：器件慢的根因排查

用户容易想到的两条：这段时间是否有 complete / 是否有读写交叉。精确定位：

- **complete 断档**：START 前最后一个 complete 与 START 后第一个 complete 的时间差。若断档长达几十 ms，说明器件是「停滞」而非「慢」。若 **issue 与 complete 同时断档**，则是设备整体停滞（写缓存耗尽触发 FTL GC/刷盘）；若只有 complete 断档而 issue 不断，则是设备消化能力下降。
- **读写交叉**：看 in-flight 请求的读写类型分布。纯写负载 + 顺序大块，是最容易打穿写缓存、触发 GC 的形态。
- **扇区离散度（LBA 连续性）**：顺序访问 LBA 连续，随机访问 LBA 大幅跳跃。离散度过高会加剧器件开销——eMMC/HDD 需要寻道、UFS 需要跨 die 随机访问，都是器件慢的常见原因。判断指标：窗口内相邻 issue 请求的 LBA 间隙（相对上一请求末尾），连续占比 vs 跳跃占比。

### 甄别：器件真慢 vs 软中断延迟（complete 时间戳被 CPU 推迟）

当 issue→complete 偏大时，**必须先区分两种可能**，否则会把「CPU 调度延迟」误判成「器件慢」：

- **器件真慢**：complete 时间戳均匀分散——每个 IO 的器件执行时间接近，issue 与 complete 各自都均匀分布，无「攒批」现象。
- **软中断延迟**：多个 IO 的 issue 分散，但 complete 时间戳「聚簇」——多个 complete 挤在极短时间（< 50us）内集中处理。这是硬件早已完成、但软中断被延迟后一次性批量处理 complete 的典型信号，此时器件耗时虚高，根因在 CPU 调度而非存储硬件。

判断步骤（三步）：
1. **看 complete 相邻间隔**：统计窗口内连续 complete 的时间间隔，若出现 < 50us 的聚簇相邻对，是软中断批量处理的强信号；
2. **看 complete 记录的进程**：若 complete 事件集中在 `ksoftirqd`（软中断溢出后交给内核线程处理）或 `swapper`/`<kernel>`（中断/idle 上下文），进一步佐证软中断延迟；
3. **对比 issue 间隔 vs complete 间隔**：issue 均匀 + complete 聚簇 = 软中断延迟；两者都均匀 = 器件真实耗时。

**专家补充（用户常漏掉的器件慢原因）**：
- **SLC cache 耗尽 → 速度悬崖**：UFS/SSD 先用 SLC 模式缓存写，写满后回退到 TLC/QLC 原生速度，IOPS 骤降（这是 30~50ms 写停滞最常见根因）；
- **写放大（Write Amplification）**：频繁覆写 + 小 IO → FTL 反复搬迁，放大写压力；
- **GC / 读干扰 / ECC 纠错**：后台 GC 与前台 IO 抢 die/bank 资源，读请求被 ECC 拖累；
- **Thermal Throttling（温度降频）**：长时间高负载设备降频，且往往伴随 issue 速率下降（可看 issue 断档）；
- **命令队列深度过浅**：并发不足以掩盖 NAND 编程延迟（tPROG ~ 数百 us 级），单 die 串行导致有效 IOPS 低；
- **后台 writeback / 脏页回写与前台 IO 竞争**：kworker flush 线程与前台请求抢设备。

## 三、根因判定决策树

1. 先算 **complete 断档时长**，与线程 D 状态 Duration 对比：
   - 高度吻合 → 卡顿 = 设备停滞，直接下「器件侧停滞」结论；
   - 不吻合 → 卡顿另有原因（可能是调度侧、或该线程等待的是其它资源的锁）。
2. 再算 in-flight 深度与堆积请求数：
   - in-flight 接近队列上限（UFS 32 / eMMC 4）+ 堆积 > 0 → 驱动已满，调度队列排队，但根因仍是器件不消化；
   - in-flight 低 + 堆积高 → 真·调度瓶颈（CPU/tag/限流）；
3. 最后看读写形态：
   - 纯顺序大块同步写 → UFS 写停滞（GC/SLC cache 耗尽）特征；
   - 读多写少 + 读尾延迟高 → 读干扰 / ECC / 寻道。

## 四、UFS 写停滞特征速查（本案例结论）

- 顺序大块同步写（如应用启动加载、视频录制、数据库 checkpoint）；
- 出现 30~50ms 级「零 issue / 零 complete」静默期；
- 静默期前 in-flight 请求数冲高到 30+（设备被灌满）；
- 恢复后 complete 集中爆发（设备一次性消化积压）；
- 结论：写缓存耗尽 → 触发 FTL GC / 缓存刷新，属设备固有行为，非单笔 IO 异常。

## 五、常见根因速查表

| 现象 | 根因 | 验证手段 |
|---|---|---|
| issue/complete 同时断档几十 ms | UFS 写停滞（GC/刷盘） | 断档前 in-flight 冲高、纯写负载 |
| 只有 complete 断档 | 设备消化能力下降 | 看 issue 是否持续、读延迟是否升高 |
| 调度耗时虚高 + 器件正常 | 排队等设备恢复 | 结合 complete 断档判断 |
| 调度耗时高 + in-flight 低 | CPU 不足 / cgroup 限流 / tag 耗尽 | 看 issue 事件 CPU、io.weight 配置 |
| 读尾延迟高（P99 >> P50） | 读干扰 / ECC / 寻道 | rw_latency 分布、LBA 离散度 |
| LBA 大幅跳跃、器件耗时高 | 扇区离散度过高（随机访问） | 相邻 LBA 间隙、跳跃占比 |
| 同步写多且卡顿 | fsync 风暴 / SQLite commit | 步骤 7 同步写分析 |

## 六、配套脚本使用

```bash
python scripts/AI/io_stall_analysis.py \
    --trace com.ss.android.ugc.aweme \
    --start 237866030786798 \
    --duration 3840677 \
    --process com.ss.android.ugc.aweme --tid 14831
```

- `--start` / `--duration`：单位纳秒；
- `--process` / `--pid` / `--tid`：三者可只给一个，用于定位责任 IO（tid 优先级最高）；
- `--pre-ms` / `--post-ms`：卡顿点前后分析窗口（默认 20 / 150 ms，complete 在窗口外时需加大）；
- 输出：`out/<trace>/stall_analysis/<start>_<dur>/io_stall_report.md`（主报告）+ `io_stall_data.txt`（原始明细，每个结论的数据来源都在报告第 7 节与明细里对齐）。