#!/usr/bin/env python3
"""多信号 MA 滤波回测分析。

读取 data/btc_v2_daily.json 的 signals_history，对做多/做空信号分别枚举各种
MA 滤波组合，统计每种组合的信号数、各前瞻期（1/3/5/7/30 日）平均收益与胜率。
用于给 PPT/演讲稿提供"多策略效用/超额收益/不足/适用场景"的实证数字。

用法：python scripts/analyze_multisignal.py
"""
import json
import sys
from itertools import combinations
from pathlib import Path

PERIODS = ["fwd_1d", "fwd_3d", "fwd_5d", "fwd_7d", "fwd_30d"]
PERIOD_LABELS = {"fwd_1d": "1日", "fwd_3d": "3日", "fwd_5d": "5日", "fwd_7d": "7日", "fwd_30d": "30日"}


def load_signals(data_dir: Path):
    path = data_dir / "btc_v2_daily.json"
    with open(path, encoding="utf-8") as f:
        j = json.load(f)
    return j.get("signals_history", [])


def stats(signals, period):
    vals = [s[period] for s in signals if s.get(period) is not None]
    if not vals:
        return None
    n = len(vals)
    avg = sum(vals) / n
    wins = sum(1 for v in vals if v > 0)
    wr = wins / n * 100
    return {"n": n, "avg": avg, "wr": wr}


def passes_long(s, filters):
    """做多滤波：要求勾选的条件为 True。"""
    if "ma110" in filters and s.get("ma110_ok") is not True:
        return False
    if "ma6_103" in filters and s.get("ma6_gt_ma103") is not True:
        return False
    if "ema50" in filters and s.get("ema50_ok") is not True:
        return False
    return True


def passes_short(s, filters):
    """做空滤波：要求价格在 MA 下方 → ma110_ok/ema50_ok 为 False。"""
    if "ma110" in filters and s.get("ma110_ok") is not False:
        return False
    if "ema50" in filters and s.get("ema50_ok") is not False:
        return False
    return True


def fmt_row(label, subset):
    cells = []
    for p in PERIODS:
        st = stats(subset, p)
        if st is None:
            cells.append(f"{'--':>16}")
        else:
            cells.append(f"{st['avg']:+6.2f}% / {st['wr']:3.0f}%")
    return f"  {label:<28} n={len(subset):<4} " + "  ".join(cells)


def enumerate_filters(keys):
    """无滤波 + 单条件 + 全条件（含两两组合）。"""
    combos = [frozenset()]
    for r in range(1, len(keys) + 1):
        for c in combinations(keys, r):
            combos.append(frozenset(c))
    return combos


def label_for(kind, fs):
    if not fs:
        return f"{kind}·无滤波"
    name = {"ma110": "价>MA110", "ma6_103": "MA6>MA103", "ema50": "价>EMA50"}
    return f"{kind}·" + "+".join(name[k] for k in ["ma110", "ma6_103", "ema50"] if k in fs)


def main():
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    script_dir = Path(__file__).resolve().parent
    data_dir = script_dir.parent / "data"
    signals = load_signals(data_dir)
    longs = [s for s in signals if s.get("type") == "long"]
    shorts = [s for s in signals if s.get("type") == "short"]

    print("=" * 100)
    print(f"多信号 MA 滤波回测   （signals_history 共 {len(signals)} 条：做多 {len(longs)} / 做空 {len(shorts)}）")
    print("每格为：平均收益% / 胜率%   前瞻期依次为 " + " ".join(PERIOD_LABELS[p] for p in PERIODS))
    print("=" * 100)

    print("\n【做多】")
    for fs in enumerate_filters(["ma110", "ma6_103", "ema50"]):
        subset = [s for s in longs if passes_long(s, fs)]
        print(fmt_row(label_for("多", fs), subset))

    print("\n【做空】（价格在 MA 下方）")
    for fs in enumerate_filters(["ma110", "ema50"]):
        subset = [s for s in shorts if passes_short(s, fs)]
        print(fmt_row(label_for("空", fs), subset))

    # 关键对比：无滤波 vs 全滤波
    print("\n" + "=" * 100)
    print("关键结论：全滤波 vs 无滤波（做多 30 日）")
    raw = stats(longs, "fwd_30d")
    full = stats([s for s in longs if passes_long(s, frozenset(["ma110", "ma6_103", "ema50"]))], "fwd_30d")
    if raw and full:
        print(f"  无滤波 30日: 均值 {raw['avg']:+.2f}%  胜率 {raw['wr']:.0f}%  (n={raw['n']})")
        print(f"  全滤波 30日: 均值 {full['avg']:+.2f}%  胜率 {full['wr']:.0f}%  (n={full['n']})")
        print(f"  超额收益: {full['avg'] - raw['avg']:+.2f} 个百分点，胜率 {full['wr'] - raw['wr']:+.0f} 个百分点")
    print("=" * 100)


if __name__ == "__main__":
    main()
