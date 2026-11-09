#!/usr/bin/env python3
"""增量更新链上指标 CSV（MVRV Ratio、Realized Price、NUPL）。

数据源优先级：
  1. CryptoQuant API（需要 CRYPTOQUANT_KEY Secret；key 撤销/过期时自动降级）
  2. btc-signal-pro 的 data/ CSV（其 CI 持有 CRYPTOQUANT_KEY 每日更新，格式与本
     仓库完全一致；raw.githubusercontent.com 免 key 拉取）
  3. bitcoin-data.com（免费、无需 key；限流 10 请求/小时，本脚本一次跑只需 3 个；
     2026-11 上旬曾断更一周——上游自身停更，不能作唯一兜底）

MVRV/RP/NUPL 三个源都有（值与 CryptoQuant 有 ~1% 方法差异，拼接处会有轻微台阶，
对趋势/阈值类分析无影响）。输出保持原格式：逗号分隔、降序、最新在最上、日期形如
2026-07-15T00:00:00Z。幂等：重复运行不会重复写入已有日期。

CryptoQuant API key 从环境变量 CRYPTOQUANT_KEY 读取（GitHub Actions 里配为 Secret）。
**绝不硬编码 key**：前端源码公开，key 只能存在于服务端 Actions 环境。

用法:
    # 按 cryptoquant → btc-signal-pro → bitcoin-data.com 顺序取第一个有新数据的源
    python scripts/update_onchain.py
    # 强制使用某个源（调试用）
    SOURCE=cryptoquant    python scripts/update_onchain.py
    SOURCE=signal-repo    python scripts/update_onchain.py
    SOURCE=bitcoin-data   python scripts/update_onchain.py
"""
import datetime
import json
import os
import sys
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
CRYPTOQUANT_API = "https://api.cryptoquant.com/v1"
BITCOIN_DATA_API = "https://bitcoin-data.com/v1"
SIGNAL_REPO_RAW = "https://raw.githubusercontent.com/hhysteric/btc-signal-pro/main/data"

# 每个指标：本地文件、表头、CryptoQuant endpoint+字段、bitcoin-data.com 路径+字段
SERIES = [
    {
        "file": "mvrv.csv",
        "header": "Datetime,MVRV Ratio",
        "cq_endpoint": "btc/market-indicator/mvrv",
        "cq_field": "mvrv",
        "bd_path": "mvrv",
        "bd_field": "mvrv",
    },
    {
        "file": "realized_price.csv",
        "header": "Datetime,Realized Price",
        "cq_endpoint": "btc/market-indicator/realized-price",
        "cq_field": "realized_price",
        "bd_path": "realized-price",
        "bd_field": "realizedPrice",
    },
    {
        "file": "nupl.csv",
        "header": "Datetime,Net Unrealized Profit/Loss (NUPL)",
        "cq_endpoint": "btc/network-indicator/nupl",
        "cq_field": "nupl",
        "bd_path": "nupl",
        "bd_field": "nupl",
    },
]


def _http_get(url, headers=None, timeout=30):
    req = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def fetch_cryptoquant(endpoint, field, key, start_date, today):
    """CryptoQuant：拉取 [start_date, today] 区间，返回 {date_iso: value}。"""
    frm = start_date.strftime("%Y%m%d")
    to = today.strftime("%Y%m%d")
    url = f"{CRYPTOQUANT_API}/{endpoint}?window=day&from={frm}&to={to}&limit=100000"
    raw = _http_get(url, headers={
        "Authorization": f"Bearer {key}",
        "User-Agent": "btc-cycle-dashboard/1.0",
    })
    data = json.loads(raw)
    if data.get("status", {}).get("code") != 200:
        raise RuntimeError(f"CryptoQuant API status {data.get('status')}")
    out = {}
    for item in data.get("result", {}).get("data", []):
        d = item.get("date")
        v = item.get(field)
        if d and v is not None:
            out[d[:10]] = v
    return out


def fetch_signal_repo(filename, start_date, today):
    """btc-signal-pro 的 data/ CSV：其 CI 持有 CRYPTOQUANT_KEY 每日更新，格式与本
    仓库完全一致（同表头、降序、T00:00:00Z 日期）。作为第二兜底——本仓库缺 key、
    bitcoin-data.com 也断更时仍能拿到 CryptoQuant 口径的值。

    注意这是跨 repo 依赖：signal-pro 的 CI 若停更，这里同样拿不到新数据（由
    兜底链末端的 bitcoin-data.com 与 workflow 的 staleness guard 兜住）。"""
    url = f"{SIGNAL_REPO_RAW}/{filename}"
    raw = _http_get(url, headers={"User-Agent": "btc-cycle-dashboard/1.0"})
    out = {}
    for line in raw.decode("utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("Datetime"):
            continue
        cols = line.split(",")
        if len(cols) < 2 or not cols[1].strip():
            continue
        d = cols[0][:10]
        try:
            dt = datetime.date.fromisoformat(d)
        except ValueError:
            continue
        if start_date <= dt <= today:
            out[d] = cols[1].strip()
    return out


def fetch_bitcoin_data(path, field, start_date, today):
    """bitcoin-data.com：返回 {date_iso: value}。

    一次返回最近 ~4 年的升序数组 [{"d": "2026-11-01", "unixTs": ..., field: v}, ...]，
    客户端自行裁切 [start_date, today]。4 年窗口足够覆盖增量更新（缺数 >4 年时
    该源无法回填，需要人工处理）。
    """
    url = f"{BITCOIN_DATA_API}/{path}"
    raw = _http_get(url, headers={"User-Agent": "btc-cycle-dashboard/1.0"})
    data = json.loads(raw)
    if not isinstance(data, list):
        raise RuntimeError(f"bitcoin-data.com 返回非数组: {str(data)[:120]}")
    out = {}
    for item in data:
        d = item.get("d")
        v = item.get(field)
        if not d or v is None:
            continue
        day = d[:10]
        try:
            dt = datetime.date.fromisoformat(day)
        except ValueError:
            continue
        if dt < start_date or dt > today:
            continue
        out[day] = v
    return out


def read_csv(path):
    with open(path, encoding="utf-8-sig") as f:
        lines = [l.rstrip("\n") for l in f if l.strip()]
    return lines[0], lines[1:]  # header, rows (descending)


def date_of(row):
    return row.split(",")[0][:10]


def newest_valid_date(rows):
    for row in rows:
        cols = row.split(",")
        if len(cols) >= 2 and cols[1].strip():
            return datetime.date.fromisoformat(date_of(row))
    return None


def append_rows(path, header, rows, fresh):
    """fresh: {date_iso: value}，追加到 CSV 顶部（保持降序）。"""
    new_rows = [f"{d}T00:00:00Z,{fresh[d]}" for d in sorted(fresh)]
    new_rows_desc = list(reversed(new_rows))
    out = [header] + new_rows_desc + rows
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out) + "\n")
    return len(new_rows), date_of(new_rows_desc[0])


def update_one(series, cq_key, today, force_source=None):
    path = os.path.join(DATA_DIR, series["file"])
    if not os.path.exists(path):
        print(f"[{series['file']}] 不存在，跳过")
        return False

    header, rows = read_csv(path)
    newest = newest_valid_date(rows)
    if newest is None:
        print(f"[{series['file']}] 无有效数据行，跳过")
        return False
    if newest >= today:
        print(f"[{series['file']}] 已是最新（{newest}），无需更新")
        return False

    start = newest + datetime.timedelta(days=1)

    # 源顺序：cryptoquant（可用时）→ btc-signal-pro → bitcoin-data.com
    sources = []
    if cq_key or force_source == "cryptoquant":
        sources.append(("cryptoquant", lambda s, t: fetch_cryptoquant(
            series["cq_endpoint"], series["cq_field"], cq_key, s, t)))
    if force_source in (None, "signal-repo"):
        sources.append(("btc-signal-pro", lambda s, t: fetch_signal_repo(
            series["file"], s, t)))
    if force_source in (None, "bitcoin-data"):
        sources.append(("bitcoin-data.com", lambda s, t: fetch_bitcoin_data(
            series["bd_path"], series["bd_field"], s, t)))

    for source_name, fetcher in sources:
        try:
            fresh = fetcher(start, today)
        except Exception as e:
            print(f"[{series['file']}] {source_name} 不可用: {e}")
            continue
        fresh = {d: v for d, v in fresh.items()
                 if datetime.date.fromisoformat(d) > newest}
        if not fresh:
            print(f"[{series['file']}] {source_name} 可达但无新数据（最新 {newest}）")
            continue
        n, newest_d = append_rows(path, header, rows, fresh)
        print(f"[{series['file']}] {source_name}: 已追加 {n} 天，最新 {newest_d}")
        return True

    print(f"[{series['file']}] 所有数据源均失败")
    return False


def main():
    cq_key = os.environ.get("CRYPTOQUANT_KEY", "").strip()
    force = os.environ.get("SOURCE", "").strip().lower() or None
    if force and force not in ("cryptoquant", "signal-repo", "bitcoin-data"):
        print(f"未知 SOURCE={force!r}，期望 cryptoquant|signal-repo|bitcoin-data")
        return 1

    today = datetime.datetime.now(datetime.timezone.utc).date()
    print(f"今日(UTC): {today}")
    if cq_key:
        print("检测到 CRYPTOQUANT_KEY，优先使用 CryptoQuant")
    else:
        print("未设置 CRYPTOQUANT_KEY，兜底顺序: btc-signal-pro → bitcoin-data.com")

    changed = False
    for series in SERIES:
        if update_one(series, cq_key, today, force):
            changed = True
    print("有更新" if changed else "无更新")
    return 0


if __name__ == "__main__":
    sys.exit(main())
