#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检测最新 JLST 多/空信号并通过 Buttondown 邮件广播提醒订阅者。

设计：
- 信号源为 data/btc_v2_daily.json（主策略 v2）的 signals_history 最后一条事件。
- 用 data/last_alert.json 记录「上次已提醒事件的日期」，仅当最新事件日期更晚时
  才发信——幂等，避免每日重复发送或漏发。
- 发信调用封装在 send_broadcast() 里，便于日后核对 Buttondown API 细节或整体
  更换服务商。API key 从环境变量 BUTTONDOWN_API_KEY 读取，未设置则只打印预览、
  跳过实际发信（本地跑不会误发）。

用法：
    BUTTONDOWN_API_KEY=xxx python scripts/notify_signals.py   # 真实发信
    python scripts/notify_signals.py                          # 仅预览（无 key）
"""
import json
import os
import sys
import urllib.request
import urllib.error
from pathlib import Path

# 信号源与状态文件
SIGNAL_FILE = "btc_v2_daily.json"
STATE_FILE = "last_alert.json"

SITE_URL = "https://hhysteric.github.io/btc-momentum-pro/"

BUTTONDOWN_API = "https://api.buttondown.email/v1/emails"

TYPE_LABEL = {
    "long": "做多",
    "short": "做空",
    "cascade": "级联强信号",
}


def _data_dir():
    return Path(__file__).resolve().parent.parent / "data"


def load_latest_event():
    """返回 signals_history 的最后一条事件，无则 None。"""
    path = _data_dir() / SIGNAL_FILE
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    history = data.get("signals_history") or []
    return history[-1] if history else None


def load_last_alert_date():
    """返回上次已提醒事件的日期字符串（YYYY-MM-DD），无则 None。"""
    path = _data_dir() / STATE_FILE
    if not path.exists():
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f).get("last_alert_date")
    except (json.JSONDecodeError, OSError):
        return None


def save_last_alert_date(date_str):
    path = _data_dir() / STATE_FILE
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"last_alert_date": date_str}, f, ensure_ascii=False, indent=2)


def build_email(event):
    """根据信号事件组装 (subject, body_markdown)。"""
    direction = TYPE_LABEL.get(event["type"], event["type"])
    date = event["date"]
    price = event.get("price")
    comp = event.get("composite")
    multi = event.get("multi_pass")

    subject = f"BTC Momentum · 新{direction}信号 {date}"

    if multi is True:
        multi_line = "三重 MA 滤波：**全部通过（高质量信号）**"
    elif multi is False:
        multi_line = "三重 MA 滤波：未全部通过（原始信号，注意逆势风险）"
    else:
        multi_line = "三重 MA 滤波：不适用"

    price_str = f"${price:,.0f}" if isinstance(price, (int, float)) else "—"
    comp_str = f"{comp:+.3f}" if isinstance(comp, (int, float)) else "—"

    body = f"""JLST 反转-动量模型在 **{date}** 触发了新的**{direction}**信号。

- 信号方向：{direction}
- 触发日价格：{price_str}
- 综合评分：{comp_str}
- {multi_line}

打开仪表盘查看 K 线、指标与信号历史：
{SITE_URL}

—— BTC Momentum

风险提示：本信号来自量化模型的历史回测框架，不构成投资建议。历史表现不代表未来收益，做空在牛市中天然逆风、胜率偏低，请自行判断仓位与风险。
"""
    return subject, body


def send_broadcast(subject, body, api_key):
    """通过 Buttondown 广播 API 给全部订阅者发信。

    注意：Buttondown API 端点/字段/状态值以实现时官方文档为准。
    """
    payload = json.dumps({
        "subject": subject,
        "body": body,
        "status": "about_to_send",
    }).encode("utf-8")

    req = urllib.request.Request(
        BUTTONDOWN_API,
        data=payload,
        method="POST",
        headers={
            "Authorization": f"Token {api_key}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status, resp.read().decode("utf-8", errors="replace")


def main():
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    event = load_latest_event()
    if event is None:
        print("[notify] signals_history 为空，跳过。")
        return

    latest_date = event["date"]
    last_alert = load_last_alert_date()

    print(f"[notify] 最新信号事件：{latest_date} ({event['type']})")
    print(f"[notify] 上次已提醒：{last_alert or '（无记录）'}")

    # 幂等：仅当最新事件日期严格晚于上次已提醒日期时才发信
    if last_alert is not None and latest_date <= last_alert:
        print("[notify] 无新信号，安静退出。")
        return

    subject, body = build_email(event)
    api_key = os.environ.get("BUTTONDOWN_API_KEY")

    if not api_key:
        print("[notify] 未设置 BUTTONDOWN_API_KEY，跳过实际发信。以下为将要发送的邮件预览：")
        print("=" * 60)
        print("Subject:", subject)
        print("-" * 60)
        print(body)
        print("=" * 60)
        print("[notify] （本地预览模式，不写回 last_alert.json）")
        return

    try:
        status, resp_body = send_broadcast(subject, body, api_key)
        print(f"[notify] Buttondown 响应：HTTP {status}")
    except urllib.error.HTTPError as e:
        print(f"[notify][error] 发信失败：HTTP {e.code} — {e.read().decode('utf-8', errors='replace')}")
        sys.exit(1)
    except urllib.error.URLError as e:
        print(f"[notify][error] 发信失败：{e.reason}")
        sys.exit(1)

    save_last_alert_date(latest_date)
    print(f"[notify] 已发送并更新状态：last_alert_date = {latest_date}")


if __name__ == "__main__":
    main()
