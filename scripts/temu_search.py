#!/usr/bin/env python3
"""
Temu 选品数据挖掘 CLI 脚本
调用店小二服务端 /api/v1/agent/tasks/temu-harvest 接口
"""

import argparse
import json
import sys
from pathlib import Path

# 添加当前脚本目录至 sys.path 以加载 dianxiaoer_auth
sys.path.insert(0, str(Path(__file__).resolve().parent))
from dianxiaoer_auth import ActionRequired, authenticated_request, DEFAULT_BASE_URL


def main():
    parser = argparse.ArgumentParser(description="调用店小二 Temu 跨境智能选品挖掘")
    parser.add_argument("--keyword", required=True, help="选品搜索关键词（如：连衣裙、无线充）")
    parser.add_argument("--price-min", type=float, default=0.0, help="最低价格（美元）")
    parser.add_argument("--price-max", type=float, default=200.0, help="最高价格（美元）")
    parser.add_argument("--min-sales", type=int, default=50, help="最低销量门槛")
    parser.add_argument("--max-items", type=int, default=20, help="抓取候选商品数量上限")
    parser.add_argument("--live", dest="live", action="store_true", default=True, help="优先启用本地 Chrome 9222 真实 CDP 实时截包采流")
    parser.add_argument("--no-live", dest="live", action="store_false", help="禁用实时 CDP，直接使用本地样本库")
    parser.add_argument("--auto-launch-chrome", action="store_true", default=False, help="若 Chrome 9222 未启动，自动拉起独立 Chrome 调试实例")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="店小二服务地址")
    args = parser.parse_args()

    payload = {
        "keyword": args.keyword,
        "price_min": args.price_min,
        "price_max": args.price_max,
        "min_sales": args.min_sales,
        "max_items": args.max_items,
        "live_scrape": args.live,
        "auto_launch_chrome": args.auto_launch_chrome
    }

    try:
        res = authenticated_request(
            endpoint_path="/api/v1/agent/tasks/temu-harvest",
            base_url=args.base_url,
            method="POST",
            body=payload
        )
    except ActionRequired as act:
        print(json.dumps(act.public_payload(), ensure_ascii=False, indent=2))
        return 2
    except Exception as err:
        print(json.dumps({"error": True, "msg": str(err)}, ensure_ascii=False), file=sys.stderr)
        return 1

    print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
