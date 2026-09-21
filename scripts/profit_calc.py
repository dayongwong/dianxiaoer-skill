#!/usr/bin/env python3
"""
综合跨境毛利与回本测算 CLI 脚本
调用店小二服务端 /api/v1/agent/tasks/calculate-profit 接口
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from dianxiaoer_auth import ActionRequired, authenticated_request, DEFAULT_BASE_URL


def main():
    parser = argparse.ArgumentParser(description="调用店小二跨境商品利润精算模型")
    parser.add_argument("--temu-price", type=float, required=True, help="Temu 目标售价/供货价（美元）")
    parser.add_argument("--supplier-price", type=float, required=True, help="1688 采购底价（人民币）")
    parser.add_argument("--exchange-rate", type=float, default=7.25, help="汇率（默认 7.25）")
    parser.add_argument("--shipping-fee", type=float, default=18.0, help="预估跨境头程+国内揽收物流费（元）")
    parser.add_argument("--commission-rate", type=float, default=0.15, help="平台扣点率（默认 15%%）")
    parser.add_argument("--packaging-cost", type=float, default=2.0, help="包装与贴标耗材成本（元）")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="店小二服务地址")
    args = parser.parse_args()

    payload = {
        "temu_price_usd": args.temu_price,
        "supplier_price_cny": args.supplier_price,
        "exchange_rate": args.exchange_rate,
        "shipping_fee_cny": args.shipping_fee,
        "commission_rate": args.commission_rate,
        "packaging_cost_cny": args.packaging_cost
    }

    try:
        res = authenticated_request(
            endpoint_path="/api/v1/agent/tasks/calculate-profit",
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
