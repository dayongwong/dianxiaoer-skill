#!/usr/bin/env python3
"""
高保真 Excel 选品与供应链决策报表导出 CLI 脚本
调用店小二服务端 /api/v1/agent/tasks/export-report 接口
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from dianxiaoer_auth import ActionRequired, authenticated_request, DEFAULT_BASE_URL


def main():
    parser = argparse.ArgumentParser(description="导出店小二选品与 1688 比价决策 Excel 报表")
    parser.add_argument("--product-ids", default="", help="指定导出的商品 ID 列表（逗号分隔，如 1,2,3）")
    parser.add_argument("--limit", type=int, default=50, help="最多导出的商品数量上限")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="店小二服务地址")
    args = parser.parse_args()

    product_ids = []
    if args.product_ids.strip():
        try:
            product_ids = [int(x.strip()) for x in args.product_ids.split(",") if x.strip()]
        except ValueError:
            print(json.dumps({"error": True, "msg": "product-ids 格式不正确，应为数字逗号分隔"}, ensure_ascii=False), file=sys.stderr)
            return 1

    payload = {
        "product_ids": product_ids if product_ids else None,
        "limit": args.limit
    }

    try:
        res = authenticated_request(
            endpoint_path="/api/v1/agent/tasks/export-report",
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
