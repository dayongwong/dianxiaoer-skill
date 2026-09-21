#!/usr/bin/env python3
"""
1688 货源图搜与批量议价 CLI 脚本
调用店小二服务端 /api/v1/agent/tasks/1688-inquiry 接口
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from dianxiaoer_auth import ActionRequired, authenticated_request, DEFAULT_BASE_URL


def main():
    parser = argparse.ArgumentParser(description="调用店小二 1688 批量图搜与自动化议价中枢")
    parser.add_argument("--items-json", default=None, help="商品信息列表 JSON 字符串或 JSON 文件路径")
    parser.add_argument("--buyer-persona", default="Temu全托管成熟大卖", help="买家画像设定")
    parser.add_argument("--title", default=None, help="询盘批次标题")
    parser.add_argument("--auto-negotiate", action="store_true", default=False, help="是否联动本地 Chrome 9222 旺旺执行真机拟人自动发信与议价")
    parser.add_argument("--sync-batch", type=int, default=None, help="同步指定 1688 批次的最新掌柜旺旺回复与底价")
    parser.add_argument("--interval-min", type=int, default=3, help="发问最小间隔秒数")
    parser.add_argument("--interval-max", type=int, default=8, help="发问最大间隔秒数")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="店小二服务地址")
    args = parser.parse_args()

    # 若指定了 --sync-batch，则执行回复同步动作
    if args.sync_batch:
        try:
            res = authenticated_request(
                endpoint_path=f"/api/v1/agent/tasks/1688-inquiry/{args.sync_batch}/sync",
                base_url=args.base_url,
                method="POST"
            )
            print(json.dumps(res, ensure_ascii=False, indent=2))
            return 0
        except ActionRequired as act:
            print(json.dumps(act.public_payload(), ensure_ascii=False, indent=2))
            return 2
        except Exception as err:
            print(json.dumps({"error": True, "msg": str(err)}, ensure_ascii=False), file=sys.stderr)
            return 1

    if not args.items_json:
        print(json.dumps({"error": True, "msg": "缺少 --items-json 参数（或使用 --sync-batch 同步批次回复）"}, ensure_ascii=False), file=sys.stderr)
        return 1

    # 支持直接传入 JSON 字符串或读取本地 JSON 文件
    raw_input = args.items_json.strip()
    if raw_input.startswith("[") or raw_input.startswith("{"):
        try:
            parsed = json.loads(raw_input)
            items = parsed if isinstance(parsed, list) else [parsed]
        except Exception as e:
            print(json.dumps({"error": True, "msg": f"解析 items-json 失败: {e}"}, ensure_ascii=False), file=sys.stderr)
            return 1
    else:
        file_path = Path(raw_input)
        if not file_path.is_file():
            print(json.dumps({"error": True, "msg": f"文件不存在: {raw_input}"}, ensure_ascii=False), file=sys.stderr)
            return 1
        try:
            parsed = json.loads(file_path.read_text(encoding="utf-8"))
            items = parsed if isinstance(parsed, list) else [parsed]
        except Exception as e:
            print(json.dumps({"error": True, "msg": f"读取商品文件失败: {e}"}, ensure_ascii=False), file=sys.stderr)
            return 1

    payload = {
        "items": items,
        "buyer_persona": args.buyer_persona,
        "title": args.title,
        "auto_negotiate": args.auto_negotiate,
        "interval_min": args.interval_min,
        "interval_max": args.interval_max
    }

    try:
        res = authenticated_request(
            endpoint_path="/api/v1/agent/tasks/1688-inquiry",
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
