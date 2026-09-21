#!/usr/bin/env python3
"""
异步任务状态与进度轮询 CLI 脚本
调用店小二服务端 /api/v1/agent/tasks/{task_no}/status 接口
感知 HITL (人机协作) 挂起状态
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from dianxiaoer_auth import ActionRequired, authenticated_request, DEFAULT_BASE_URL


def main():
    parser = argparse.ArgumentParser(description="查询店小二异步任务进度及结果")
    parser.add_argument("--task-no", required=True, help="任务唯一编号 (TASK-XXXX)")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="店小二服务地址")
    args = parser.parse_args()

    try:
        res = authenticated_request(
            endpoint_path=f"/api/v1/agent/tasks/{args.task_no}/status",
            base_url=args.base_url,
            method="GET"
        )
        # 检查是否触发人机协作阻断
        data = res.get("data", {})
        if data.get("actionRequired"):
            raise ActionRequired(
                message=data.get("msg") or "任务需人工交互后继续",
                jump_url=data.get("jumpUrl", "http://127.0.0.1:9222"),
                action=data.get("hitlAction", "CAPTCHA_DETECTED"),
                pending=True
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
