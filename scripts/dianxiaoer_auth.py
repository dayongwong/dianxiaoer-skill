#!/usr/bin/env python3
"""
店小二 (Dianxiaoer) 客户端统一设备码认证、并发文件锁与请求封装模块
纯 Python 标准库实现，零外部重量级依赖，支持 Linux/macOS/Windows 跨平台运行。
"""

import argparse
import errno
import json
import os
import sys
import tempfile
import time
import uuid
from contextlib import ExitStack, contextmanager
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

# 默认后端微服务地址（已绑定云端域名）
DEFAULT_BASE_URL = os.getenv("DIANXIAOER_BASE_URL", "http://dianxiaoer-skill.ivector.cn")


AUTH_CHALLENGE_ENDPOINT = "/api/v1/agent/auth/challenge"
TOKEN_ENDPOINT = "/api/v1/agent/auth/token"

STATE_VERSION = 1
AUTH_FILE_NAME = "agent-auth.json"
AUTH_STATE_DIR = ".dianxiaoer"
CLIENT_NAME = "Antigravity Agent"


class ActionRequired(Exception):
    """
    当操作被阻断（未登录、等待用户授权、配额不足、或触发滑块风控）时抛出。
    提供结构化的公开 JSON 载荷供 Agent 或 CLI 消费。
    """
    def __init__(
        self,
        message: str,
        jump_url: str = "",
        action: str = "ACTION_REQUIRED",
        expires_in: int = 0,
        pending: bool = False
    ):
        super().__init__(message)
        self.message = message
        self.jump_url = jump_url
        self.action = action
        self.expires_in = expires_in
        self.pending = pending

    def public_payload(self):
        return {
            "actionRequired": True,
            "actionPending": self.pending,
            "action": self.action,
            "msg": self.message,
            "jumpUrl": self.jump_url,
            "expiresIn": self.expires_in,
        }


class ResolvedStore:
    def __init__(self, path: Path, kind: str):
        self.path = _absolute_path(path)
        self.kind = kind


def _absolute_path(path) -> Path:
    path = Path(path).expanduser()
    try:
        return path.resolve(strict=False)
    except (OSError, RuntimeError):
        return Path(os.path.abspath(os.fspath(path)))


def _user_config_state_path() -> Path:
    # 用户系统级配置目录: ~/.config/Dianxiaoer/agent-auth.json 或 AppData
    if os.name == "nt":
        base = os.getenv("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    else:
        base = os.getenv("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return _absolute_path(Path(base) / "Dianxiaoer" / AUTH_FILE_NAME)


def _skill_state_path() -> Path:
    return _absolute_path(Path(__file__).parent.parent / AUTH_STATE_DIR / AUTH_FILE_NAME)


def _workspace_state_path() -> Path:
    return _absolute_path(Path(os.getcwd()) / AUTH_STATE_DIR / AUTH_FILE_NAME)


def _resolve_stores():
    candidates = (
        ResolvedStore(_user_config_state_path(), "user-config-directory"),
        ResolvedStore(_skill_state_path(), "skill-directory"),
        ResolvedStore(_workspace_state_path(), "working-directory"),
    )
    stores = []
    seen = set()
    for s in candidates:
        key = os.path.normcase(os.fspath(s.path))
        if key not in seen:
            seen.add(key)
            stores.append(s)
    return tuple(stores)


def _empty_state():
    return {
        "version": STATE_VERSION,
        "storeId": uuid.uuid4().hex,
        "servers": {},
    }


def _normalize_state(payload):
    if not isinstance(payload, dict):
        raise ValueError("登录状态文件格式错误")
    servers = payload.get("servers", {})
    if not isinstance(servers, dict):
        servers = {}
    return {
        "version": STATE_VERSION,
        "storeId": payload.get("storeId") or uuid.uuid4().hex,
        "servers": servers,
    }


def _restrict_permissions(path, mode):
    if os.name == "posix":
        try:
            os.chmod(os.fspath(path), mode)
        except (OSError, AttributeError):
            pass


def _probe_file_lock(handle):
    if os.name == "nt":
        import msvcrt
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        try:
            return
        finally:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            return
        finally:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


@contextmanager
def _state_lock(store):
    lock_path = store.path.with_name(f".{store.path.name}.lock")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    _restrict_permissions(lock_path.parent, 0o700)
    with open(lock_path, "a+b") as handle:
        _restrict_permissions(lock_path, 0o600)
        if os.name == "nt":
            import msvcrt
            handle.seek(0, os.SEEK_END)
            if handle.tell() == 0:
                handle.write(b"\0")
                handle.flush()
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_LOCK, 1)
            try:
                yield
            finally:
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX)
            try:
                yield
            finally:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _write_json_file(path, payload, prefix):
    path.parent.mkdir(parents=True, exist_ok=True)
    _restrict_permissions(path.parent, 0o700)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            prefix=prefix,
            delete=False
        ) as handle:
            temp_path = Path(handle.name)
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.flush()
        os.replace(temp_path, path)
        _restrict_permissions(path, 0o600)
    finally:
        if temp_path and temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass


def _load_state():
    stores = _resolve_stores()
    for s in stores:
        if s.path.is_file():
            try:
                payload = json.loads(s.path.read_text(encoding="utf-8"))
                return _normalize_state(payload)
            except Exception:
                continue
    return _empty_state()


def _update_state(callback):
    stores = _resolve_stores()
    target_store = stores[0]
    with _state_lock(target_store):
        payload = _load_state()
        changed, result = callback(payload)
        if changed:
            _write_json_file(target_store.path, _normalize_state(payload), ".agent-auth-")
        return result


def _server_key(base_url: str) -> str:
    return base_url.rstrip("/")


def _server_state(payload, base_url: str):
    return payload["servers"].setdefault(_server_key(base_url), {})


def complete_pending_login(base_url: str, timeout: float = 15.0):
    """
    检查是否存在待用户确认的 Device Flow，若存在则向服务端轮询 Token
    """
    payload = _load_state()
    server_key = _server_key(base_url)
    server = payload["servers"].get(server_key, {})
    pending = server.get("pending")
    now = int(time.time())

    if not isinstance(pending, dict):
        return False

    if pending.get("expiresAt", 0) <= now:
        def clear_expired_pending(p):
            s = p["servers"].get(server_key, {})
            s.pop("pending", None)
            return True, None
        _update_state(clear_expired_pending)
        return False

    endpoint = f"{server_key}{TOKEN_ENDPOINT}"
    body = json.dumps({"device_code": pending["deviceCode"]}).encode("utf-8")
    req = Request(endpoint, data=body, headers={"Content-Type": "application/json", "Accept": "application/json"}, method="POST")

    try:
        with urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            token_data = data.get("data", {})
            access_token = token_data.get("accessToken")
            expires_in = token_data.get("expiresIn", 2592000)

            def save_token(latest):
                srv = latest["servers"].setdefault(server_key, {})
                srv["accessToken"] = access_token
                srv["accessTokenExpiresAt"] = now + expires_in
                srv.pop("pending", None)
                return True, True
            _update_state(save_token)
            return True
    except HTTPError as error:
        try:
            err_data = json.loads(error.read().decode("utf-8"))
        except Exception:
            err_data = {}
        detail = err_data.get("detail", {}) if isinstance(err_data.get("detail"), dict) else {}
        err_code = detail.get("error")

        if error.code == 202 or err_code == "AUTHORIZATION_PENDING":
            raise ActionRequired(
                detail.get("msg") or "等待用户在网页端完成授权。授权完成后请再次调用原指令。",
                pending.get("jumpUrl", ""),
                action="AUTHORIZATION_PENDING",
                expires_in=max(0, pending.get("expiresAt", now) - now),
                pending=True
            )
        elif error.code in (400, 410):
            def clear_invalid(latest):
                srv = latest["servers"].setdefault(server_key, {})
                srv.pop("pending", None)
                return True, None
            _update_state(clear_invalid)
        raise


def _get_access_token(base_url: str) -> str:
    payload = _load_state()
    server = payload["servers"].get(_server_key(base_url), {})
    token = server.get("accessToken")
    now = int(time.time())
    if token and server.get("accessTokenExpiresAt", 0) > now:
        return token
    return ""


def _trigger_new_device_flow(base_url: str, timeout: float = 15.0):
    """
    向服务端申请全新的设备授权挑战，并记录到本地状态文件
    """
    endpoint = f"{_server_key(base_url)}{AUTH_CHALLENGE_ENDPOINT}"
    body = json.dumps({"client_name": CLIENT_NAME, "expires_seconds": 600}).encode("utf-8")
    req = Request(endpoint, data=body, headers={"Content-Type": "application/json", "Accept": "application/json"}, method="POST")

    with urlopen(req, timeout=timeout) as resp:
        res = json.loads(resp.read().decode("utf-8"))
        data = res.get("data", {})
        device_code = data.get("deviceCode")
        user_code = data.get("userCode")
        jump_url = data.get("jumpUrl")
        expires_in = data.get("expiresIn", 600)

        def save_challenge(latest):
            srv = latest["servers"].setdefault(_server_key(base_url), {})
            srv.pop("accessToken", None)
            srv["pending"] = {
                "deviceCode": device_code,
                "userCode": user_code,
                "jumpUrl": jump_url,
                "expiresAt": int(time.time()) + expires_in
            }
            return True, None
        _update_state(save_challenge)

        raise ActionRequired(
            f"首次使用请点击链接完成授权绑定（授权短码：{user_code}）。完成授权后请回复“继续”。",
            jump_url=jump_url,
            action="AUTH_REQUIRED",
            expires_in=expires_in,
            pending=True
        )


def authenticated_request(
    endpoint_path: str,
    base_url: str = DEFAULT_BASE_URL,
    method: str = "GET",
    body: dict = None,
    headers: dict = None,
    timeout: float = 30.0
) -> dict:
    """
    通用认证请求封装：
    1. 优先自动处理待授权轮询
    2. 校验并自动附带 Bearer Token
    3. 遇到 401 自动触发 Device Flow 并引导用户
    4. 遇到 402 提取充值 jumpUrl 并阻断
    """
    # 步骤 1: 尝试完成正在 pending 的授权
    try:
        complete_pending_login(base_url, timeout)
    except ActionRequired:
        raise

    token = _get_access_token(base_url)
    if not token:
        # 无 Token，自动触发全新 Device Flow
        _trigger_new_device_flow(base_url, timeout)

    req_headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "X-Timestamp": str(int(time.time())),
        "X-Nonce": uuid.uuid4().hex[:16]
    }
    if body is not None:
        req_headers["Content-Type"] = "application/json"
    if headers:
        req_headers.update(headers)

    url = f"{_server_key(base_url)}{endpoint_path}"
    data_bytes = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
    req = Request(url, data=data_bytes, headers=req_headers, method=method)

    try:
        with urlopen(req, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
            return payload
    except HTTPError as error:
        try:
            err_payload = json.loads(error.read().decode("utf-8"))
        except Exception:
            err_payload = {}

        if isinstance(err_payload.get("detail"), dict):
            detail = err_payload.get("detail")
        elif isinstance(err_payload.get("detail"), str):
            detail = {"msg": err_payload.get("detail")}
        else:
            detail = {}
        err_code = detail.get("error")

        if error.code == 401 or err_code in ("AUTH_REQUIRED", "AUTH_EXPIRED"):
            # Token 失效或 3 天已过期，清除失效凭据并重新拉起设备授权
            def clear_token(p):
                s = p["servers"].get(_server_key(base_url), {})
                s.pop("accessToken", None)
                s.pop("accessTokenExpiresAt", None)
                return True, None
            _update_state(clear_token)
            _trigger_new_device_flow(base_url, timeout)

        if error.code == 402 or err_code == "INSUFFICIENT_QUOTA":
            # 配额不足
            clean_base = _server_key(base_url)
            raw_jump = detail.get("jumpUrl", "http://127.0.0.1:8089/#/tenant/quotas")
            jump_url = raw_jump if raw_jump.startswith("http") else f"{clean_base}{raw_jump}"
            raise ActionRequired(
                message=detail.get("msg") or "企业共享配额已耗尽，请充值后重试。",
                jump_url=jump_url,
                action="INSUFFICIENT_QUOTA"
            )

        if error.code == 403 and err_code in ("MEMBERSHIP_REQUIRED", "MEMBERSHIP_EXPIRED", "TENANT_INACTIVE", "TENANT_NOT_FOUND"):
            # 会员未开通、过期或租户停用
            clean_base = _server_key(base_url)
            raw_jump = detail.get("jumpUrl", "http://127.0.0.1:8089/#/membership/pricing")
            jump_url = raw_jump if raw_jump.startswith("http") else f"{clean_base}{raw_jump}"
            raise ActionRequired(
                message=detail.get("msg") or "当前功能需具备有效店小二企业会员资格，请开通或续费后继续使用。",
                jump_url=jump_url,
                action=err_code or "MEMBERSHIP_REQUIRED"
            )

        msg = detail.get("msg") or err_payload.get("msg") or f"HTTP {error.code}: {error.reason}"
        raise ValueError(msg) from error


def main():
    parser = argparse.ArgumentParser(description="店小二客户端认证与本地状态管理器")
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("storage-status", help="检查登录状态本地存储")
    subparsers.add_parser("clear", help="清除所有已缓存的登录凭据")
    args = parser.parse_args()

    stores = _resolve_stores()
    if args.command == "clear":
        for s in stores:
            if s.path.exists():
                try:
                    s.path.unlink()
                except OSError:
                    pass
        print(json.dumps({"code": 0, "msg": "本地凭据已清除"}, ensure_ascii=False, indent=2))
        return 0

    if args.command == "storage-status":
        status_list = []
        for s in stores:
            status_list.append({
                "type": s.kind,
                "path": str(s.path),
                "exists": s.path.is_file()
            })
        print(json.dumps({"code": 0, "stores": status_list}, ensure_ascii=False, indent=2))
        return 0


if __name__ == "__main__":
    sys.exit(main())
