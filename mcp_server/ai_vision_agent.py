import base64
import json
import logging
import asyncio
import urllib.request
import sys
import uuid
from typing import Dict, Any, Tuple, Optional

logger = logging.getLogger(__name__)

async def sample_agent_vision(base64_image: str) -> Optional[Dict[str, int]]:
    """
    通过 MCP Sampling 反向调用 Antigravity (Agent) 的多模态视觉能力，获取坐标。
    由于目前 mcp_server 是阻塞读取的，我们在 tool 执行期间发送 RPC 请求并强制内联读取 stdin。
    """
    req_id = str(uuid.uuid4())
    req = {
        "jsonrpc": "2.0",
        "id": req_id,
        "method": "sampling/createMessage",
        "params": {
            "messages": [
                {
                    "role": "user",
                    "content": {
                        "type": "text",
                        "text": "这是 1688 商品详情页的截图。找到代表“客服”、“联系客服”或“咨询”的按钮。假设屏幕左上角坐标为(0,0)，估算该按钮中心点的大致绝对像素坐标 (X, Y)。请只返回严格的 JSON 格式：{\"x\": 1024, \"y\": 768}"
                    }
                },
                {
                    "role": "user",
                    "content": {
                        "type": "image",
                        "data": base64_image,
                        "mimeType": "image/jpeg"
                    }
                }
            ],
            "maxTokens": 300
        }
    }
    
    # 打印日志（发到 stderr 不影响 stdio RPC）
    logger.info(f"[MCP Sampling] 正在向 Agent 抛出视觉求助请求 (ReqID: {req_id})...")
    
    # 向 Agent 发送请求
    sys.stdout.write(json.dumps(req, ensure_ascii=False) + "\n")
    sys.stdout.flush()
    
    # 阻塞当前协程/线程，强制从 stdin 读取 Agent 的响应
    # 注意：这在单线程 MCP 架构下是合法的，因为主循环正在 handle_tool_call 内被阻塞等待此函数返回。
    while True:
        line = sys.stdin.readline()
        if not line:
            logger.error("[MCP Sampling] Agent stdin 断开连接。")
            return None
        line = line.strip()
        if not line:
            continue
        try:
            resp = json.loads(line)
            # 如果收到的是别的通知（比如 ping），先忽略
            if resp.get("id") == req_id:
                # the result format for sampling/createMessage is:
                # { "content": { "type": "text", "text": "..." }, "role": "assistant" }
                result = resp.get("result", {})
                content = result.get("content", {})
                
                # Check for MCP standard content object OR Anthropic/OpenAI array format
                if isinstance(content, dict) and content.get("type") == "text":
                    text_val = content.get("text", "")
                elif isinstance(result, dict) and "text" in result:
                     text_val = result.get("text")
                else:
                    text_val = str(content)
                    
                # fallback text search for json
                import re
                match = re.search(r'\{.*\}', text_val, re.DOTALL)
                if match:
                    try:
                        return json.loads(match.group(0))
                    except:
                        pass
                
                try:
                    return json.loads(text_val)
                except:
                    logger.error(f"[MCP Sampling] Agent 返回的不是合法 JSON: {text_val}")
                    return None
        except Exception as e:
            logger.error(f"[MCP Sampling] 解析 Agent 响应失败: {e}, {line[:100]}")


async def local_ai_fallback_handler(cdp_ws, cdp_base) -> Optional[Dict[str, Any]]:
    """
    执行视觉坐标点击，无任何 OpenAI 依赖，全靠白嫖 Agent
    """
    logger.info("[MCP AI Agent] 启动零配置视觉自愈 (MCP Sampling)...")
    try:
        await cdp_ws.send(json.dumps({"id": 9001, "method": "Page.captureScreenshot"}))
        shot_resp = None
        for _ in range(10):
            r = json.loads(await asyncio.wait_for(cdp_ws.recv(), timeout=5.0))
            if r.get("id") == 9001:
                shot_resp = r
                break
        
        if shot_resp and shot_resp.get("result", {}).get("data"):
            b64 = shot_resp["result"]["data"]
            
            # 反向调用 Agent
            coords = await sample_agent_vision(b64)
            if not coords:
                logger.error("[MCP AI Agent] 未能从 Agent 处获取到坐标。")
                return None
                
            tx, ty = int(coords.get("x", 0)), int(coords.get("y", 0))
            if tx > 0 and ty > 0:
                logger.info(f"[MCP AI Agent] Agent 视觉识别坐标: ({tx}, {ty})，注入点击...")
                await cdp_ws.send(json.dumps({"id": 8801, "method": "Input.dispatchMouseEvent", "params": {"type": "mousePressed", "button": "left", "x": tx, "y": ty, "clickCount": 1}}))
                await asyncio.sleep(0.1)
                await cdp_ws.send(json.dumps({"id": 8802, "method": "Input.dispatchMouseEvent", "params": {"type": "mouseReleased", "button": "left", "x": tx, "y": ty, "clickCount": 1}}))
                
                # 检查旺旺弹出
                for poll in range(8):
                    await asyncio.sleep(1.0)
                    req = urllib.request.Request(f"{cdp_base}/json/list")
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        tabs = json.loads(resp.read().decode("utf-8"))
                    im_tab = next((p for p in tabs if ("def_cbu_web_im" in p.get("url", "") or "wangwang" in p.get("url", "").lower() or "im.1688.com" in p.get("url", "")) and p.get("url", "") != "about:blank"), None)
                    if im_tab:
                        logger.info(f"[MCP AI Agent] ✅ Agent 视觉自愈成功！")
                        return im_tab
    except Exception as e:
        logger.error(f"[MCP AI Agent] 视觉自愈执行失败: {e}")

    return None
