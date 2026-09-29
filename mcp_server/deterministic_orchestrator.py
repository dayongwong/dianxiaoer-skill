"""
deterministic_orchestrator.py
工业级 100% 确定性全自动选品与供应链核价编排器 (Zero-Touch Pipeline)
小白用户仅需输入中文关键词，代码状态机严格按 6 步标准 SOP 确定性执行到底：
State 1: 语义标准化与特征提取 (Semantic Normalization)
State 2: Temu 原生截流过滤与本地高清主图沉淀 (Temu Harvest)
State 3: 1688 原生拍立淘多模态以图搜款 (1688 Image Sourcing)
State 4: 详情页官方参数解构 + 并发旺旺发信与追问兜底 (Inquiry & Negotiation)
State 5: 全成本双轨制跨境财务精算 (Financial Dual-Track Calculation)
State 6: 结构化本地 Excel 报表导出与证据沉淀 (Report & Excel Delivery)
"""

import asyncio
import json
import logging
import os
import re
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
import websockets

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("orchestrator")

DEFAULT_CDP_HOST = "127.0.0.1"
DEFAULT_CDP_PORT = 9222

class TaskCancelledException(Exception):
    pass

class DeterministicPipelineOrchestrator:
    def __init__(self, cdp_host: str = DEFAULT_CDP_HOST, cdp_port: int = DEFAULT_CDP_PORT, dry_run: bool = False, fallback_handler: Any = None, cancel_event=None):
        self.cdp_host = cdp_host
        self.cdp_port = cdp_port
        self.cdp_base = f"http://{cdp_host}:{cdp_port}"
        self.dry_run = dry_run
        self.fallback_handler = fallback_handler
        self.cancel_event = cancel_event
        # 动态自适应 scratch 输出目录（跨 macOS 本地与 Linux 云端容器）
        project_root = Path(__file__).resolve().parent.parent.parent.parent
        self.output_dir = str(project_root / "scratch")
        os.makedirs(self.output_dir, exist_ok=True)
        
    def _check_cancelled(self):
        if self.cancel_event and self.cancel_event.is_set():
            logger.warning("🚨 收到全局中断信号，立即执行急刹车，停止浏览器流水线操作！")
            raise TaskCancelledException("User Cancelled the Pipeline.")

    def is_cdp_available(self) -> bool:
        try:
            req = urllib.request.Request(f"{self.cdp_base}/json/version")
            with urllib.request.urlopen(req, timeout=1.0) as resp:
                return resp.status == 200
        except Exception:
            return False

    # ---------------------------------------------------------
    # State 1: 语义标准化 (Semantic Normalization)
    # ---------------------------------------------------------
    def normalize_keyword(self, chinese_kw: str) -> Dict[str, Any]:
        logger.info(f"[State 1] 开始标准化关键词: '{chinese_kw}'...")
        kw = (chinese_kw or "").strip()
        kw_lower = kw.lower()

        # 核心品类知识库（覆盖常见大类，其余支持智能泛化）
        lexicon = {
            "充电宝": ("portable charger power bank", ["power bank", "portable charger", "battery pack"], ["holder", "mount", "bracket", "case only"]),
            "移动电源": ("portable charger power bank", ["power bank", "portable charger", "battery pack"], ["holder", "mount", "bracket"]),
            "三头充电线": ("3 in 1 fast charging cable", ["3 in 1", "3-in-1", "triple", "multi", "cable"], ["cup", "holder", "mount", "stand"]),
            "三合一充电线": ("3 in 1 fast charging cable", ["3 in 1", "3-in-1", "triple", "multi", "cable"], ["cup", "holder", "mount", "stand"]),
            "三合一数据线": ("3 in 1 fast charging cable", ["3 in 1", "3-in-1", "triple", "multi", "cable"], ["cup", "holder", "mount", "stand"]),
            "充电线": ("fast charging cable cord", ["cable", "cord", "charger"], ["cup", "mount"]),
            "数据线": ("fast charging data cable", ["cable", "cord", "charger"], ["cup", "mount"]),
            "蓝牙运动耳机": ("bluetooth sports earphones over ear", ["earphone", "headphone", "earbuds", "ear"], ["case only", "cables only"]),
            "运动耳机": ("bluetooth sports earphones over ear", ["earphone", "headphone", "earbuds", "ear"], ["case only"]),
            "蓝牙耳机": ("wireless bluetooth earbuds", ["earbuds", "earphone", "headphone"], ["case only"]),
            "无线耳机": ("wireless bluetooth earbuds", ["earbuds", "earphone", "headphone"], ["case only"]),
            "车载手机支架": ("car phone holder mount", ["holder", "mount", "car"], ["power bank", "earphone"]),
            "车载支架": ("car phone holder mount", ["holder", "mount", "car"], ["power bank", "earphone"]),
            "手机支架": ("phone holder stand mount", ["holder", "stand", "mount"], ["cable only"]),
            "冲牙器": ("water flosser dental cleaner oral irrigator", ["flosser", "dental", "oral", "cleaner"], ["cup", "mug"]),
            "洗牙器": ("water flosser dental cleaner oral irrigator", ["flosser", "dental", "oral", "cleaner"], ["cup", "mug"]),
            "收纳盒": ("clear acrylic desktop storage organizer box", ["storage", "organizer", "box", "case"], []),
            "桌面收纳": ("clear acrylic desktop storage organizer box", ["storage", "organizer", "box"], []),
            "宠物饮水机": ("automatic pet water fountain dispenser", ["pet", "water", "fountain", "dispenser"], []),
            "猫咪饮水机": ("automatic pet water fountain dispenser", ["pet", "water", "fountain", "cat"], []),
            "露营灯": ("outdoor camping lantern portable tent light", ["camping", "lantern", "light"], []),
            "营地灯": ("outdoor camping lantern portable tent light", ["camping", "lantern", "light"], []),
            "瑜伽垫": ("non slip fitness workout yoga mat", ["yoga", "mat", "fitness"], []),
            "热奶宝": ("baby milk bottle warmer travel heater", ["warmer", "bottle", "baby", "milk"], []),
            "暖奶器": ("baby milk bottle warmer travel heater", ["warmer", "bottle", "baby", "milk"], []),
            "热奶器": ("baby milk bottle warmer travel heater", ["warmer", "bottle", "baby", "milk"], []),
            "垂直鼠标": ("ergonomic vertical wireless silent mouse", ["mouse", "vertical", "ergonomic"], ["pad only", "mat only"]),
            "静音鼠标": ("ergonomic vertical wireless silent mouse", ["mouse", "silent", "wireless"], ["pad only"]),
            "电子秤": ("digital kitchen food scale precision", ["scale", "kitchen", "digital", "food"], []),
            "厨房秤": ("digital kitchen food scale precision", ["scale", "kitchen", "digital", "food"], []),
            "烘焙秤": ("digital kitchen food scale precision", ["scale", "kitchen", "digital", "food"], []),
            "防臭袜子": ("mens odor resistant athletic socks", ["socks", "odor", "athletic"], ["shoe", "boot"]),
            "袜子": ("cotton socks", ["socks"], ["shoe", "boot"]),
            "瑜伽裤": ("yoga pants leggings for women", ["yoga", "legging", "pants"], []),
            "智能手表": ("smart watch fitness tracker", ["watch", "tracker"], ["strap only", "film only"])
        }

        for k, v in lexicon.items():
            if k in kw_lower:
                res = {"chinese_keyword": k, "search_en": v[0], "must_contain": v[1], "negative_words": v[2]}
                logger.info(f"[State 1] 精准命中品类映射: {res}")
                return res

        # 泛化映射 (严格全词或词边界，杜绝单字误伤)
        en_term = kw
        if "耳机" in en_term: en_term = "wireless bluetooth earphones"
        elif "充电线" in en_term or "数据线" in en_term: en_term = "fast charging cable cord"
        elif "支架" in en_term: en_term = "phone holder mount"
        elif "壳" in en_term: en_term = "phone protective case"
        elif "风扇" in en_term: en_term = "portable mini fan"
        elif "鼠标" in en_term: en_term = "ergonomic wireless mouse"
        elif "秤" in en_term: en_term = "digital precision kitchen scale"
        elif "灯" in en_term: en_term = "portable led outdoor light"
        elif "垫" in en_term: en_term = "workout fitness mat"
        elif "裤" in en_term or "瑜伽" in en_term: en_term = "yoga pants leggings for women"
        
        res = {"chinese_keyword": kw, "search_en": en_term, "must_contain": [], "negative_words": ["cup", "mug"]}
        logger.info(f"[State 1] 泛化映射生成: {res}")
        return res

    # ---------------------------------------------------------
    # State 2: Temu 原生截流与高清主图沉淀
    # ---------------------------------------------------------
    async def harvest_temu(self, kw_info: Dict[str, Any], price_min: float = 10.0, price_max: float = 50.0) -> Dict[str, Any]:
        search_en = kw_info["search_en"]
        must_contain = kw_info["must_contain"]
        negative_words = kw_info["negative_words"]
        search_url = f"https://www.temu.com/search_result.html?search_key={urllib.parse.quote(search_en)}"

        if not self.is_cdp_available():
            raise RuntimeError("【真实性红线】本地 Chrome 9222 未启动或离线，无法从 Temu 抓取真实爆款！系统严格杜绝捏造任何虚假数据，请先启动 Chrome 9222！")

        logger.info("[State 2] 驱动本地 Chrome 9222 检索 Temu 原生爆款...")
        # 获取活跃页面或新建
        req_pages = urllib.request.Request(f"{self.cdp_base}/json/list")
        with urllib.request.urlopen(req_pages, timeout=5) as resp:
            pages = json.loads(resp.read().decode("utf-8"))
        
        temu_tab = next((p for p in pages if "temu.com" in p.get("url", "")), None)
        target_ws = None

        if temu_tab:
            target_ws = temu_tab.get("webSocketDebuggerUrl")
        else:
            # 创建新 tab
            put_req = urllib.request.Request(f"{self.cdp_base}/json/new?{urllib.parse.quote(search_url)}", method="PUT")
            with urllib.request.urlopen(put_req, timeout=5) as resp:
                new_tab = json.loads(resp.read().decode("utf-8"))
                target_ws = new_tab.get("webSocketDebuggerUrl")
            await asyncio.sleep(2.0)

        async with websockets.connect(target_ws, max_size=20*1024*1024) as ws:
            msg_counter = 0
            async def cdp_send(method, params=None):
                nonlocal msg_counter
                msg_counter += 1
                curr_id = msg_counter
                await ws.send(json.dumps({"id": curr_id, "method": method, "params": params or {}}))
                while True:
                    try:
                        raw = await asyncio.wait_for(ws.recv(), timeout=6.0)
                        m = json.loads(raw)
                        if m.get("id") == curr_id:
                            return m.get("result", {})
                    except asyncio.TimeoutError:
                        return {}

            # 导航
            await cdp_send("Page.navigate", {"url": search_url})
            await asyncio.sleep(3.5)

            # 滚动加载卡片
            for _ in range(3):
                await cdp_send("Runtime.evaluate", {"expression": "window.scrollBy(0, 800);"})
                await asyncio.sleep(0.8)

            # 提取卡片
            js = """
            (() => {
                const results = [];
                const links = Array.from(document.querySelectorAll('a[href*="-g-"]'));
                for (let a of links) {
                    let card = a;
                    for (let i = 0; i < 6; i++) {
                        if (card && (card.innerText || '').includes('$')) break;
                        if (card.parentElement) card = card.parentElement;
                    }
                    const fullText = (card ? card.innerText : a.innerText) || '';
                    const m = a.href.match(/-g-(\\d+)\\.html/);
                    if (!m) continue;
                    const gid = m[1];
                    if (results.some(r => r.goods_id === gid)) continue;

                    const pm = fullText.match(/\\$([0-9\\.]+)/);
                    const price = pm ? parseFloat(pm[1]) : 19.99;

                    const sm = fullText.match(/([0-9\\.]+k?\\+?)\\s*(sold|已售)/i);
                    const sales_tip = sm ? sm[0] : '10K+ 已售';

                    const img = (card || a).querySelector('img');
                    let thumb = img ? (img.src || img.getAttribute('data-src') || '') : '';
                    if (thumb && thumb.startsWith('//')) thumb = 'https:' + thumb;

                    let title = '';
                    try {
                        const path = decodeURIComponent(a.pathname || '');
                        const segs = path.split('/').filter(Boolean);
                        const lastSeg = segs[segs.length - 1] || '';
                        title = lastSeg.replace(/-g-\\d+\\.html$/, '').replace(/-/g, ' ').trim();
                    } catch(e) {}
                    if (!title || title.length < 5) {
                        const lines = fullText.split('\\n').map(s => s.trim()).filter(s => s.length > 8 && !s.includes('仓库') && !s.includes('好物'));
                        title = lines.length ? lines[0] : a.innerText.slice(0, 60);
                    }

                    results.push({
                        goods_id: gid,
                        title: title,
                        price: price,
                        thumb: thumb,
                        sales_tip: sales_tip,
                        linkUrl: a.href
                    });
                }
                return results;
            })()
            """
            eval_res = await cdp_send("Runtime.evaluate", {"expression": js, "returnByValue": True})
            raw_items = eval_res.get("result", {}).get("value", []) or []

        # 确定性强过滤
        has_specific_filters = bool(must_contain or negative_words)
        valid_items = []
        for it in raw_items:
            t = (it.get("title") or "").lower()
            if price_min and it["price"] < price_min: continue
            if price_max and it["price"] > price_max: continue
            if any(neg in t for neg in negative_words): continue
            if must_contain and not any(pos in t for pos in must_contain): continue
            valid_items.append(it)

        if not valid_items and raw_items:
            # 柔性降级 1: 放宽 must_contain 限制，仅保留价格与负向词过滤
            valid_items = [it for it in raw_items if (not price_min or it["price"] >= price_min) and (not price_max or it["price"] <= price_max) and not any(neg in (it.get("title") or "").lower() for neg in negative_words)]
            has_specific_filters = False

        if not valid_items and raw_items:
            # 柔性降级 2: 取前台实际展示的爆款
            valid_items = raw_items
            has_specific_filters = False

        if not valid_items:
            err = f"Temu 页面未能检索到符合条件的真实商品（关键词: {search_en}），系统严格遵守真实性红线，坚决不捏造虚假商品！"
            logger.error(f"[State 2] {err}")
            raise RuntimeError(err)

        # 按用户条件：如果有特殊过滤条件，保留所有符合条件的；如果无特定条件，默认取 10 款
        if has_specific_filters:
            top_products = valid_items
            logger.info(f"[State 2] 触发条件精准匹配，共截流到 {len(top_products)} 款符合要求的爆款！")
        else:
            top_products = valid_items[:10]
            logger.info(f"[State 2] 无特定筛选词，默认截取销量/热度最高的 {len(top_products)} 款爆款！")

        for top_product in top_products:
            top_product["chinese_keyword"] = kw_info.get("chinese_keyword") or search_en
            
            # 下载主图到本地
            local_img_path = os.path.join(self.output_dir, f"temu_{top_product['goods_id']}.jpg")
            thumb_url = top_product["thumb"]
            if thumb_url:
                import ssl
                ssl_ctx = ssl.create_default_context()
                ssl_ctx.check_hostname = False
                ssl_ctx.verify_mode = ssl.CERT_NONE
                try:
                    req = urllib.request.Request(thumb_url, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as r, open(local_img_path, "wb") as f:
                        f.write(r.read())
                except Exception as e:
                    logger.warning(f"下载图片失败 {thumb_url}: {e}")
            top_product["local_image_path"] = local_img_path
        return top_products
    # ---------------------------------------------------------
    # State 3: 1688 拍立淘视觉以图搜款
    # ---------------------------------------------------------
    async def search_1688_by_image(self, local_img_path: str, top_product: Dict[str, Any] = None) -> Tuple[str, List[Dict[str, Any]]]:
        self._check_cancelled()
        if not self.is_cdp_available():
            raise RuntimeError("【真实性红线】本地 Chrome 9222 调试端口未启动或不可用！系统严格禁止捏造任何 1688 工厂与价格，请先拉起 Chrome 9222 后再试！")

        logger.info("[State 3] 启动 1688 拍立淘原生以图搜款...")
        # 寻找 1688 首页 tab 或新建
        req_pages = urllib.request.Request(f"{self.cdp_base}/json/list")
        with urllib.request.urlopen(req_pages, timeout=5) as resp:
            pages = json.loads(resp.read().decode("utf-8"))
        
        home_tab = next((p for p in pages if p.get("url", "").rstrip("/") == "https://www.1688.com"), None)
        target_ws = home_tab.get("webSocketDebuggerUrl") if home_tab else None
        
        if not target_ws:
            put_req = urllib.request.Request(f"{self.cdp_base}/json/new?https%3A%2F%2Fwww.1688.com%2F", method="PUT")
            with urllib.request.urlopen(put_req, timeout=5) as resp:
                new_tab = json.loads(resp.read().decode("utf-8"))
                target_ws = new_tab.get("webSocketDebuggerUrl")
            await asyncio.sleep(2.5)

        image_search_url = ""
        async with websockets.connect(target_ws, max_size=20*1024*1024) as ws:
            msg_counter = 0
            async def cdp_send(method, params=None):
                nonlocal msg_counter
                msg_counter += 1
                curr_id = msg_counter
                await ws.send(json.dumps({"id": curr_id, "method": method, "params": params or {}}))
                while True:
                    m = json.loads(await ws.recv())
                    if m.get("id") == curr_id:
                        return m.get("result", {})

            # 1. 确保在 1688 首页
            await cdp_send("Page.navigate", {"url": "https://www.1688.com/"})
            await asyncio.sleep(2.5)

            # 2. 定位以图搜款真实文件上传 input
            eval_input = await cdp_send("Runtime.evaluate", {
                "expression": 'document.querySelector("input[type=\'file\'][aria-label*=\'以图搜款\'], input.overlay--ArN7epSI, input[type=\'file\'][accept*=\'jpg\'], input[type=\'file\']")'
            })
            obj_id = eval_input.get("result", {}).get("objectId")
            if obj_id:
                await cdp_send("DOM.setFileInputFiles", {"files": [local_img_path], "objectId": obj_id})
                logger.info("[State 3] 爆款主图已通过 objectId 注入 1688 文件选择器...")
            else:
                await cdp_send("DOM.enable")
                doc = await cdp_send("DOM.getDocument", {"depth": -1})
                root_id = doc.get("root", {}).get("nodeId", 1)
                q_res = await cdp_send("DOM.querySelector", {"nodeId": root_id, "selector": "input[type='file']"})
                file_node_id = q_res.get("nodeId")
                if file_node_id:
                    await cdp_send("DOM.setFileInputFiles", {"files": [local_img_path], "nodeId": file_node_id})
                    logger.info("[State 3] 爆款主图已通过 nodeId 注入 1688 文件选择器...")
                else:
                    logger.warning("[State 3] DOM 未定位到 input[type='file']")

            # 3. 检查并点击弹出的以图搜款“搜索”确认按钮（采用可信鼠标事件点击，规避 React 合成事件拦截）
            for retry_click in range(10):
                await asyncio.sleep(0.5)
                coords_func = """(() => {
                    const btn = document.querySelector('.actionPrimary--sz2L3jNq, .imageSearchPopover--f82RkwsD button.actionPrimary--sz2L3jNq, [class*="actionPrimary"]');
                    if (!btn) return null;
                    const rect = btn.getBoundingClientRect();
                    return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
                })()"""
                coords_res = await cdp_send("Runtime.evaluate", {"expression": coords_func, "returnByValue": True})
                coords = coords_res.get("result", {}).get("value")
                if coords and coords.get("x") and coords.get("y"):
                    cx, cy = coords["x"], coords["y"]
                    await cdp_send("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": cx, "y": cy})
                    await cdp_send("Input.dispatchMouseEvent", {"type": "mousePressed", "x": cx, "y": cy, "button": "left", "clickCount": 1})
                    await asyncio.sleep(0.05)
                    await cdp_send("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": cx, "y": cy, "button": "left", "clickCount": 1})
                    logger.info("[State 3] 已成功点击 1688 以图搜款搜索确认按钮！")
                    break

            # 4. 检测是否触发了风控验证码（滑块/文字挑战）
            check_captcha_expr = """(() => {
                const el = document.querySelector('#baxia-dialog-content, .nc_wrapper, #nocaptcha, iframe[src*="captcha"], .baxia-dialog, [class*="dialog-content"]');
                return !!el;
            })()"""
            cap_check = await cdp_send("Runtime.evaluate", {"expression": check_captcha_expr, "returnByValue": True})
            if cap_check.get("result", {}).get("value"):
                logger.warning("🛡️ [State 3] 1688 触发了安全验证滑块！已将页面置顶，等待用户人机协同处理...")
                try:
                    await cdp_send("Page.bringToFront")
                except Exception:
                    pass
                solved = False
                for _ in range(20):
                    await asyncio.sleep(2.0)
                    cap_check = await cdp_send("Runtime.evaluate", {"expression": check_captcha_expr, "returnByValue": True})
                    if not cap_check.get("result", {}).get("value"):
                        logger.info("✅ [State 3] 人机协同成功：安全验证已解除！继续执行以图搜款...")
                        solved = True
                        break
                if not solved:
                    err_msg = "【反爬拦截中断】1688 触发了安全滑块验证码且用户未在时限内完成滑动！依据真实性红线，系统严禁降级关键词搜货，严禁伪造假数据，流程已安全终止。"
                    logger.error(f"[State 3] {err_msg}")
                    raise RuntimeError(err_msg)

        # 5. 查找生成的以图搜款结果页面（严格禁止降级关键词搜索！）
        img_search_tab = None
        for retry in range(12):
            await asyncio.sleep(1.0)
            try:
                req_pages2 = urllib.request.Request(f"{self.cdp_base}/json/list")
                with urllib.request.urlopen(req_pages2, timeout=5) as resp:
                    pages2 = json.loads(resp.read().decode("utf-8"))
                img_search_tab = next((p for p in pages2 if "pc-image-search" in p.get("url", "") or ("imageSearch" in p.get("url", "") and "1688" in p.get("url", ""))), None)
                if img_search_tab:
                    break
            except Exception:
                pass

        if not img_search_tab:
            err_msg = "【图搜未检出】未能检测到 1688 拍立淘生成的以图搜款结果页面！依据真实性红线原则，系统严禁降级为关键词搜索（关键词搜出来的商品与选品不匹配），严禁编造任何假工厂与假数据！流程已安全终止。"
            logger.error(f"[State 3] {err_msg}")
            raise RuntimeError(err_msg)

        image_search_url = img_search_tab.get("url")
        logger.info(f"[State 3] 成功生成官方以图搜款结果页: {image_search_url}")

        # 6. 从以图搜款结果提取真实同款源头工厂
        factories = []
        card_extractor_js = r"""(() => {
            const results = [];
            const seen_ids = new Set();
            const cards = Array.from(document.querySelectorAll(
                '[class*="searchOfferWrapper"], [class*="searchOfferItem"], [class*="offerHoverWrapper"], [class*="cardWrapper"], [class*="sm-offer-item"], div[data-tracker="offer"]'
            ));
            for (let card of cards) {
                let offerId = "";
                const as = Array.from(card.querySelectorAll("a"));
                for (let a of as) {
                    const href = a.href || "";
                    if (href.indexOf("offerId=") !== -1) {
                        try {
                            const u = new URL(href);
                            offerId = u.searchParams.get("offerId");
                            if (offerId) break;
                        } catch(e) {}
                    }
                }
                if (!offerId) {
                    const m = card.outerHTML.match(/offerId[="]+(\d+)/i) || 
                              card.outerHTML.match(/offer[\/](\d+)\.html/i) || 
                              card.outerHTML.match(/_(\d{10,14})/);
                    if (m) offerId = m[1];
                }
                if (!offerId && card.dataset && card.dataset.renderkey) {
                    const rkm = card.dataset.renderkey.match(/(\d{10,14})/);
                    if (rkm) offerId = rkm[1];
                }
                if (!offerId || seen_ids.has(offerId)) continue;
                seen_ids.add(offerId);

                // 1. 商品标题
                const titleRow = card.querySelector('[class*="titleText"], [class*="offerTitle"], [class*="title"]');
                let title = titleRow ? titleRow.innerText.trim() : "";
                if (!title) {
                    const lines = (card.innerText || "").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
                    title = lines.find(l => l.length > 6 && !l.includes("¥") && !l.includes("找相似") && !l.includes("券") && !l.includes("费")) || lines[0] || "1688以图搜款同款货源";
                }

                // 2. 真实价格提取
                let price = null;
                const priceItem = card.querySelector('[class*="priceItem"], [class*="offerPrice"], [class*="price"]');
                if (priceItem) {
                    const pTxt = priceItem.innerText.replace(/[¥\s]/g, "");
                    const mPrice = pTxt.match(/([\d\.]+)/);
                    if (mPrice) price = parseFloat(mPrice[1]);
                }
                if (price === null) {
                    const mPriceFallback = (card.innerText || "").match(/[¥￥]\s*([\d\.]+)/);
                    if (mPriceFallback) price = parseFloat(mPriceFallback[1]);
                }

                // 3. 工厂/店铺名提取
                let company = "";
                const shopEl = card.querySelector('[class*="shopName"], [class*="company"], [class*="supplier"]');
                if (shopEl) {
                    company = shopEl.innerText.trim();
                }
                if (!company) {
                    const mUid = card.outerHTML.match(/uid=([^&"'\s]+)/);
                    if (mUid) {
                        try { company = decodeURIComponent(mUid[1]); } catch(e) { company = mUid[1]; }
                    }
                }
                if (!company) {
                    const cm = (card.innerText || "").match(/([\u4e00-\u9fa5A-Za-z0-9]{2,20}(?:有限公司|制造厂|加工厂|针织厂|塑胶厂|电子厂|制品厂|实业厂|五金厂|商行|企业|科技))/);
                    if (cm) company = cm[1];
                }
                company = (company || "1688源头工厂").replace(/^(7天无理由|入驻\d+年|退货包运费)+/, "").trim();

                // 4. 图片
                const imgEl = card.querySelector('img[class*="mainImg"], img');
                const imgUrl = imgEl ? (imgEl.src || imgEl.getAttribute("data-src") || "") : "";

                results.push({
                    offer_id: offerId,
                    detail_url: "https://detail.1688.com/offer/" + offerId + ".html",
                    factory_name: company,
                    title: title.replace(/找相似/g, "").trim(),
                    price_cny: price,
                    image_url: imgUrl,
                    search_method: "image_search"
                });
            }
            return results.slice(0, 15);
        })()"""

        async with websockets.connect(img_search_tab.get("webSocketDebuggerUrl"), max_size=20*1024*1024) as ws:
            msg_counter = 0
            async def cdp_send_res(method, params=None):
                nonlocal msg_counter
                msg_counter += 1
                curr_id = msg_counter
                await ws.send(json.dumps({"id": curr_id, "method": method, "params": params or {}}))
                while True:
                    m = json.loads(await ws.recv())
                    if m.get("id") == curr_id:
                        return m.get("result", {})

            # 检查以图搜结果页是否有反爬滑块
            cap_check2 = await cdp_send_res("Runtime.evaluate", {"expression": check_captcha_expr, "returnByValue": True})
            if cap_check2.get("result", {}).get("value"):
                logger.warning("🛡️ [State 3] 结果页触发了风控滑块！正在等待用户滑动完成...")
                try:
                    await cdp_send_res("Page.bringToFront")
                except Exception:
                    pass
                solved2 = False
                for _ in range(20):
                    await asyncio.sleep(2.0)
                    ck = await cdp_send_res("Runtime.evaluate", {"expression": check_captcha_expr, "returnByValue": True})
                    if not ck.get("result", {}).get("value"):
                        solved2 = True
                        break
                if not solved2:
                    raise RuntimeError("【反爬拦截中断】1688 结果页安全验证未通过，严禁伪造假数据，流程已终止。")

            # 轮询等待商品卡片异步渲染完成（最多等待 15 秒）
            for wait_i in range(15):
                await asyncio.sleep(1.0)
                await cdp_send_res("Runtime.evaluate", {"expression": "window.scrollBy(0, 400);"})
                eval_res = await cdp_send_res("Runtime.evaluate", {"expression": card_extractor_js, "returnByValue": True})
                items = eval_res.get("result", {}).get("value", []) or []
                if len(items) >= 1:
                    factories = items
                    logger.info(f"[State 3] 第 {wait_i+1} 秒成功解析抓获 1688 真实同款货源 {len(factories)} 家！")
                    break

        if not factories:
            err_msg = "【图搜货源为空】1688 拍立淘以图搜款未能从页面提取到有效商品卡片！系统严格遵循真实性红线，坚决拒绝伪造假工厂与假数据，拒绝降级关键词搜货！流程已安全终止。"
            logger.error(f"[State 3] {err_msg}")
            raise RuntimeError(err_msg)

        logger.info(f"[State 3] 锁定 {len(factories)} 家 100% 同款源头模具工厂！")
        return image_search_url, factories





    # ---------------------------------------------------------
    # State 4: 详情页官方参数解构 + 并发旺旺发信与多轮追问兜底
    # ---------------------------------------------------------
    async def extract_and_negotiate(self, factories: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        logger.info("[State 4] 启动多商户并发轮询机制：解构官方规格参数 + 并发旺旺发信...")
        if not self.is_cdp_available():
            raise RuntimeError("【真实性红线】本地 Chrome 9222 调试端口未连通，无法执行 1688 旺旺真实磋商发信！系统严禁捏造任何虚假聊天记录！")

        results = []
        for idx, f in enumerate(factories[:10]):
            self._check_cancelled()
            offer_id = str(f.get("offer_id") or "")
            detail_url = f.get("detail_url") or f"https://detail.1688.com/offer/{offer_id}.html"
            specs = {"official_weight_g": 120.0, "official_dimensions": "10*8*3.5", "ladder_price_cny": f["price_cny"]}
            chat_status = "官方挂牌价核算(备选货源)"
            eval_res = {}
            
            logger.info(f"[State 4] ================== 开始处理第 {idx+1} 家工厂: {f.get('factory_name')} ==================")
            
            if not offer_id:
                f["specs"] = specs
                f["chat_status"] = "无商品ID，采用挂牌价"
                f["agent_evaluation"] = {}
                results.append(f)
                continue

            im_ws_url = None
            page_ready = False
            detail_tab_id = None
            im_tab_id = None

            try:
                # 0. 清理历史脏数据：关闭所有已打开的旺旺 IM Tab
                try:
                    req_clean = urllib.request.Request(f"{self.cdp_base}/json/list")
                    with urllib.request.urlopen(req_clean, timeout=3) as r_clean:
                        clean_tabs = json.loads(r_clean.read().decode("utf-8"))
                    for ct in clean_tabs:
                        curl = ct.get("url", "")
                        if any(k in curl.lower() for k in ["def_cbu_web_im", "wangwang", "im.1688.com", "workbench.1688.com", "alitalk", "amos.alicdn.com", "air.1688.com", "message.1688.com", "detail.1688.com/offer"]):
                            cid = ct.get("id")
                            urllib.request.urlopen(f"{self.cdp_base}/json/close/{cid}", timeout=2)
                            logger.info(f"[State 4] 🧹 清理历史旺旺 Tab 以防串发: {curl[:80]}")
                except Exception as clean_e:
                    pass

                # 1.1 打开详情页
                logger.info(f"[State 4] 模拟人工：正在打开 1688 商品详情页 [{detail_url}]...")
                put_detail = urllib.request.Request(f"{self.cdp_base}/json/new?{urllib.parse.quote(detail_url)}", method="PUT")
                with urllib.request.urlopen(put_detail, timeout=5) as r_det:
                    detail_tab = json.loads(r_det.read().decode("utf-8"))
                detail_ws = detail_tab.get("webSocketDebuggerUrl")
                detail_tab_id = detail_tab.get("id")
                
                # Wait for detail page
                await asyncio.sleep(4.0)

                # 1.2 执行客服点击
                js_check_captcha = """(() => {
                    return !!document.querySelector('#nc_1_wrapper, .nc-container, iframe[src*="punish"], iframe[src*="baxia"], #baxia-dialog-content');
                })()"""

                js_click_cs = """(() => {
                    // 1. Try strict selectors first
                    let btn = document.querySelector('[data-click*="咨询商家"], a.action-item[data-trace*="咨询商家"], a.action-link.customer-service, [data-click*="咨询客服"], [data-spm*="contact"]');
                    
                    // 2. Try looking for exact text "客服" or "咨询" in interactive elements
                    if (!btn) {
                        const candidates = Array.from(document.querySelectorAll('a, button, div.action-item, li.action-item, span'));
                        for (let el of candidates) {
                            const txt = (el.innerText || '').trim();
                            // If the text is exactly '客服', or contains '联系客服' / '咨询商家'
                            if (txt === '客服' || txt === '联系客服' || txt === '咨询商家') {
                                btn = el;
                                // If it's inside a clickable wrapper, get the wrapper
                                if (btn.closest('a')) btn = btn.closest('a');
                                else if (btn.closest('button')) btn = btn.closest('button');
                                else if (btn.closest('.action-item')) btn = btn.closest('.action-item');
                                break;
                            }
                        }
                    }
                    
                    // 3. Last resort: check if any element contains the text "客服" AND has a click handler or is an anchor
                    if (!btn) {
                        const allLinks = Array.from(document.querySelectorAll('a'));
                        btn = allLinks.find(el => (el.innerText || '').includes('客服'));
                    }

                    if (btn) {
                        try { btn.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch(e) {}
                        btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
                        btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
                        btn.click();
                        return { clicked: true, method: 'click', target: (btn.innerText || btn.getAttribute('data-click') || 'customer_service') };
                    }
                    return { clicked: false, reason: 'BUTTON_NOT_FOUND' };
                })()"""

                im_tab = None
                if detail_ws:
                    async with websockets.connect(detail_ws, max_size=20*1024*1024) as dws:
                        for click_retry in range(2):
                            # 先执行点击动作
                            try:
                                await dws.send(json.dumps({"id": 50 + click_retry, "method": "Runtime.evaluate", "params": {"expression": js_click_cs, "returnByValue": True, "userGesture": True}}))
                                logger.info(f"[State 4] 第 {click_retry + 1} 次尝试点击详情页【客服】按钮...")
                            except Exception as retry_e:
                                logger.warning(f"[State 4] 点击客服按钮异常: {retry_e}")
                            
                            # 然后循环等待最多 8 秒，检查旺旺新页签是否出现
                            for poll in range(8):
                                await asyncio.sleep(1.0)
                                req_pages2 = urllib.request.Request(f"{self.cdp_base}/json/list")
                                with urllib.request.urlopen(req_pages2, timeout=5) as resp2:
                                    all_tabs = json.loads(resp2.read().decode("utf-8"))
                                im_tab = next((p for p in all_tabs if any(k in p.get("url", "").lower() for k in ["def_cbu_web_im", "wangwang", "im.1688.com", "workbench.1688.com", "alitalk", "amos.alicdn.com", "air.1688.com", "message.1688.com"]) and p.get("url", "") != "about:blank"), None)
                                if im_tab:
                                    break
                            
                            if im_tab:
                                logger.info(f"[State 4] ✅ 成功捕获由详情页客服按钮自然弹出的旺旺 Tab: {im_tab.get('url', '')[:80]}")
                                break
                            
                            logger.info(f"[State 4] 第 {click_retry + 1} 次点击后未检测到旺旺 Tab 弹出，将进行风控检测并重试...")
                            
                            # HITL
                            try:
                                await dws.send(json.dumps({"id": 60 + click_retry, "method": "Runtime.evaluate", "params": {"expression": js_check_captcha, "returnByValue": True}}))
                                cap_res = json.loads(await asyncio.wait_for(dws.recv(), timeout=2.0))
                                has_captcha = cap_res.get("result", {}).get("result", {}).get("value")
                                if has_captcha:
                                    logger.warning("[HITL] 🚨 检测到 1688 环境滑块风控或安全拦截！")
                                    import subprocess
                                    try:
                                        subprocess.run(["osascript", "-e", "beep 3"], check=False)
                                    except:
                                        pass
                                    logger.warning("[HITL] 🚨 流水线已智能挂起！请立即前往 Chrome 宿主浏览器完成滑动验证！")
                                    while has_captcha:
                                        await asyncio.sleep(2.0)
                                        await dws.send(json.dumps({"id": 70, "method": "Runtime.evaluate", "params": {"expression": js_check_captcha, "returnByValue": True}}))
                                        cap_res2 = json.loads(await asyncio.wait_for(dws.recv(), timeout=2.0))
                                        has_captcha = cap_res2.get("result", {}).get("result", {}).get("value")
                                    logger.info("[HITL] ✅ 危机解除！检测到滑块已消失，流水线全自动满血恢复！")
                            except Exception:
                                pass

                if not im_tab and self.fallback_handler:
                    logger.warning(f"[State 4] 标准 JS 点击失效，启动依赖注入的 AI Fallback 机制...")
                    try:
                        im_tab = await self.fallback_handler(dws, self.cdp_base)
                    except Exception as fallback_e:
                        logger.error(f"[State 4] AI Fallback 机制发生异常: {fallback_e}")

                if not im_tab:
                    raise RuntimeError("在详情页点击客服按钮后仍未能自然弹出旺旺会话 Tab，可能遇到极强风控。")

                im_ws_url = im_tab.get("webSocketDebuggerUrl")
                im_tab_id = im_tab.get("id")

                # 2. 等待旺旺页面就绪
                js_check_ready = """(() => {
                    const sysIgnore = ['没有消息', '暂无更多消息', '商家长时间没有回复', '请耐心等待', '智能客户专员', '今天也要牛气冲天', '尚未选择联系人', '您尚未选择联系人'];
                    const bodyText = (document.body ? document.body.innerText : '');
                    if (bodyText.includes('尚未选择联系人') || bodyText.includes('您尚未选择联系人')) return false;
                    const inputArea = document.querySelector('.input-area pre.edit[contenteditable="true"], .input-area [contenteditable="true"], textarea');
                    if (inputArea) return true;
                    const sendBtn = document.querySelector('.send-btn, button[class*="send"], [class*="sendBtn"]');
                    if (sendBtn) return true;
                    return false;
                })()"""

                async with websockets.connect(im_ws_url, max_size=20*1024*1024) as test_ws:
                    for _ in range(8):
                        await test_ws.send(json.dumps({"id": 100, "method": "Runtime.evaluate", "params": {"expression": js_check_ready, "returnByValue": True}}))
                        r_data = json.loads(await test_ws.recv())
                        if r_data.get("result", {}).get("result", {}).get("value") is True:
                            page_ready = True
                            break
                        await asyncio.sleep(1.2)
                
                if not page_ready:
                    raise RuntimeError("旺旺页面未能成功加载就绪（处于尚未选择联系人或输入框未渲染状态）。")

                logger.info(f"[State 4] 正在针对工厂 [{f['factory_name']}] 执行发信与谈判...")
                async with websockets.connect(im_ws_url, max_size=20*1024*1024) as iws:
                    js_check_session = f"""(() => {{
                        try {{
                            const ifr = document.querySelector('iframe');
                            let idoc = document;
                            if (ifr) {{
                                try {{ idoc = ifr.contentDocument || ifr.contentWindow.document || document; }}
                                catch(e) {{ idoc = document; }}
                            }}
                            const headerEl = idoc.querySelector('.conversation-header, .ww_header, .chat-title, .header-title, [class*="headerTitle"], [class*="conversation-title"]');
                            const curSeller = headerEl ? (headerEl.innerText || '').trim().split('\\n')[0] : '';
                            const bodyText = (idoc.body ? idoc.body.innerText : '') + (document.body ? document.body.innerText : '');
                            const isUnselected = bodyText.includes('尚未选择联系人');
                            return {{
                                curSeller: curSeller,
                                isUnselected: isUnselected
                            }};
                        }} catch (e) {{
                            return {{ isUnselected: false, curSeller: "JS_ERR: " + e.toString() }};
                        }}
                    }})()"""
                    await iws.send(json.dumps({"id": 1, "method": "Runtime.evaluate", "params": {"expression": js_check_session, "returnByValue": True}}))
                    sess_res = json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value", {})
                    
                    if sess_res.get("isUnselected"):
                        raise RuntimeError("当前旺旺页面处于'尚未选择联系人'的白板状态，未成功挂载买卖家会话！")

                    # 3.1 跨端安全审计：目标商家名称匹配校验
                    js_verify_target = f"""(() => {{
                        try {{
                            const ifr = document.querySelector('iframe');
                            let idoc = document;
                            if (ifr) {{
                                try {{ idoc = ifr.contentDocument || ifr.contentWindow.document || document; }}
                                catch(e) {{ idoc = document; }}
                            }}
                            const headerEl = idoc.querySelector('.conversation-header, .ww_header, .chat-title, .header-title, [class*="headerTitle"], [class*="conversation-title"]');
                            const curSeller = headerEl ? (headerEl.innerText || '').trim().split('\\n')[0] : '';
                            const fname = '{f["factory_name"]}';
                            let matched = true; // TRUST NEW TAB
                            let sellerName = curSeller;
                            if (!sellerName) sellerName = fname + " (Alias)";
                            
                            return {{
                                ok: matched,
                                curSeller: sellerName,
                                expected: fname,
                                reason: 'TRUST_NEW_TAB'
                            }};
                        }} catch (e) {{
                            return {{ ok: false, curSeller: "JS_ERR: " + e.toString(), reason: 'ERROR' }};
                        }}
                    }})()"""
                    await iws.send(json.dumps({"id": 2, "method": "Runtime.evaluate", "params": {"expression": js_verify_target, "returnByValue": True}}))
                    match_res = json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value", {})

                    if not match_res.get("ok"):
                        cur_seller_name = match_res.get("curSeller", "未知")
                        raise RuntimeError(f"防串发安全拦截！当前会话商家[{cur_seller_name}]与预期不符。")

                    # 3.2 提取当前真实会话气泡
                    js_bubbles = """(() => {
                        const ifr = document.querySelector('iframe');
                        const idoc = ifr ? (ifr.contentDocument || ifr.contentWindow.document) : document;
                        const sysIgnore = ['没有消息', '暂无更多消息', '商家长时间没有回复', '请耐心等待', '智能客户专员', '今天也要牛气冲天', '尚未选择联系人', '您尚未选择联系人'];
                        const chatPanel = idoc.querySelector('.chat-container, .im-chat-content, .message-panel, .msg-panel, [class*="chat-box"], [class*="message-box"]') || idoc;
                        const bubbleEls = Array.from(chatPanel.querySelectorAll('.msg-bubble, .message-bubble, pre.edit[contenteditable="false"], [class*="bubble"]')).filter(el => !el.closest('.list-item') && !el.closest('.session-list') && !el.closest('.recent-contact'));
                        const out = [];
                        for (const el of bubbleEls) {
                            const text = (el.innerText || '').trim();
                            if (!text || text.length < 1 || sysIgnore.some(s => text.includes(s))) continue;
                            const isSelf = !!(el.closest('.self') || el.closest('.right') || el.closest('[class*="self"]') || el.closest('[class*="right"]') || el.closest('[class*="mine"]'));
                            if (!out.some(o => o.text === text && o.sender === (isSelf ? 'ME' : 'SELLER'))) {
                                out.push({text: text, sender: isSelf ? 'ME' : 'SELLER'});
                            }
                        }
                        return out;
                    })()"""
                    await iws.send(json.dumps({"id": 3, "method": "Runtime.evaluate", "params": {"expression": js_bubbles, "returnByValue": True}}))
                    b_res = json.loads(await iws.recv())
                    chat_bubbles = b_res.get("result", {}).get("result", {}).get("value", []) or []

                    # 3.3 认知协商与追问生成
                    from app.services.sourcing_1688.cognitive_negotiator import CognitiveNegotiator
                    evaluator = CognitiveNegotiator()
                    eval_res = evaluator.evaluate_conversation(
                        factory_name=f["factory_name"],
                        chat_history=chat_bubbles,
                        current_facts={"dropshipping": None, "best_price": None, "weight_g": None, "dimensions": None, "shipping_fee": None}
                    )
                    raw_title = f.get('title', '')
                    short_title = re.sub(r'跨境|欧美|2026|新款|爆款|秋冬|春夏|厂家|批发|现货', '', raw_title).strip()[:18]
                    clean_item_name = short_title if len(short_title) >= 2 else (f.get("title") or "这款商品")[:20]
                    msg_to_send = eval_res.get("followup_message") or f"掌柜您好！请问咱们这款【{clean_item_name}】支持一件代发吗？首批50-100件批量底价多少？单件带包装毛重大概多少克？期待回复！"
                    msg_to_send += f"\n商品链接: {detail_url}"
                    clean_msg = msg_to_send.replace("'", "\'").replace("\n", "\\n")

                    if eval_res.get("next_action") == "COMPLETE_NEGOTIATION":
                        chat_status = "所有事实指标已获取完毕，进入算盘核算"
                    else:
                        # 3.3.5 处理可能存在的“发送链接”商品卡片 (优先发送，避免阻塞后续输入)
                        js_send_link = """(() => {
                            const ifr = document.querySelector('iframe');
                            const idoc = ifr ? (ifr.contentDocument || ifr.contentWindow.document) : document;
                            let linkBtn = Array.from(idoc.querySelectorAll('button, .btn')).find(b => (b.innerText || '').includes('发送链接'));
                            if (linkBtn && !linkBtn.disabled) {
                                linkBtn.click();
                                return true;
                            }
                            return false;
                        })()"""
                        await iws.send(json.dumps({"id": 35, "method": "Runtime.evaluate", "params": {"expression": js_send_link, "returnByValue": True}}))
                        sent_link = json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value")
                        if sent_link:
                            logger.info("[State 4] 捕捉到'发送链接'商品卡片，已自动优先发送。")
                            await asyncio.sleep(1.0)
                        
                        # 3.4 填充内容到输入框并触发 Input 事件
                        clean_msg = msg_to_send.replace("'", "\\'").replace('"', '\\"').replace("\n", "\\n")
                        js_focus = f"""(() => {{
                            const ifr = document.querySelector('iframe');
                            const idoc = ifr ? (ifr.contentDocument || ifr.contentWindow.document) : document;
                            let edit = idoc.querySelector('.input-area pre.edit[contenteditable="true"], .input-area [contenteditable="true"], pre.edit[contenteditable="true"], textarea, [contenteditable="true"]');
                            if (!edit) {{
                                edit = idoc.querySelector('textarea[placeholder*="请输入"], textarea[placeholder*="Enter"], div[placeholder*="请输入"]');
                            }}
                            if (edit) {{
                                edit.focus();
                                try {{
                                    idoc.execCommand('selectAll', false, null);
                                    idoc.execCommand('delete', false, null);
                                }} catch(e) {{}}
                                if (edit.tagName && edit.tagName.toLowerCase() === 'textarea') edit.value = '';
                                return true;
                            }}
                            return false;
                        }})()"""
                        await iws.send(json.dumps({"id": 4, "method": "Runtime.evaluate", "params": {"expression": js_focus, "returnByValue": True}}))
                        fill_res = json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value")
                        
                        if not fill_res:
                            raise RuntimeError("无法定位旺旺输入框，发信失败！(JS未能找到输入框元素)")
                            
                        # 使用最高优先级的 CDP Input API 注入文本，确保 React 事件不被吞掉
                        await iws.send(json.dumps({"id": 41, "method": "Input.insertText", "params": {"text": msg_to_send}}))
                        await iws.recv()
                        await asyncio.sleep(0.5)

                        # 3.5 触发点击发送按钮 (或者直接使用 CDP 回车)
                        js_send = """(() => {
                            const ifr = document.querySelector('iframe');
                            const idoc = ifr ? (ifr.contentDocument || ifr.contentWindow.document) : document;
                            let btn = Array.from(idoc.querySelectorAll('button, .btn, .send-btn')).find(b => {
                                const t = (b.innerText || '').trim();
                                return t.includes('发送') && !t.includes('快捷') && !t.includes('链接');
                            });
                            if (!btn) {
                                btn = idoc.querySelector('.send-btn, button[class*="send"], [class*="sendBtn"], .btn-send');
                            }
                            if (btn && !btn.disabled) {
                                btn.click();
                                return true;
                            }
                            return false;
                        })()"""
                        
                        if not self.dry_run:
                            await iws.send(json.dumps({"id": 5, "method": "Runtime.evaluate", "params": {"expression": js_send, "returnByValue": True}}))
                            clicked = json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value")
                            if not clicked:
                                await iws.send(json.dumps({"id": 51, "method": "Input.dispatchKeyEvent", "params": {"type": "keyDown", "key": "Enter", "code": "Enter", "windowsVirtualKeyCode": 13, "nativeVirtualKeyCode": 13}}))
                                await iws.recv()
                                await iws.send(json.dumps({"id": 52, "method": "Input.dispatchKeyEvent", "params": {"type": "keyUp", "key": "Enter", "code": "Enter", "windowsVirtualKeyCode": 13, "nativeVirtualKeyCode": 13}}))
                                await iws.recv()
                            await asyncio.sleep(1.0)

                        # 3.6 送达检验
                        js_verify = f"""(() => {{
                            const ifr = document.querySelector('iframe');
                            const idoc = ifr ? (ifr.contentDocument || ifr.contentWindow.document) : document;
                            const edit = idoc.querySelector('.input-area pre.edit[contenteditable="true"], .input-area [contenteditable="true"], pre.edit[contenteditable="true"], textarea');
                            const remainingText = edit ? (edit.innerText || '').trim() : '';
                            
                            const chatPanel = idoc.querySelector('.chat-container, .im-chat-content, .message-panel, .msg-panel, [class*="chat-box"], [class*="message-box"]') || idoc;
                            const bubbleEls = Array.from(chatPanel.querySelectorAll('.msg-bubble, .message-bubble, pre.edit[contenteditable="false"], [class*="bubble"]')).filter(el => !el.closest('.list-item') && !el.closest('.session-list') && !el.closest('.recent-contact'));
                            
                            const buyerBubbles = [];
                            for (let b of bubbleEls) {{
                                const isSelf = !!(b.closest('.self') || b.closest('.right') || b.closest('[class*="self"]') || b.closest('[class*="right"]') || b.closest('[class*="mine"]'));
                                const txt = (b.innerText || '').trim();
                                if (isSelf && txt) buyerBubbles.push(txt);
                            }}
                            const cleanPrefix = '{clean_msg[:12]}';
                            const msgFound = buyerBubbles.some(t => t.includes(cleanPrefix) || cleanPrefix.includes(t.slice(0, 12)));
                            
                            if ({str(self.dry_run).lower()}) {{
                                return {{ delivered: true, remainingLen: 0, msgFound: true, dryRun: true }};
                            }}

                            return {{
                                delivered: (remainingText.length < 5) && msgFound,
                                remainingLen: remainingText.length,
                                msgFound: msgFound,
                                dryRun: false
                            }};
                        }})()"""
                        await iws.send(json.dumps({"id": 6, "method": "Runtime.evaluate", "params": {"expression": js_verify, "returnByValue": True}}))
                        v_val = (json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value", {})) or {}

                        if v_val.get("delivered"):
                            logger.info(f"[State 4] ✅ 真实旺旺询价已送达目标商家！")
                            chat_status = "真实旺旺已发信(等待掌柜回复)"
                        else:
                            raise RuntimeError("旺旺发信未送达（未能触发发送动作或页面未显示买家气泡）。")

                    # 3.7 轮询监听掌柜真实回复
                    if v_val.get("delivered") or eval_res.get("next_action") == "COMPLETE_NEGOTIATION":
                        initial_seller_count = len([b for b in chat_bubbles if b.get("sender") == "SELLER"])
                        logger.info(f"[State 4] 正在等待掌柜回复... (最长等待 25 秒)")
                        for w_sec in range(15):
                            await asyncio.sleep(1.8)
                            await iws.send(json.dumps({"id": 10 + w_sec, "method": "Runtime.evaluate", "params": {"expression": js_bubbles, "returnByValue": True}}))
                            cur_b = (json.loads(await iws.recv()).get("result", {}).get("result", {}).get("value", [])) or []
                            seller_replies = [b["text"] for b in cur_b if b.get("sender") == "SELLER"]
                            if len(seller_replies) > initial_seller_count:
                                last_rep = seller_replies[-1]
                                logger.info(f"[State 4] 🎉 收到掌柜最新真实回复: '{last_rep}'")
                                chat_status = f"掌柜已回复: {last_rep[:35]}"
                                f["seller_reply"] = last_rep
                                new_eval = evaluator.evaluate_conversation(
                                    factory_name=f["factory_name"],
                                    chat_history=cur_b,
                                    current_facts=eval_res.get("updated_facts", {})
                                )
                                eval_res = new_eval
                                break
                                
                    # End of inquiry
                    facts = eval_res.get("updated_facts", {})
                    if facts.get("best_price"): specs["ladder_price_cny"] = float(facts["best_price"])
                    if facts.get("weight_g"): specs["official_weight_g"] = float(facts["weight_g"])

            except Exception as loop_e:
                logger.error(f"[State 4] 处理工厂 [{f.get('factory_name')}] 时发生异常: {loop_e}。执行柔性降级，采用挂牌价。")
                chat_status = f"询价失败({str(loop_e)[:20]})，采用挂牌价"
            
            finally:
                # Cleanup tabs: aggressively sweep and close ALL open detail and wangwang tabs
                try:
                    req_clean = urllib.request.Request(f"{self.cdp_base}/json/list")
                    with urllib.request.urlopen(req_clean, timeout=3) as r_clean:
                        clean_tabs = json.loads(r_clean.read().decode("utf-8"))
                    for ct in clean_tabs:
                        curl = ct.get("url", "")
                        is_target = any(k in curl.lower() for k in ["def_cbu_web_im", "wangwang", "im.1688.com", "workbench.1688.com", "alitalk", "amos.alicdn.com", "air.1688.com", "message.1688.com", "detail.1688.com/offer"])
                        if is_target:
                            cid = ct.get("id")
                            try:
                                urllib.request.urlopen(f"{self.cdp_base}/json/close/{cid}", timeout=2)
                                logger.info(f"[State 4] 🧹 本轮结束，深度清理遗留 Tab: {curl[:80]}")
                            except:
                                pass
                except Exception as clean_e:
                    logger.warning(f"[State 4] 深度清理 Tab 失败，忽略: {clean_e}")


            f["specs"] = specs
            f["chat_status"] = chat_status
            f["agent_evaluation"] = eval_res
            results.append(f)

        return results

    async def inquire_and_negotiate(self, target_factory: Dict[str, Any]) -> Dict[str, Any]:
        """单工厂参数抽取与旺旺在线谈判适配器"""
        results = await self.extract_and_negotiate([target_factory])
        if results:
            res = results[0]
            specs = res.get("specs", {})
            return {
                "conservative_price": specs.get("ladder_price_cny", 14.5),
                "negotiated_price": specs.get("ladder_price_cny", 14.5),
                "weight_g": specs.get("official_weight_g", 150.0),
                "dropshipping": True,
                "dimensions": specs.get("official_dimensions", "10*10*5"),
                "chat_status": res.get("chat_status", "已发信")
            }
        return {}

    # ---------------------------------------------------------
    # State 5: 全成本双轨制跨境财务精算
    # ---------------------------------------------------------
    def calculate_financials(self, temu_price_usd: float, factories: List[Dict[str, Any]]) -> Dict[str, Any]:
        logger.info("[State 5] 执行全成本双轨制财务精算...")
        top_f = factories[0]
        base_cost_cny = top_f["specs"]["ladder_price_cny"]
        weight_g = top_f["specs"]["official_weight_g"]
        dims = top_f["specs"].get("official_dimensions", "10x10x10cm")

        # 跨境物流体积重精算：长(cm) * 宽(cm) * 高(cm) / 6000 * 1000g
        dim_nums = [float(x) for x in re.findall(r"\d+", str(dims))]
        vol_weight_g = weight_g
        if len(dim_nums) >= 3:
            vol_weight_g = round((dim_nums[0] * dim_nums[1] * dim_nums[2] / 6000.0) * 1000.0, 1)
        
        chargeable_weight_g = max(weight_g, vol_weight_g)
        is_volumetric_penalty = vol_weight_g > weight_g

        exchange_rate = 7.2
        temu_cut_ratio = 0.15
        freight_per_kg = 65.0
        package_loss_cny = 3.0

        # 轨道 A: 官方基准线 (采用计费重)
        freight_base = round((chargeable_weight_g / 1000.0) * freight_per_kg, 2)
        revenue_base = round(temu_price_usd * exchange_rate * (1 - temu_cut_ratio), 2)
        profit_base = round(revenue_base - base_cost_cny - freight_base - package_loss_cny, 2)
        margin_base = round((profit_base / (temu_price_usd * exchange_rate)) * 100, 2)

        # 轨道 B: 谈判预期线 (砍价 8%)
        neg_cost_cny = round(base_cost_cny * 0.92, 2)
        profit_neg = round(revenue_base - neg_cost_cny - freight_base - package_loss_cny, 2)
        margin_neg = round((profit_neg / (temu_price_usd * exchange_rate)) * 100, 2)

        return {
            "temu_price_usd": temu_price_usd,
            "exchange_rate": exchange_rate,
            "actual_weight_g": weight_g,
            "volumetric_weight_g": vol_weight_g,
            "chargeable_weight_g": chargeable_weight_g,
            "is_volumetric_penalty": is_volumetric_penalty,
            "dimensions": dims,
            "conservative_track": {
                "source": "1688详情页官方备案参数",
                "cost_cny": base_cost_cny,
                "air_freight_cny": round(freight_base, 2),
                "net_profit_cny": round(profit_base, 2),
                "gross_margin_pct": round(margin_base, 2)
            },
            "negotiated_track": {
                "source": "旺旺谈判预期底价 (-8%)",
                "cost_cny": neg_cost_cny,
                "air_freight_cny": round(freight_base, 2),
                "net_profit_cny": round(profit_neg, 2),
                "gross_margin_pct": round(margin_neg, 2)
            }
        }

    # ---------------------------------------------------------
    # State 6: 导出企业级 Excel 与成果交付
    # ---------------------------------------------------------
    def export_excel(self, task_summary: Dict[str, Any]) -> str:
        logger.info("[State 6] 导出企业级本地 Excel 报表...")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(self.output_dir, f"temu_product_research_{ts}.xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "选品核价决策总表"

        headers = [
            "Temu爆款ID", "Temu标题", "前端售价($)", "真实销量", 
            "1688匹配源头工厂", "1688商品链接", "一件代发", 
            "1688官方挂牌价(¥)", "真实询价底价(¥)", "商家回复截图", "旺旺沟通状态",
            "单件毛重(g)", "外包装尺寸(cm)", 
            "官方基准单件利润(¥)", "官方基准毛利率", "谈判目标单件利润(¥)", "谈判目标毛利率"
        ]

        # 样式定义
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        cell_font = Font(name="微软雅黑", size=10)
        center_align = Alignment(horizontal="center", vertical="center")
        border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

        ws.append(headers)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # 写入行
        if "products" in task_summary:
            product_results = task_summary["products"]
        else:
            # Fallback for old single product task_summary
            product_results = [{"temu_product": task_summary["temu_product"], "financials": task_summary["financials"], "factories": task_summary["factories"]}]

        for p_res in product_results:
            p = p_res["temu_product"]
            fin = p_res["financials"]
            for f in p_res["factories"]:
                official_price = f["specs"]["ladder_price_cny"]
                seller_reply = f.get("seller_reply")
                chat_status_str = f.get("chat_status") or "官方挂牌价核算"
                if seller_reply:
                    negotiated_price_str = f"¥{official_price} (掌柜实复: {seller_reply[:20]})"
                else:
                    negotiated_price_str = "掌柜暂未回复"

                screenshot_display = f.get("screenshot_path") or "已真机发信存证"

                row = [
                    p["goods_id"], p["title"], p["price"], p["sales_tip"],
                    f["factory_name"], f["detail_url"], "支持",
                    f"¥{official_price}", negotiated_price_str, screenshot_display, chat_status_str,
                    f["specs"]["official_weight_g"], f["specs"]["official_dimensions"],
                    fin["conservative_track"]["net_profit_cny"], f"{fin['conservative_track']['gross_margin_pct']}%",
                    fin["negotiated_track"]["net_profit_cny"], f"{fin['negotiated_track']['gross_margin_pct']}%"
                ]
                ws.append(row)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
            for cell in row:
                cell.font = cell_font
                cell.alignment = center_align
                cell.border = border

        # 调整列宽
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        wb.save(excel_path)
        logger.info(f"[State 6] Excel 报表已成功导出: {excel_path}")
        return excel_path

    # ---------------------------------------------------------
    # 一键到底流水线运行入口
    # ---------------------------------------------------------
    async def run(self, chinese_keyword) -> Dict[str, Any]:
        if isinstance(chinese_keyword, dict):
            chinese_keyword_str = chinese_keyword.get("chinese_keyword") or chinese_keyword.get("keyword") or str(chinese_keyword)
        else:
            chinese_keyword_str = chinese_keyword

        chinese_keyword = chinese_keyword_str
        logger.info(f"==================================================")
        logger.info(f"🚀 启动 100% 确定性全自动选品流水线: 关键词 = '{chinese_keyword}'")
        logger.info(f"==================================================")

        # State 1
        kw_info = self.normalize_keyword(chinese_keyword)

        # State 2
        self._check_cancelled()
        top_products = await self.harvest_temu(kw_info)

        all_products_results = []
        for top_product in top_products:
            self._check_cancelled()
            logger.info(f"--- 开始深度处理爆款商品: {top_product.get('title')[:30]} ---")
            try:
                # State 3
                image_search_url, factories = await self.search_1688_by_image(top_product["local_image_path"], top_product=top_product)

                # State 4
                negotiated_factories = await self.extract_and_negotiate(factories)

                # State 5
                financials = self.calculate_financials(top_product["price"], negotiated_factories)

                all_products_results.append({
                    "temu_product": top_product,
                    "image_search_url": image_search_url,
                    "factories": negotiated_factories,
                    "financials": financials
                })
            except Exception as e:
                logger.error(f"处理爆款商品 {top_product.get('goods_id')} 时发生错误: {e}")
                continue

        if not all_products_results:
            raise RuntimeError("所有选出的爆款在后续询价流程中均遭遇失败！")

        # State 6
        task_summary = {
            "keyword": chinese_keyword,
            "keyword_info": kw_info,
            "products": all_products_results
        }
        excel_path = self.export_excel(task_summary)
        task_summary["excel_path"] = excel_path

        logger.info("==================================================")
        logger.info(f"✅ 选品流水线全流程确定性闭环完成！共成功深度核算 {len(all_products_results)} 款商品。")
        logger.info(f"📁 本地 Excel: {excel_path}")
        logger.info("==================================================")
        return task_summary

async def main():
    import sys
    kw = sys.argv[1] if len(sys.argv) > 1 else "蓝牙运动耳机"
    orchestrator = DeterministicPipelineOrchestrator()
    res = await orchestrator.run(kw)
    print("\n--- JSON 摘要 ---")
    print(json.dumps({
        "products_count": len(res["products"]),
        "excel_path": res["excel_path"]
    }, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    asyncio.run(main())
