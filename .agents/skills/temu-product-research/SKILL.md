---
name: temu-product-research
description: 针对 Temu 跨境电商的极简全自动选品与 1688 供应链核价 Skill。当用户输入“选品”、“找爆款”、“挖掘/搜索商品”、“Temu选品”、“1688寻源”或任何跨境电商选品需求时立即触发。
---

# Temu 智能选品与 1688 供应链全自动执行规约

## 核心执行铁律：
收到任何选品核价需求时，必须且只能立即发起 MCP 工具调用：

- 推荐调用：
  call_mcp_tool(server_name="dianxiaoer", tool_name="dianxiaoer_run_sourcing_pipeline", arguments={"keyword": "用户关键词"})
- 或直接调用：
  dianxiaoer_run_sourcing_pipeline({"keyword": "用户关键词"})

### 🛑 绝对红线：
1. 严禁使用 run_command 在终端启动或测试 server.py！MCP Server 已由宿主系统自动托管！
2. 严禁使用 view_file, grep_search 查看本地 Python 源码或系统配置文件！
3. 工具返回后，直接呈递真实的 1688 工厂、Offer 链接、旺旺发信状态与 Excel 决策报表！
