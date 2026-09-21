# 店小二 MCP 工具直接调用铁律

1. **直接调用 MCP 工具，禁止阅读源码**：
   当用户提出跨境选品、Temu 商品检索/挖掘、1688 寻源核价、查询商户配额、利润精算等相关需求时，你必须且只能直接调用注册好的 `dianxiaoer_*` MCP 工具（如 `dianxiaoer_run_sourcing_pipeline`、`dianxiaoer_temu_search` 等）。
2. **严禁文件探索与代码阅读**：
   严禁使用 `view_file`、`grep_search`、`list_dir`、`run_command` 去查看、搜索或阅读本地任何 Python 文件（例如 `server.py`、`deterministic_orchestrator.py`、`run_pipeline.py`、`agent_workflow.py` 等）。
   MCP Server 工具已经封装了全部底层业务，不需要了解实现细节，直接传入参数调用工具即可！
3. **首选全自动选品工具**：
   当用户输入类似“选品 [关键词]”、“找爆款 [关键词]”、“在 Temu 挖掘 [关键词]”时，必须调用唯一的端到端选品工具：
   ```json
   dianxiaoer_run_sourcing_pipeline({
     "keyword": "用户关键词"
   })
   ```
