# Excel 表格助手（WorkBuddy Skill）

这是一个普通 WorkBuddy Skill 包，不使用 Connector。

## 包内内容

- `SKILL.md`：技能说明和操作规程
- `scripts/excel_mcp_bridge.py`：轻量 MCP 调用桥
- `scripts/start_server.py`：下载、校验并离线启动本地 Excel 引擎
- `runtime-lock.json`：依赖版本和 SHA-256 锁定清单
- `LICENSE`：上游 MIT 许可及原版权声明

## 当前运行支持

Windows x64 + CPython 3.13。首次使用会从维护者 GitHub Release 下载约 22.3 MB 运行依赖备份；解压前验证整包 SHA-256，之后逐个验证依赖哈希，并缓存到用户本机。Excel 文件本身通过本机进程处理，不会被上传到 GitHub。

## 安装

在 WorkBuddy「技能」页面选择「添加技能」并上传完整 ZIP。ZIP 内保留 `excel-workbook-assistant/` 目录及其全部文件。安装后用一个测试用 xlsx 做只读查询，再做一次复制文件的写入测试。

## 上游归属

本 Skill 使用基于 Haris Musa 的 `excel-mcp-server` v2.2.6 改编的本地引擎，遵守上游 MIT 许可。该引擎不是维护者原创。
