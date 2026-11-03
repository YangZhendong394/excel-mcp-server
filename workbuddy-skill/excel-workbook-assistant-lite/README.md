# Excel 表格助手（轻量版原型）

该目录是独立的 WorkBuddy 普通 Skill 原型，不覆盖原 MCP 版目录。

当前原型包含本地 Excel CLI、Skill 指令和第三方依赖打包说明。发布 ZIP 时由 GitHub Actions 将固定版本 openpyxl 3.1.5 与 et-xmlfile 2.0.0 的纯 Python 包源码放入 vendor/，用户运行时不执行 pip，也不下载运行依赖。

当前只实现新建、结构检查、区域读取，以及将单元格写入新副本。复杂透视表、切片器、VBA 深度编辑和底层 XML 精细修复尚未实现。
