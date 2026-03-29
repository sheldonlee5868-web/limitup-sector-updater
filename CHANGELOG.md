# Changelog

## v0.1.0
- Initial public release
- Added project screenshots
- Improved repository documentation



- ## v0.1.1 - 2026-03-29

  ### Fixed
  - Fixed the issue where empty limit-up pool results could be incorrectly written as 0
  - Fixed the hidden risk where historical rebuilds could fail while still appearing “normal” in the output
  - Fixed the refresh workflow when Sheet1 and Dashboard were out of sync

  ### Added
  - Added three modes: `update`, `rebuild`, and `dashboard`
  - Added `--proxy-mode auto/direct` for better proxy/network compatibility
  - Added automatic log file generation and run summary output
  - Added automatic Dashboard rebuild, including line charts, bar chart, and 60-day sector heatmap

  ### Improved
  - Improved index history fetching by trying Eastmoney direct API first and falling back to akshare if needed
  - Improved abnormal date column detection and ignore logic
  - Improved usability under Windows + PowerShell / Git Bash
  - Improved visibility and traceability for historical rebuild failures

- ## v0.1.1 - 2026-03-29

  ### Fixed
  - 修复涨停股池返回空数据时可能被误写为 0 的问题
  - 修复历史区间重算时“抓取失败但结果看似正常”的隐性风险
  - 修复 Sheet1 与 Dashboard 不一致时的刷新链路问题

  ### Added
  - 新增 `update / rebuild / dashboard` 三种运行模式
  - 新增 `--proxy-mode auto/direct` 参数，增强代理环境兼容性
  - 新增自动日志文件与运行摘要输出
  - 新增 Dashboard 自动重建能力，包括折线图、柱状图和最近 60 日热力表

  ### Improved
  - 优化指数历史数据抓取逻辑，优先尝试东财直连接口，失败后回退到 akshare
  - 优化异常日期列识别与忽略逻辑
  - 提升 Windows + PowerShell / Git Bash 环境下的可用性
  - 提高历史重算失败时的可见性和可追踪性