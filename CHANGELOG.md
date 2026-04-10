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

## v0.1.2 - 2026-04-10

### Improved
- Enhanced the Dashboard heatmap to make sector activity more vivid and easier to read
- Added clearer heat-level color segmentation for recent 60-day sector limit-up counts
- Added a heatmap legend for faster interpretation
- Added daily totals and sector summary indicators to improve readability
- Added a latest trading day sector ranking area
- Improved the overall Dashboard layout and presentation quality

### Compatibility
- No changes to command-line usage
- No changes to `update`, `rebuild`, or `dashboard` mode invocation
- Existing Excel output workflow remains compatible

## v0.1.2 - 2026-04-10

### 优化
- 优化 Dashboard 热力图展示效果，使板块热度分布更生动、更易读
- 为最近 60 日板块涨停数增加更清晰的热度分级配色
- 新增热力图图例，便于快速理解颜色含义
- 新增单日合计与板块汇总指标，提升信息可读性
- 新增最新交易日板块排行区域
- 整体优化 Dashboard 的布局与展示效果

### 兼容性
- 不改变命令行使用方式
- 不影响 `update`、`rebuild`、`dashboard` 三种模式的调用方式
- 兼容现有 Excel 输出流程

