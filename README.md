# LimitUp Sector Updater

A Windows-friendly Python tool for tracking daily limit-up counts across 28 A-share industry sectors, along with combined turnover of the Shanghai Composite Index and Shenzhen Component Index.

这是一个适用于 Windows 10 本地运行的 Python 工具，用于：
- 基于本地行业成分股映射表，统计 28 个行业板块的每日涨停数量
- 识别并展示二板、三板及以上连板股名称
- 统计上证指数与深证成指总成交额，以及相较前一交易日的变化
- 输出 Excel 结果，并自动生成 Dashboard 图表
- 支持增量更新与区间重算

---

## Features

- Read the latest `TDX_Industry_merged.xlsx` as the sector constituent source
- Track daily limit-up counts for 28 sectors
- Exclude `ST`, `*ST`, and related special-treatment stocks
- Show consecutive limit-up stock names for 2-board, 3-board, and higher
- Calculate combined turnover of:
  - Shanghai Composite Index (`000001`)
  - Shenzhen Component Index (`399001`)
- Generate Excel output and dashboard charts
- Support:
  - incremental update
  - rebuild for a date range
  - dashboard-only refresh
- Friendly for Windows users with `.bat` launcher

---

## Project Structure

```text
limitup-sector-updater/
├─ README.md
├─ LICENSE
├─ .gitignore
├─ requirements.txt
├─ limitup_sector_updater.py
├─ run_limitup_update.bat
├─ docs/
│  ├─ screenshot-main.png
│  └─ screenshot-dashboard.png
└─ examples/
   └─ config_example.txt
```

## Requirements

- Python 3.11+
- Windows 10 / Windows 11 recommended
- Internet access for fetching public market data
- Recommended packages are listed in `requirements.txt`

> Note:
>  This project depends on public market data interfaces. Temporary network instability or upstream interface changes may affect data fetching.

------

## Installation

Clone or download this repository, then install dependencies:

```
pip install -U -r requirements.txt
```

------

## Input Files

This tool expects the following input files from the user locally:

1. `TDX_Industry_merged.xlsx`
    Contains the latest 28-sector constituent stock mapping.
2. A statistics workbook template, for example:
    `2026-01 涨停板块统计（涨停数大于3标红加粗）.xlsx`

These files are **not included** in this repository because they may contain user-specific or continuously updated data.

------

## Usage

### 1. Incremental Update

```
python limitup_sector_updater.py --proxy-mode direct --industry TDX_Industry_merged.xlsx --stats "2026-01 涨停板块统计（涨停数大于3标红加粗）.xlsx" --output "涨停板块统计_自动更新.xlsx"
```

### 2. Rebuild for a Date Range

```
python limitup_sector_updater.py --mode rebuild --start-date 2026-02-01 --end-date 2026-02-28 --proxy-mode direct --industry TDX_Industry_merged.xlsx --stats "2026-01 涨停板块统计（涨停数大于3标红加粗）.xlsx" --output "涨停板块统计_重算.xlsx"
```

### 3. Refresh Dashboard Only

```
python limitup_sector_updater.py --mode dashboard --industry TDX_Industry_merged.xlsx --stats "涨停板块统计_自动更新.xlsx" --output "涨停板块统计_自动更新.xlsx"
```

### 4. Use the Windows Batch Launcher

Double-click:

```
run_limitup_update.bat
```

Or modify it to match your local file paths.

------

## Output

The tool generates an Excel workbook that includes:

- Daily sector-level limit-up counts
- Consecutive limit-up stock names
- Combined turnover of the Shanghai Composite Index and Shenzhen Component Index
- Day-over-day turnover change
- A `Dashboard` worksheet with charts

------

## Screenshots

> Add your screenshots into the `docs/` folder, then keep the links below.

### Main Excel Output







### Dashboard







------

## Notes

- This repository publishes the **tooling code only**
- User data files, generated Excel workbooks, logs, and private datasets should not be committed
- Some public interfaces may occasionally return SSL / connection / retry errors
- The script includes retry and fallback logic, but upstream instability may still occur

------

## Roadmap

-  Incremental update mode
-  Rebuild mode
-  Dashboard generation
-  Windows batch launcher
-  Retry and log output
-  Configurable sector alias matching
-  Better fallback sources for public market data
-  Optional command-line config file
-  Optional packaging for PyPI or Docker

------

## Example Maintenance Issue

You can create issues such as:

- Add configurable sector name aliases
- Improve retry strategy for unstable data source
- Add command-line config file support

This helps show the project is actively maintained.