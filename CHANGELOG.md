# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.1] - 2025-01-12

### Fixed
- `run.bat` now auto-installs dependencies if missing
- Better error messages for Python not found
- Added dependency check before launching application

### Changed
- Improved launcher script with clearer user feedback
- Added Python version display on startup

## [1.1.0] - 2025-01-12

### Added
- Portable version without matplotlib dependency to fix DLL compatibility issues
- Pure tkinter Canvas-based charts (bar charts, pie charts)
- `fix_dll_errors.bat` script for resolving Windows DLL issues
- Support for Python 3.10 and 3.11 compatibility

### Changed
- Replaced matplotlib charts with native tkinter Canvas implementation
- Updated numpy to 1.24.3 for better Windows compatibility
- Updated pandas to 2.0.3 for stability
- Simplified requirements.txt with pinned versions

### Fixed
- "ordinal 380 could not be located" DLL error on Windows
- Visual C++ Redistributable compatibility issues
- PyInstaller build failures with matplotlib

## [1.0.0] - 2025-01-12

### Added
- Initial release of Trading Report Builder
- Excel (.xlsx) and CSV file import for Questrade transactions
- FIFO (First-In-First-Out) profit/loss calculation engine
- Multi-tab interface: Raw Transactions, Trades Analysis, Dividends, P&L Summary, Charts
- Stock categorization: TSX Mining (40%), Dividend (10%), Tech (30%), Blue Chip (20%)
- Filtering by date range, category, action, currency, and symbol
- Quick filter reports: Top 10 Gainers, Top 10 Losers, Biggest Trades, Most Active
- Export to Excel with multiple sheets and formatting
- Export to PDF with formatted tables
- Export to HTML with interactive styling
- Print report functionality via browser
- Per-stock and per-category P&L summaries
- Dividend tracking and aggregation
- Monthly summary reports
- Interactive charts: P&L by Category, Top Performers, Monthly Trades, Trade Volume
- Windows executable build support via PyInstaller
- Keyboard shortcuts (Ctrl+O, Ctrl+E, Ctrl+P)

### Technical
- Built with Python 3.10+ and tkinter
- Uses pandas for data processing
- Uses openpyxl for Excel operations
- Uses reportlab for PDF generation
- Uses matplotlib for charting (v1.0.0 only)

## [0.1.0] - 2025-01-12

### Added
- Sample data generator for Questrade transactions
- Generated 6,306 transactions across 258 trading days
- Realistic price movements using geometric Brownian motion
- 52 stocks across 4 categories
- Quarterly dividend payments for dividend stocks

---

## Git Commit Messages

### v1.1.1
```
fix: auto-install dependencies in run.bat launcher

- Add dependency check before launching application
- Auto-run pip install if packages are missing
- Improve error messages for missing Python
- Display Python version on startup

Fixes: ModuleNotFoundError when running from source
```

### v1.1.0
```
feat: add portable version without matplotlib dependency

- Replace matplotlib with pure tkinter Canvas charts
- Add fix_dll_errors.bat for Windows compatibility
- Pin numpy==1.24.3 and pandas==2.0.3 for stability
- Fix "ordinal 380 could not be located" DLL errors
- Add horizontal/vertical bar charts and pie charts
- Update build script with --collect-all flags

Fixes #1: DLL compatibility issues on Windows
```

### v1.0.0
```
feat: initial release of Trading Report Builder

- Add Questrade transaction import (Excel/CSV)
- Implement FIFO P&L calculation engine
- Create multi-tab interface for analysis
- Add stock categorization (Mining, Dividend, Tech, Blue Chip)
- Implement filtering and quick reports
- Add export to Excel, PDF, HTML
- Add print functionality
- Create PyInstaller build configuration

BREAKING CHANGE: First stable release
```

### v0.1.0
```
feat: add sample data generator

- Generate realistic Questrade transaction data
- Create 6,306 transactions across 258 trading days
- Implement price simulation with geometric Brownian motion
- Add 52 stocks across 4 categories
- Include quarterly dividend payments
```

[Unreleased]: https://github.com/hmofet/Financial-Analyst/compare/v1.1.1...HEAD
[1.1.1]: https://github.com/hmofet/Financial-Analyst/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/hmofet/Financial-Analyst/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/hmofet/Financial-Analyst/compare/v0.1.0...v1.0.0
[0.1.0]: https://github.com/hmofet/Financial-Analyst/releases/tag/v0.1.0
