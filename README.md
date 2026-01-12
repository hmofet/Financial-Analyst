# Financial Analyst - Trading Report Builder

A Windows desktop application for analyzing Questrade trading transactions with FIFO P&L calculations, charts, and comprehensive reporting.

![Version](https://img.shields.io/badge/version-1.1.1-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-green)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

### Data Import
- **Excel Import** (.xlsx, .xls) - Import Questrade activity exports
- **CSV Import** (.csv) - Import CSV transaction files
- Automatic data parsing and validation
- Support for both CAD and USD transactions

### Analysis & Reports
- **FIFO P&L Calculation** - Automatic First-In-First-Out profit/loss calculation
- **Per-Stock Summary** - Revenue, cost basis, realized gains, dividends
- **Category Analysis** - TSX Mining, Tech, Dividend, Blue Chip stocks
- **Quick Filters**:
  - Top 10 Gainers
  - Top 10 Losers
  - Biggest Trades
  - Most Active Stocks
  - Monthly Summary

### Charts & Visualization
- P&L by Category
- Top 10 Performers
- Top 10 Losers
- Dividend Distribution
- Trades by Category

### Export Options
- **Excel** (.xlsx) - Multi-sheet workbook with formatting
- **PDF** - Formatted report with tables
- **HTML** - Interactive web report
- **Print** - Browser-based printing

## Installation

### Option 1: Quick Start (Recommended)

1. Install Python 3.10 or 3.11 from [python.org](https://python.org/downloads/)
   - **Important:** Check "Add Python to PATH" during installation

2. Double-click `run.bat`
   - Dependencies will be installed automatically on first run
   - Application will launch after setup completes

### Option 2: Run from Source (Manual)

1. Install Python 3.10 or 3.11 (recommended)
   ```
   https://python.org/downloads/
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python trading_report_builder.py
   ```

### Option 3: Build Standalone Executable

1. Run the build script:
   ```batch
   build_windows.bat
   ```

2. Find the executable at `dist\TradingReportBuilder.exe`

3. The .exe runs without Python installed

## Usage

### Importing Data

1. Click **File → Import Excel** or **File → Import CSV**
2. Select your Questrade activity export file
3. Data will be automatically parsed and analyzed

### Expected File Format

The application expects Questrade activity export format:

| Column | Description |
|--------|-------------|
| Transaction Date | Date/time of transaction |
| Settlement Date | Settlement date |
| Action | Buy, Sell, or DIV |
| Symbol | Stock ticker symbol |
| Description | Transaction description |
| Quantity | Number of shares |
| Price | Price per share |
| Gross Amount | Total before commission |
| Commission | Trading commission |
| Net Amount | Total after commission |
| Currency | CAD or USD |
| Account # | Account number |
| Activity Type | Trades or Dividends |
| Account Type | Account type (e.g., TFSA) |

### Stock Categories

| Category | Allocation | Symbols |
|----------|------------|---------|
| TSX Mining | 40% | ABX.TO, CCO.TO, TECK-B.TO, NTR.TO, FM.TO, FNV.TO, AGI.TO, AEM.TO, K.TO, WPM.TO, LUN.TO, IVN.TO, NXE.TO, CS.TO, B2GOLD.TO |
| Dividend | 10% | ENB.TO, SU.TO, BCE.TO, JNJ, ABBV, PFE, KO, PG, T.TO, BNS.TO |
| Tech | 30% | AAPL, MSFT, NVDA, GOOGL, META, AMZN, TSLA, AMD, CRM, SHOP.TO, ADBE, INTC, CSCO, ORCL |
| Blue Chip | 20% | JPM, WMT, V, UNH, LLY, MRK, BMY, CAT, HD, MA, DIS, XOM, CVX |

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O | Import Excel file |
| Ctrl+E | Export to Excel |
| Ctrl+P | Print Report |

## Troubleshooting

### "ordinal 380 could not be located" Error

This DLL error is caused by numpy/Windows library mismatches. Solutions:

1. **Install Visual C++ Redistributable**
   ```
   https://aka.ms/vs/17/release/vc_redist.x64.exe
   ```

2. **Run the fix script**
   ```batch
   fix_dll_errors.bat
   ```

3. **Use Python 3.10 or 3.11** (not 3.12+)

4. **Create fresh virtual environment**
   ```batch
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

### Other Issues

| Issue | Solution |
|-------|----------|
| ModuleNotFoundError | Run `pip install -r requirements.txt` |
| PDF export fails | Run `pip install reportlab` |
| Excel import fails | Run `pip install openpyxl` |
| Build fails | Use fresh virtual environment |

## Sample Data

The `sample_data/` folder contains a generated Questrade transaction file with:
- 6,306 transactions
- 258 trading days (2025)
- 52 stocks across 4 categories
- Average 25 trades per day
- Quarterly dividend payments

## Project Structure

```
financial-analyst/
├── trading_report_builder.py   # Main application
├── requirements.txt            # Python dependencies
├── build_windows.bat          # Build script for .exe
├── fix_dll_errors.bat         # DLL error fix script
├── run.bat                    # Quick run script
├── CHANGELOG.md               # Version history
├── README.md                  # This file
└── sample_data/
    └── Questrade_Activities_2025.xlsx
```

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pandas | 2.0.3 | Data processing |
| numpy | 1.24.3 | Numerical operations |
| openpyxl | 3.1.2 | Excel file operations |
| reportlab | 4.0.4 | PDF generation |
| pyinstaller | 6.3.0 | Executable building |

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'feat: add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

MIT License - See LICENSE file for details.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history.

---

**Version 1.1.0** | Portable Edition (No matplotlib dependency)
