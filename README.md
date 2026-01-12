# Trading Report Builder - Portable Version

A Windows desktop application for analyzing Questrade trading transactions.

**This version uses pure Tkinter for charts (no matplotlib) to avoid DLL compatibility issues.**

## Quick Start

### Option 1: Run Directly (Requires Python)

```batch
pip install -r requirements.txt
python trading_report_builder.py
```

### Option 2: Build Executable

```batch
build_windows.bat
```

The executable will be at `dist\TradingReportBuilder.exe`

## Fixing "ordinal 380 could not be located" Error

This error typically occurs due to numpy/DLL version mismatches. Try these solutions:

### Solution 1: Install Visual C++ Redistributable

Download and install from:
https://aka.ms/vs/17/release/vc_redist.x64.exe

### Solution 2: Run the Fix Script

```batch
fix_dll_errors.bat
```

### Solution 3: Use a Clean Virtual Environment

```batch
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python trading_report_builder.py
```

### Solution 4: Use Python 3.10 or 3.11

Python 3.12+ can have compatibility issues. Use Python 3.10 or 3.11 instead.

## Features

- **Import**: Excel (.xlsx) and CSV files
- **FIFO P&L**: Automatic profit/loss calculation
- **Filtering**: Date, category, action, currency, symbol
- **Charts**: Built-in tkinter canvas charts (no matplotlib needed)
- **Export**: Excel, PDF, HTML
- **Print**: Browser-based printing

## Stock Categories

| Category | Symbols |
|----------|---------|
| TSX Mining | ABX.TO, CCO.TO, TECK-B.TO, NTR.TO, FM.TO, FNV.TO, etc. |
| Dividend | ENB.TO, SU.TO, BCE.TO, JNJ, ABBV, PFE, KO, PG, etc. |
| Tech | AAPL, MSFT, NVDA, GOOGL, META, AMZN, TSLA, AMD, etc. |
| Blue Chip | JPM, WMT, V, UNH, LLY, MRK, BMY, CAT, HD, MA, etc. |

## Dependencies

- Python 3.10 or 3.11 (recommended)
- pandas 2.0.3
- numpy 1.24.3
- openpyxl 3.1.2
- reportlab 4.0.4 (for PDF export)

## Troubleshooting

| Error | Solution |
|-------|----------|
| ordinal 380 not found | Run fix_dll_errors.bat or install VC++ Redistributable |
| ModuleNotFoundError | Run `pip install -r requirements.txt` |
| PDF export fails | Run `pip install reportlab` |
| Executable crashes | Build in a fresh virtual environment |

---

**Version 1.1 Portable** | No matplotlib - maximum compatibility
