# CUSIP Transform

A lightweight toolkit for transforming CUSIP identifiers in Excel files. Reads a spreadsheet with CUSIPs in column A, extracts the first 3 characters using `LEFT(A,3)`, and writes the result to column B.

Built with Claude as a first project exploring AI-assisted Excel automation.

---

## Project Structure

```
cusip-transform/
├── README.md
├── requirements.txt
├── .gitignore
├── src/
│   ├── cusip_transform_macro.py   # Python macro (primary tool)
│   └── TransformCusip.bas         # Excel VBA macro (alternative)
├── data/
│   └── sample/
│       └── Cusip.xlsx             # Sample input file
├── output/
│   └── Cusip_Transformed.xlsx     # Sample output
└── docs/
    └── USAGE.md                   # Detailed usage guide
```

## Quick Start

### Python (recommended)

```bash
pip install openpyxl

python src/cusip_transform_macro.py  data/sample/Cusip.xlsx  output/result.xlsx
```

### Excel VBA (alternative)

1. Open a blank workbook in Excel
2. Press `Alt+F11` → **File → Import File** → select `src/TransformCusip.bas`
3. Save as `.xlsm`
4. Press `Alt+F8` → Run **TransformCusip**
5. Select your input file when prompted

## How It Works

| Step | Action |
|------|--------|
| 1 | Reads the input `.xlsx` file |
| 2 | Detects all data rows in column A (CUSIP) |
| 3 | Writes `=LEFT(A{row},3)` formula in column B |
| 4 | Saves the output file with live formulas |

The output file retains Excel formulas so column B updates automatically if column A changes.

## Requirements

- Python 3.8+
- `openpyxl` library

## License

MIT
