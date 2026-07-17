# Finance Management System

A simple personal finance tracker that collects income sources, budget categories, and actual expenses.

## Setup

1. Create a virtual environment:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
2. Install dependencies:
   ```powershell
   pip install -r "Finance management system\requirements.txt"
   ```

## Usage

Run the tracker from the project folder:
```powershell
python "Finance management system\main.py"
```

Follow the prompts to enter income, budget categories, and expenses.

## Excel Export

If `pandas` and `openpyxl` are installed, the script can export a summary to `Finance management system\finance_report.xlsx`.
