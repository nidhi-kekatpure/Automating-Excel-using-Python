# Excel Transaction Processor

Python script that applies 10% discount to transaction prices in Excel files and generates bar charts using openpyxl.

## Features

- Reads Excel transaction data from `transactions.xlsx`
- Applies 10% discount to prices (calculates 90% of original price)
- Adds discounted prices as a new column
- Generates bar chart visualization from the processed data
- Saves results to a new Excel file

## Requirements

```bash
pip install openpyxl
```

## Usage

1. Place your `transactions.xlsx` file in the project directory
2. Run the script:

```bash
python finalapp.py
```

3. Check the generated `new_transactions.xlsx` file with discounted prices and chart

## File Structure

- `app.py` - Initial development version with detailed comments
- `finalapp.py` - Clean, production-ready version with function-based approach
- `transactions.xlsx` - Input Excel file (not included)
- `new_transactions.xlsx` - Generated output file

## Input Format

Your Excel file should have:
- Sheet named "Sheet1"
- Column 3: Original prices
- Headers in row 1

The script will add discounted prices in column 4 and a bar chart at cell E2.
