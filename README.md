# Automated Profit & Loss (P&L) and Balance Sheet Tool

This project automates the generation of Profit & Loss statements and Balance Sheets using trial balance data from an Excel file as input. It categorizes accounts, performs necessary calculations, and validates the balance sheet equation.

## Features

- **Data Validation:** Cleans and validates the input Excel data.
- **Account Categorization:** Maps trial balance accounts to predefined categories (Revenue, Expenses, Assets, Liabilities, Equity).
- **Profit & Loss Calculation:** Automatically calculates Net Profit.
- **Balance Sheet Verification:** Validates that Assets = Liabilities + Equity.
- **Excel Export:** Saves the results as separate Excel files for Profit & Loss and Balance Sheet.

## Getting Started

### Prerequisites

- Python 3.8 or later
- Required libraries:
  - `pandas`

### Input Data Requirements

The input data must be an Excel file (.xlsx) with the following structure:

- **Columns**: Account, Debit, Credit  
- **Category Headers**: Rows like Assets, Liabilities, etc., should be excluded or identified as headers to be ignored.

### How to Use 

1. Place your trial balance Excel file in the same directory as the script.  
2. Update the `file_path` variable with the path to your file (e.g., `trial_balance.xlsx`).  
3. Run the script.  
4. The results will be saved as Excel files in the same directory.

### Contribution

Contributions, issues, and feature requests are welcome!  
Feel free to fork the repository and submit pull requests.


