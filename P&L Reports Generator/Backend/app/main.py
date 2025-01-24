import pandas as pd

file_path = "trial_balance.xlsx" # Sample file path


#Load and validate the Excel file data.
def load_and_validate(file_path):
    df = pd.read_excel(file_path)

    df.columns = df.columns.str.strip().str.lower()
    # Drop rows where the "Account" field is empty or a category headers based on known conditions
    category_headers = ["Assets", "Liabilities", "Equity", "Revenue", "Expenses", "Totals"]
    df = df[~df["account"].isin(category_headers)]  # Remove rows with headers

    # Fill missing or NaN values in Debit and Credit columns with 0
    if "debit" in df.columns:
        df["debit"] = df["debit"].fillna(0)
    else:
        print("The 'Debit' column is missing. Please check your file.")

    if "credit" in df.columns:
        df["credit"] = df["credit"].fillna(0)
    else:
        print("The 'Credit' column is missing. Please check your file.")
    
    # Convert to a list of dictionaries
    accounts = df.to_dict(orient='records')
    return accounts


# Mapping of trial balance account types to predefined categories
account_mapping = {
        "Cash": "Assets",
        "Accounts Receivable": "Assets",
        "Inventory": "Assets",
        "Prepaid Insurance": "Assets",
        "Supplies": "Assets",
        "Equipment": "Assets",
        "Accumulated Depreciation - Equipment": "Liabilities",
        "Accounts Payable": "Liabilities",
        "Salaries Payable": "Liabilities",
        "Unearned Revenue": "Liabilities",
        "Notes Payable (Short-term)": "Liabilities",
        "Common Stock": "Equity",  
        "Retained Earnings": "Equity",
        "Dividends": "Equity",    
        "Sales Revenue": "Revenue",
        "Cost of Goods Sold": "Expenses",
        "Salaries Expense": "Expenses",
        "Rent Expense": "Expenses",
        "Utilities Expense": "Expenses",
        "Depreciation Expense": "Expenses",
        "Advertising Expense": "Expenses",
        "Insurance Expense": "Expenses",
        "Supplies Expense": "Expenses",
        "Interest Expense": "Expenses"
    }

 
#initializing account categories to 0 in a dictionary
categories = {
    "Revenue": 0,
    "Expenses": 0,
    "Assets": 0,
    "Liabilities": 0,
    "Equity": 0
}


# Categorize accounts and perform P&L and Balance Sheet calculations
def categorize_and_calculate(categories, accounts):

    #Categorizing accounts
    
    for account in accounts:
        account_type = account["account"]

        debit = account["debit"]
        credit = account["credit"]

        mapped_category = account_mapping.get(account_type)

        if mapped_category in categories:
            amount = debit if debit != 0 else credit
            categories[mapped_category] += amount
        
        else:
            print(f"Unmatched account type: {account_type}")
    
    #Calculate Profit and Loss
    net_profit = categories["Revenue"] - categories["Expenses"]
    # Calculate Balance Sheet with the balance sheet equation (Assets = Liabilities + Equity)

    print(net_profit)
    balance_sheet = {
        "Assets": categories["Assets"],
        "Liabilities": categories["Liabilities"],
        "Equity": categories["Equity"]
    }

    # Check to ensure the balance sheet balances (Assets = Liabilities + Equity)
    if categories["Assets"] == (categories["Liabilities"] + categories["Equity"]):
        print("Successfully balanced!")
    else:
        print("Not balanced!")

    print(categories["Assets"])
    print(categories["Liabilities"])
    print(categories["Equity"])

    #Return results
    return {
        "Profit and Loss": {
            "Revenue": categories["Revenue"],
            "Expenses": categories["Expenses"],
            "Net_profit": net_profit
        },
        "Balance Sheet": balance_sheet
    }


# SAVE THE RESULTS TO EXCEL FILES
def save_files(results):
    # Create DataFrames for P&L and Balance Sheet
    pnl_df = pd.DataFrame([results["Profit and Loss"]])
    balance_sheet_df = pd.DataFrame(["Balance Sheet"])

    print(pnl_df)
    print(balance_sheet_df)

    # Save to Excel files
    pnl_df.to_excel("profit_and_loss.xlsx", index=False)
    balance_sheet_df.to_excel("balance_sheet.xlsx")

# Main execution
accounts = load_and_validate(file_path)
results = categorize_and_calculate(categories, accounts)
save_files(results)
