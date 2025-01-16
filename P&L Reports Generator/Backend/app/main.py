import pandas as pd
import openpyxl

file_path = "trial_balance.xlsx"

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

        amount = debit if debit != 0 else credit
        
        if account_type in categories:
            categories[account_type] += amount
    
    #Calculate Profit and Loss
    net_profit = categories["Revenue"] - categories["Expenses"]
    # Calculate Balance Sheet with the balance sheet equation (Assets = Liabilities + Equity)
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

    #Return results
    return {
        "Profit and Loss": {
            "Revenue": categories["Revenue"],
            "Expenses": categories["Expenses"],
            "Net_profit": net_profit
        },
        "Balance Sheet": balance_sheet
    }

accounts = load_and_validate(file_path)
categorize_and_calculate(categories, accounts)

