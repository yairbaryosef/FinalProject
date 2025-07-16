

def current_ratio(balance_sheet):
    # Extract current assets (רכוש שוטף) and current liabilities (התחייבות שוטפות)
    current_assets = balance_sheet.loc['Current Assets'].iloc[0]  # Most recent year
    current_liabilities = balance_sheet.loc['Current Liabilities'].iloc[0]  # Most recent year

    # Calculate the current ratio (יחס שוטף)
    current_ratio = current_assets / current_liabilities
    return current_ratio

def quick_ratio(balance_sheet):
    current_assets = balance_sheet.loc['Current Assets'].iloc[0]  # Most recent year
    current_liabilities = balance_sheet.loc['Current Liabilities'].iloc[0]
    inventory = balance_sheet.loc['Inventory'].iloc[0]


    # Calculate Quick Ratio
    quick_ratio = (current_assets - inventory) / current_liabilities
    return quick_ratio

def calc_liquidity_ratio(balance_sheet):
    try:
        # Get the balance sheet data

        # Extract the required values for calculation
        cash_and_equiv = balance_sheet.loc['Cash And Cash Equivalents'].iloc[0]
        marketable_securities = balance_sheet.loc['Other Short Term Investments'].iloc[0]
        current_liabilities = balance_sheet.loc['Current Liabilities'].iloc[0]

        # Calculate the immediate liquidity ratio
        immediate_liquidity_ratio = (cash_and_equiv + marketable_securities) / current_liabilities

        return immediate_liquidity_ratio

    except Exception as e:
        print(f"Error calculating liquidity ratio: {e}")
        return 0


def calc_cashflow_to_sales_ratio(cash_flow, income_statement):
    # Extract relevant data
    operating_cash_flow = cash_flow.loc['Operating Cash Flow'].iloc[0]
    total_revenue = income_statement.loc['Total Revenue'].iloc[0]

    # Calculate Cash Flow to Sales Ratio
    cashflow_to_sales_ratio = operating_cash_flow / total_revenue
    return cashflow_to_sales_ratio