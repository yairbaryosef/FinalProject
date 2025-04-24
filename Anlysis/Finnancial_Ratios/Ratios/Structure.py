def calc_leverage_ratio(balance_sheet, market_value):
  try:
    # Extract relevant data
    net_financial_liabilities = balance_sheet.loc['Net Debt'].iloc[0]  # התחייבויות פיננסיות נטו
    cash_and_equivalents = balance_sheet.loc['Cash And Cash Equivalents'].iloc[0]  # מזומנים ושווי מזומנים


    # Calculate net financial liabilities
    net_financial_liabilities = net_financial_liabilities - cash_and_equivalents

    # Calculate market capitalization


    # Calculate leverage ratio
    leverage_ratio = net_financial_liabilities / (net_financial_liabilities + market_value) if (
                                                                                                           net_financial_liabilities + market_value) != 0 else None

    return leverage_ratio
  except:
      0

def calc_equity_to_assets_ratio(balance_sheet):
    # Extract relevant data
    total_equity = balance_sheet.loc['Common Stock Equity'].iloc[0]  # הון עצמי
    total_assets = balance_sheet.loc['Total Assets'].iloc[0]  # סך כל המאזן

    # Calculate equity to total assets ratio
    equity_to_assets_ratio = total_equity / total_assets if total_assets != 0 else None

    return equity_to_assets_ratio
