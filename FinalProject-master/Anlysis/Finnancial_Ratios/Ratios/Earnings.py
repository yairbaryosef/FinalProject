def calc_net_profit_margin(income_statement):
    try:
        # שליפת הרווח הנקי וההכנסות הכוללות מהדוח הכספי
        net_income = income_statement.loc['Net Income'].iloc[0]  # Most recent year
        total_revenue = income_statement.loc['Total Revenue'].iloc[0]  # Most recent year

        # חישוב שיעור הרווח הנקי
        net_profit_margin = net_income / total_revenue

        return net_profit_margin
    except Exception as e:
        print(f"Error calculating Net Profit Margin: {e}")
        return 0

def calc_operating_profit_margin(income_statement):
    try:
        # שליפת הרווח התפעולי וההכנסות הכוללות מהדוח הכספי
        operating_income = income_statement.loc['Operating Income'].iloc[0]

        total_revenue = income_statement.loc['Total Revenue'].iloc[0]


        # חישוב שיעור הרווח התפעולי
        operating_profit_margin = operating_income / total_revenue

        return operating_profit_margin
    except Exception as e:
        print(f"Error calculating Operating Profit Margin: {e}")
        return 0

def calc_ebitda_ratio(income_statement):
    # Extract relevant data
    ebitda = income_statement.loc['EBITDA'].iloc[0]  # Most recent year
    total_revenue = income_statement.loc['Total Revenue'].iloc[0]  # Most recent year

    # Calculate EBITDA ratio
    ebitda_ratio = ebitda / total_revenue
    return ebitda_ratio


import numpy as np


def calc_roe(balance_sheet, income_statement):
    # Extract relevant data
    net_income = income_statement.loc['Net Income'].iloc[0]  # Most recent year
    total_equity = 0
    count = 0

    # Loop through Stockholders Equity to calculate total and count, skipping NaN values
    for equity in balance_sheet.loc['Stockholders Equity'].iloc:
        if not np.isnan(equity):  # Check if equity is not NaN
            total_equity += equity
            count += 1

    # Calculate average shareholder's equity
   # average_equity = total_equity / count if count > 0 else 0
    average_equity=balance_sheet.loc['Stockholders Equity'].iloc[0]
    # Calculate ROE
    roe = net_income / average_equity if average_equity != 0 else None  # Avoid division by zero

    return roe


import numpy as np


def calc_roa(balance_sheet, income_statement):
    # Extract relevant data
    operating_income = income_statement.loc['Operating Income'].iloc[0]  # Most recent year
    total_assets = balance_sheet.loc['Total Assets'].iloc[0]



    # Calculate ROA
    roa = operating_income / total_assets # Avoid division by zero

    return roa
