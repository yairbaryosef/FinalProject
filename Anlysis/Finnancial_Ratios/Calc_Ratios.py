import pandas as pd

from Anlysis.Finnancial_Ratios import Calc_Altman


def calculate_financial_ratios(file_path):
    # קריאת הנתונים מקובץ ה-Excel
    df = pd.read_excel(file_path)

    # חישוב יחסי נזילות
    df['Current Ratio'] = df['Current Assets'] / df['Current Liabilities']
    df['Quick Ratio'] = (df['Current Assets'] - df['Inventory']) / df['Current Liabilities']
    df['Immediate Liquidity'] = (df['Cash And Cash Equivalents'] + df['Other Short Term Investments']) / df['Current Liabilities']
    df['Cash Flow to Sales'] = df['Operating Cash Flow'] / df['Total Revenue']

    # חישוב יחסי רווחיות
    df['Net Profit Margin'] = df['Net Income'] / df['Total Revenue']
    df['Operating Margin'] = df['EBIT'] / df['Total Revenue']
    df['EBITDA Margin'] = df['Normalized EBITDA'] / df['Total Revenue']
    df['ROE'] = df['Net Income'] / df['Common Stock Equity']
    df['ROA'] = df['Net Income'] / df['Total Assets']

    # חישוב יחסי מבנה ההון
    df['Leverage Ratio'] = df['Net Debt'] / (df['Net Debt'] + df['Common Stock Equity'])
    df['Equity to Total Assets'] = df['Common Stock Equity'] / df['Total Assets']

    # חישוב יחסי יעילות תפעולית
    df['Receivables Ratio'] = (df['Accounts Receivable'] + df['Other Receivables']) / df['Total Revenue']
    df['Days Sales Outstanding'] = (df['Accounts Receivable'] / df['Total Revenue']) * 365
    df['Inventory Ratio'] = df['Inventory'] / df['Cost Of Revenue']
    df['Inventory Turnover'] = df['Cost Of Revenue'] / df['Inventory']
    df['Days Inventory'] = (df['Inventory'] / df['Cost Of Revenue']) * 365
    df['Days Payable'] = (df['Accounts Payable'] / df['Cost Of Revenue']) * 365

    # Recalculate Altman Z-Score based on provided formula
    # If a column has missing values, we will replace them with 0 during calculation

    # Prepare the formula for Altman Z-Score calculation while handling missing values (NaNs)

    df['Current Assets'] = df['Current Assets'].fillna(0)
    df['Current Liabilities'] = df['Current Liabilities'].fillna(0)
    df['Total Assets'] = df['Total Assets'].fillna(0)
    df['Retained Earnings'] = df['Retained Earnings'].fillna(0)
    df['EBIT'] = df['EBIT'].fillna(0)
    df['Common Stock Equity'] = df['Common Stock Equity'].fillna(0)
    df['Total Debt'] = df['Total Debt'].fillna(0)
    df['Total Revenue'] = df['Total Revenue'].fillna(0)
    # Initialize empty lists to store results
    scores = []
    statuses = []

    # Loop over all symbols and try to calculate Altman Z-Score
    for symbol in df['Symbol']:
        try:
            score, status = Calc_Altman.CalcAltman(symbol)
            scores.append(score)
            statuses.append(status)
        except Exception as e:
            # Append None or a default value if an error occurs
            scores.append(None)
            statuses.append(None)
            print(f"Error calculating Altman Z-Score for {symbol}: {e}")

    # Add the results to the DataFrame
    df['Altman Z-Score'] = scores
    df['Altman Status'] = statuses

    # Display the Altman Z-Score results for all rows
    df[['Year', 'Altman Z-Score', 'Altman Status']]

    # שמירת התוצאות בקובץ חדש
    output_file = '../../AI/Net income Prediction/financial_ratios_with_altman.xlsx'
    df.to_excel(output_file, index=False)
    print(f"החישובים הושלמו והתוצאות נשמרו בקובץ '{output_file}'")


# שימוש בפונקציה
file_path = 'combined_financial_data_all_stocks.xlsx'  # הקובץ שלך
calculate_financial_ratios(file_path)
