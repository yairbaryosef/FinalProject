import yfinance as yf
import pandas as pd
from colorama import Style, Fore

from Anlysis.Finnancial_Ratios.Ratios.Operations_Efficencey import calc_receivables_ratio, calc_days_sales_outstanding, \
    calc_inventory_ratio, calc_inventory_turnover_ratio, calc_inventory_days, calc_payables_days
from Anlysis.Finnancial_Ratios.Ratios.Structure import calc_leverage_ratio, calc_equity_to_assets_ratio
from Anlysis.Finnancial_Ratios.Ratios.Liquidity import current_ratio,quick_ratio,calc_liquidity_ratio,calc_cashflow_to_sales_ratio
from Anlysis.Finnancial_Ratios.Ratios.Earnings import calc_net_profit_margin, calc_operating_profit_margin, calc_ebitda_ratio, calc_roe, calc_roa

# Define the symbol for the company
symbol = 'AVGL.TA'   # Example: Aura

# Download financial data
stock = yf.Ticker(symbol)

# Historical financial statements


def calc_stock_price(stock,symbol):
    # Download historical data for the last 5 days using the history function
    stock_data = stock.history(period="5d")

    print(symbol)
    # Print the last 5 closing prices
    closing_prices = stock_data['Close']
    print("Last 5 closing prices for:")
    print(closing_prices)

def calc_Ratios(stock,symbol):
    balance_sheet = stock.balance_sheet

    income_statement = stock.financials

    cash_flow = stock.cashflow
    print(balance_sheet)
    #יחסי נזילות
    current_ratio_value = current_ratio(balance_sheet)
    quick_ratio_value = quick_ratio(balance_sheet)
    immediate_liquidity_ratio=calc_liquidity_ratio(balance_sheet)
    cashflow_to_sales_ratio=calc_cashflow_to_sales_ratio(cash_flow,income_statement)
    # Print the
    print("יחסי נזילות:")
    print(f"{Fore.GREEN}Current Ratio (יחס שוטף) for {symbol}: {current_ratio_value:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Quick Ratio (יחס מהיר) for {symbol}: {quick_ratio_value:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Immediate Liquidity Ratio (רמת נזילות מיידית) for {symbol}: {immediate_liquidity_ratio:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Cash Flow to Sales Ratio for {symbol}: {cashflow_to_sales_ratio:.2f}{Style.RESET_ALL}")


    #יחסי רווחיות
    net_profit_margin_value = calc_net_profit_margin(income_statement)
    operating_profit_margin_value = calc_operating_profit_margin(income_statement)
    ebitda_ratio = calc_ebitda_ratio(income_statement)
    roe_value = calc_roe(balance_sheet, income_statement)
    roa_value = calc_roa(balance_sheet, income_statement)
    print("\nיחסי רווחיות")
    print(f"{Fore.GREEN}Net Profit Margin (שיעור הרווח הנקי) for {symbol}: {net_profit_margin_value:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Operating Profit Margin (שיעור הרווח התפעולי) for {symbol}: {operating_profit_margin_value:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}EBITDA to Sales Ratio for {symbol}: {ebitda_ratio:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}ROE for {symbol}: {roe_value:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}ROA for {symbol}: {roa_value:.2f}{Style.RESET_ALL}")
    #יחסי מבני הון
    market_value_of_equity = stock.info['marketCap']
    calc_leverage_ratio_rate=calc_leverage_ratio(balance_sheet,market_value_of_equity)
    calc_equity_to_assets_ratio_rate=calc_equity_to_assets_ratio(balance_sheet)
    print("\nיחסי מבני הון")
    print(f"{Fore.GREEN}Leverage ratio for {symbol}: {calc_leverage_ratio_rate:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Equity to Assets ratio for {symbol}: {calc_equity_to_assets_ratio_rate:.2f}{Style.RESET_ALL}")

    #יחסי יעילות תפעולית
    calc_receivables_ratio_rate=calc_receivables_ratio(balance_sheet,income_statement)
    calc_customers_ratio_rate = calc_days_sales_outstanding(balance_sheet, income_statement)
    inventory_ratio = calc_inventory_ratio(balance_sheet, income_statement)
    inventory_turnover_ratio = calc_inventory_turnover_ratio(balance_sheet, income_statement)
    inventory_days = calc_inventory_days(balance_sheet, income_statement)
    payables_days = calc_payables_days(balance_sheet, income_statement)
    print("\nיחסי יעילות תפעולית")
    print(f"{Fore.GREEN}Receivables ratio to Assets ratio for {symbol}: {calc_receivables_ratio_rate:.2f}{Style.RESET_ALL}")
    print( f"{Fore.GREEN}Customers days ratio ratio to Assets ratio for {symbol}: {calc_customers_ratio_rate:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Inventory ratio ratio to Assets ratio for {symbol}: {inventory_ratio:.2f}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Inventory turnover ratio ratio to Assets ratio for {symbol}: {inventory_turnover_ratio:.2f}{Style.RESET_ALL}")
    print( f"{Fore.GREEN}Inventory days ratio ratio to Assets ratio for {symbol}: {inventory_days:.2f}{Style.RESET_ALL}")
    print( f"{Fore.GREEN}Payables Days ratio ratio to Assets ratio for {symbol}: {payables_days:.2f}{Style.RESET_ALL}")


from colorama import Fore, Style


# פונקציה לעדכון צבעים בהתאם לערכים פיננסיים
def get_color_by_value(value, growth_threshold, stable_threshold):
    if value > growth_threshold:
        return Fore.GREEN  # צמיחה
    elif stable_threshold <= value <= growth_threshold:
        return Fore.YELLOW  # יציבות
    else:
        return Fore.RED  # דעיכה


import pandas as pd
import yfinance as yf
from openpyxl import Workbook


# Step 2: Define the calc_Ratios_with_growth function
def calc_Ratios_with_growth(stock, symbol):
    balance_sheet = stock.balance_sheet
    income_statement = stock.financials
    cash_flow = stock.cashflow

    # Liquidity ratios

    import numpy as np

    def safe_calc(func, *args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception:
            return np.nan

    # Liquidity ratios
    current_ratio_value = safe_calc(current_ratio, balance_sheet)
    quick_ratio_value = safe_calc(quick_ratio, balance_sheet)
    immediate_liquidity_ratio = safe_calc(calc_liquidity_ratio, balance_sheet)
    cashflow_to_sales_ratio = safe_calc(calc_cashflow_to_sales_ratio, cash_flow, income_statement)

    # Profitability ratios
    net_profit_margin_value = safe_calc(calc_net_profit_margin, income_statement)
    operating_profit_margin_value = safe_calc(calc_operating_profit_margin, income_statement)
    ebitda_ratio = safe_calc(calc_ebitda_ratio, income_statement)
    roe_value = safe_calc(calc_roe, balance_sheet, income_statement)
    roa_value = safe_calc(calc_roa, balance_sheet, income_statement)

    # Capital structure ratios
    market_value_of_equity = stock.info.get('marketCap', np.nan)
    leverage_ratio_value = safe_calc(calc_leverage_ratio, balance_sheet, market_value_of_equity)
    equity_to_assets_ratio_value = safe_calc(calc_equity_to_assets_ratio, balance_sheet)

    # Efficiency ratios
    receivables_ratio = safe_calc(calc_receivables_ratio, balance_sheet, income_statement)
    customers_ratio = safe_calc(calc_days_sales_outstanding, balance_sheet, income_statement)
    inventory_ratio = safe_calc(calc_inventory_ratio, balance_sheet, income_statement)
    inventory_turnover_ratio = safe_calc(calc_inventory_turnover_ratio, balance_sheet, income_statement)
    inventory_days = safe_calc(calc_inventory_days, balance_sheet, income_statement)
    payables_days = safe_calc(calc_payables_days, balance_sheet, income_statement)

    # Return all calculated ratios as a dictionary
    return {
        'current_ratio': current_ratio_value,
        'quick_ratio': quick_ratio_value,
        'immediate_liquidity_ratio': immediate_liquidity_ratio,
        'cashflow_to_sales_ratio': cashflow_to_sales_ratio,
        'net_profit_margin': net_profit_margin_value,
        'operating_profit_margin': operating_profit_margin_value,
        'ebitda_ratio': ebitda_ratio,
        'roe': roe_value,
        'roa': roa_value,
        'leverage_ratio': leverage_ratio_value,
        'equity_to_assets_ratio': equity_to_assets_ratio_value,
        'receivables_ratio': receivables_ratio,
        'customers_ratio': customers_ratio,
        'inventory_ratio': inventory_ratio,
        'inventory_turnover_ratio': inventory_turnover_ratio,
        'inventory_days': inventory_days,
        'payables_days': payables_days
    }

if __name__ == '__main__':
   r''' # Step 1: Load the symbols from the Excel file into a list
 file_path = r'C:\Users\yairb\PycharmProjects\Projectduhot\SYMBOLS.xlsx'
 df = pd.read_excel(file_path)
 symbols = df['symb'].tolist()  # Assuming the column is named 'symb'

    # Step 3: Iterate over the symbols and calculate the ratios
 results = []
 for symbol in symbols:
  try:
    stock = yf.Ticker(symbol)
    ratios = calc_Ratios_with_growth(stock, symbol)
    ratios['symb'] = symbol  # Add the symbol to the ratios
    results.append(ratios)
    print(results)
  except:
     pass

# Step 4: Save the results to a new Excel file
 output_df = pd.DataFrame(results)
 output_file_path = r'C:\Users\yairb\PycharmProjects\Projectduhot\financial_ratios.xlsx'
 output_df.to_excel(output_file_path, index=False)

 print(f"Financial ratios saved to {output_file_path}")


'''
   stock = yf.Ticker("EVGN.TA")
   ratios = calc_Ratios_with_growth(stock, symbol)
   print(ratios)





