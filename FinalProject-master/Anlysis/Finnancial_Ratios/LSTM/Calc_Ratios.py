from datetime import datetime
import pandas as pd
import numpy as np
import yfinance as yf
import sys
import os

# Add the project root to sys.path
project_root = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "../../..")
)
sys.path.append(project_root)

# Now you can import modules

from Anlysis.Finnancial_Ratios.Ratios.Operations_Efficencey import calc_receivables_ratio, calc_days_sales_outstanding, \
    calc_inventory_ratio, calc_inventory_turnover_ratio, calc_inventory_days, calc_payables_days
from Anlysis.Finnancial_Ratios.Ratios.Structure import calc_leverage_ratio, calc_equity_to_assets_ratio
from Anlysis.Finnancial_Ratios.Ratios.Liquidity import current_ratio,quick_ratio,calc_liquidity_ratio,calc_cashflow_to_sales_ratio
from Anlysis.Finnancial_Ratios.Ratios.Earnings import calc_net_profit_margin, calc_operating_profit_margin, calc_ebitda_ratio, calc_roe, calc_roa

# Asset Turnover Function
def calc_asset_turnover_ratio(balance_sheet, income_statement, current_period, previous_period):
    try:
        revenue = income_statement.loc['Total Revenue'].loc[current_period]
        total_assets_current = balance_sheet.loc['Total Assets'].loc[current_period]
        total_assets_prev = balance_sheet.loc['Total Assets'].loc[previous_period] if previous_period in balance_sheet.columns else np.nan
        avg_assets = (total_assets_current + total_assets_prev) / 2
        return revenue / avg_assets if avg_assets != 0 else np.nan
    except Exception:
        return 1

def calc_days_sales_outstanding(balance_sheet, income_statement):
    try:
        ar = balance_sheet.loc["Receivables"].values[0]
        revenue = income_statement.loc["Total Revenue"].values[0]
        if revenue == 0 or np.isnan(ar) or np.isnan(revenue):
            return np.nan
        return (ar / revenue) * 90
    except Exception as e:
        print(f"Error calculating ratio: {e}")
        return np.nan

def calc_debt_to_equity_ratio(balance_sheet):
    try:
        total_liabilities = balance_sheet.loc['Total Liabilities Net Minority Interest'].iloc[0]
        total_equity = balance_sheet.loc['Common Stock Equity'].iloc[0]
        return total_liabilities / total_equity if total_equity != 0 else None
    except:
        return None

def calc_ebit_margin(income_statement):
    try:
        ebit = income_statement.loc['EBIT'].iloc[0]
        revenue = income_statement.loc['Total Revenue'].iloc[0]
        return ebit / revenue if revenue != 0 else None
    except:
        return None

def calc_gross_margin(income_statement):
    try:
        gross_profit = income_statement.loc['Gross Profit'].iloc[0]
        revenue = income_statement.loc['Total Revenue'].iloc[0]
        return gross_profit / revenue if revenue != 0 else None
    except:
        return None

def calc_long_term_debt_to_capital_ratio(balance_sheet):
    try:
        long_term_debt = balance_sheet.loc['Long Term Debt'].iloc[0]
        total_equity = balance_sheet.loc['Common Stock Equity'].iloc[0]
        capital = long_term_debt + total_equity
        return long_term_debt / capital if capital != 0 else None
    except:
        return None

def calc_pre_tax_profit_margin(income_statement):
    try:
        pre_tax_income = income_statement.loc['Pretax Income'].iloc[0]
        revenue = income_statement.loc['Total Revenue'].iloc[0]
        return pre_tax_income / revenue if revenue != 0 else None
    except:
        return None

def calc_return_on_tangible_equity(balance_sheet, income_statement):
    try:
        net_income = income_statement.loc['Net Income'].iloc[0]
        total_equity = balance_sheet.loc['Common Stock Equity'].iloc[0]
        goodwill = balance_sheet.loc['Goodwill'].iloc[0] if 'Goodwill' in balance_sheet.index else 0
        intangible_assets = balance_sheet.loc['Other Intangible Assets'].iloc[0] if 'Other Intangible Assets' in balance_sheet.index else 0
        tangible_equity = total_equity - goodwill - intangible_assets
        if tangible_equity == 0 or np.isnan(tangible_equity):
            return None
        return net_income / tangible_equity
    except:
        return None
def calc_receivables_ratio(balance_sheet, income_statement):
    try:
        # Try 'Other Receivables' first, fallback to 'Receivables' if available
        if 'Other Receivables' in balance_sheet.index:
            receivables = balance_sheet.loc['Other Receivables'].values[0]
        elif 'Receivables' in balance_sheet.index:
            receivables = balance_sheet.loc['Receivables'].values[0]
        else:
            print("Receivables field not found in balance sheet.")
            return np.nan

        revenue = income_statement.loc["Total Revenue"].values[0]
        if revenue == 0 or np.isnan(receivables) or np.isnan(revenue):
            return np.nan
        return revenue / receivables
    except Exception as e:
        print(f"Error in calc_receivables_ratio: {e}")
        return np.nan

def calc_ratios_for_quarter(stock, current_period, previous_period):
    try:
        bs = stock.quarterly_balance_sheet
        is_ = stock.quarterly_financials
        cf = stock.quarterly_cashflow

        if bs.empty or is_.empty or cf.empty:
            print("returned None due to empty financials")
            return None

        bs_q = bs.loc[:, bs.columns == current_period]
        is_q = is_.loc[:, is_.columns == current_period]
        cf_q = cf.loc[:, cf.columns == current_period]
        print(current_period)
        print(bs_q)
        if bs_q.empty or is_q.empty or cf_q.empty:
            print("returned None due to missing current quarter data")
            return None

        def safe_calc(func, *args):
            try:
                return func(*args)
            except Exception as e:
                print(f"Error in {func.__name__}: {e}")
                return np.nan

        return {
            'Asset Turnover': safe_calc(calc_asset_turnover_ratio, bs, is_, current_period, previous_period),
            'Current Ratio': safe_calc(current_ratio, bs_q),
            'Days Sales In Receivables': safe_calc(calc_days_sales_outstanding, bs_q, is_q),
            'Debt/Equity Ratio': safe_calc(calc_debt_to_equity_ratio, bs_q),
            'EBIT Margin': safe_calc(calc_ebit_margin, is_q),
            'EBITDA Margin': safe_calc(calc_ebitda_ratio, is_q),
            'Gross Margin': safe_calc(calc_gross_margin, is_q),
            'Inventory Turnover Ratio': safe_calc(calc_inventory_turnover_ratio, bs_q, is_q),
            'Long-term Debt / Capital': safe_calc(calc_long_term_debt_to_capital_ratio, bs_q),
            'Net Profit Margin': safe_calc(calc_net_profit_margin, is_q),
            'Operating Margin': safe_calc(calc_operating_profit_margin, is_q),
            'Pre-Tax Profit Margin': safe_calc(calc_pre_tax_profit_margin, is_q),
            'ROA - Return On Assets': safe_calc(calc_roa, bs_q, is_q),
            'ROE - Return On Equity': safe_calc(calc_roe, bs_q, is_q),
            'Receiveable Turnover': safe_calc(calc_receivables_ratio, bs_q, is_q),
            'Return On Tangible Equity': safe_calc(calc_return_on_tangible_equity, bs_q, is_q)
        }

    except Exception as e:
        print(f"Error calculating ratios for {stock.ticker} on {current_period}: {e}")
        return None
