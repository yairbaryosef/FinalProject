import yfinance

import Anlysis.Finnancial_Ratios.scraperyahoo
from update import *

import yfinance as yf

import yfinance as yf
import math

import yfinance as yf
import math

def get_timeseries_data(ticker: str):
    stock = yf.Ticker(ticker)

    income = stock.financials
    balance = stock.balance_sheet
    cashflow = stock.cashflow

    years = list(income.columns)

    def safe_get(df, row, y):
        try:
            val = df.loc[row, y]
            return float(val) if val is not None and not math.isnan(val) else None
        except:
            return None

    raw_data = {
        'Revenue': [safe_get(income, 'Total Revenue', y) for y in years],
        'Net Income': [safe_get(income, 'Net Income', y) for y in years],
        'Operating Income': [safe_get(income, 'Operating Income', y) for y in years],
        'Total Assets': [safe_get(balance, 'Total Assets', y) for y in years],
        'Total Liabilities': [safe_get(balance, 'Current Liabilities', y) for y in years],
        'Free Cash Flow': [safe_get(cashflow, 'Free Cash Flow', y) for y in years]
    }

    # המרה למיליונים
    for key in raw_data:
        raw_data[key] = [round(v / 1e6, 2) if v is not None else None for v in raw_data[key]]

    # סינון שנים שבהן כל הערכים הם None
    filtered_years = []
    filtered_data = {key: [] for key in raw_data}

    for i, y in enumerate(years):
        column_values = [raw_data[key][i] for key in raw_data]
        if any(v is not None for v in column_values):
            filtered_years.append(str(y.year) if hasattr(y, "year") else str(y))
            for key in raw_data:
                filtered_data[key].append(raw_data[key][i])
    filtered_years = filtered_years[::-1]
    return filtered_data, filtered_years

def Present():
    lines = [
        "Company: DemoTech Inc.",
        "Ticker: DMTC",
        "Industry: Artificial Intelligence",
        "Market Cap: $120 Billion",

        "Key Business Segments:",
        "• Develops AI chips and infrastructure",
        "• Provides enterprise-grade ML solutions",
        "• Partners with cloud providers for global expansion"
    ]
    stock = yfinance.Ticker('AAPL')
    new_ratios = Anlysis.Finnancial_Ratios.scraperyahoo.calc_Ratios_with_growth(stock, "AAPL")
    print(new_ratios)
    timeseries_data, years = get_timeseries_data("AAPL")

    sentiment = {
            "Positive": 45,
            "Neutral": 35,
            "Negative": 20
        }
    section_data_positive={
            "Strengths": [
                "Strong brand reputation and customer loyalty",
                "High R&D investment driving innovation",
                "Diversified product portfolio"
            ],
            "Opportunities": [
                "Expansion into emerging markets",
                "Growing demand for AI and automation",
                "Strategic partnerships and acquisitions"
            ]
        }
    section_data_negative={
            "Weaknesses": [
                "Dependence on a limited number of suppliers",
                "High production costs in core segments",
                "Limited presence in specific geographic areas"
            ],
            "Threats": [
                "Intensifying competition in global markets",
                "Regulatory pressures and compliance risks",
                "Economic downturn affecting consumer spending"
            ]
        }
    forecast = {
            "3 Days": 0.8,
            "1 Month": 2.4,
            "3 Months": 4,
            "1 Year": 5.7
        }
    recommendation = 'Buy'
    reasons=[
            "Strong financial ratios and profitability margins",
            "Positive sentiment dominates the market tone",
            "Consistent CAGR growth across major KPIs",
            "1-Year forecast shows strong potential upside",
            "SWOT analysis indicates strengths outweigh risks"
        ]
    Create(lines,new_ratios, timeseries_data,years, sentiment, section_data_positive,section_data_negative, forecast, recommendation, reasons)
Present()