import os
import json
import re
import math
import yfinance as yf

from Anlysis.Finnancial_Ratios.scraperyahoo import calc_Ratios_with_growth
from Presentation.UpdatePre.design.update import Create


# --------- פונקציה לשליפת נתוני זמן מ-YFinance ---------
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



# --------- פונקציה לחילוץ sentiment ו-SWOT ---------
def extract_data_from_files(sentiment_path, swot_path):
    with open(sentiment_path, "r", encoding="utf-8") as f:
        sentiment_data = json.load(f)


    sentiment_counter = {"Positive": 0, "Neutral": 0, "Negative": 0}

    for item in sentiment_data:
        sentiment = item.get("predicted_sentiment", "").capitalize()
        if sentiment in sentiment_counter:
            sentiment_counter[sentiment] += 1

    with open(swot_path, "r", encoding="utf-8") as f:
        swot_text = f.read()

    import re

    def extract_section(text, header):
        # Step 1: Try ## Header ## markdown-style
        pattern_md = rf"##\s*{re.escape(header)}\s*##[\n\r]+((?:\d+\.\s+.+[\n\r]*)+)"
        match = re.search(pattern_md, text, flags=re.IGNORECASE)

        if not match:
            # Step 2: Fallback to "header are:" style
            pattern_fallback = rf"{header.lower()}s? are:\s*((?:\d+\.\s+.+\n?)*)"
            match = re.search(pattern_fallback, text.lower(), flags=re.IGNORECASE)

        if not match:
            return []

        lines = match.group(1).strip().splitlines()
        return [re.sub(r"^\d+\.\s*", "", line.strip()) for line in lines if line.strip()]

    section_data_positive = {
        "Strengths": extract_section(swot_text, "Strengths"),
        "Opportunities": extract_section(swot_text, "Opportunities")
    }

    section_data_negative = {
        "Weaknesses": extract_section(swot_text, "Weaknesses"),
        "Threats": extract_section(swot_text, "Threats")
    }

    return sentiment_counter, section_data_positive, section_data_negative


# --------- פונקציה לחילוץ lines מתוך קובץ טקסט ---------
def extract_lines_dict_from_textfile(filepath, symbol, market_cap):
    with open(filepath, 'r', encoding='utf-8') as f:
        text = f.read().split("Summery About The Company: <summery about the company>")[1]


    company_match = re.search(r"Company:\s*(.*)", text)

    ticker_match = symbol
    industry_match = re.search(r"Industry:\s*(.*)", text)
    market_cap_match = market_cap
    segments_match = re.search(r"Key Business Segments:\s*((?:•\s?.*\n?)*)", text)
    summary_match = re.search(r"Summery About The Company:\s*(.*)", text)
    print(summary_match)
    segments = []
    if segments_match:
        segments_raw = segments_match.group(1).strip().splitlines()
        segments = [seg.strip() for seg in segments_raw if seg.strip().startswith("•")]

    lines = {
        "Company": company_match.group(1).strip() if company_match else None,
        "Ticker": symbol,
        "Industry": industry_match.group(1).strip() if industry_match else None,
        "Market Cap": market_cap

    }
    '''  "Key Business Segments": segments,
           "Summery Of The Company": summary_match.group(1).strip() if summary_match else ""
           '''

    return lines

def format_market_cap(market_cap):
    try:
        cap = float(market_cap)
        if cap >= 1e12:
            return f"{cap / 1e12:.2f} T$"
        elif cap >= 1e9:
            return f"{cap / 1e9:.2f} B$"
        elif cap >= 1e6:
            return f"{cap / 1e6:.2f} M$"
        elif cap >= 1e3:
            return f"{cap / 1e3:.2f} K$"
        else:
            return f"{cap:.2f}$"
    except (ValueError, TypeError):
        return "N/A"



# --------- פונקציית Present ---------
def Present(symbol,sentiment_file, swot_output, llm_lines_path,output):

    stock = yf.Ticker(symbol)
    raw_cap = stock.info.get('marketCap')  # או stock.fast_info["marketCap"] אם זה זמין
    market_cap_str = format_market_cap(raw_cap)
    print()
    print(market_cap_str)  # למשל: "23.45B$"
    lines = extract_lines_dict_from_textfile(llm_lines_path, symbol, market_cap_str)
    print(lines)





    new_ratios = calc_Ratios_with_growth(stock, symbol)
    timeseries_data, years = get_timeseries_data(symbol)

    sentiment, section_data_positive, section_data_negative = extract_data_from_files(sentiment_file, swot_output)


    forecast = {
        "3 Days": 0.8,
        "1 Month": 2.4,
        "3 Months": 4.0,
        "1 Year": 5.7
    }

    current_dir = os.path.dirname(__file__)
    base_dir = os.path.abspath(os.path.join(current_dir, "..", "..", ".."))  # תלוי בעומק
    results_path = os.path.join(base_dir, "output", "results.txt")
    print(results_path)

    with open(results_path, "r", encoding="utf-8") as f:
        result_text = f.read()
        result_text = result_text.split('\n')[:-1]
        reasons = result_text

    print(result_text)
    test_rec = result_text[0].split(': ')[1][:-1]

    if float(test_rec) > 0:
        recommendation = 'Buy'
    else:
        recommendation = 'Sell'



    Create(
        lines,
        new_ratios,
        timeseries_data,
        years,
        sentiment,
        section_data_positive,
        section_data_negative,
        forecast,
        recommendation,
        reasons,output
    )


# --------- הרצה ---------
if __name__ == "__main__":



    '''submit_lstm_job(hostname='slurm.bgu.ac.il', port=22, username='yairbary', password='yairYAIR0_0',
                    model='LSTM', symbol='TSLA')'''
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    print(base_dir)
    files_dir = os.path.join(base_dir, "Files")
    swot_output = os.path.join(files_dir, "swot_output.txt")

    sentiment_file = os.path.join(files_dir, "sentiment_results.json")
    llm_lines_path = os.path.join(files_dir, "llm_lines.txt")

    Present(sentiment_file, swot_output, llm_lines_path)
