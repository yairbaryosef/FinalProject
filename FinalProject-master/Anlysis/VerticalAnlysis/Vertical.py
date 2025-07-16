import yfinance as yf
import pandas as pd
from fontTools.misc.symfont import green


def filter_income_statement(df):
    # Define the relevant financial statement labels to filter
    relevant_labels = [
        'Total Revenue',
        'Cost Of Revenue',
        'Gross Profit',
        'Selling And Marketing Expense',
        'General And Administrative Expense',
        'Operating Income',
        'Interest Income',
        'Interest Expense',
        'Pretax Income',
        'Tax Provision',
        'Net Income'
    ]
    # Filter DataFrame to include only rows with the relevant labels
    filtered_df = df[df.index.isin(relevant_labels)]
    return filtered_df


def retriveData(stock):
    # Assume stock.income_stmt retrieves the financial statement as a DataFrame
    df = filter_income_statement(stock.income_stmt)
    df = df.fillna('N/A')  # Replace missing values with 'N/A'
    df = df.drop(df.columns[-1], axis=1)  # Drop the last column if unnecessary

    # Generate LaTeX code for the table
    latex_code = r"""
    \begin{table}[ht]
    \centering
    \tiny
    \begin{tabular}{lcccccc}
    \toprule
     & \textbf{2023-12-31} & \textbf{2022-12-31} & \textbf{2021-12-31} & \textbf{2020-12-31} \\
    \midrule
    """

    # Append each row from the DataFrame to the LaTeX table
    for index, row in df.iterrows():
        formatted_row = " & ".join(
            f"{value:,}" if isinstance(value, (int, float)) else str(value) for value in row.values)
        latex_code += f"\\textbf{{{index}}} & {formatted_row} \\\\\n"

    # Close the LaTeX table structure
    latex_code += r"""
    \bottomrule
    \end{tabular}
    \caption{Financial Data from 2023 to 2019}
    \label{tab:financial_data}
    \end{table}
    """

    # Print LaTeX code (for debugging purposes)
    print(latex_code)

    return latex_code

def vertical(stock):
    # Retrieve the financial statement and filter relevant rows
    df = filter_income_statement(stock.income_stmt)

    # Ensure Total Revenue is present as the base for analysis
    if 'Total Revenue' not in df.index:
        raise ValueError("Total Revenue not found in the data for vertical analysis.")

    # Fill missing values with 'N/A' and drop the last column (if unnecessary)
    df = df.fillna('N/A')
    df = df.drop(df.columns[-1], axis=1)

    # Perform vertical analysis: calculate percentage of each item relative to Total Revenue
    try:
        df_percentage = (df / df.loc['Total Revenue']) * 100
    except ZeroDivisionError:
        raise ValueError("Total Revenue is zero for vertical analysis. Cannot compute percentages.")

    # Generate LaTeX code for the table with adjusted column widths
    latex_code = r"""
    \begin{table}[ht]
    \centering
    \tiny
    \begin{tabular}{lcccccc}
    \toprule
     & \textbf{2023-12-31} & \textbf{2022-12-31} & \textbf{2021-12-31} & \textbf{2020-12-31} \\
    \midrule
    """

    # Append each row from the DataFrame to the LaTeX table
    for index, row in df_percentage.iterrows():
        formatted_row = " & ".join(
            f"{value:.2f}\%" if pd.notnull(value) and value != 'N/A' else 'N/A' for value in row.values)
        latex_code += f"\\textbf{{{index}}} & {formatted_row} \\\\\n"

    # Close the LaTeX table structure
    latex_code += r"""
    \bottomrule
    \end{tabular}
    \caption{Vertical Analysis of Financial Data from 2023 to 2019}
    \label{tab:vertical_data}
    \end{table}
    """

    # Print LaTeX code (for debugging purposes)
    print(latex_code)

    return latex_code


if __name__ == '__main__':
    # Instantiate the stock object for AURA.TA
    stock = yf.Ticker('AURA.TA')
    # Call the function and generate LaTeX output
    vertical(stock)
