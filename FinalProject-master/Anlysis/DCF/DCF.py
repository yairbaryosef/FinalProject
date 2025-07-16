import numpy as np
import matplotlib.pyplot as plt

# Assumptions for the DCF calculation
initial_revenue = 1000e6  # Starting revenue in dollars
growth_rate = 0.01 # 5% annual revenue growth rate
discount_rate = 0.1 # 10% discount rate (WACC)
terminal_growth_rate = 0.02  # 2% terminal growth rate
projection_years = 5  # Projecting for 5 years

# Assumption on free cash flow margin
fcf_margin = 0.15  # 15% of revenue

# Projected Free Cash Flows (FCF)
fcfs = []
for year in range(projection_years):
    revenue = initial_revenue * (1 + growth_rate) ** year
    fcf = revenue * fcf_margin
    fcfs.append(fcf)

# Terminal Value Calculation
last_year_fcf = fcfs[-1] * (1 + terminal_growth_rate)
terminal_value = last_year_fcf / (discount_rate - terminal_growth_rate)

# Discounting Free Cash Flows to Present Value
discounted_fcfs = [fcf / ((1 + discount_rate) ** (year + 1)) for year, fcf in enumerate(fcfs)]
discounted_terminal_value = terminal_value / ((1 + discount_rate) ** projection_years)

# Calculate DCF Value (Enterprise Value)
dcf_value = sum(discounted_fcfs) + discounted_terminal_value

# Market Value of Assets (Net Asset Value, using assumed asset data)
total_assets = 1500e6  # Total assets in dollars
total_liabilities = 800e6  # Total liabilities in dollars
market_value_of_assets = total_assets - total_liabilities

# Market Cap for comparison (assuming from external source or financial data)
market_cap = 1200e6  # Hypothetical market capitalization in dollars

# Plotting Results
labels = ['DCF Value', 'Market Value of Assets', 'Market Cap']
values = [dcf_value, market_value_of_assets, market_cap]

plt.figure(figsize=(10, 6))
plt.bar(labels, values, color=['blue', 'green', 'red'])
plt.title('Comparison of DCF Value, Market Value of Assets, and Market Cap for Aura')
plt.ylabel('Value (in millions)')
plt.xlabel('Valuation Metric')
plt.savefig("valuation_comparison.png")
plt.show()
