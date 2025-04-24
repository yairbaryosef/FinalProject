def calc_receivables_ratio(balance_sheet, income_statement, vat_rate=0.17):
    # Extract relevant data
    receivables = balance_sheet.loc['Accounts Receivable'].iloc[0]  # חייבים
    short_term_loans = balance_sheet.loc['Other Receivables'].iloc[0]  # שטל"ק (שטרי לקוחות קצרים)
    sales = income_statement.loc['Total Revenue'].iloc[0]  # מכירות

    # Calculate total receivables (subtract VAT)
    total_receivables = receivables + short_term_loans
    receivables_without_vat = total_receivables / (1 + vat_rate)

    # Calculate receivables ratio
    receivables_ratio = receivables_without_vat / sales if sales != 0 else None

    return receivables_ratio

def calc_days_sales_outstanding(balance_sheet, income_statement, vat_rate=0.17):
    # Extract relevant data
    beginning_receivables = balance_sheet.loc['Accounts Receivable'].iloc[1]  # חייבים בתחילת התקופה (2022)
    ending_receivables = balance_sheet.loc['Accounts Receivable'].iloc[0]  # חייבים בסוף התקופה (2023)
    sales = income_statement.loc['Total Revenue'].iloc[0]  # מכירות

    # Calculate average receivables (subtract VAT)
    avg_receivables = (beginning_receivables + ending_receivables) / 2
    avg_receivables_without_vat = avg_receivables / (1 + vat_rate)

    # Calculate days sales outstanding (DSO)
    dso = (avg_receivables_without_vat * 365) / sales if sales != 0 else None

    return dso

def calc_inventory_ratio(balance_sheet, income_statement):
    # Extract relevant data
    inventory = balance_sheet.loc['Inventory'].iloc[0]  # מלאי בסוף התקופה (2023)
    cost_of_goods_sold = income_statement.loc['Cost Of Revenue'].iloc[0]  # עלות המכר

    # Calculate inventory ratio
    inventory_ratio = inventory / cost_of_goods_sold if cost_of_goods_sold != 0 else None

    return inventory_ratio

def calc_inventory_turnover_ratio(balance_sheet, income_statement):
    # Extract relevant data
    inventory = balance_sheet.loc['Inventory'].iloc[0]  # מלאי בסוף התקופה (2023)
    cost_of_goods_sold = income_statement.loc['Cost Of Revenue'].iloc[0]  # עלות המכר

    # Calculate inventory turnover ratio
    inventory_turnover_ratio = cost_of_goods_sold / inventory if inventory != 0 else None

    return inventory_turnover_ratio
def calc_inventory_days(balance_sheet, income_statement):
    # Extract relevant data
    beginning_inventory = balance_sheet.loc['Inventory'].iloc[1]  # מלאי בתחילת התקופה (2022)
    ending_inventory = balance_sheet.loc['Inventory'].iloc[0]  # מלאי בסוף התקופה (2023)
    cost_of_goods_sold = income_statement.loc['Cost Of Revenue'].iloc[0]  # עלות המכר

    # Calculate the average inventory
    avg_inventory = (beginning_inventory + ending_inventory) / 2

    # Calculate inventory days
    inventory_days = (avg_inventory *365) / cost_of_goods_sold if cost_of_goods_sold != 0 else None

    return inventory_days

def calc_payables_days(balance_sheet, income_statement, vat_rate=0.17):
    # Extract relevant data
    beginning_payables = balance_sheet.loc['Accounts Payable'].iloc[1]  # יתרת ספקים בתחילת התקופה (2022)
    ending_payables = balance_sheet.loc['Accounts Payable'].iloc[0]  # יתרת ספקים בסוף התקופה (2023)
    cost_of_goods_sold = income_statement.loc['Cost Of Revenue'].iloc[0]  # עלות המכר

    # ניטרול מע"מ מיתרת הספקים
    beginning_payables_ex_vat = beginning_payables / (1 + vat_rate)
    ending_payables_ex_vat = ending_payables / (1 + vat_rate)

    # Calculate the average payables (excluding VAT)
    avg_payables_ex_vat = (beginning_payables_ex_vat + ending_payables_ex_vat) / 2

    # Calculate payables days
    payables_days = (avg_payables_ex_vat * 365) / cost_of_goods_sold if cost_of_goods_sold != 0 else None

    return payables_days
