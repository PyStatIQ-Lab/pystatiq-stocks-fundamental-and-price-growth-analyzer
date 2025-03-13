import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# File path for the stocklist.xlsx
FILE_PATH = 'stocklist.xlsx'

# Function to fetch financial data and calculate growth
def get_financial_data(symbol):
    stock = yf.Ticker(symbol)
    income_statement = stock.financials.T  # Transposed for easier access
    quarterly_income_statement = stock.quarterly_financials.T  # Quarterly data
    balance_sheet = stock.balance_sheet.T  # Transposed for easier access
    cashflow = stock.cashflow.T  # Transposed for easier access

    try:
        # Safe extraction method using .get() to handle missing data
        revenue = income_statement.get('Total Revenue', [np.nan])[0]
        cogs = income_statement.get('Cost Of Revenue', [np.nan])[0]
        gross_profit = income_statement.get('Gross Profit', [np.nan])[0]
        operating_income = income_statement.get('Operating Income', [np.nan])[0]
        net_income = income_statement.get('Net Income', [np.nan])[0]
        
        # Extract balance sheet data
        cash_equivalents = balance_sheet.get('Cash And Cash Equivalents', [np.nan])[0]
        accounts_receivable = balance_sheet.get('Accounts Receivable', [np.nan])[0]
        inventory = balance_sheet.get('Inventory', [np.nan])[0]
        pp_and_e = balance_sheet.get('Property, Plant and Equipment', [np.nan])[0]
        goodwill = balance_sheet.get('Goodwill', [np.nan])[0]
        total_assets = balance_sheet.get('Total Assets', [np.nan])[0]
        total_liabilities = balance_sheet.get('Total Liabilities Net Minority Interest', [np.nan])[0]
        equity = balance_sheet.get('Total Equity Gross Minority Interest', [np.nan])[0]

        # Profitability Ratios
        gross_profit_margin = (gross_profit / revenue) * 100 if revenue else np.nan
        operating_profit_margin = (operating_income / revenue) * 100 if revenue else np.nan
        net_profit_margin = (net_income / revenue) * 100 if revenue else np.nan
        roe = (net_income / equity) * 100 if equity else np.nan
        roa = (net_income / total_assets) * 100 if total_assets else np.nan

        # Liquidity Ratios
        current_ratio = (balance_sheet.get('Current Assets', [np.nan])[0] /
                         balance_sheet.get('Current Liabilities', [np.nan])[0]) if balance_sheet.get('Current Liabilities', [np.nan])[0] else np.nan
        quick_ratio = ((balance_sheet.get('Current Assets', [np.nan])[0] - inventory) /
                       balance_sheet.get('Current Liabilities', [np.nan])[0]) if balance_sheet.get('Current Liabilities', [np.nan])[0] else np.nan

        # Leverage Ratios
        debt_to_equity = (total_liabilities / equity) if equity else np.nan
        debt_to_assets = (total_liabilities / total_assets) if total_assets else np.nan

        # Efficiency Ratios
        asset_turnover = (revenue / total_assets) if total_assets else np.nan
        inventory_turnover = (cogs / inventory) if inventory else np.nan
        receivables_turnover = (revenue / accounts_receivable) if accounts_receivable else np.nan

        # Valuation Ratios
        pe_ratio = stock.info.get('trailingPE', np.nan)
        pb_ratio = stock.info.get('priceToBook', np.nan)
        dividend_yield = stock.info.get('dividendYield', np.nan)

        # Cash Flow Ratios
        operating_cash_flow = cashflow.get('Operating Cash Flow', [np.nan])[0]
        free_cash_flow = operating_cash_flow - cashflow.get('Capital Expenditures', [np.nan])[0] if operating_cash_flow and cashflow.get('Capital Expenditures', [np.nan])[0] else np.nan

        # Calculate growth for each financial metric (annual growth)
        revenue_growth = ((revenue - income_statement.get('Total Revenue', [np.nan])[1]) / income_statement.get('Total Revenue', [np.nan])[1]) * 100 if len(income_statement) > 1 else np.nan
        cogs_growth = ((cogs - income_statement.get('Cost Of Revenue', [np.nan])[1]) / income_statement.get('Cost Of Revenue', [np.nan])[1]) * 100 if len(income_statement) > 1 else np.nan
        gross_profit_growth = ((gross_profit - income_statement.get('Gross Profit', [np.nan])[1]) / income_statement.get('Gross Profit', [np.nan])[1]) * 100 if len(income_statement) > 1 else np.nan
        operating_income_growth = ((operating_income - income_statement.get('Operating Income', [np.nan])[1]) / income_statement.get('Operating Income', [np.nan])[1]) * 100 if len(income_statement) > 1 else np.nan
        net_income_growth = ((net_income - income_statement.get('Net Income', [np.nan])[1]) / income_statement.get('Net Income', [np.nan])[1]) * 100 if len(income_statement) > 1 else np.nan

        return {
            'Symbol': symbol,
            'Revenue': revenue,
            'Revenue Growth': revenue_growth,
            'COGS Growth': cogs_growth,
            'Gross Profit Growth': gross_profit_growth,
            'Operating Income Growth': operating_income_growth,
            'Net Income Growth': net_income_growth,
            'Gross Profit Margin': gross_profit_margin,
            'Operating Profit Margin': operating_profit_margin,
            'Net Profit Margin': net_profit_margin,
            'Return on Equity (ROE)': roe,
            'Return on Assets (ROA)': roa,
            'Current Ratio': current_ratio,
            'Quick Ratio': quick_ratio,
            'Debt-to-Equity': debt_to_equity,
            'Debt-to-Assets': debt_to_assets,
            'Asset Turnover': asset_turnover,
            'Inventory Turnover': inventory_turnover,
            'Receivables Turnover': receivables_turnover,
            'P/E Ratio': pe_ratio,
            'P/B Ratio': pb_ratio,
            'Dividend Yield': dividend_yield,
            'Operating Cash Flow to Sales': (operating_cash_flow / revenue) * 100 if operating_cash_flow and revenue else np.nan,
            'Free Cash Flow': free_cash_flow
        }
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
        return None

# Function to fetch stock price performance over the past period
def get_stock_price_performance(symbol, period):
    stock = yf.Ticker(symbol)
    end_date = datetime.now()
    start_date = end_date - timedelta(days=period)
    
    data = stock.history(start=start_date, end=end_date)
    if len(data) > 0:
        price_change = (data['Close'].iloc[-1] - data['Close'].iloc[0]) / data['Close'].iloc[0]
        return price_change
    else:
        return np.nan

# Main function to analyze all stocks
def analyze_and_display(sheet_name):
    xl = pd.ExcelFile(FILE_PATH)
    df = xl.parse(sheet_name)
    stock_symbols = df['Symbol'].dropna().tolist()
    
    results = []
    for symbol in stock_symbols:
        financial_data = get_financial_data(symbol)
        if financial_data:
            for period in [30, 180, 365]:  # 1 month, 6 months, 1 year
                price_performance = get_stock_price_performance(symbol, period)
                financial_data[f'{period} Day Price Performance'] = price_performance
            
            results.append(financial_data)
    
    # Convert results to DataFrame and display in Streamlit
    results_df = pd.DataFrame(results)
    st.write(f"### Analysis Results for {sheet_name}")
    st.dataframe(results_df, use_container_width=True)

# Streamlit UI
def main():
    st.title("Stock Financial Analysis")
    
    # Sheet names from the Excel file
    sheet_names = ['NIFTY50', 'NIFTYNEXT50', 'NIFTY100', 'NIFTY20', 'NIFTY500']
    
    sheet_name = st.selectbox("Select a sheet to analyze", sheet_names)
    
    if st.button("Analyze Stocks"):
        analyze_and_display(sheet_name)

if __name__ == "__main__":
    main()
