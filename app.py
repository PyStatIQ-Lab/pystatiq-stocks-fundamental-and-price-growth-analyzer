import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas

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

        # Price performance (30-day)
        price_history = stock.history(period="30d")
        price_performance = ((price_history['Close'][-1] - price_history['Close'][0]) / price_history['Close'][0]) * 100

        return {
            'Symbol': symbol,
            'Revenue': revenue,
            'Revenue Growth': revenue_growth,
            'COGS Growth': cogs_growth,
            'Gross Profit Growth': gross_profit_growth,
            'Operating Income Growth': operating_income_growth,
            'Net Income Growth': net_income_growth,
            '30 Day Price Performance': price_performance,
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

# Function to generate PDF report
def generate_pdf(results_df, sheet_name, analyst_name):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    
    # Title and metadata
    c.setFont("Helvetica-Bold", 18)
    c.drawString(30, 750, "Stock Financial Growth and Price Movement Analysis Dashboard")
    c.setFont("Helvetica", 12)
    c.drawString(30, 730, f"Sheet Name: {sheet_name}")
    c.drawString(30, 710, f"Research Analyst: {analyst_name}")
    c.drawString(30, 690, f"Date of Analysis: {datetime.now().strftime('%Y-%m-%d')}")
    c.drawString(30, 670, "Disclaimer: This analysis is for informational purposes only. It does not constitute investment advice.")
    
    # Insert a table
    c.setFont("Helvetica", 10)
    y_position = 630
    
    # Add header
    headers = results_df.columns.tolist()
    for i, header in enumerate(headers):
        c.drawString(30 + i*80, y_position, header)
    y_position -= 20
    
    # Add data rows
    for index, row in results_df.iterrows():
        for i, col in enumerate(headers):
            c.drawString(30 + i*80, y_position, str(row[col]))
        y_position -= 20
        if y_position < 100:
            c.showPage()
            y_position = 750

    # Save PDF to buffer
    c.save()
    packet.seek(0)
    
    return packet

# Function to fetch stock symbols from the selected sheet
def get_stock_symbols(sheet_name):
    xl = pd.ExcelFile(FILE_PATH)
    df = xl.parse(sheet_name)
    return df['Symbol'].dropna().tolist()

# Main function to analyze and display
def analyze_and_display(sheet_name):
    stock_symbols = get_stock_symbols(sheet_name)
    results = []
    
    for symbol in stock_symbols:
        financial_data = get_financial_data(symbol)
        if financial_data:
            results.append(financial_data)
    
    results_df = pd.DataFrame(results)
    return results_df

# Streamlit UI
def main():
    st.title("Stock Financial Analysis Dashboard")
    
    sheet_names = ['NIFTY50', 'NIFTYNEXT50', 'NIFTY100', 'NIFTY20', 'NIFTY500']
    sheet_name = st.selectbox("Select a sheet to analyze", sheet_names)
    analyst_name = st.text_input("Research Analyst Name", "")
    
    if st.button("Analyze Stocks"):
        if analyst_name:
            results_df = analyze_and_display(sheet_name)
            
            # Show Top 5 Stocks with Highest Revenue Growth
            st.subheader(f"Top 5 Stocks with Highest Revenue Growth")
            top_stocks_revenue_growth = results_df.sort_values(by='Revenue Growth', ascending=False).head(10)
            st.dataframe(top_stocks_revenue_growth[['Symbol', 'Revenue Growth', 'Net Income Growth', 'Operating Income Growth', 'Gross Profit Growth']], use_container_width=True)
            
            # Show Top 5 Stocks with Lowest Price Performance in Last 1 Month
            st.subheader(f"Top 5 Stocks with Lowest Price Performance in Last 1 Month")
            top_stocks_price_performance = results_df.sort_values(by='30 Day Price Performance').head(10)
            st.dataframe(top_stocks_price_performance[['Symbol', '30 Day Price Performance', 'Revenue Growth', 'Net Income Growth', 'Operating Income Growth']], use_container_width=True)

            # Visualizing Revenue Growth for the top 10 stocks
            st.subheader("Revenue Growth Visualization")
            st.bar_chart(results_df['Revenue Growth'].head(10))
            
            # Generate PDF button
            

if __name__ == "__main__":
    main()
