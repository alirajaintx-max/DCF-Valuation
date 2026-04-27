DCF Valuation Model
A Python tool that values any public company using a Discounted Cash Flow model with live financial data and sensitivity analysis.
What it does

Fetches live revenue, market cap, and share price data using yfinance
Projects Free Cash Flows over 5 years using configurable growth assumptions
Calculates Enterprise Value using discounted FCFs and Terminal Value
Compares intrinsic value vs current market price to give an undervalued or overvalued verdict
Generates a 5x5 sensitivity table across WACC and terminal growth assumptions
Exports a full report to Excel with 4 structured sheets

Results (Apple — AAPL)

Intrinsic Value: $139.74 per share
Enterprise Value: $2,152B
PV of Terminal Value: 85% of total EV
Sensitivity range: $89 to $241 depending on assumptions

How to run it

Install dependencies: pip install yfinance pandas numpy matplotlib openpyxl
Run the script: python dcf_valuation.py
Change the company by editing TICKER = "AAPL" at the top of the script

Technologies
Python · yfinance · NumPy · Pandas · Matplotlib · openpyxl
