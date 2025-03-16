# Technical Analysis Indicators

This project is **NOT** intended for providing professional financial advice or making investment decisions. It's a practice repository created to identify potential patterns in stock data. The analysis methods used here are for educational and practice purposes only. 

## Project Overview
This project analyzes stock data using various technical indicators like EMA, MACD, RSI, Volume indicators, Bollinger Bands, and Candlestick patterns. The workflow follows a sequential process from data scraping to generating a final summary report.

## Workflow

### Step 1: Data Scraping
Navigate to the folder `1. scrape_from_scratch` and run the Python script inside.

This will scrape stock data and create an Excel file which will be automatically copied to the next folder.

**Note:** You must run this step if you don't have existing data. This is the starting point of the workflow.

### Step 2: Append to Existing Data
Navigate to the folder `2. append_to_existing_data` and run the Python script inside.

This will update the existing data with the latest information and create an updated Excel file with data up to the current date.

### Step 3: Analyze Shares
Navigate to the folder `3. analyse_shares` and run each Python script **one by one** in the following order. Make sure each script completes execution before running the next one:

- `1. stock_analysis_ema.py`
- `2. stock_analysis_macd.py`
- `3. stock_analysis_rsi.py`
- `4. candlestick_patterns_custom.py`
- `5. bollinger_band.py`
- `6. volume_indicator.py`

### Step 4: Access Analysis Results
After all scripts have executed, navigate to the folder `4. analysis_result`.

Inside this folder, you'll find the following subfolders:
- bollinger
- candlestick pattern
- ema
- macd
- rsi
- volume

Within each subfolder, results are organized by date (YYYY-MM-DD format). Find the most recent date folder in each category to access the latest results.

#### Scoring Rules:
1. **EMA Result**:
   - If both "5 EMA crosses 13 EMA" and "5 EMA crosses 26 EMA" are "Yes": +2
   - If only "5 EMA crosses 13 EMA" is "Yes": +1
   - Otherwise: 0

2. **MACD Result**:
   - If "MACD crosses Signal Line" is "Yes": +1
   - Otherwise: 0

3. **RSI Result**:
   - If "RSI Status" is "uptrend": +1
   - If "RSI Status" is "sideways": 0
   - If "RSI Status" is "downtrend": -1

4. **Volume Result**:
   - If "Increasing Volume" is "Yes": +1
   - Otherwise: 0

6. **Bollinger Band Result**:
   - If "Crossed Band in ALL Last 3 Days" is "Yes", then write "Yes" else "No"

7. **Candlestick Pattern Result**:
   - Each pattern listed in "Bullish Reversal" and "Bullish Continuation"
   - Each pattern listed in "Bearish Reversal" and "Bearish Continuation"
   - I am not satisfied with the accuracy of this code in identifying pattern, so I tried TA-LIB but it is not that great either. Hence, these identified patterns can be of no great use. 

## Important Notes:
1. This is just my fun project. But I would love to improve it if I get time in the future. 
