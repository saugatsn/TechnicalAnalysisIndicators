from selenium import webdriver
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
import time
import chromedriver_autoinstaller as chromedriver
import os
import glob
import shutil
import re
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("scraper.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

# Install ChromeDriver
chromedriver.install()

def find_latest_data_file():
    """
    Find the most recent Excel file with the pattern shares_data_till_yyyy_mm_dd.xlsx
    Returns the file path and the date from the filename
    """
    logger.info("Looking for latest data file...")
    files = glob.glob("shares_data_till_*.xlsx")
    if not files:
        logger.info("No existing data files found.")
        return None, None
    
    # Extract dates from filenames and find the most recent one
    latest_file = None
    latest_date = None
    
    for file in files:
        match = re.search(r'shares_data_till_(\d{4})_(\d{2})_(\d{2})\.xlsx', file)
        if match:
            year, month, day = map(int, match.groups())
            file_date = datetime(year, month, day).replace(hour=0, minute=0, second=0, microsecond=0)
            
            if latest_date is None or file_date > latest_date:
                latest_date = file_date
                latest_file = file
    
    if latest_file:
        logger.info(f"Found latest file: {latest_file} (date: {latest_date.strftime('%Y-%m-%d')})")
    return latest_file, latest_date

def find_latest_common_date(file_path):
    """
    Analyze the first 10 sheets of the Excel file to find the latest date 
    that appears in the most sheets
    """
    logger.info("Analyzing Excel file to find latest common date...")
    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names[:10]  # Consider only first 10 sheets
    
    # Dictionary to count date occurrences across sheets
    date_counts = {}
    latest_dates = {}
    
    for sheet in sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            if 'Date' in df.columns and not df['Date'].empty:
                # Find the latest date in this sheet, ignoring time component
                df['Date'] = pd.to_datetime(df['Date']).dt.normalize()
                latest_date = df['Date'].max()
                if latest_date not in date_counts:
                    date_counts[latest_date] = 0
                date_counts[latest_date] += 1
                latest_dates[sheet] = latest_date
        except Exception as e:
            logger.error(f"Error reading sheet {sheet}: {e}")
    
    # Find the date with the highest occurrence
    if date_counts:
        most_common_date = max(date_counts.items(), key=lambda x: x[1])[0]
        logger.info(f"Latest common date found: {most_common_date.strftime('%Y-%m-%d')}")
        return most_common_date
    
    return None

def format_date_for_search(date):
    """Format date object as YYYY-MM-DD for website search"""
    return date.strftime('%Y-%m-%d')

def format_date_for_filename(date):
    """Format date object as YYYY_MM_DD for filename"""
    return date.strftime('%Y_%m_%d')

def search(driver, date_str):
    """
    Search by date on the website
    Returns True if data is found, False otherwise
    """
    try:
        logger.info(f"Searching website for date: {date_str}...")
        driver.get("https://www.sharesansar.com/today-share-price")
        
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@id='fromdate']"))
        )
        date_input = driver.find_element("xpath", "//input[@id='fromdate']")
        
        # Clear any existing value
        date_input.clear()
        time.sleep(1)
        
        # Enter the date in YYYY-MM-DD format
        date_input.send_keys(date_str)
        time.sleep(1)
        
        # Find and click the search button
        search_btn = driver.find_element("xpath", "//button[@id='btn_todayshareprice_submit']")
        search_btn.click()
        
        # Wait for results or no-data message
        time.sleep(3)
        
        # Check if "No Record Found" message appears
        no_data_elements = driver.find_elements("xpath", "//*[contains(text(), 'No Record Found')]")
        if no_data_elements:
            logger.info(f"No record found for date: {date_str}")
            return False
        
        # Check if "Could not find floorsheet" message appears
        no_floorsheet_elements = driver.find_elements("xpath", "//*[contains(text(), 'Could not find floorsheet matching the search criteria')]")
        if no_floorsheet_elements:
            logger.info(f"No floorsheet found for date: {date_str}")
            return False
        
        # Check if the table with data is present
        table_elements = driver.find_elements("xpath", "//table[contains(@class, 'table-bordered')]")
        if not table_elements:
            logger.info(f"No data table found for date: {date_str}")
            return False
            
        logger.info(f"Data found for date: {date_str}")
        return True
        
    except Exception as e:
        logger.error(f"Search error for date {date_str}: {str(e)}")
        return False

def get_page_table(driver, table_class):
    """
    Extract table data from the current page
    """
    try:
        logger.info("Extracting table data from current page...")
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='floatThead-wrapper']"))
        )
        soup = BeautifulSoup(driver.page_source, 'lxml')
        table = soup.find("table", {"class": table_class})
        
        if not table:
            logger.warning("No table found on current page.")
            return pd.DataFrame()
            
        tab_data = [[cell.text.replace('\r', '').replace('\n', '') for cell in row.find_all(["th","td"])]
                            for row in table.find_all("tr")]
        df = pd.DataFrame(tab_data)
        logger.info(f"Extracted {len(df)-1} rows of data from current page")
        return df
        
    except Exception as e:
        logger.error(f"Table extraction error: {str(e)}")
        return pd.DataFrame()

def scrape_data(driver, date_str):
    """
    Scrape data for a specific date, handling pagination
    """
    logger.info(f"Starting data scraping for date: {date_str}")
    if not search(driver, date_str):
        return None
    
    df = pd.DataFrame()
    page_count = 0
    
    try:
        while True:
            page_count += 1
            logger.info(f"Scraping page {page_count} for date {date_str}")
            
            page_table_df = get_page_table(driver, table_class="table table-bordered table-striped table-hover dataTable compact no-footer")
            
            if not page_table_df.empty:
                df = pd.concat([df, page_table_df], ignore_index=True)
            
            try:
                next_btn = driver.find_element(By.LINK_TEXT, 'Next')
                if next_btn.is_enabled():
                    driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(2)  # Wait for page to load
                else:
                    break
            except NoSuchElementException:
                logger.info("No more pages to scrape.")
                break
            except Exception as e:
                logger.error(f"Pagination error: {str(e)}")
                break
        
        logger.info(f"Completed scraping for date {date_str}. Total rows: {len(df)-1}")
        return df
        
    except Exception as e:
        logger.error(f"Scraping error: {str(e)}")
        return None


def clean_df(df, date_str):
    """
    Clean and format the dataframe with correct column mapping
    """
    if df is None or df.empty:
        logger.info(f"No data to clean for date: {date_str}")
        return None
        
    try:
        logger.info(f"Cleaning data for date: {date_str}")
        new_df = df.drop_duplicates(keep='first')  # drop all duplicates
        
        # Check if the dataframe has content
        if new_df.shape[0] <= 1:  # Only header or empty
            logger.warning("Dataframe has only headers or is empty after cleaning.")
            return None
            
        new_header = new_df.iloc[0]  # grabbing the first row for the header
        new_df = new_df[1:]  # taking the data lower than the header row
        new_df.columns = new_header  # setting the header row as the df header
        
        # Create a new dataframe with only the columns we need
        result_df = pd.DataFrame()
        
        # Add Date column with the current date we're scraping (date only, no time)
        result_df["Date"] = [date_str] * len(new_df)  # Ensure date is added to each row
        
        # Map the columns from website to Excel format
        # Improved mapping to handle various column names that might appear on the website
        excel_columns = {
            "Open": ["Open"],  # Excel column: [possible website column names]
            "High": ["High"],
            "Low": ["Low"],
            "Ltp": ["Close", "LTP", "Ltp", "Last"],  # Map any of these to Ltp
            "% Change": ["Diff %", "% Change", "Change %", "Change"],
            "Qty": ["Vol", "Volume", "Qty", "Quantity"],
            "Turnover": ["Turnover"]
        }
        
        # Handle Symbol column separately
        if "Symbol" in new_df.columns:
            result_df["Symbol"] = new_df["Symbol"]
        elif "Traded Companies" in new_df.columns:
            result_df["Symbol"] = new_df["Traded Companies"]
        else:
            # Look for any column that might contain symbol information
            for col in new_df.columns:
                if "symbol" in col.lower() or "compan" in col.lower() or "script" in col.lower():
                    result_df["Symbol"] = new_df[col]
                    logger.info(f"Using column '{col}' as Symbol")
                    break
            else:
                logger.error("Could not find a Symbol column in the data")
                return None
        
        # Map columns using the flexible mapping
        website_columns = set(new_df.columns)
        logger.info(f"Website columns found: {', '.join(website_columns)}")
        
        for excel_col, possible_web_cols in excel_columns.items():
            mapped = False
            for web_col in possible_web_cols:
                if web_col in website_columns:
                    result_df[excel_col] = new_df[web_col]
                    logger.info(f"Mapped website column '{web_col}' to Excel column '{excel_col}'")
                    mapped = True
                    break
            
            if not mapped:
                logger.warning(f"Could not find a match for Excel column '{excel_col}'")
                result_df[excel_col] = ""
        
        logger.info(f"Data cleaning completed. Rows: {len(result_df)}")
        return result_df
        
    except Exception as e:
        logger.error(f"Data cleaning error: {str(e)}")
        return None

def sanitize_sheet_name(sheet_name):
    """
    Sanitize sheet name to remove characters not allowed in Excel
    """
    # Excel doesn't allow these characters in sheet names: / \ ? * [ ]
    invalid_chars = ['/', '\\', '?', '*', '[', ']', ':', ' ']
    
    # Replace invalid characters with underscore
    sanitized_name = sheet_name
    for char in invalid_chars:
        sanitized_name = sanitized_name.replace(char, '_')
    
    # Excel sheet names have a 31 character limit
    if len(sanitized_name) > 31:
        sanitized_name = sanitized_name[:31]
    
    # Sheet name can't be empty
    if not sanitized_name:
        sanitized_name = "Unknown"
    
    # Log if the name was changed
    if sanitized_name != sheet_name:
        logger.info(f"Sanitized sheet name from '{sheet_name}' to '{sanitized_name}'")
    
    return sanitized_name


def update_excel_file(existing_file, new_data_dict):
    """
    Update the Excel file with new data for each symbol/sheet
    """
    try:
        logger.info(f"Updating Excel file: {existing_file}")
        logger.info(f"New data available for {len(new_data_dict)} symbols")
        
        # Load the existing Excel file
        with pd.ExcelFile(existing_file) as xls:
            sheet_names = xls.sheet_names
        
        # Create a writer for the output Excel file
        temp_file = "temp_updated_file.xlsx"
        writer = pd.ExcelWriter(temp_file, engine='openpyxl')
        
        # Create a mapping from original sheet names to sanitized sheet names
        # and a reverse mapping to maintain the connection to original symbols
        sanitized_to_original = {}
        original_to_sanitized = {}
        
        # Track statistics
        updated_sheets = 0
        unchanged_sheets = 0
        new_sheets = 0
        
        # First sanitize all existing sheet names
        for sheet in sheet_names:
            sanitized_sheet = sanitize_sheet_name(sheet)
            original_to_sanitized[sheet] = sanitized_sheet
            sanitized_to_original[sanitized_sheet] = sheet
        
        # Then sanitize all new symbol names
        for symbol in new_data_dict.keys():
            sanitized_symbol = sanitize_sheet_name(symbol)
            original_to_sanitized[symbol] = sanitized_symbol
            sanitized_to_original[sanitized_symbol] = symbol
        
        # Process each existing sheet
        for sheet in sheet_names:
            try:
                # Get the sanitized sheet name
                sanitized_sheet = original_to_sanitized[sheet]
                
                # Read the existing sheet data
                existing_df = pd.read_excel(existing_file, sheet_name=sheet)
                
                # If we have new data for this symbol, append it
                if sheet in new_data_dict:
                    # Convert Date column to string format if it's not already
                    if 'Date' in existing_df.columns:
                        # Ensure dates are in yyyy-mm-dd string format
                        existing_df['Date'] = pd.to_datetime(existing_df['Date']).dt.strftime('%Y-%m-%d')
                    
                    # Ensure dates in new data are in string format as well
                    new_data = new_data_dict[sheet].copy()
                    if 'Date' in new_data.columns:
                        new_data['Date'] = pd.to_datetime(new_data['Date']).dt.strftime('%Y-%m-%d')
                    
                    # Append new data and sort by date
                    combined_df = pd.concat([existing_df, new_data], ignore_index=True)
                    combined_df = combined_df.drop_duplicates(subset=['Date'], keep='last')
                    combined_df = combined_df.sort_values('Date')
                    
                    # Write the updated data to the sheet
                    combined_df.to_excel(writer, sheet_name=sanitized_sheet, index=False)
                    updated_sheets += 1
                    logger.info(f"Updated sheet: {sanitized_sheet} with new data")
                else:
                    # No new data for this symbol, keep the sheet as is
                    existing_df.to_excel(writer, sheet_name=sanitized_sheet, index=False)
                    unchanged_sheets += 1
            
            except Exception as e:
                logger.error(f"Error processing sheet {sheet}: {str(e)}")
                # If there's an error, preserve the original sheet
                try:
                    existing_df = pd.read_excel(existing_file, sheet_name=sheet)
                    existing_df.to_excel(writer, sheet_name=sanitized_sheet, index=False)
                except Exception as inner_e:
                    logger.error(f"Could not preserve original sheet {sheet}: {str(inner_e)}")
        
        # Add any new symbols that weren't in the original file
        for symbol, data in new_data_dict.items():
            if symbol not in sheet_names:
                sanitized_symbol = original_to_sanitized[symbol]
                data.to_excel(writer, sheet_name=sanitized_symbol, index=False)
                new_sheets += 1
                logger.info(f"Added new sheet for symbol: {sanitized_symbol} (original: {symbol})")
        
        # Save the Excel file
        writer.close()
        
        # Replace the original file with the temp file
        os.replace(temp_file, existing_file)
        
        logger.info(f"Excel update summary:")
        logger.info(f"- Updated sheets: {updated_sheets}")
        logger.info(f"- Unchanged sheets: {unchanged_sheets}")
        logger.info(f"- New sheets added: {new_sheets}")
        logger.info(f"- Total sheets: {updated_sheets + unchanged_sheets + new_sheets}")
        
        return True
        
    except Exception as e:
        logger.error(f"Excel update error: {str(e)}")
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
                logger.info(f"Removed temporary file: {temp_file}")
            except:
                pass
        return False

def rename_and_copy_file(current_file, latest_date):
    """
    Rename the file with the new latest date and copy to analysis directory
    """
    try:
        # Format the date for the filename
        date_str = format_date_for_filename(latest_date)
        new_filename = f"shares_data_till_{date_str}.xlsx"
        
        logger.info(f"Renaming file to: {new_filename}")
        # Rename the file
        os.rename(current_file, new_filename)
        
        # Create path to analysis directory (one level up, then into "3. analyse_shares")
        analysis_dir = Path("..") / "3. analyse_shares"
        
        # Ensure the directory exists
        if not os.path.exists(analysis_dir):
            logger.warning(f"Analysis directory not found: {analysis_dir}")
            return False
        
        logger.info(f"Copying file to analysis directory: {analysis_dir}")
        # Check for existing files with similar pattern and remove them
        existing_files = glob.glob(str(analysis_dir / "shares_data_till_*.xlsx"))
        for file in existing_files:
            try:
                os.remove(file)
                logger.info(f"Removed existing file: {file}")
            except Exception as e:
                logger.error(f"Error removing file {file}: {str(e)}")
        
        # Copy the new file
        shutil.copy2(new_filename, analysis_dir)
        logger.info(f"File successfully copied to: {analysis_dir / new_filename}")
        
        return True
        
    except Exception as e:
        logger.error(f"File renaming error: {str(e)}")
        return False

def setup_webdriver():
    """
    Set up and configure the Chrome WebDriver with appropriate options
    to minimize browser messages
    """
    options = Options()
    options.add_argument("--headless=new")  # Modern headless mode
    options.add_argument("--disable-gpu")  # Disable GPU acceleration
    options.add_argument("--window-size=1920,1080")  # Set window size
    options.add_argument("--no-sandbox")  # Bypass OS security model
    options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
    options.add_argument("--log-level=3")  # Suppress most Chrome logging
    options.add_argument("--silent")  # Silent mode
    
    # Disable WebGL, SwiftShader, and other notifications
    options.add_argument("--disable-webgl")
    options.add_argument("--disable-software-rasterizer")
    
    # Use a realistic user agent
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.115 Safari/537.36")
    
    # Add experimental options to suppress console messages
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(120)
    return driver

def main():
    try:
        logger.info("=== Share Price Data Scraper ===")
        logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Find the latest data file
        latest_file, file_date = find_latest_data_file()
        if not latest_file:
            logger.error("ERROR: No existing data file found. Please create an initial file first.")
            return
        
        # Find the latest common date in the file
        latest_common_date = find_latest_common_date(latest_file)
        if not latest_common_date:
            logger.error("ERROR: Could not determine the latest common date in the file.")
            return
        
        # Set up the date range to scrape (from day after latest_common_date to today)
        start_date = latest_common_date + timedelta(days=1)
        start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        
        end_date = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
        
        if start_date > end_date:
            logger.info("No new data to fetch - the latest date in the file is already up to date.")
            return
        
        logger.info(f"Will scrape data from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        
        # Set up the webdriver with improved options to reduce browser messages
        logger.info("Initializing web driver...")
        driver = setup_webdriver()
        
        # Dictionary to store new data by symbol
        all_new_data = {}
        latest_date_with_data = None
        dates_processed = 0
        dates_with_data = 0
        
        try:
            # Process each date in the range
            current_date = start_date
            while current_date <= end_date:
                date_str = format_date_for_search(current_date)
                logger.info(f"\n[{dates_processed+1}/{(end_date-start_date).days+1}] Processing date: {date_str}")
                
                # Scrape data for this date
                scraped_df = scrape_data(driver, date_str)
                
                if scraped_df is not None and not scraped_df.empty:
                    # Clean and format the data
                    cleaned_df = clean_df(scraped_df, date_str)
                    
                    if cleaned_df is not None and not cleaned_df.empty:
                        # Record this as the latest date with data
                        latest_date_with_data = current_date
                        dates_with_data += 1
                        
                        # Group data by symbol
                        symbols_count = 0
                        if 'Symbol' in cleaned_df.columns:
                            for symbol, group in cleaned_df.groupby('Symbol'):
                                # Skip empty or invalid symbols
                                if not symbol or pd.isna(symbol) or symbol.strip() == '':
                                    continue
                                    
                                # Create a copy of the group without modifying the original
                                symbol_data = group.copy()
                                
                                # Use original symbol as key but remove it from the data
                                if symbol not in all_new_data:
                                    all_new_data[symbol] = pd.DataFrame()
                                
                                # Remove Symbol column from the data we'll store
                                symbol_data = symbol_data.drop(columns=['Symbol'])
                                
                                # Append this data to the symbol's dataframe
                                all_new_data[symbol] = pd.concat([all_new_data[symbol], symbol_data], ignore_index=True)
                                symbols_count += 1
                            
                            logger.info(f"Successfully processed data for {date_str} - {symbols_count} symbols")
                        else:
                            logger.error(f"ERROR: 'Symbol' column missing in cleaned data for {date_str}")
                    else:
                        logger.warning(f"No usable data found for {date_str} after cleaning")
                else:
                    logger.info(f"No data found for {date_str} - skipping")
                
                # Move to next date
                current_date += timedelta(days=1)
                dates_processed += 1
                
        except KeyboardInterrupt:
            logger.warning("\nProcess interrupted by user. Saving current progress...")
        except Exception as e:
            logger.error(f"\nError during scraping: {str(e)}")
        finally:
            # Close the webdriver
            try:
                driver.quit()
                logger.info("Web driver closed")
            except:
                pass
        
        # Print summary of scraping
        logger.info("\n=== Scraping Summary ===")
        logger.info(f"Dates processed: {dates_processed}/{(end_date-start_date).days+1}")
        logger.info(f"Dates with data: {dates_with_data}")
        logger.info(f"Symbols with new data: {len(all_new_data)}")
        
        # Update the Excel file with new data if we have any
        if all_new_data and latest_date_with_data:
            logger.info(f"\nUpdating Excel file with new data up to {latest_date_with_data.strftime('%Y-%m-%d')}")
            
            # Update the Excel file
            if update_excel_file(latest_file, all_new_data):
                # Rename and copy the file
                if rename_and_copy_file(latest_file, latest_date_with_data):
                    logger.info("\nProcess completed successfully!")
                else:
                    logger.error("\nERROR: Could not rename and copy the file.")
            else:
                logger.error("\nERROR: Could not update the Excel file.")
        else:
            logger.info("\nNo new data was found to update the file.")
            
        logger.info(f"Ended at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
    except Exception as e:
        logger.critical(f"CRITICAL ERROR: {str(e)}")
        import traceback
        logger.critical(traceback.format_exc())
        logger.error("Process terminated due to error.")

if __name__ == "__main__":
    main()