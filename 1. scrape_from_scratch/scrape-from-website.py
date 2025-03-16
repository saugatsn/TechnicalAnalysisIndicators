import random
import time
import pandas as pd
import os
import logging
import traceback
import re
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    WebDriverException,
    StaleElementReferenceException
)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraper.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# List of companies 
companies = [
  "ACLBSL", "ADBL", "ADBLD83", "AHL", "AHPC", "AKJCL", "AKPL", "ALBSL", "ALICL", "ANLB", 
  "API", "AVYAN", "BARUN", "BBC", "BEDC", "BFC", "BGWT", "BHDC", "BHL", "BHPL", 
  "BNHC", "BNT", "BPCL", "C30MF", "CBBL", "CBLD88", "CFCL", "CGH", "CHCL", "CHDC", 
  "CHL", "CIT", "CITY", "CIZBD86", "CKHL", "CLI", "CMF2", "CORBL", "CYCL", "CZBIL", 
  "DDBL", "DHPL", "DLBS", "DOLTI", "DORDI", "EBL", "EBLD85", "EBLD86", "EDBL", "EHPL", 
  "ENL", "FMDBL", "FOWAD", "GBBD85", "GBBL", "GBILD84/85", "GBIME", "GBIMEP", "GBLBS", "GCIL", 
  "GFCL", "GHL", "GIBF1", "GILB", "GLBSL", "GLH", "GMFBS", "GMFIL", "GMLI", "GRDBL", 
  "GSY", "GUFL", "GVL", "H8020", "HATHY", "HBL", "HBLD83", "HBLD86", "HDHPC", "HDL", 
  "HEI", "HEIP", "HHL", "HIDCL", "HIDCLP", "HLBSL", "HLI", "HPPL", "HRL", "HURJA", 
  "ICFC", "ICFCD88", "IGI", "IHL", "ILBS", "ILI", "JBBL", "JBLB", "JFL", "JOSHI", 
  "JSLBB", "KBL", "KBLD89", "KBSH", "KDBY", "KDL", "KEF", "KKHC", "KMCDB", "KPCL", 
  "KSBBL", "KSBBLD87", "KSY", "LBBL", "LBBLD89", "LEC", "LICN", "LLBS", "LSL", "LUK", 
  "LVF2", "MAKAR", "MANDU", "MATRI", "MBJC", "MBL", "MBLD87", "MCHL", "MDB", "MEHL", 
  "MEL", "MEN", "MERO", "MFIL", "MFLD85", "MHCL", "MHL", "MHNL", "MKCL", "MKHC", 
  "MKHL", "MKJC", "MLBBL", "MLBL", "MLBS", "MLBSL", "MMF1", "MMKJL", "MNBBL", "MND84/85", 
  "MNMF1", "MPFL", "MSHL", "MSLB", "NABBC", "NABIL", "NADEP", "NBF2", "NBF3", "NBL", 
  "NBLD82", "NBLD87", "NESDO", "NFS", "NGPL", "NHDL", "NHPC", "NIBD84", "NIBLGF", "NIBLSTF", 
  "NIBSF2", "NICA", "NICAD8283", "NICBF", "NICD88", "NICFC", "NICGF", "NICGF2", "NICL", "NICLBSL", 
  "NICSF", "NIFRA", "NIL", "NIMB", "NIMBD90", "NIMBPO", "NLG", "NLIC", "NLICL", "NMB", 
  "NMB50", "NMBMF", "NMFBS", "NMLBBL", "NRIC", "NRM", "NRN", "NSIF2", "NTC", "NUBL", 
  "NWCL", "NYADI", "OHL", "PBD84", "PBD85", "PBD88", "PBLD84", "PBLD87", "PCBL", "PFL", 
  "PHCL", "PMHPL", "PMLI", "PPCL", "PPL", "PRIN", "PROFL", "PRSF", "PRVU", "PSF", 
  "RADHI", "RAWA", "RBBD83", "RBCL", "RBCLPO", "RFPL", "RHGCL", "RHPL", "RIDI", "RLFL", 
  "RMF1", "RMF2", "RNLI", "RSDC", "RURU", "SADBL", "SAGF", "SAHAS", "SALICO", "SAMAJ", 
  "SAND2085", "SANIMA", "SAPDBL", "SARBTM", "SBCF", "SBI", "SBID83", "SBL", "SCB", "SDBD87", 
  "SEF", "SFCL", "SFEF", "SFMF", "SGHC", "SGIC", "SHEL", "SHINE", "SHIVM", "SHL", 
  "SHLB", "SHPC", "SICL", "SIFC", "SIGS2", "SIGS3", "SIKLES", "SINDU", "SJCL", "SJLIC", 
  "SKBBL", "SLBBL", "SLBSL", "SLCF", "SMATA", "SMB", "SMFBS", "SMH", "SMHL", "SMJC", 
  "SMPDA", "SNLI", "SONA", "SPC", "SPDL", "SPHL", "SPIL", "SPL", "SRLI", "SSHL", 
  "STC", "SWBBL", "SWMF", "TAMOR", "TPC", "TRH", "TSHL", "TVCL", "UAIL", "UHEWA", 
  "ULBSL", "ULHC", "UMHL", "UMRH", "UNHPL", "UNL", "UNLB", "UPCL", "UPPER", "USHEC", 
  "USHL", "USLB", "VLBS", "VLUCL", "WNLB"
]
# Example list of user agents for rotation
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"
]

# Example list of proxies (format: "ip:port") - replace with real proxies if available
PROXIES = [
    # "123.123.123.123:8080",
    # "234.234.234.234:8000",
]

def get_random_proxy():
    """Return a random proxy from the list, or None if no proxies available."""
    return random.choice(PROXIES) if PROXIES else None

def sanitize_sheet_name(name):
    """
    Sanitize sheet name to be valid for Excel:
    - Replace invalid characters with underscores
    - Ensure name is not longer than 31 characters
    - Ensure name doesn't end with a space
    """
    # Replace invalid characters: [ ] : * ? / \
    sanitized = re.sub(r'[\[\]:*?/\\]', '_', name)
    
    # Trim to 31 characters
    sanitized = sanitized[:31]
    
    # Remove trailing spaces
    sanitized = sanitized.rstrip()
    
    # If empty after sanitizing, use a default name
    if not sanitized:
        sanitized = f"Sheet_{datetime.now().strftime('%H%M%S')}"
        
    return sanitized

def create_driver():
    """Create and return a configured WebDriver instance with error handling."""
    try:
        # Set up Chrome options with more complete headless settings
        options = Options()
        
        # Rotate a random User-Agent
        user_agent = random.choice(USER_AGENTS)
        options.add_argument(f'user-agent={user_agent}')
        
        # Enhanced headless setup to prevent any popups
        options.add_argument('--headless=new')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-notifications')
        options.add_argument('--window-size=1920,1080')
        
        # Use a proxy if available
        proxy = get_random_proxy()
        if proxy:
            options.add_argument(f'--proxy-server={proxy}')
        
        # Initialize the WebDriver
        return webdriver.Chrome(options=options)
    except Exception as e:
        logger.error(f"Failed to create WebDriver: {e}")
        raise

def safe_find_element(driver, by, value, timeout=10):
    """Safely find an element with WebDriverWait and proper error handling."""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return element
    except TimeoutException:
        logger.warning(f"Timeout waiting for element '{value}'")
        return None
    except NoSuchElementException:
        logger.warning(f"Element '{value}' not found")
        return None
    except Exception as e:
        logger.warning(f"Error finding element '{value}': {e}")
        return None

def scrape_share_data(company):
    """Scrape share data for a given company with comprehensive error handling."""
    data = []
    driver = None
    retry_count = 0
    max_retries = 3
    
    while retry_count < max_retries:
        try:
            if driver:
                driver.quit()  # Quit any existing driver session before creating a new one
                
            driver = create_driver()
            
            company_url = f"https://www.sharesansar.com/company/{company.lower()}"
            driver.get(company_url)
            
            # Wait until the page loads
            body = safe_find_element(driver, By.TAG_NAME, "body")
            if not body:
                raise Exception("Page body not found, possibly blocked or unavailable")
            
            # Click the "Price History" tab
            price_history_tab = safe_find_element(driver, By.XPATH, "//a[contains(text(), 'Price History')]")
            if not price_history_tab:
                raise Exception("Price History tab not found")
            
            price_history_tab.click()
            
            # Wait for the price history table to appear by its ID
            table = safe_find_element(driver, By.ID, "myTableCPriceHistory")
            if not table:
                raise Exception("Price history table not found")
            
            # Use JavaScript to safely change the entries dropdown to 50
            try:
                driver.execute_script("""
                    const selectElement = document.querySelector('select[name="myTableCPriceHistory_length"]');
                    if (selectElement) {
                        selectElement.value = '50';
                        selectElement.dispatchEvent(new Event('change'));
                    }
                """)
                # Wait for the table to reload with 50 entries
                time.sleep(random.uniform(2, 3))
            except Exception as e:
                logger.warning(f"Failed to change entries to 50: {e}")
                # Continue anyway as we can still scrape whatever entries are showing
            
            # First get the headers to know all columns
            headers = table.find_elements(By.XPATH, ".//thead/tr/th")
            column_names = [header.text.strip() for header in headers if header.text.strip()]
            
            # Filter out the S.N. column if it exists
            sn_index = None
            for i, col_name in enumerate(column_names):
                if col_name.lower() in ["s.n.", "sn", "s.n"]:
                    sn_index = i
                    break
            
            # Create a filtered list of column names without S.N.
            filtered_column_names = [col for i, col in enumerate(column_names) if i != sn_index]
            
            # Extract data from the table rows in tbody
            rows = table.find_elements(By.XPATH, ".//tbody/tr")
            
            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 6:
                        record = {}
                        # Add all cell data based on column position, skipping the S.N. column
                        for idx, cell in enumerate(cells):
                            # Skip the S.N. column
                            if idx != sn_index and idx < len(column_names):
                                column_idx = idx if (sn_index is None or idx < sn_index) else idx - 1
                                if column_idx < len(filtered_column_names):
                                    try:
                                        record[filtered_column_names[column_idx]] = cell.text.strip()
                                    except StaleElementReferenceException:
                                        # Handle stale element reference
                                        logger.warning(f"Stale element reference for {company} at row {idx}")
                                        break
                        if record:  # Only append non-empty records
                            data.append(record)
                except StaleElementReferenceException:
                    logger.warning(f"Stale element reference for {company} when processing rows")
                    continue
                except Exception as row_error:
                    logger.warning(f"Error processing row for {company}: {row_error}")
                    continue
            
            # Reverse the order of data so oldest entries come first
            data.reverse()
            
            logger.info(f"{company}: Data scraped successfully with {len(data)} entries.")
            break  # Exit the retry loop on success
            
        except TimeoutException:
            logger.warning(f"Timeout error for {company}. Retry {retry_count + 1}/{max_retries}")
            retry_count += 1
            time.sleep(random.uniform(5, 10))  # Longer delay between retries
            
        except WebDriverException as wde:
            logger.error(f"WebDriver error for {company}: {wde}")
            retry_count += 1
            time.sleep(random.uniform(5, 10))
            
        except Exception as e:
            logger.error(f"Error scraping {company}: {e}")
            traceback.print_exc()
            retry_count += 1
            time.sleep(random.uniform(5, 10))
            
        finally:
            if retry_count >= max_retries:
                logger.error(f"Failed to scrape {company} after {max_retries} attempts.")
            
            if driver:
                try:
                    driver.quit()
                except Exception as e:
                    logger.warning(f"Error closing driver for {company}: {e}")
            
            # Random delay to mimic human browsing behavior
            time.sleep(random.uniform(2, 5))
    
    return company, data

def save_data_in_batches(all_data, batch_size=50):
    """Save data in batches to prevent data loss if an error occurs in the middle."""
    total_companies = len(all_data)
    
    # Create a directory for the output files if it doesn't exist
    output_dir = "share_data_output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Get current timestamp for batch files
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Clear out old batch files before creating new ones
    for existing_file in os.listdir(output_dir):
        if existing_file.startswith("shares_data_batch_") and existing_file.endswith(".xlsx"):
            try:
                os.remove(os.path.join(output_dir, existing_file))
                logger.info(f"Removed old batch file: {existing_file}")
            except Exception as e:
                logger.warning(f"Failed to remove old batch file {existing_file}: {e}")
    
    # Process in batches
    for batch_start in range(0, total_companies, batch_size):
        batch_end = min(batch_start + batch_size, total_companies)
        batch_companies = list(all_data.keys())[batch_start:batch_end]
        
        batch_data = {company: all_data[company] for company in batch_companies}
        
        # Save this batch to an Excel file
        batch_file = f"{output_dir}/shares_data_batch_{batch_start+1}_to_{batch_end}_{timestamp}.xlsx"
        
        try:
            with pd.ExcelWriter(batch_file, engine='openpyxl') as writer:
                for company, records in batch_data.items():
                    try:
                        # Sanitize sheet name to avoid Excel errors
                        sheet_name = sanitize_sheet_name(company)
                        
                        if records:
                            df = pd.DataFrame(records)
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        else:
                            # Create an empty sheet
                            pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                            
                        logger.info(f"Saved data for {company} to sheet '{sheet_name}'")
                        
                    except Exception as sheet_error:
                        logger.error(f"Error saving {company} data: {sheet_error}")
                        # Continue with next company if one fails
                        continue
            
            logger.info(f"Successfully saved batch {batch_start+1} to {batch_end} to {batch_file}")
            
        except Exception as batch_error:
            logger.error(f"Error saving batch {batch_start+1} to {batch_end}: {batch_error}")
    
    # After all batches are processed, we don't need to create individual CSV files anymore
    # as we're handling the consolidated file separately
    logger.info(f"All batch data has been processed.")

def save_consolidated_excel(all_data):
    """Save all company data to a single Excel file with the current date."""
    # Path is now one directory up, then into "2. append_to_existing_data"
    output_dir = os.path.join(os.path.dirname(os.getcwd()), "2. append_to_existing_data")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Get current date in yyyy_mm_dd format
    current_date = datetime.now().strftime("%Y_%m_%d")
    file_name = f"shares_data_till_{current_date}.xlsx"
    consolidated_file = os.path.join(output_dir, file_name)
    
    # Check if any previous share_data files exist and remove them
    for existing_file in os.listdir(output_dir):
        if existing_file.startswith("shares_data_till_") and existing_file.endswith(".xlsx"):
            try:
                os.remove(os.path.join(output_dir, existing_file))
                logger.info(f"Removed existing file: {existing_file}")
            except Exception as e:
                logger.error(f"Failed to remove existing file {existing_file}: {e}")
    
    logger.info(f"Creating consolidated Excel file: {consolidated_file}")
    
    try:
        with pd.ExcelWriter(consolidated_file, engine='openpyxl') as writer:
            for company, records in all_data.items():
                try:
                    # Sanitize sheet name to avoid Excel errors
                    sheet_name = sanitize_sheet_name(company)
                    
                    if records:
                        df = pd.DataFrame(records)
                        # Sort by date if a date column exists (assuming 'Date' is the column name)
                        date_columns = [col for col in df.columns if 'date' in col.lower()]
                        if date_columns:
                            # Use the first column that contains 'date'
                            date_col = date_columns[0]
                            # Convert to datetime, handling various formats
                            try:
                                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                                # Sort by date in ascending order (oldest first)
                                df = df.sort_values(by=date_col)
                            except Exception as date_error:
                                logger.warning(f"Could not sort {company} by date: {date_error}")
                        
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"Added {company} with {len(records)} records to consolidated file")
                    else:
                        # Create an empty sheet
                        pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.info(f"Added empty sheet for {company} to consolidated file")
                        
                except Exception as sheet_error:
                    logger.error(f"Error adding {company} to consolidated file: {sheet_error}")
                    # Continue with next company if one fails
                    continue
                
        logger.info(f"Successfully created consolidated Excel file: {consolidated_file}")
        return consolidated_file
        
    except Exception as e:
        logger.error(f"Error creating consolidated Excel file: {e}")
        return None

def main():
    logger.info("Starting share data scraping process")
    
    # Dictionary to store data for each company
    all_data = {}
    processed_count = 0
    skipped_companies = []
    failed_companies = []
    
    # Save progress file path
    progress_file = "scraping_progress.txt"
    
    # Check if we have a progress file to resume from
    processed_companies = set()
    if os.path.exists(progress_file):
        with open(progress_file, 'r') as f:
            processed_companies = set(line.strip() for line in f.readlines())
        logger.info(f"Resuming from progress file. {len(processed_companies)} companies already processed.")
    
    # Process each company sequentially
    try:
        for i, company in enumerate(companies):
            # Skip if already processed
            if company in processed_companies:
                logger.info(f"Skipping already processed company: {company}")
                skipped_companies.append(company)
                continue
                
            logger.info(f"Processing {company}... ({i+1}/{len(companies)})")
            
            try:
                company_name, data = scrape_share_data(company)
                all_data[company_name] = data
                processed_count += 1
                
                # Save progress after each successful scrape
                with open(progress_file, 'a') as f:
                    f.write(f"{company}\n")
                    
                # Save data in batches of 10 companies to prevent data loss
                if processed_count % 10 == 0:
                    logger.info(f"Saving intermediate batch after {processed_count} companies")
                    current_batch = {k: all_data[k] for k in list(all_data.keys())[-10:]}
                    save_data_in_batches(current_batch, batch_size=10)
                    
            except Exception as company_error:
                logger.error(f"Fatal error processing {company}: {company_error}")
                failed_companies.append(company)
                continue
                
    except KeyboardInterrupt:
        logger.warning("Process interrupted by user. Saving current progress...")
    except Exception as e:
        logger.error(f"Unexpected error in main process: {e}")
        traceback.print_exc()
    finally:
        # Save all the scraped data in batches
        if all_data:
            logger.info(f"Saving all collected data for {len(all_data)} companies")
            save_data_in_batches(all_data)
            
            # Create the consolidated Excel file with all company data
            consolidated_file = save_consolidated_excel(all_data)
            if consolidated_file:
                logger.info(f"Consolidated data saved to: {consolidated_file}")
            
        # Generate a summary report
        logger.info("\n----- SCRAPING SUMMARY -----")
        logger.info(f"Total companies: {len(companies)}")
        logger.info(f"Successfully processed: {processed_count}")
        logger.info(f"Skipped (already processed): {len(skipped_companies)}")
        logger.info(f"Failed: {len(failed_companies)}")
        
        if failed_companies:
            logger.info("Failed companies:")
            for company in failed_companies:
                logger.info(f"  - {company}")
                
        logger.info("---------------------------")
        logger.info("Scraping process completed")

if __name__ == "__main__":
    main()