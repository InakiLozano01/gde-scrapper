import os
import logging
import time
import csv
import asyncio
import re
import sys
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def wait_and_find_element(driver, by, value, timeout=15):
    """Wait for an element to be present and return it."""
    try:
        element = WebDriverWait(driver, timeout, poll_frequency=0.5).until(
            EC.presence_of_element_located((by, value))
        )
        time.sleep(1)
        return element
    except TimeoutException:
        logger.error(f"Element not found: {value}")
        raise

def wait_for_downloads(directory, timeout=60):
    """Wait for downloads to complete with frequent polling."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        downloading_files = [f for f in os.listdir(directory) if f.endswith('.crdownload') or f.endswith('.tmp')]
        if not downloading_files:
            return True
        time.sleep(0.2)
    return False

def get_number_of_docs():
    """Prompt the user to enter the number of documents to process."""
    logger.info("Starting document input process")
    print("Enter number of documents to process (e.g., 35): ", end='', flush=True)
    try:
        num = input().strip()
        logger.info(f"Received input: {num}")
        return int(num)
    except ValueError as e:
        logger.error(f"Invalid input: {e}")
        raise

def go_to_next_page(driver):
    """Attempt to click the next page button."""
    try:
        for attempt in range(3):
            try:
                # Use a CSS selector that selects the next button which is not disabled.
                next_button = wait_and_find_element(
                    driver,
                    By.CSS_SELECTOR,
                    "a.z-paging-button.z-paging-next:not([disabled])",
                    timeout=5
                )
                if next_button.get_attribute("disabled") == "true":
                    return False
                try:
                    driver.execute_script("arguments[0].click();", next_button)
                except Exception:
                    try:
                        ActionChains(driver).move_to_element(next_button).click().perform()
                    except Exception:
                        next_button.click()
                WebDriverWait(driver, 10, poll_frequency=0.2).until(
                    lambda d: len(d.find_elements(By.XPATH, "//tr[contains(@class, 'z-listitem')]")) > 0
                )
                WebDriverWait(driver, 10, poll_frequency=0.2).until(
                    lambda d: not any(
                        'z-loading' in cls and 'display: none' not in indicator.get_attribute('style')
                        for indicator in d.find_elements(By.CLASS_NAME, "z-loading")
                        for cls in indicator.get_attribute("class").split()
                    )
                )
                time.sleep(0.5)
                return True
            except Exception as e:
                if attempt == 2:
                    logger.error(f"Failed to go to next page after retries: {str(e)}")
                    return False
                time.sleep(1)
        return False
    except Exception as e:
        logger.error(f"Error navigating to next page: {str(e)}")
        return False

def wait_for_all_downloads(directory, timeout=300):
    """Wait for all downloads to complete with a longer timeout."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        downloading_files = [f for f in os.listdir(directory) if f.endswith('.crdownload') or f.endswith('.tmp')]
        if not downloading_files:
            time.sleep(5)
            return True
        logger.info(f"Still waiting for downloads: {downloading_files}")
        time.sleep(5)
    return False

def get_expediente_number(row):
    """Extract the expediente number from the third cell of a row."""
    try:
        expediente_cell = row.find_element(By.XPATH, ".//td[contains(@class, 'z-listcell')][3]")
        expediente_text = expediente_cell.get_attribute("title")
        return expediente_text.strip()
    except Exception as e:
        logger.error(f"Error getting expediente number: {str(e)}")
        return None

def handle_login(driver):
    """Handle login with retries and error handling."""
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            username_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[placeholder='Usuario/Cuil/Cuit']"))
            )
            time.sleep(1)
            username_field.clear()
            user = os.getenv("USERNAME")
            username_field.send_keys(user)
            logger.info("Username entered")
            password_field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
            )
            time.sleep(1)
            password_field.clear()
            password = os.getenv("PASSWORD")
            password_field.send_keys(password)
            logger.info("Password entered")
            login_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Acceder')]"))
            )
            time.sleep(1)
            login_button.click()
            logger.info("Login button clicked")
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".glyphicon-th"))
            )
            return True
        except Exception as e:
            logger.error(f"Login attempt {retry_count + 1} failed: {str(e)}")
            retry_count += 1
            if retry_count < max_retries:
                logger.info("Refreshing page and retrying login...")
                driver.refresh()
                time.sleep(3)
            else:
                raise Exception("Failed to login after multiple attempts")

def wait_for_downloads_to_complete(driver, directory, timeout=60):
    """Wait for both the browser loading indicators and file downloads to complete."""
    try:
        start_time = time.time()
        while time.time() - start_time < timeout:
            loading_indicators = driver.find_elements(By.CLASS_NAME, "z-loading")
            browser_loading = any(
                not 'display: none' in indicator.get_attribute('style')
                for indicator in loading_indicators if indicator.is_displayed()
            )
            downloading_files = [f for f in os.listdir(directory) if f.endswith('.crdownload') or f.endswith('.tmp')]
            if not browser_loading and not downloading_files:
                time.sleep(0.5)
                return True
            time.sleep(0.2)
        return False
    except Exception as e:
        logger.error(f"Error checking download status: {str(e)}")
        return False

def verify_download_completed(directory, filename_pattern, timeout=60):
    """Verify that a file matching the given pattern has been downloaded."""
    try:
        def _check_file():
            files = [f for f in os.listdir(directory) if filename_pattern in f and not f.endswith('.crdownload') and not f.endswith('.tmp')]
            return len(files) > 0
        wait = WebDriverWait(None, timeout, poll_frequency=0.5)
        return wait.until(lambda _: _check_file())
    except TimeoutException:
        return False

def wait_for_element_with_retry(driver, by, value, condition_type="presence", timeout=30, retries=3):
    """Wait for an element with more frequent polling."""
    for attempt in range(retries):
        try:
            wait = WebDriverWait(driver, timeout / retries, poll_frequency=0.2)
            if condition_type == "presence":
                element = wait.until(EC.presence_of_element_located((by, value)))
            elif condition_type == "clickable":
                element = wait.until(EC.element_to_be_clickable((by, value)))
            elif condition_type == "visible":
                element = wait.until(EC.visibility_of_element_located((by, value)))
            WebDriverWait(driver, 2, poll_frequency=0.2).until(
                lambda d: not any(
                    'z-loading' in cls and 'display: none' not in indicator.get_attribute('style')
                    for indicator in d.find_elements(By.CLASS_NAME, "z-loading")
                    for cls in indicator.get_attribute("class").split()
                )
            )
            return element
        except Exception as e:
            if attempt == retries - 1:
                raise
            logger.info(f"Retry {attempt + 1}/{retries} for element {value}")
            time.sleep(0.5)
    raise TimeoutException(f"Element {value} not found after {retries} retries")

def safe_click_with_retry(driver, element, retries=3):
    """Safely click an element with retries."""
    for attempt in range(retries):
        try:
            WebDriverWait(driver, 5, poll_frequency=0.5).until(
                lambda d: not any(
                    'z-loading' in cls and 'display: none' not in indicator.get_attribute('style')
                    for indicator in d.find_elements(By.CLASS_NAME, "z-loading")
                    for cls in indicator.get_attribute("class").split()
                )
            )
            try:
                driver.execute_script("arguments[0].click();", element)
            except Exception:
                try:
                    ActionChains(driver).move_to_element(element).click().perform()
                except Exception:
                    element.click()
            return True
        except Exception as e:
            if attempt == retries - 1:
                raise
            logger.info(f"Click retry {attempt + 1}/{retries}")
            time.sleep(1)
    return False

def wait_for_download_with_verification(driver, directory, expediente, timeout=60):
    """Wait for download with verification that the expected file appears."""
    try:
        WebDriverWait(driver, timeout/2, poll_frequency=0.5).until(
            lambda d: not any(
                'z-loading' in cls and 'display: none' not in indicator.get_attribute('style')
                for indicator in d.find_elements(By.CLASS_NAME, "z-loading")
                for cls in indicator.get_attribute("class").split()
            )
        )
        filename_pattern = f"Documentos-{expediente}"
        return verify_download_completed(directory, filename_pattern, timeout=timeout/2)
    except Exception as e:
        logger.error(f"Error verifying download: {str(e)}")
        return False

def handle_modal_download(driver, expediente, downloads_dir):
    """
    Handle the modal download process for a given expediente.
    Returns a tuple (downloaded, error_message).
    """
    for attempt in range(3):
        try:
            modal = WebDriverWait(driver, 15, poll_frequency=0.2).until(
                EC.presence_of_element_located((By.CLASS_NAME, "z-window-modal"))
            )
            download_button = WebDriverWait(modal, 5, poll_frequency=0.2).until(
                EC.element_to_be_clickable((By.XPATH, ".//button[contains(text(), 'Descargar todos los Documentos')]"))
            )
            try:
                driver.execute_script("arguments[0].click();", download_button)
            except Exception:
                try:
                    ActionChains(driver).move_to_element(download_button).click().perform()
                except Exception:
                    download_button.click()
            logger.info(f"Starting download for expediente: {expediente}")
            download_start_time = time.time()
            last_file_check = time.time()
            download_logged = False
            modal_closed = False
            while time.time() - download_start_time < 60:
                current_time = time.time()
                if current_time - last_file_check >= 0.2:
                    downloading_files = [f for f in os.listdir(downloads_dir) if f.endswith('.crdownload') or f.endswith('.tmp')]
                    if not downloading_files and not download_logged:
                        logger.info(f"Download completed for expediente: {expediente}")
                        download_logged = True
                    last_file_check = current_time
                if download_logged and not modal_closed:
                    try:
                        close_button = WebDriverWait(driver, 2, poll_frequency=0.2).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, ".z-window-modal .z-icon-times"))
                        )
                        driver.execute_script("arguments[0].click();", close_button)
                        WebDriverWait(driver, 2, poll_frequency=0.2).until(
                            EC.invisibility_of_element_located((By.CLASS_NAME, "z-window-modal"))
                        )
                        modal_closed = True
                        break
                    except Exception:
                        pass
                time.sleep(0.1)
            if not download_logged:
                raise TimeoutException("Download timeout")
            if not modal_closed:
                try:
                    close_button = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".z-window-modal .z-icon-times"))
                    )
                    driver.execute_script("arguments[0].click();", close_button)
                    modal_closed = True
                except Exception:
                    pass
            return True, None
        except Exception as e:
            if attempt == 2:
                error_msg = f"Error in modal handling: {str(e)}"
                logger.error(f"Error in modal handling for expediente {expediente}: {str(e)}")
                return False, error_msg
            logger.warning(f"Retry {attempt + 1} for modal handling: {str(e)}")
            time.sleep(1)
    return False, "Maximum retries exceeded"

def click_visualizar_option(driver, row, expediente, retries=3):
    """
    Attempt to click the combobox button in the given row and then select the
    "Visualizar" option from the resulting popup.
    """
    for attempt in range(retries):
        try:
            # Re-find the combobox button
            combo_btn = row.find_element(By.CSS_SELECTOR, "a.z-combobox-button")
            driver.execute_script("arguments[0].click();", combo_btn)
            time.sleep(0.5)
            # Wait for the popup option "Visualizar" to become clickable.
            visualizar_option = WebDriverWait(driver, 30, poll_frequency=0.2).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(@class, 'z-combobox-popup') and not(contains(@style, 'display: none'))]//li[contains(@class, 'z-comboitem')]//span[contains(text(), 'Visualizar')]")
                )
            )
            driver.execute_script("arguments[0].click();", visualizar_option)
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.warning(f"Retry {attempt+1}/{retries} for combobox selection on expediente {expediente}: {str(e)}")
            time.sleep(1)
    raise Exception(f"Failed to select 'Visualizar' option for expediente {expediente} after {retries} retries")

def setup_chrome_options():
    """Set up Chrome options for the webdriver."""
    chrome_options = Options()
    
    # Basic Chrome settings
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--headless=new')  # Using the new headless mode
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--start-maximized')
    
    # Additional options for reliability
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-web-security')
    chrome_options.add_argument('--allow-running-insecure-content')
    
    # Set user agent to look more like a regular browser
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36')
    
    # Experimental options
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Set download preferences
    download_dir = os.getenv('DOWNLOAD_DIR', '/app/downloads')
    prefs = {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'profile.default_content_setting_values.automatic_downloads': 1,
        'profile.content_settings.exceptions.automatic_downloads.*.setting': 1,
        'credentials_enable_service': False,
        'profile.password_manager_enabled': False
    }
    chrome_options.add_experimental_option('prefs', prefs)
    
    logger.info(f"Chrome options configured with download directory: {download_dir}")
    return chrome_options

def initialize_driver():
    """Initialize and return the Chrome webdriver."""
    try:
        chrome_options = setup_chrome_options()
        service = ChromeService(executable_path="/usr/local/bin/chromedriver")
        
        logger.info("Initializing Chrome driver...")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(30)
        logger.info("Chrome driver initialized successfully")
        
        return driver
    except Exception as e:
        logger.error(f"Failed to initialize Chrome driver: {str(e)}")
        raise

async def async_main():
    try:
        # Load environment variables
        load_dotenv()
        url = os.getenv("URL")
        if not url:
            raise ValueError("URL environment variable is not set")
        if not url.startswith("http"):
            raise ValueError(f"Invalid URL format: {url}")
        
        logger.info(f"Using URL: {url}")
        
        # Create data directory if it doesn't exist
        data_dir = os.path.join(os.getcwd(), "data")
        os.makedirs(data_dir, exist_ok=True)
        
        # Create CSV file for logging processed expedientes
        csv_file = os.path.join(data_dir, "expedientes.csv")
        with open(csv_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Expediente', 'Downloaded', 'Error'])

        num_docs = get_number_of_docs()
        logger.info(f"Will process {num_docs} documents")

        downloads_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(downloads_dir, exist_ok=True)
        logger.info(f"Downloads directory set to: {downloads_dir}")

        # Initialize the driver using our helper function
        driver = initialize_driver()
        
        try:
            # Login and navigation
            logger.info("Navigating to login page...")
            driver.get(url)
            await asyncio.to_thread(handle_login, driver)

            modules_button = await asyncio.to_thread(wait_and_find_element, driver, By.CSS_SELECTOR, ".glyphicon-th")
            await asyncio.to_thread(driver.execute_script, "arguments[0].click();", modules_button)
            logger.info("Modules button clicked")

            ee_button = await asyncio.to_thread(wait_and_find_element, driver, By.CSS_SELECTOR, ".z-icon-copy")
            await asyncio.to_thread(driver.execute_script, "arguments[0].click();", ee_button)
            logger.info("EE button clicked")

            consultas_tab = await asyncio.to_thread(wait_and_find_element, driver, By.XPATH, "//span[contains(text(),'Consultas')]")
            await asyncio.to_thread(driver.execute_script, "arguments[0].click();", consultas_tab)
            logger.info("Consultas tab clicked")
            await asyncio.sleep(2)

            jurisdiccion_radio = await asyncio.to_thread(wait_and_find_element, driver, By.XPATH, 
                "//span[@class='z-radio']//input[@type='radio'][following-sibling::label[contains(text(), 'Tramitados por mi Jurisdicción')]]")
            await asyncio.to_thread(driver.execute_script, "arguments[0].click();", jurisdiccion_radio)
            logger.info("Jurisdiccion filter applied")
            await asyncio.sleep(2)

            # Apply the "Guarda Temporal" filter (multiple strategies)
            try:
                guarda_label = await asyncio.to_thread(wait_and_find_element, driver, By.XPATH, "//span[contains(text(), ' Guarda Temporal')]")
                checkbox = guarda_label.find_element(By.XPATH, "..//input[@type='checkbox']")
                await asyncio.to_thread(driver.execute_script, "arguments[0].click();", checkbox)
            except Exception:
                try:
                    checkbox = await asyncio.to_thread(wait_and_find_element, driver, By.XPATH, "//div[contains(.//span/text(), ' Guarda Temporal')]//input[@type='checkbox']")
                    await asyncio.to_thread(driver.execute_script, "arguments[0].click();", checkbox)
                except Exception:
                    guarda_label = await asyncio.to_thread(wait_and_find_element, driver, By.XPATH, "//label[following-sibling::span[contains(text(), ' Guarda Temporal')]]")
                    await asyncio.to_thread(driver.execute_script, "arguments[0].click();", guarda_label)
            logger.info("Guarda Temporal filter applied")
            await asyncio.sleep(3)

            # Perform search
            logger.info("Looking for search button...")
            search_buttons = driver.find_elements(By.XPATH, "//button[.//i[contains(@class, 'z-icon-search')]]")
            if not search_buttons:
                raise Exception("Search button not found")
            search_button = search_buttons[-1]
            await asyncio.to_thread(driver.execute_script, "arguments[0].scrollIntoView(true);", search_button)
            await asyncio.sleep(1)
            await asyncio.to_thread(driver.execute_script, "arguments[0].click();", search_button)
            logger.info("Search button clicked")
            await asyncio.sleep(5)  # wait for the results to load

            docs_processed = 0
            current_page = 1

            # Main processing loop: while there are still documents to process
            while docs_processed < num_docs:
                try:
                    rows = await asyncio.to_thread(
                        lambda: WebDriverWait(driver, 10, poll_frequency=0.2).until(
                            EC.presence_of_all_elements_located((By.XPATH, "//tr[contains(@class, 'z-listitem')]"))
                        )
                    )
                except Exception as e:
                    logger.error(f"No rows found on page {current_page}: {str(e)}")
                    break

                rows_count = len(rows)
                if rows_count == 0:
                    logger.info(f"No results on page {current_page}. Ending processing.")
                    break

                remaining_docs = num_docs - docs_processed
                docs_this_page = min(remaining_docs, rows_count)
                logger.info(f"Processing {docs_this_page} documents on page {current_page}")

                for doc_num in range(1, docs_this_page + 1):
                    logger.info(f"Processing document {doc_num} on page {current_page} (total processed: {docs_processed + 1})")
                    try:
                        row = await asyncio.to_thread(wait_and_find_element, driver, By.XPATH,
                                                      f"//div[contains(@class, 'z-listbox-body')]//tr[contains(@class, 'z-listitem')][{doc_num}]")
                        expediente = get_expediente_number(row)
                        downloaded = False
                        error_message = None

                        if expediente:
                            logger.info(f"Processing expediente: {expediente}")
                            try:
                                # Use the helper function with retry to click the combobox and select "Visualizar"
                                await asyncio.to_thread(click_visualizar_option, driver, row, expediente)
                                await asyncio.sleep(0.5)
                                # Handle the modal download for the expediente.
                                result = await asyncio.to_thread(handle_modal_download, driver, expediente, downloads_dir)
                                downloaded, error_message = result
                            except Exception as e:
                                error_message = str(e)
                                logger.error(f"Error processing expediente {expediente}: {str(e)}")

                            with open(csv_file, 'a', newline='') as f:
                                writer = csv.writer(f)
                                writer.writerow([expediente, downloaded, error_message if not downloaded else ""])
                            await asyncio.sleep(1)

                        docs_processed += 1
                        if docs_processed >= num_docs:
                            break
                    except Exception as e:
                        logger.error(f"Error processing document {doc_num} on page {current_page}: {str(e)}")
                        continue

                if docs_processed < num_docs:
                    try:
                        next_button = await asyncio.to_thread(wait_and_find_element, driver, By.CSS_SELECTOR, "a.z-paging-button.z-paging-next:not([disabled])")
                        await asyncio.to_thread(driver.execute_script, "arguments[0].click();", next_button)
                        current_page += 1
                        logger.info(f"Moved to page {current_page}")
                        await asyncio.sleep(3)
                    except Exception as e:
                        logger.error(f"Failed to move to next page: {str(e)}")
                        break

            logger.info(f"Processed {docs_processed} documents in total")
            logger.info("Waiting for all downloads to complete...")
            downloads_complete = await asyncio.to_thread(wait_for_all_downloads, downloads_dir)
            if downloads_complete:
                logger.info("All downloads completed successfully")
            else:
                logger.warning("Download wait timeout - some downloads may not have completed")
        finally:
            await asyncio.to_thread(driver.quit)

        files = os.listdir(downloads_dir)
        logger.info(f"Files in downloads directory: {files}")
        print("\nAutomation completed. Please check your downloads folder for files.")

    except Exception as e:
        logger.error(f"An error occurred: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    asyncio.run(async_main())
