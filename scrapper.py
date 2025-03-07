import os
import logging
import time
import csv
import asyncio
import pandas as pd
import requests
import zipfile
import io
import platform
import shutil
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# ---------------------------------------------
# CONFIGURATION / LOGGING
# ---------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def setup_chrome_options():
    """Set up Chrome options for the webdriver."""
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument('--disable-web-security')
    chrome_options.add_argument('--allow-running-insecure-content')
    chrome_options.add_argument(
        '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36'
    )
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Set up downloads directory with absolute path
    download_dir = os.path.abspath("./downloads")
    os.makedirs(download_dir, exist_ok=True)
    logger.info(f"Using downloads directory: {download_dir}")
    
    prefs = {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'profile.default_content_setting_values.automatic_downloads': 1,
        'profile.content_settings.exceptions.automatic_downloads.*.setting': 1,
        'credentials_enable_service': False,
        'profile.password_manager_enabled': False,
        'download.default_directory': download_dir.replace('/', '\\'),  # Windows path format
        'savefile.default_directory': download_dir.replace('/', '\\'),  # Windows path format
    }
    chrome_options.add_experimental_option('prefs', prefs)
    logger.info(f"Chrome options configured with download directory: {download_dir}")
    return chrome_options, download_dir

def get_chrome_driver_path():
    """
    Download the correct Chrome driver using Chrome for Testing API.
    This ensures compatibility with the installed Chrome browser.
    """
    # Create driver directory if it doesn't exist
    driver_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "drivers")
    os.makedirs(driver_dir, exist_ok=True)
    
    driver_path = os.path.join(driver_dir, "chromedriver.exe")
    
    # If driver already exists, return its path
    if os.path.exists(driver_path):
        logger.info(f"Using existing Chrome driver at {driver_path}")
        return driver_path
    
    try:
        # Get the latest Chrome driver info from Chrome for Testing
        logger.info("Getting latest Chrome driver version information...")
        api_url = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json"
        response = requests.get(api_url)
        response.raise_for_status()  # Raise exception for HTTP errors
        driver_data = response.json()
        
        # Get the stable version info
        stable_version = driver_data["channels"]["Stable"]["version"]
        logger.info(f"Latest stable Chrome version: {stable_version}")
        
        # Download the driver
        logger.info(f"Downloading Chrome driver for version {stable_version}...")
        download_url = f"https://storage.googleapis.com/chrome-for-testing-public/{stable_version}/win64/chromedriver-win64.zip"
        response = requests.get(download_url)
        response.raise_for_status()  # Raise exception for HTTP errors
        
        # Extract the zip file
        with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
            # Extract all files from the zip
            zip_file.extractall(driver_dir)
        
        # Move the chromedriver.exe from the extracted folder to drivers directory
        extracted_driver = os.path.join(driver_dir, "chromedriver-win64", "chromedriver.exe")
        if os.path.exists(extracted_driver):
            shutil.copy(extracted_driver, driver_path)
            logger.info(f"Chrome driver downloaded and installed at {driver_path}")
            return driver_path
        else:
            logger.error(f"Could not find chromedriver.exe in extracted files")
            return None
    except Exception as e:
        logger.error(f"Failed to download Chrome driver: {str(e)}")
        return None

def initialize_driver():
    """Initialize and return the Chrome webdriver."""
    try:
        chrome_options, download_dir = setup_chrome_options()
        
        # Try to get manually downloaded driver first
        driver_path = get_chrome_driver_path()
        
        if driver_path and os.path.exists(driver_path):
            service = Service(executable_path=driver_path)
            logger.info(f"Using Chrome driver at: {driver_path}")
        else:
            # This fallback should only happen if our manual download failed completely
            logger.warning("Manual driver download failed, using ChromeDriverManager")
            service = Service(ChromeDriverManager(version="stable").install())
            
        logger.info("Initializing Chrome driver...")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(30)
        logger.info("Chrome driver initialized successfully")
        return driver, download_dir
    except Exception as e:
        logger.error(f"Failed to initialize Chrome driver: {str(e)}")
        raise

# ---------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------
def wait_and_find_element(driver, by, value, timeout=15):
    """Wait for an element to be present and return it."""
    try:
        time.sleep(2)  # Adding a small delay for better visibility
        element = WebDriverWait(driver, timeout, poll_frequency=0.5).until(
            EC.presence_of_element_located((by, value))
        )
        # Scroll element into view
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(0.5)  # Wait for scroll
        return element
    except TimeoutException:
        logger.error(f"Element not found: {value}")
        raise

def type_into_field(driver, element, text):
    """Type text into a field using JavaScript and direct input."""
    try:
        # Clear using JavaScript
        driver.execute_script("arguments[0].value = '';", element)
        time.sleep(0.5)
        
        # Type using JavaScript
        driver.execute_script(f"arguments[0].value = '{text}';", element)
        time.sleep(0.5)
        
        # Trigger input event
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", element)
        time.sleep(0.5)
        
        # Also try direct sendKeys as backup
        element.send_keys(text)
        time.sleep(0.5)
        
        # Verify the text was entered
        actual_value = driver.execute_script("return arguments[0].value;", element)
        if actual_value != text:
            logger.warning(f"Text verification failed. Expected: {text}, Got: {actual_value}")
            # One more attempt with direct sendKeys
            element.clear()
            time.sleep(0.5)
            element.send_keys(text)
            time.sleep(0.5)
    except Exception as e:
        logger.error(f"Error typing into field: {str(e)}")
        raise

def check_for_loading(driver):
    """Wait for any loading indicators to disappear."""
    WebDriverWait(driver, 10, poll_frequency=0.2).until(
        lambda d: all(
            'display: none' in indicator.get_attribute('style')
            for indicator in d.find_elements(By.CLASS_NAME, "z-loading")
            if indicator.is_displayed()
        )
    )

def check_session_expired(driver):
    """Return True if login fields are visible (session expired)."""
    return bool(driver.find_elements(By.CSS_SELECTOR, "[placeholder='Usuario/Cuil/Cuit']"))

def handle_login(driver):
    """Perform login with retries."""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            username_field = wait_and_find_element(driver, By.CSS_SELECTOR, "[placeholder='Usuario/Cuil/Cuit']", timeout=15)
            time.sleep(1)
            username_field.clear()
            user = os.getenv("USERNAME")
            username_field.send_keys(user)
            logger.info("Username entered")

            password_field = wait_and_find_element(driver, By.CSS_SELECTOR, "input[type='password']", timeout=15)
            time.sleep(1)
            password_field.clear()
            password = os.getenv("PASSWORD")
            password_field.send_keys(password)
            logger.info("Password entered")

            login_button = wait_and_find_element(driver, By.XPATH, "//button[contains(.,'Acceder')]", timeout=15)
            time.sleep(1)
            login_button.click()
            logger.info("Login button clicked")
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".glyphicon-th"))
            )
            return
        except Exception as e:
            logger.error(f"Login attempt {attempt+1} failed: {str(e)}")
            if attempt < max_retries - 1:
                driver.refresh()
                time.sleep(3)
            else:
                raise Exception("Failed to login after multiple attempts")

def reapply_navigation(driver):
    """Reapply navigation after page refresh or reset"""
    try:
        logger.info("Attempting to reapply navigation...")
        
        # First check if we're on the login page
        try:
            # Short timeout here because we just want to check quickly
            username_input = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            
            # We found a username field, so we're on the login page
            logger.info("On login page, re-logging in...")
            
            # Re-login if needed
            username = driver.find_element(By.ID, "username")
            password = driver.find_element(By.ID, "password")
            submit = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
            
            # Fill login form
            username.send_keys(os.environ.get("GDE_USERNAME", ""))
            password.send_keys(os.environ.get("GDE_PASSWORD", ""))
            submit.click()
            
            # Wait for login to complete - don't look for specific elements,
            # just wait for the page to load and stabilize
            time.sleep(5)
            logger.info("Successfully logged in after page refresh")
            
        except TimeoutException:
            # Not on login page, which is expected after successful login
            logger.info("Not on login page, continuing navigation...")
        
        # Try to find and click "Consulta de expediente" if it exists
        try:
            # Check for the search link
            search_link = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Consulta de expediente"))
            )
            search_link.click()
            logger.info("Clicked 'Consulta de expediente' link")
            
            # Wait for the search page to load
            time.sleep(3)
            
            # Check if we reached the search page
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "textInput"))
                )
                logger.info("Successfully navigated to search page")
                return True
            except:
                logger.warning("Could not find textInput after clicking 'Consulta de expediente'")
        except:
            logger.info("'Consulta de expediente' link not found, checking for alternative navigation")
        
        # Alternative: If we're already on a page with search functionality
        try:
            input_field = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
            )
            logger.info("Found text input field, may already be on search page")
            return True
        except:
            logger.warning("No text input field found")
        
        # Take a screenshot for debugging
        try:
            screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "current_page.png")
            driver.save_screenshot(screenshot_path)
            logger.info(f"Current page screenshot saved to {screenshot_path}")
        except:
            pass
            
        # Last resort - check various elements that might indicate what page we're on
        page_source = driver.page_source.lower()
        if "expediente" in page_source:
            logger.info("Page contains 'expediente', may be on correct page")
            return True
        elif "error" in page_source or "invalidar" in page_source:
            logger.warning("Page may contain error or session invalidation message")
            # Try to refresh the page
            driver.refresh()
            time.sleep(5)
        
        logger.info("Navigation reapplied with fallback approach")
        return True  # We'll continue with the process even if navigation is uncertain
            
    except Exception as e:
        logger.error(f"Error in reapply_navigation: {e}")
        # Take a screenshot to debug
        try:
            screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "navigation_error.png")
            driver.save_screenshot(screenshot_path)
            logger.info(f"Screenshot saved to {screenshot_path}")
        except:
            pass
        # Don't raise the exception - let the script continue
        return False

def check_and_relogin(driver):
    """
    If session expired, perform login and always reapply navigation.
    """
    if check_session_expired(driver):
        logger.info("Session expired detected. Re-logging in.")
        handle_login(driver)
        if not reapply_navigation(driver):
            logger.warning(f"Failed to reapply navigation after session expired for expediente {expediente}")
            # Try to continue anyway
        return True
    return False

def wait_for_single_result(driver, timeout=10):
    """
    Wait for exactly one row in the results table.
    """
    WebDriverWait(driver, timeout, poll_frequency=0.2).until(
        lambda d: len(d.find_elements(By.XPATH, "//tr[contains(@class, 'z-listitem')]")) == 1
    )

def click_visualizar_option(driver, expediente, max_retries=3):
    """
    Click the combobox button in the result row and select "Visualizar."
    """
    for attempt in range(max_retries):
        try:
            check_for_loading(driver)
            row = driver.find_element(By.XPATH, "//tr[contains(@class, 'z-listitem')]")
            combo_btn = row.find_element(By.CSS_SELECTOR, "a.z-combobox-button")
            driver.execute_script("arguments[0].scrollIntoView(true);", combo_btn)
            driver.execute_script("arguments[0].click();", combo_btn)
            time.sleep(0.7)
            visualizar_option = WebDriverWait(driver, 20, poll_frequency=0.2).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(@class, 'z-combobox-popup') and not(contains(@style, 'display: none'))]"
                               "//li[contains(@class, 'z-comboitem')]//span[contains(text(), 'Visualizar')]"))
            )
            driver.execute_script("arguments[0].click();", visualizar_option)
            time.sleep(0.7)
            return
        except StaleElementReferenceException as e:
            logger.warning(f"Stale element on combo for {expediente}, attempt {attempt+1}: {str(e)}")
            time.sleep(1)
        except Exception as e:
            logger.warning(f"Combo box attempt {attempt+1}/{max_retries} for {expediente} failed: {str(e)}")
            time.sleep(1)
    raise Exception(f"Failed to select 'Visualizar' for expediente {expediente} after {max_retries} retries")

async def wait_for_download_with_verification(driver, download_dir, expediente, timeout=90):
    """
    Wait for a download to complete and verify the downloaded file.
    Returns (True, None) on success or (False, error_message) on failure.
    """
    normalized_expediente = "".join(expediente.split())
    start_time = time.time()
    completed_time = None
    
    try:
        # Get initial file list
        initial_files = set(os.listdir(download_dir))
        
        # Set up polling interval
        poll_interval = 1.0  # seconds
        
        # Wait for download to complete
        while time.time() - start_time < timeout:
            # Wait for a moment
            await asyncio.sleep(poll_interval)
            
            try:
                # Get current file list
                current_files = set(os.listdir(download_dir))
                new_files = current_files - initial_files
                
                # Check for temp download files
                temp_files = [f for f in new_files if f.endswith('.crdownload') or f.endswith('.tmp')]
                
                # Check for files matching the expediente
                matching_files = []
                for f in new_files:
                    # Skip temporary files
                    if f.endswith('.crdownload') or f.endswith('.tmp'):
                        continue
                        
                    # Process filename for comparison
                    file_for_comparison = f
                    
                    # Handle URL-encoded characters
                    normalized_exp = normalized_expediente.replace('#', '%')
                    
                    # Try multiple matching methods
                    if (normalized_exp in "".join(file_for_comparison.split()) or
                        normalized_exp.replace('%', '') in "".join(file_for_comparison.split()) or
                        normalized_expediente in "".join(file_for_comparison.split())):
                        matching_files.append(f)
                        
                    # Also look for a match with spaces preserved
                    elif (expediente.replace('#', '%') in file_for_comparison or
                          expediente in file_for_comparison):
                        matching_files.append(f)
                
                if matching_files:
                    # If we found a matching file and it's not temporary
                    if not completed_time:
                        completed_time = time.time()
                        logger.info(f"Download appears complete, waiting to confirm: {matching_files}")
                    elif time.time() - completed_time >= 2:
                        logger.info(f"Download confirmed complete: {matching_files}")
                        return True, None
                else:
                    # Reset completed_time if we detect in-progress downloads
                    completed_time = None
                    
                # Log status periodically
                elapsed = time.time() - start_time
                if elapsed % 5 < poll_interval:
                    if temp_files:
                        logger.info(f"Download in progress... ({int(elapsed)}/{timeout}s)")
                    elif new_files:
                        logger.info(f"New files found but none match expediente: {new_files}")
                    else:
                        logger.info(f"Waiting for download to start... ({int(elapsed)}/{timeout}s)")
                        
                # Check for error messages in the UI
                try:
                    error_elements = driver.find_elements(By.CLASS_NAME, "z-messagebox-error")
                    for element in error_elements:
                        if element.is_displayed():
                            error_text = element.text
                            logger.warning(f"Error message displayed: {error_text}")
                            # Try to dismiss error message
                            dismiss_buttons = driver.find_elements(By.CSS_SELECTOR, ".z-messagebox-button")
                            for button in dismiss_buttons:
                                if button.is_displayed():
                                    button.click()
                                    logger.info("Dismissed error message")
                            return False, f"Error message: {error_text}"
                except Exception as e:
                    pass  # Ignore errors checking for error messages
                    
            except Exception as e:
                logger.warning(f"Error checking download status: {str(e)}")
        
        # Timeout reached, check one more time for matching files
        try:
            current_files = set(os.listdir(download_dir))
            new_files = current_files - initial_files
            
            # Use same enhanced matching as above
            matching_files = []
            for f in new_files:
                # Skip temporary files
                if f.endswith('.crdownload') or f.endswith('.tmp'):
                    continue
                    
                # Process filename for comparison
                file_for_comparison = f
                
                # Handle URL-encoded characters
                normalized_exp = normalized_expediente.replace('#', '%')
                
                # Try multiple matching methods
                if (normalized_exp in "".join(file_for_comparison.split()) or
                    normalized_exp.replace('%', '') in "".join(file_for_comparison.split()) or
                    normalized_expediente in "".join(file_for_comparison.split())):
                    matching_files.append(f)
                    
                # Also look for a match with spaces preserved
                elif (expediente.replace('#', '%') in file_for_comparison or
                      expediente in file_for_comparison):
                    matching_files.append(f)
            
            if matching_files:
                logger.info(f"Found matching files after timeout: {matching_files}")
                return True, None
        except Exception as e:
            logger.error(f"Error logging final directory state: {str(e)}")
        
        return False, "Download timed out"
    except Exception as e:
        logger.error(f"Error in wait_for_download: {str(e)}")
        return False, f"Download error: {str(e)}"

async def handle_modal_download(driver, expediente, download_dir):
    """
    In the modal, click the download button, wait for the file,
    then close the modal. Returns (True, None) on success or (False, error_message) on failure.
    """
    # Check if file already exists before downloading using normalized comparison
    normalized_expediente = "".join(expediente.split())
    existing_files = [f for f in os.listdir(download_dir) if normalized_expediente in "".join(f.split()) and not (f.endswith('.crdownload') or f.endswith('.tmp'))]
    if existing_files:
        logger.info(f"File for expediente {expediente} already exists: {existing_files}. Closing modal without downloading again.")
        # Forcefully remove the modal element from the DOM using modern JavaScript
        driver.execute_script("document.querySelector('.z-window-modal')?.remove();")
        logger.info("Modal removed from DOM because file already exists.")
        time.sleep(1)
        return True, None

    max_modal_attempts = 3
    for attempt in range(max_modal_attempts):
        try:
            # Wait for modal and download button
            modal = WebDriverWait(driver, 15, poll_frequency=0.2).until(
                EC.presence_of_element_located((By.CLASS_NAME, "z-window-modal"))
            )
            download_button = WebDriverWait(modal, 5, poll_frequency=0.2).until(
                EC.element_to_be_clickable((By.XPATH, ".//button[contains(text(), 'Descargar todos los Documentos')]"))
            )
            
            # Ensure button is in view and click
            driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", download_button)
            logger.info(f"Starting download for expediente: {expediente}")
            
            # Wait for download with verification
            logger.info(f"Waiting for download to complete for {expediente}")
            download_success, error = await wait_for_download_with_verification(driver, download_dir, expediente)
            
            if not download_success:
                logger.warning(f"Download failed: {error}")
                # Try next attempt
                continue
            
            logger.info(f"Download successful for expediente {expediente}, closing modal")
            
            # Close modal after successful download
            try:
                # First try to click the close button using reliable class selectors
                try:
                    # Use the most reliable class-based selectors - avoid any ID-based selectors
                    close_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".z-window-modal .z-window-icon.z-window-close"))
                    )
                    # Try JavaScript click which is more reliable
                    driver.execute_script("arguments[0].click();", close_button)
                    logger.info("Modal close button clicked with JavaScript")
                    time.sleep(1)
                except Exception as e:
                    logger.warning(f"Could not click modal close button: {e}")
                    
                    # Try clicking the i tag inside the close button
                    try:
                        close_icon = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, ".z-window-modal .z-icon-times"))
                        )
                        driver.execute_script("arguments[0].click();", close_icon)
                        logger.info("Modal close icon clicked with JavaScript")
                        time.sleep(1)
                    except Exception as e2:
                        logger.warning(f"Could not click close icon: {e2}")
                
                # Check if modal is still present
                try:
                    modals = driver.find_elements(By.CLASS_NAME, "z-window-modal")
                    if modals:
                        logger.warning(f"Modal still present after clicking close button, using DOM removal")
                        # Use JS to remove all modals
                        driver.execute_script("""
                            // Remove all modal-related elements
                            document.querySelectorAll('.z-window-modal').forEach(el => el.parentNode.removeChild(el));
                            document.querySelectorAll('.z-modal-mask').forEach(el => el.parentNode.removeChild(el));
                            
                            // For ZK framework, try more aggressive cleanup
                            if (window.zk) {
                                try {
                                    zk.Widget.$(document.body).children.forEach(function(w) {
                                        if (w.$instanceof(zk.wnd.Window) && w.isVisible())
                                            w.close();
                                    });
                                } catch(e) {
                                    console.error('ZK cleanup error:', e);
                                }
                            }
                        """)
                        logger.info("Modal forcefully removed from DOM")
                except Exception as e3:
                    logger.warning(f"Error checking/removing modal: {e3}")
                
                # Always send escape key as failsafe
                try:
                    ActionChains(driver).send_keys(webdriver.Keys.ESCAPE).perform()
                    logger.info("Sent escape key to close modal")
                except Exception as e4:
                    logger.warning(f"Failed to send escape key: {e4}")
                
                # Verify the modal is actually gone
                try:
                    modal_gone = WebDriverWait(driver, 5).until_not(
                        EC.presence_of_element_located((By.CLASS_NAME, "z-window-modal"))
                    )
                    if modal_gone:
                        logger.info("Confirmed modal is closed")
                    else:
                        logger.warning("Modal may still be present after close attempts")
                except:
                    logger.warning("Could not verify if modal is closed")
                
            except Exception as close_err:
                logger.warning(f"All modal closing attempts failed: {close_err}. Using most aggressive method")
                
                # Most aggressive approach: inject page-level reset code
                driver.execute_script("""
                    // Clear any modal or overlay elements
                    document.querySelectorAll('.z-window-modal, .z-modal-mask, .z-window-shadow, .z-window').forEach(el => {
                        if (el.parentNode) el.parentNode.removeChild(el);
                    });
                    
                    // Clear overlay styles from body
                    document.body.style.overflow = '';
                    
                    // Release any event handlers and reset UI state
                    document.body.click();
                    
                    // Try to force garbage collection of event handlers
                    setTimeout(function() { 
                        if (window.gc) window.gc();
                    }, 100);
                """)
                
                # Force page refresh as last resort
                try:
                    driver.execute_script("location.reload();")
                    logger.info("Reloaded page to clear modal state")
                    time.sleep(3)
                    if not reapply_navigation(driver):
                        logger.warning(f"Failed to reapply navigation after refreshing page for expediente {expediente}")
                        # Try to continue anyway
                    time.sleep(2)
                except Exception as reload_err:
                    logger.warning(f"Error during page reload: {reload_err}")
                    pass
            
            time.sleep(2)  # Extended wait for modal to fully close
            return True, None
            
        except Exception as e:
            logger.warning(f"Modal attempt {attempt+1}/{max_modal_attempts} for {expediente}: {str(e)}")
            if attempt == max_modal_attempts - 1:
                error_msg = f"Error in modal handling: {str(e)}"
                logger.error(error_msg)
                # Try to force close the modal even on error
                try:
                    driver.execute_script("document.querySelector('.z-window-modal')?.remove();")
                except:
                    pass
                return False, error_msg
            time.sleep(1)
    
    return False, "Max attempts exceeded"

async def clear_search_state(driver):
    """
    Clear the search input field and reset the search state.
    This prevents getting stuck in a loop with the same expediente.
    """
    logger.info("Clearing search state to prepare for next expediente")
    
    try:
        # First try to find and click the "Limpiar" (Clear) button if present
        try:
            clear_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Limpiar')]"))
            )
            driver.execute_script("arguments[0].click();", clear_button)
            logger.info("Clicked 'Limpiar' button to clear search")
            await asyncio.sleep(1)
            return True
        except Exception as e:
            logger.info(f"Limpiar button not found or not clickable: {e}")
        
        # Try to clear the input field directly
        try:
            search_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "textInput"))
            )
            # Clear the input field
            search_input.clear()
            # Also send ctrl+a and delete as a more thorough way to clear
            search_input.send_keys(webdriver.Keys.CONTROL + "a")
            search_input.send_keys(webdriver.Keys.DELETE)
            logger.info("Cleared search input field directly")
            await asyncio.sleep(1)
            return True
        except Exception as e:
            logger.warning(f"Could not clear search input field: {e}")
        
        # If all else fails, use JavaScript to reset the input field
        try:
            driver.execute_script("""
                // Reset the search input
                var inputs = document.querySelectorAll('input[type="text"]');
                for(var i=0; i<inputs.length; i++) {
                    inputs[i].value = '';
                }
                
                // Try to find specific search input by ID
                var searchInput = document.getElementById('textInput');
                if(searchInput) {
                    searchInput.value = '';
                }
            """)
            logger.info("Reset search input using JavaScript")
            await asyncio.sleep(1)
            return True
        except Exception as e:
            logger.warning(f"JavaScript reset of search input failed: {e}")
            
        # Last resort: refresh the page, but this is expensive
        try:
            logger.warning("Using page refresh as last resort to clear search state")
            driver.refresh()
            time.sleep(3)
            if not reapply_navigation(driver):
                logger.warning("Failed to reapply navigation after page refresh in clear_search_state")
            return True
        except Exception as e:
            logger.error(f"Page refresh failed: {e}")
            
        return False
    except Exception as e:
        logger.error(f"Error in clear_search_state: {e}")
        return False

def type_and_search(driver, expediente, max_attempts=3):
    """Type text and click search with retry logic."""
    input_selector = 'input.z-textbox:not([style*="display:none"])'
    for attempt in range(max_attempts):
        try:
            time.sleep(5)  # Increased initial wait for page load
            try:
                WebDriverWait(driver, 10).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "z-loading"))
                )
            except Exception:
                pass
            # Always fetch a fresh input element
            gde_input = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, input_selector))
            )
            # Inner loop to handle stale element issues during typing
            for inner_attempt in range(2):
                try:
                    gde_input.click()
                    time.sleep(1)
                    gde_input.clear()
                    time.sleep(1)
                    gde_input.send_keys(expediente)
                    time.sleep(2)
                    if gde_input.get_attribute('value') == expediente:
                        break  # Successfully entered text
                    else:
                        raise Exception("Text verification failed in inner loop")
                except StaleElementReferenceException:
                    logger.info("Input field went stale during typing; re-finding element.")
                    gde_input = driver.find_element(By.CSS_SELECTOR, input_selector)
            if gde_input.get_attribute('value') != expediente:
                raise Exception(f"Failed to enter text exactly. Expected: '{expediente}', Got: '{gde_input.get_attribute('value')}'")
            logger.info(f"Successfully typed '{expediente}' into search box")
            # Find and click search button
            search_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="Buscar"]'))
            )
            search_btn.click()
            logger.info("Clicked search button")
            time.sleep(3)
            for msg in driver.find_elements(By.CLASS_NAME, "z-notification-content"):
                if "debe ingresar un valor" in msg.text.lower():
                    raise Exception("Input validation error: Must enter a value")
            # Optionally verify the input value if possible
            try:
                gde_input_check = driver.find_element(By.CSS_SELECTOR, input_selector)
                if gde_input_check.get_attribute('value') != expediente:
                    logger.warning(f"Input value changed after search click. Expected: '{expediente}', Got: '{gde_input_check.get_attribute('value')}'. This might be expected behavior.")
            except StaleElementReferenceException:
                logger.info("Input element became stale after search; which is expected.")
            return True
        except Exception as e:
            logger.warning(f"Attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_attempts - 1:
                driver.refresh()
                time.sleep(5)
                continue
            raise Exception("Failed to type and search after multiple attempts")
    return False

# ---------------------------------------------
# MAIN ASYNC WORKFLOW
# ---------------------------------------------
async def async_main():
    try:
        load_dotenv()
        url = os.getenv("URL")
        if not url:
            raise ValueError("URL environment variable is not set")
        if not url.startswith("http"):
            raise ValueError(f"Invalid URL format: {url}")
        logger.info(f"Using URL: {url}")

        # Ask user for CSV or XLSX path
        print("Enter the path to the CSV or XLSX file with expedientes (skip header row): ", end="", flush=True)
        input_path = input().strip()
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"File not found: {input_path}")

        # Handle ZIP files
        if input_path.lower().endswith(".zip"):
            logger.info(f"Detected ZIP file: {input_path}")
            import zipfile
            import tempfile
            
            # Create a temporary directory to extract files
            with tempfile.TemporaryDirectory() as temp_dir:
                logger.info(f"Extracting ZIP to temporary directory: {temp_dir}")
                try:
                    with zipfile.ZipFile(input_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                    
                    # Find CSV or XLSX files in the extracted contents
                    extracted_files = [f for f in os.listdir(temp_dir) if f.lower().endswith(('.csv', '.xlsx'))]
                    
                    if not extracted_files:
                        raise ValueError("No CSV or XLSX files found in the ZIP archive")
                    
                    if len(extracted_files) > 1:
                        print("Multiple files found in ZIP. Please select one:")
                        for i, file in enumerate(extracted_files):
                            print(f"{i+1}. {file}")
                        selection = int(input("Enter number: ")) - 1
                        if selection < 0 or selection >= len(extracted_files):
                            raise ValueError("Invalid selection")
                        selected_file = extracted_files[selection]
                    else:
                        selected_file = extracted_files[0]
                    
                    # Update input_path to point to the extracted file
                    input_path = os.path.join(temp_dir, selected_file)
                    logger.info(f"Using extracted file: {selected_file}")
                    
                except Exception as e:
                    logger.error(f"Error extracting ZIP file: {e}")
                    raise
        
        # Check if file is Excel and verify openpyxl is installed
        if input_path.lower().endswith(".xlsx"):
            try:
                import openpyxl
            except ImportError:
                logger.error("The openpyxl package is required to read Excel files.")
                print("\nPlease install openpyxl using one of these commands:")
                print("pip install openpyxl")
                print("- or -")
                print("conda install openpyxl")
                raise ImportError("Missing required package: openpyxl")
            df = pd.read_excel(input_path)
        else:
            # Try different encodings if UTF-8 fails
            try:
                df = pd.read_csv(input_path, encoding='utf-8', sep=',', engine='python')
            except UnicodeDecodeError:
                try:
                    logger.info("UTF-8 encoding failed, trying latin-1")
                    df = pd.read_csv(input_path, encoding='latin-1', sep=',', engine='python')
                except Exception as e:
                    logger.info("latin-1 encoding failed, trying with auto detection")
                    import chardet
                    with open(input_path, 'rb') as f:
                        result = chardet.detect(f.read())
                    detected_encoding = result['encoding']
                    logger.info(f"Detected encoding: {detected_encoding}")
                    df = pd.read_csv(input_path, encoding=detected_encoding, sep=',', engine='python')

        if "Número Expediente" not in df.columns:
            raise ValueError("Column 'Número Expediente' not found in file")

        data_dir = os.path.join(os.getcwd(), "data")
        os.makedirs(data_dir, exist_ok=True)
        results_csv = os.path.join(data_dir, "expedientes.csv")
        with open(results_csv, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(["Expediente", "Downloaded", "Error"])

        driver, download_dir = initialize_driver()

        try:
            logger.info(f"Navigating to {url}")
            driver.get(url)
            await asyncio.to_thread(handle_login, driver)
            
            # Try to reapply navigation, but continue even if it fails
            if not reapply_navigation(driver):
                logger.warning("Navigation reapplication failed, will attempt to continue anyway")
                # Give the page some time to stabilize
                time.sleep(5)
                
                # Take a screenshot to see where we are
                try:
                    screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "post_login_state.png")
                    driver.save_screenshot(screenshot_path)
                    logger.info(f"Current page state saved to {screenshot_path}")
                except Exception as ss_err:
                    logger.error(f"Failed to save screenshot: {ss_err}")
                
                # Check if we need to refresh the page
                try:
                    if "error" in driver.page_source.lower() or "expired" in driver.page_source.lower():
                        logger.warning("Page appears to have error or session expired message, refreshing")
                        driver.refresh()
                        time.sleep(5)
                except:
                    pass
            
            # For each expediente in the file
            for index, row in df.iterrows():
                # Get the expediente exactly as it is in Excel, just strip any leading/trailing whitespace
                expediente = str(row["Número Expediente"]).strip()
                logger.info(f"==> Processing expediente: {expediente} (Row {index+1})")
                
                # Refresh the page every 5 expedientes to keep UI state clean
                if index > 0 and index % 5 == 0:
                    logger.info("Refreshing page to maintain clean UI state...")
                    driver.refresh()
                    time.sleep(5)
                    if not reapply_navigation(driver):
                        logger.warning(f"Failed to reapply navigation after refreshing page for expediente {expediente}")
                        # Try to continue anyway
                    time.sleep(2)
                
                attempt_count = 0
                max_attempts = 5
                downloaded = False
                error_message = ""
                while attempt_count < max_attempts:
                    try:
                        attempt_count += 1
                        if check_and_relogin(driver):
                            if not reapply_navigation(driver):
                                logger.warning(f"Failed to reapply navigation after session expired for expediente {expediente}")
                                # Try to continue anyway
                        
                        # Type and search with retry logic
                        if not await asyncio.to_thread(type_and_search, driver, expediente):
                            raise Exception("Failed to type and search after multiple attempts")
                        
                        # Add delay after search to ensure page loads
                        time.sleep(5)  # Wait for search results
                        
                        # --- Wait for exactly one row in results ---
                        await asyncio.to_thread(wait_for_single_result, driver, timeout=15)
                        logger.info("One result row detected.")
                        
                        # Add delay before clicking combo
                        time.sleep(2)
                        
                        # --- Click the combo to select "Visualizar" ---
                        await asyncio.to_thread(click_visualizar_option, driver, expediente)
                        # --- In the modal, click download and then close modal ---
                        res, err = await handle_modal_download(driver, expediente, download_dir)
                        if not res:
                            error_message = err
                            raise Exception(f"Modal download error: {err}")
                        
                        # Clear search state to prevent looping on the same expediente
                        # await clear_search_state(driver)  # Commented out to prevent automatic clearing of search input
                        
                        logger.info(f"Successfully processed expediente: {expediente}")
                        downloaded = True
                        break  # Exit the retry loop on success
                    except Exception as ex:
                        logger.warning(f"Attempt {attempt_count} for expediente {expediente} failed: {str(ex)}")
                        if check_session_expired(driver):
                            handle_login(driver)
                            if not reapply_navigation(driver):
                                logger.warning(f"Failed to reapply navigation after session expired for expediente {expediente}")
                                # Try to continue anyway
                        time.sleep(3)
                if not downloaded and error_message == "":
                    error_message = f"Exceeded {max_attempts} attempts"
                with open(results_csv, 'a', newline='', encoding='utf-8') as f:
                    csv.writer(f).writerow([expediente, downloaded, error_message])
            # Wait for downloads to finish
            logger.info("Waiting for downloads to complete...")
            end_time = time.time() + 60
            while time.time() < end_time:
                if not [f for f in os.listdir(download_dir) if f.endswith('.crdownload') or f.endswith('.tmp')]:
                    break
                time.sleep(0.5)
        finally:
            driver.quit()
            logger.info("Driver quit.")

        final_files = os.listdir(download_dir)
        logger.info(f"Files in downloads directory: {final_files}")
        print("\nAutomation completed. Please check your downloads folder for files.")
    except Exception as e:
        logger.error(f"An error occurred: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    asyncio.run(async_main())
