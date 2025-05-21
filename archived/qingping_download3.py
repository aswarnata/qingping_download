        self.progress_data = {}import tkinter as tk
from tkinter import ttk, messagebox, StringVar, IntVar, BooleanVar
import threading
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, 
    StaleElementReferenceException, 
    WebDriverException,
    NoSuchElementException
)
from datetime import datetime
import os
import json
import random

class AutoQingPing:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.driver = None
        self.log_callback = print  # Default to print
        self.last_page_refresh = time.time()
        self.max_retries = 3

    def set_log_callback(self, callback):
        """Set a callback function for logging"""
        self.log_callback = callback

    def log(self, message):
        """Log a message using the callback"""
        self.log_callback(message)

    def login(self, email, password, headless=False):
        """Login to QingPing IoT with retry logic"""
        options = Options()
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-extensions")
        
        # If very laggy, headless mode can help
        if headless:
            options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")  # Helps with memory issues
        options.add_argument("--disable-browser-side-navigation")  # Reduces crashes
        options.add_argument("--window-size=1920,1080")  # Set a standard window size
        
        # Improve performance
        options.add_argument("--disable-features=NetworkService")
        options.add_argument("--disable-features=VizDisplayCompositor")
        
        # Increase timeouts for laggy sites
        options.add_experimental_option("detach", True)  # Keep browser open
        
        # Add page load strategy
        options.page_load_strategy = 'eager'  # 'eager' is faster than 'normal'
        
        # Initialize WebDriver with retry logic
        retry_count = 0
        while retry_count < self.max_retries:
            try:
                self.log("🔄 Initializing Chrome WebDriver...")
                self.driver = webdriver.Chrome(options=options)
                self.driver.set_page_load_timeout(60)  # Set page load timeout to 60 seconds
                self.driver.set_script_timeout(60)  # Set script timeout to 60 seconds
                break
            except WebDriverException as e:
                retry_count += 1
                self.log(f"⚠️ WebDriver initialization failed (attempt {retry_count}/{self.max_retries}): {e}")
                if retry_count >= self.max_retries:
                    raise Exception("Failed to initialize WebDriver after multiple attempts") from e
                time.sleep(3)  # Wait before retrying
        
        # Login with retry logic
        retry_count = 0
        while retry_count < self.max_retries:
            try:
                self.log("🔐 Opening login page...")
                self.driver.get("https://qingpingiot.com/user/login")
                
                # Create a longer wait time for laggy sites
                wait = WebDriverWait(self.driver, 30)
                
                # Click on the email login tab
                self.log("⏱️ Waiting for login elements to load...")
                email_tab = self._retry_find_element(wait, By.XPATH, "//div[@data-node-key='email']", "email tab")
                email_tab.click()
                
                # Fill in email
                email_input = self._retry_find_element(wait, By.XPATH, "//input[@placeholder='Email']", "email input")
                email_input.clear()
                email_input.send_keys(email)
                
                # Press Tab to move to password input
                email_input.send_keys(Keys.TAB)
                time.sleep(1)  # Small pause to ensure focus shifts
                
                # Type the password
                webdriver.ActionChains(self.driver).send_keys(password).perform()
                time.sleep(1)
                
                # Click the "Log In" button
                login_button = self._retry_find_element(
                    wait, By.XPATH, "//button[@type='button' and .//span[text()='Log In']]", "login button"
                )
                login_button.click()
                
                self.log("✅ Login button clicked. Waiting for page load or verification (15s)...")
                time.sleep(15)  # Increased wait time for laggy sites
                
                # Verify we're logged in by checking for a dashboard element
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-cardItem")))
                    self.log("🎉 Successfully logged in!")
                    self.last_page_refresh = time.time()  # Reset refresh timer
                    return self.driver
                except TimeoutException:
                    # We might not be on the dashboard page yet
                    self.log("⚠️ Couldn't verify dashboard loaded. Checking if we need to handle verification...")
                    
                    # Check for verification elements or errors
                    if "login" in self.driver.current_url.lower():
                        raise Exception("Login failed - still on login page")
                    
                    # If we're not on the login page, assume we're logged in somewhere
                    self.log("➡️ Navigating to main dashboard...")
                    self.driver.get("https://qingpingiot.com/welcome")
                    time.sleep(5)
                
                return self.driver
                
            except Exception as e:
                retry_count += 1
                self.log(f"⚠️ Login attempt {retry_count}/{self.max_retries} failed: {e}")
                
                if retry_count >= self.max_retries:
                    raise Exception("Failed to login after multiple attempts") from e
                    
                # Close and recreate driver for a fresh attempt
                if self.driver:
                    self.driver.quit()
                    self.driver = webdriver.Chrome(options=options)
                
                time.sleep(3)  # Wait before retrying
    
    def _retry_find_element(self, wait, by, selector, element_name, max_attempts=5):
        """Retry finding an element with exponential backoff"""
        for attempt in range(max_attempts):
            try:
                element = wait.until(EC.element_to_be_clickable((by, selector)))
                return element
            except (TimeoutException, StaleElementReferenceException) as e:
                if attempt == max_attempts - 1:
                    self.log(f"❌ Failed to find {element_name} after {max_attempts} attempts")
                    raise e
                
                backoff = 2 ** attempt + random.uniform(0, 1)  # Exponential backoff with jitter
                self.log(f"⏱️ Waiting {backoff:.2f}s before retrying to find {element_name}...")
                time.sleep(backoff)
    
    def check_and_refresh_if_needed(self, force=False, max_age_seconds=300):
        """Check if page needs a refresh based on time or force parameter"""
        current_time = time.time()
        time_since_refresh = current_time - self.last_page_refresh
        
        if force or time_since_refresh > max_age_seconds:
            self.log(f"🔄 Page refresh needed (age: {time_since_refresh:.1f}s). Refreshing...")
            try:
                current_url = self.driver.current_url
                self.driver.refresh()
                # Wait for the page to load after refresh
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                self.log("✅ Page refreshed successfully")
                self.last_page_refresh = time.time()
                return True
            except Exception as e:
                self.log(f"⚠️ Error during page refresh: {e}")
                
                # If refresh fails, try to navigate to the current URL again
                try:
                    current_url = self.driver.current_url
                    self.driver.get(current_url)
                    time.sleep(5)
                    self.log("✅ Navigated to current URL as fallback")
                    self.last_page_refresh = time.time()
                    return True
                except Exception as e2:
                    self.log(f"❌ Navigation fallback also failed: {e2}")
                    return False
        
        return False
    
    def get_devices(self, driver, mode="all", page_size=40, start_page=1, end_page=None):
        """Get devices with pagination support for large dashboards"""
        all_data = []
        current_page = start_page
        self.log(f"📱 Starting device collection from page {start_page}")
        
        # Calculate total pages if end_page is not specified
        if end_page is None:
            try:
                # Navigate to the devices page first
                driver.get("https://qingpingiot.com/devices")
                time.sleep(5)
                
                # Look for pagination info
                wait = WebDriverWait(driver, 10)
                pagination = self._retry_find_element(wait, By.CSS_SELECTOR, ".ant-pagination", "pagination", 2)
                
                # Get the last page button text
                page_items = pagination.find_elements(By.CSS_SELECTOR, ".ant-pagination-item")
                if page_items:
                    last_page_text = page_items[-1].text
                    end_page = int(last_page_text)
                    self.log(f"📄 Detected {end_page} total pages")
                else:
                    # Default to 1 page if we can't detect pagination
                    end_page = 1
                    self.log("⚠️ Could not detect pagination, assuming 1 page")
            except Exception as e:
                self.log(f"⚠️ Error detecting pagination: {e}")
                end_page = 1  # Default to 1 page
        
        # Process each page
        while current_page <= end_page:
            retry_count = 0
            success = False
            
            while retry_count < self.max_retries and not success:
                try:
                    # Check if we need to refresh the page
                    if retry_count > 0:
                        self.check_and_refresh_if_needed(force=True)
                    
                    self.log(f"📄 Processing page {current_page} of {end_page}...")
                    
                    # Navigate to the specific page (for pages after page 1)
                    if current_page > 1:
                        # Ensure we're on the devices page
                        if "devices" not in driver.current_url:
                            driver.get("https://qingpingiot.com/devices")
                            time.sleep(5)
                        
                        wait = WebDriverWait(driver, 20)
                        
                        # Find pagination and click on the specific page
                        try:
                            pagination = self._retry_find_element(wait, By.CSS_SELECTOR, ".ant-pagination", "pagination")
                            page_link = self._retry_find_element(
                                wait, 
                                By.XPATH, 
                                f"//li[contains(@class, 'ant-pagination-item') and @title='{current_page}']//a",
                                f"page {current_page} link"
                            )
                            self.log(f"➡️ Clicking on page {current_page}...")
                            page_link.click()
                            time.sleep(5)  # Wait for page content to load
                        except Exception as e:
                            self.log(f"⚠️ Error navigating to page {current_page}: {e}")
                            
                            # Try direct URL approach as a fallback
                            driver.get(f"https://qingpingiot.com/devices?current={current_page}")
                            time.sleep(5)
                            self.log(f"✓ Navigated to page {current_page} via direct URL")
                    
                    # Wait for cards to load
                    wait = WebDriverWait(driver, 20)
                    wait.until(EC.presence_of_element_located((
                        By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-cardItem"
                    )))
                    
                    # Get all cards
                    cards = driver.find_elements(By.CSS_SELECTOR, 
                        ".antd-pro-pages-welcome-less-index-cardItem.antd-pro-pages-welcome-less-index-cardItemHiddenStatus"
                    )
                    
                    # Filter based on mode
                    if mode == "online":
                        filtered = [card for card in cards if "offlineCardItem" not in card.get_attribute("class")]
                    elif mode == "offline":
                        filtered = [card for card in cards if "offlineCardItem" in card.get_attribute("class")]
                    else:
                        filtered = cards
                    
                    self.log(f"📦 Found {len(filtered)} device(s) on page {current_page}")
                    
                    # Process each card
                    page_data = []
                    for i, card in enumerate(filtered):
                        try:
                            # Extract card data with retry logic
                            device_data = self._extract_card_data(card, i)
                            if device_data:
                                page_data.append(device_data)
                                
                        except StaleElementReferenceException:
                            self.log(f"⚠️ Stale element for card {i+1}, refreshing page...")
                            self.check_and_refresh_if_needed(force=True)
                            break  # Break the loop and retry the current page
                            
                        except Exception as e:
                            self.log(f"⚠️ Error processing card {i+1}: {e}")
                            continue
                    
                    # If we processed cards successfully, add to all_data
                    if page_data:
                        all_data.extend(page_data)
                        self.log(f"✅ Added {len(page_data)} devices from page {current_page}")
                        success = True
                    else:
                        raise Exception(f"No devices extracted from page {current_page}")
                    
                except Exception as e:
                    retry_count += 1
                    self.log(f"⚠️ Error processing page {current_page} (attempt {retry_count}/{self.max_retries}): {e}")
                    
                    if retry_count >= self.max_retries:
                        self.log(f"❌ Failed to process page {current_page} after {self.max_retries} attempts")
                        # Continue to the next page instead of failing the entire process
                        success = True  # Mark as "success" to move on
                    else:
                        # Wait before retrying
                        time.sleep(5)
            
            # Move to the next page
            current_page += 1
        
        # Convert to DataFrame
        df = pd.DataFrame(all_data)
        self.log(f"📊 Collected data for {len(df)} devices across {end_page - start_page + 1} pages")
        return df
    
    def _extract_card_data(self, card, index):
        """Extract data from a device card with error handling"""
        try:
            device_name = card.find_element(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-deviceNameText").text
            
            try:
                update_time = card.find_element(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-updateTime").text
            except NoSuchElementException:
                update_time = "Unknown"
            
            try:
                link = card.find_element(By.TAG_NAME, "a").get_attribute("href")
            except NoSuchElementException:
                link = ""
            
            # Get all data points (with more robust error handling)
            data_points = {}
            try:
                titles = card.find_elements(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-title")
                values = card.find_elements(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-dataShow p")
                
                for title, value in zip(titles, values):
                    try:
                        title_text = title.text.strip()
                        value_text = value.text.strip()
                        if title_text and value_text:  # Only add non-empty data
                            data_points[title_text] = value_text
                    except Exception:
                        continue
            except Exception:
                # If we can't get data points, just continue with basic info
                pass
            
            # Create and return the data dictionary
            card_data = {
                "device_name": device_name,
                "update_time": update_time,
                "link": link,
                **data_points
            }
            
            self.log(f"[{index+1}] Scraped: {device_name}")
            return card_data
            
        except Exception as e:
            self.log(f"[CARD ERROR {index+1}] {e}")
            return None  # Return None for failed cards
    
    def automate_date_selection(self, driver, url, start_date, end_date, retry_count=0):
        """Automate date selection with retry logic for laggy pages"""
        if retry_count >= self.max_retries:
            self.log(f"❌ Max retries reached ({self.max_retries}) for date selection")
            return False
            
        try:
            # Navigate to the detail page
            self.log(f"📄 Opening detail page: {url}")
            driver.get(url)
            
            # Wait for the page to load
            wait = WebDriverWait(driver, 30)
            
            # Check if the page is responsive
            try:
                # Look for a key element to verify the page loaded properly
                wait.until(EC.presence_of_element_located((
                    By.CSS_SELECTOR, "span.antd-pro-pages-history-chart-chart-export_text"
                )))
            except TimeoutException:
                self.log("⚠️ Page took too long to load, refreshing...")
                self.check_and_refresh_if_needed(force=True)
                # Try again with incremented retry count
                return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
            
            # Step 1: Click Export button
            self.log("🔍 Looking for Export button...")
            try:
                export_button = self._retry_find_element(
                    wait, 
                    By.CSS_SELECTOR, 
                    "span.antd-pro-pages-history-chart-chart-export_text", 
                    "export button"
                )
                export_button.click()
                self.log("📦 Export button clicked")
            except Exception as e:
                self.log(f"⚠️ Error clicking export button: {e}")
                self.check_and_refresh_if_needed(force=True)
                return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
            
            # Step 2: Wait for modal to show
            try:
                modal_wrapper = wait.until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ant-modal-content"))
                )
                self.log("🪟 Export modal appeared")
            except TimeoutException:
                self.log("⚠️ Export modal didn't appear, retrying...")
                self.check_and_refresh_if_needed(force=True)
                return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
            
            # Step 3: Open date range picker
            try:
                range_picker_wrapper = modal_wrapper.find_element(By.CLASS_NAME, "ant-picker-range")
                range_picker_wrapper.click()
                self.log("📆 Opened date range picker")
            except Exception as e:
                self.log(f"⚠️ Error opening date picker: {e}")
                # Close the modal and retry
                try:
                    close_button = modal_wrapper.find_element(By.CSS_SELECTOR, ".ant-modal-close")
                    close_button.click()
                except:
                    pass
                time.sleep(1)
                return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
            
            # Step 4: Type the start and end dates
            try:
                formatted_start = start_date
                formatted_end = end_date
                
                inputs = modal_wrapper.find_elements(By.CSS_SELECTOR, "input[placeholder]")
                if len(inputs) >= 2:
                    start_input, end_input = inputs[:2]
                    
                    # Clear and enter start date
                    start_input.send_keys(Keys.CONTROL + "a")
                    start_input.send_keys(Keys.BACKSPACE)
                    start_input.send_keys(formatted_start)
                    time.sleep(0.5)
                    
                    # Clear and enter end date
                    end_input.send_keys(Keys.CONTROL + "a")
                    end_input.send_keys(Keys.BACKSPACE)
                    end_input.send_keys(formatted_end)
                    time.sleep(0.5)
                    
                    # Press ENTER to confirm date range
                    end_input.send_keys(Keys.ENTER)
                    time.sleep(1)
                    
                    self.log(f"📅 Dates filled: {formatted_start} → {formatted_end}")
                else:
                    raise Exception("⚠️ Couldn't locate both date input fields")
            except Exception as e:
                self.log(f"⚠️ Error filling dates: {e}")
                # Close modal and retry
                try:
                    close_button = modal_wrapper.find_element(By.CSS_SELECTOR, ".ant-modal-close")
                    close_button.click()
                except:
                    pass
                time.sleep(1)
                return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
            
            # Step 5: Click Confirm button
            try:
                confirm_button = self._retry_find_element(
                    WebDriverWait(modal_wrapper, 10),
                    By.CSS_SELECTOR, 
                    "button.ant-btn.ant-btn-primary", 
                    "confirm button"
                )
                confirm_button.click()
                self.log("✅ Date selection confirmed")
                time.sleep(2)  # Wait for the export process to start
                return True
            except Exception as e:
                self.log(f"⚠️ Error clicking confirm: {e}")
                # Close modal and retry
                try:
                    close_button = modal_wrapper.find_element(By.CSS_SELECTOR, ".ant-modal-close")
                    close_button.click()
                except:
                    pass
                time.sleep(1)
                return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
            
        except Exception as e:
            self.log(f"❌ Error during date selection: {e}")
            self.check_and_refresh_if_needed(force=True)
            return self.automate_date_selection(driver, url, start_date, end_date, retry_count + 1)
    
    def scrape_filenames(self, driver, page_size=100, page_number=1, retry_count=0):
        """Scrape filenames from the export page with retry logic"""
        if retry_count >= self.max_retries:
            self.log(f"❌ Max retries reached ({self.max_retries}) for scraping filenames")
            return []
            
        try:
            url = "https://qingpingiot.com/export"
            self.log(f"📁 Navigating to export page: {url}")
            driver.get(url)
            
            wait = WebDriverWait(driver, 30)
            
            # Check if page loaded properly
            try:
                # Wait for table or any key element to appear
                wait.until(EC.presence_of_element_located((
                    By.CSS_SELECTOR, ".ant-table"
                )))
            except TimeoutException:
                self.log("⚠️ Export page took too long to load, refreshing...")
                self.check_and_refresh_if_needed(force=True)
                return self.scrape_filenames(driver, page_size, page_number, retry_count + 1)
            
            # Change page size if needed
            if page_size > 40:
                try:
                    # Find and click the page size dropdown
                    page_size_dropdown = self._retry_find_element(
                        wait,
                        By.XPATH, 
                        "//span[@class='ant-select-selection-item' and contains(@title, '/ page')]",
                        "page size dropdown"
                    )
                    page_size_dropdown.click()
                    time.sleep(1)
                    
                    # Select the closest available page size option (20, 40, 100)
                    if page_size <= 20:
                        target_size = "20 / page"
                    elif page_size <= 40:
                        target_size = "40 / page"
                    else:
                        target_size = "100 / page"
                        
                    option = self._retry_find_element(
                        wait,
                        By.XPATH, 
                        f"//div[contains(@class, 'ant-select-item-option') and .//div[text()='{target_size}']]",
                        f"page size option: {target_size}"
                    )
                    
                    # Scroll into view and click
                    driver.execute_script("arguments[0].scrollIntoView(true);", option)
                    time.sleep(0.5)
                    option.click()
                    time.sleep(2)  # Wait for page to refresh with new size
                    
                    self.log(f"📄 Changed page size to {target_size}")
                except Exception as e:
                    self.log(f"⚠️ Error changing page size: {e}")
                    # Continue with default page size
            
            # Navigate to specific page if needed
            if page_number > 1:
                try:
                    # Wait for pagination to be present
                    pagination = self._retry_find_element(wait, By.CSS_SELECTOR, ".ant-pagination", "pagination")
                    
                    # Try to click directly on the page number
                    try:
                        page_button = self._retry_find_element(
                            wait,
                            By.XPATH, 
                            f"//a[@rel='nofollow' and text()='{page_number}']",
                            f"page {page_number} button"
                        )
                        self.log(f"➡️ Clicking page {page_number}...")
                        page_button.click()
                    except Exception:
                        # If direct click fails, try using next button repeatedly
                        self.log(f"⚠️ Direct page navigation failed, using next button...")
                        
                        # Find current page
                        try:
                            current_page_elem = pagination.find_element(
                                By.CSS_SELECTOR, "li.ant-pagination-item-active"
                            )
                            current_page = int(current_page_elem.get_attribute("title"))
                        except:
                            current_page = 1
                        
                        # Click next until we reach target page
                        while current_page < page_number:
                            next_button = pagination.find_element(
                                By.CSS_SELECTOR, "li.ant-pagination-next"
                            )
                            next_button.click()
                            time.sleep(2)
                            
                            # Verify page changed
                            try:
                                current_page_elem = pagination.find_element(
                                    By.CSS_SELECTOR, "li.ant-pagination-item-active"
                                )
                                current_page = int(current_page_elem.get_attribute("title"))
                                self.log(f"📄 Now on page {current_page}")
                            except:
                                self.log("⚠️ Could not verify page change")
                                break
                    
                    # Wait for page content to load
                    time.sleep(3)
                    
                except Exception as e:
                    self.log(f"⚠️ Error navigating to page {page_number}: {e}")
                    # Try direct URL navigation as fallback
                    driver.get(f"https://qingpingiot.com/export?current={page_number}")
                    time.sleep(3)
            
            # Scrape the filenames
            try:
                # Wait for elements to be present
                wait.until(EC.presence_of_element_located((
                    By.XPATH, "//span[contains(text(), '.csv') or contains(text(), '.xlsx')]"
                )))
                
                # Find all filename elements
                file_elements = driver.find_elements(
                    By.XPATH, "//span[contains(text(), '.csv') or contains(text(), '.xlsx')]"
                )
                
                # Extract filenames
                filenames = [
                    elem.text.strip() 
                    for elem in file_elements 
                    if elem.text.strip().endswith(('.csv', '.xlsx'))
                ]
                
                self.log(f"✅ Found {len(filenames)} files on page {page_number}")
                
                # Log the first few filenames
                for i, name in enumerate(filenames[:5]):
                    self.log(f"- {name}")
                
                if len(filenames) > 5:
                    self.log(f"- ... and {len(filenames) - 5} more files")
                
                return filenames
                
            except TimeoutException:
                self.log("⚠️ No files found or page took too long to load")
                self.check_and_refresh_if_needed(force=True)
                return self.scrape_filenames(driver, page_size, page_number, retry_count + 1)
                
            except Exception as e:
                self.log(f"❌ Error scraping filenames: {e}")
                return []
                
        except Exception as e:
            self.log(f"❌ Error in scrape_filenames: {e}")
            self.check_and_refresh_if_needed(force=True)
            return self.scrape_filenames(driver, page_size, page_number, retry_count + 1)
    
    def click_files(self, driver, filenames, page_number=1, retry_count=0):
        """Click on files to download them with retry logic"""
        if not filenames:
            self.log("⚠️ No files to click")
            return False
            
        if retry_count >= self.max_retries:
            self.log(f"❌ Max retries reached ({self.max_retries}) for clicking files")
            return False
        
        try:
            # Navigate to the export page
            url = "https://qingpingiot.com/export"
            driver.get(url)
            
            wait = WebDriverWait(driver, 30)
            
            # Navigate to the specific page if needed
            if page_number > 1:
                try:
                    # Wait for pagination to be present
                    pagination = self._retry_find_element(wait, By.CSS_SELECTOR, ".ant-pagination", "pagination")
                    
                    # Try direct page button click
                    try:
                        page_button = self._retry_find_element(
                            wait,
                            By.XPATH, 
                            f"//a[@rel='nofollow' and text()='{page_number}']",
                            f"page {page_number} button"
                        )
                        self.log(f"➡️ Navigating to page {page_number}...")
                        page_button.click()
                    except Exception:
                        # If direct click fails, try using next button repeatedly
                        self.log(f"⚠️ Direct page navigation failed, using next button...")
                        
                        # Find current page
                        try:
                            current_page_elem = pagination.find_element(
                                By.CSS_SELECTOR, "li.ant-pagination-item-active"
                            )
                            current_page = int(current_page_elem.get_attribute("title"))
                        except:
                            current_page = 1
                        
                        # Click next until we reach target page
                        while current_page < page_number:
                            next_button = pagination.find_element(
                                By.CSS_SELECTOR, "li.ant-pagination-next"
                            )
                            next_button.click()
                            time.sleep(2)
                            
                            # Verify page changed
                            try:
                                current_page_elem = pagination.find_element(
                                    By.CSS_SELECTOR, "li.ant-pagination-item-active"
                                )
                                current_page = int(current_page_elem.get_attribute("title"))
                                self.log(f"📄 Now on page {current_page}")
                            except:
                                self.log("⚠️ Could not verify page change")
                                break
                    
                    # Wait for page content to load
                    time.sleep(5)
                    
                except Exception as e:
                    self.log(f"⚠️ Error navigating to page {page_number}: {e}")
                    # Try direct URL navigation as fallback
                    driver.get(f"https://qingpingiot.com/export?current={page_number}")
                    time.sleep(5)
            
            # Wait for the page to fully load
            self.log(f"⏳ Waiting for page {page_number} to fully load...")
            time.sleep(10)
            
            # Verify files are present
            try:
                # Wait for at least one file to be present
                wait.until(EC.presence_of_element_located((
                    By.XPATH, "//span[contains(text(), '.csv') or contains(text(), '.xlsx')]"
                )))
            except TimeoutException:
                self.log("⚠️ Files not found on page, refreshing...")
                self.check_and_refresh_if_needed(force=True)
                return self.click_files(driver, filenames, page_number, retry_count + 1)
            
            self.log(f"🖱️ Starting to click on {len(filenames)} files...")
            
            # Process files in reverse order to maintain expected behavior
            success_count = 0
            for i in range(len(filenames)):
                if not filenames:
                    break
                    
                # Get the file to click (in reverse order)
                filename = filenames[len(filenames) - 1 - i]
                
                try:
                    # Find and click the file
                    self.log(f"🔍 Looking for file: {filename}")
                    
                    # Use advanced retry finding for each file
                    file_element = self._retry_find_element(
                        wait,
                        By.XPATH, 
                        f"//span[text()='{filename}']",
                        f"file: {filename}",
                        max_attempts=3  # Fewer retries per file to avoid spending too much time
                    )
                    
                    # Scroll to the element and click
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", file_element)
                    time.sleep(1)
                    
                    file_element.click()
                    self.log(f"✅ [{i+1}/{len(filenames)}] Clicked: {filename}")
                    success_count += 1
                    
                    # Add a delay between clicks to avoid overwhelming the browser
                    time.sleep(3)
                    
                    # Check if we need to refresh the page (every 10 files)
                    if (i + 1) % 10 == 0 and i < len(filenames) - 1:
                        self.log("🔄 Refreshing page after 10 downloads...")
                        self.check_and_refresh_if_needed(force=True)
                        
                        # Navigate back to the export page and current page
                        driver.get(f"https://qingpingiot.com/export?current={page_number}")
                        time.sleep(5)
                    
                except TimeoutException:
                    self.log(f"⚠️ Couldn't find file: {filename}")
                    continue
                    
                except StaleElementReferenceException:
                    self.log(f"⚠️ Element became stale for: {filename}")
                    # Refresh and retry from current position
                    self.check_and_refresh_if_needed(force=True)
                    
                    # Create a new list with remaining files
                    remaining_files = filenames[len(filenames) - 1 - i:]
                    return self.click_files(driver, remaining_files, page_number, retry_count)
                    
                except Exception as e:
                    self.log(f"⚠️ Error clicking on {filename}: {e}")
                    continue
            
            self.log(f"📊 Successfully clicked {success_count} out of {len(filenames)} files")
            return success_count > 0
            
        except Exception as e:
            self.log(f"❌ Error in click_files: {e}")
            self.check_and_refresh_if_needed(force=True)
            
            # If we've had multiple failures, try a smaller batch
            if retry_count > 0 and len(filenames) > 5:
                self.log("⚠️ Reducing batch size for better reliability")
                # Try with just the first 5 files
                return self.click_files(driver, filenames[:5], page_number, retry_count + 1)
            
            return self.click_files(driver, filenames, page_number, retry_count + 1)
    
    def close_driver(self):
        """Safely close the WebDriver"""
        if self.driver:
            try:
                self.driver.quit()
                self.log("🔒 Browser closed")
            except Exception as e:
                self.log(f"⚠️ Error closing browser: {e}")
            finally:
                self.driver = None


class QingPingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("QingPing IoT Automation")
        self.root.geometry("900x700")
        self.root.minsize(900, 700)
        
        self.automation = None
        self.running = False
        # Create a mapping between devices and filenames
        self.device_filename_map = {}
        
        # Create variables
        self.email_var = StringVar(value="")
        self.password_var = StringVar(value="")
        self.start_date_var = StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.end_date_var = StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.mode_var = StringVar(value="all")
        self.device_count_var = IntVar(value=100)
        self.headless_var = BooleanVar(value=False)
        self.start_page_var = IntVar(value=1)
        self.end_page_var = IntVar(value=12)  # Default to your 12 pages
        self.batch_size_var = IntVar(value=20)
        self.auto_refresh_var = BooleanVar(value=True)
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.tab_control = ttk.Notebook(self.main_frame)
        
        # Settings tab
        self.settings_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.settings_tab, text="Settings")
        
        # Advanced tab
        self.advanced_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.advanced_tab, text="Advanced")
        
        # Log tab
        self.log_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.log_tab, text="Log")
        
        # Progress tab
        self.progress_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.progress_tab, text="Progress")
        
        self.tab_control.pack(fill=tk.BOTH, expand=True)
        
        # Setup tabs
        self.setup_settings_tab()
        self.setup_advanced_tab()
        self.setup_log_tab()
        self.setup_progress_tab()
        
        # Setup bottom buttons
        self.setup_buttons()
        
        # Load previous settings
        self.load_settings()
        
    def setup_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.settings_tab, text="QingPing IoT Settings", padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Login credentials
        cred_frame = ttk.LabelFrame(settings_frame, text="Login Credentials", padding=10)
        cred_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(cred_frame, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(cred_frame, textvariable=self.email_var, width=40).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(cred_frame, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        password_entry = ttk.Entry(cred_frame, textvariable=self.password_var, width=40, show="*")
        password_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Date range
        date_frame = ttk.LabelFrame(settings_frame, text="Date Range", padding=10)
        date_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(date_frame, text="Start Date (YYYY/MM/DD):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(date_frame, textvariable=self.start_date_var, width=15).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(date_frame, text="End Date (YYYY/MM/DD):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(date_frame, textvariable=self.end_date_var, width=15).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Device settings
        device_frame = ttk.LabelFrame(settings_frame, text="Device Settings", padding=10)
        device_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(device_frame, text="Mode:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        mode_combo = ttk.Combobox(device_frame, textvariable=self.mode_var, values=["all", "online", "offline"], state="readonly", width=15)
        mode_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Page settings
        ttk.Label(device_frame, text="Start Page:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(device_frame, from_=1, to=100, textvariable=self.start_page_var, width=5).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(device_frame, text="End Page:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(device_frame, from_=1, to=100, textvariable=self.end_page_var, width=5).grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Headless mode checkbox
        ttk.Checkbutton(
            device_frame, 
            text="Run in Headless Mode (faster but no visible browser)", 
            variable=self.headless_var
        ).grid(row=2, column=0, columnspan=4, sticky=tk.W, padx=5, pady=5)
        
    def setup_advanced_tab(self):
        adv_frame = ttk.LabelFrame(self.advanced_tab, text="Advanced Settings", padding=10)
        adv_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Batch processing
        batch_frame = ttk.LabelFrame(adv_frame, text="Batch Processing", padding=10)
        batch_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(batch_frame, text="Process Devices in Batches of:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(batch_frame, from_=1, to=100, textvariable=self.batch_size_var, width=5).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(batch_frame, text="(smaller batches are more reliable but slower)").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        # Performance settings
        perf_frame = ttk.LabelFrame(adv_frame, text="Performance & Reliability", padding=10)
        perf_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(
            perf_frame, 
            text="Auto-refresh browser when page becomes unresponsive",
            variable=self.auto_refresh_var
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Help text
        ttk.Label(
            adv_frame, 
            text="Note: For large numbers of devices (450+), use smaller batch sizes (10-20) and enable auto-refresh\n" +
                 "for better reliability. Headless mode can improve performance if your system is struggling.",
            wraplength=600
        ).pack(padx=5, pady=10)
        
    def setup_log_tab(self):
        log_frame = ttk.Frame(self.log_tab, padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create log text widget with scrollbar
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=20, width=80)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure tags for colored text
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("success", foreground="green")
        self.log_text.tag_configure("info", foreground="blue")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("download", foreground="purple")
        
        # Make the log text read-only
        self.log_text.configure(state=tk.DISABLED)
        
    def setup_progress_tab(self):
        progress_frame = ttk.Frame(self.progress_tab, padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create a treeview to display progress
        columns = ("device", "status", "timestamp", "filename")
        self.progress_tree = ttk.Treeview(progress_frame, columns=columns, show="headings")
        
        # Define headings
        self.progress_tree.heading("device", text="Device Name")
        self.progress_tree.heading("status", text="Status")
        self.progress_tree.heading("timestamp", text="Timestamp")
        self.progress_tree.heading("filename", text="Export Filename")
        
        # Define columns
        self.progress_tree.column("device", width=250)
        self.progress_tree.column("status", width=100)
        self.progress_tree.column("timestamp", width=150)
        self.progress_tree.column("filename", width=200)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(progress_frame, orient="vertical", command=self.progress_tree.yview)
        self.progress_tree.configure(yscrollcommand=scrollbar.set)
        
        self.progress_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add stats at the bottom
        self.stats_frame = ttk.LabelFrame(progress_frame, text="Statistics", padding=10)
        self.stats_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
        
        self.stats_var = StringVar(value="Total: 0 | Exported: 0 | Downloaded: 0 | Failed: 0 | Pending: 0")
        ttk.Label(self.stats_frame, textvariable=self.stats_var, font=("Arial", 10, "bold")).pack()
        
    def setup_buttons(self):
        button_frame = ttk.Frame(self.main_frame, padding=10)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.start_button = ttk.Button(button_frame, text="Start Automation", command=self.start_automation)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.resume_button = ttk.Button(button_frame, text="Resume Automation", command=self.resume_automation, state=tk.DISABLED)
        self.resume_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop Automation", command=self.stop_automation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save Settings", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.exit_app).pack(side=tk.RIGHT, padx=5)
        
    def log_message(self, message, tag=None):
        """Add a message to the log text widget"""
        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if tag:
            self.log_text.insert(tk.END, f"[{timestamp}] ", "info")
            self.log_text.insert(tk.END, f"{message}\n", tag)
        else:
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            
        self.log_text.see(tk.END)  # Scroll to the end
        self.log_text.configure(state=tk.DISABLED)
        
        # Update the UI
        self.root.update_idletasks()
        
    def clear_log(self):
        """Clear the log text widget"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        
    def update_progress(self, device_name, status, filename=""):
        """Update progress in the treeview"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Check if device already exists in the tree
        existing_items = self.progress_tree.get_children()
        for item_id in existing_items:
            if self.progress_tree.item(item_id, "values")[0] == device_name:
                # If updating to "Downloaded", keep the filename from the previous status
                if status == "Downloaded" and not filename:
                    filename = self.progress_tree.item(item_id, "values")[3]
                # Update existing item
                self.progress_tree.item(item_id, values=(device_name, status, timestamp, filename))
                break
        else:
            # Add new item
            self.progress_tree.insert("", "end", values=(device_name, status, timestamp, filename))
        
        # Apply color tags based on status
        for item_id in self.progress_tree.get_children():
            values = self.progress_tree.item(item_id, "values")
            if values[0] == device_name:
                if status == "Exported":
                    self.progress_tree.item(item_id, tags=("exported",))
                elif status == "Downloaded":
                    self.progress_tree.item(item_id, tags=("downloaded",))
                elif status == "Failed":
                    self.progress_tree.item(item_id, tags=("failed",))
                elif status == "Pending":
                    self.progress_tree.item(item_id, tags=("pending",))
        
        # Configure the tags if not already done
        try:
            self.progress_tree.tag_configure("exported", background="#FFF9C4")  # Light yellow
            self.progress_tree.tag_configure("downloaded", background="#C8E6C9")  # Light green
            self.progress_tree.tag_configure("failed", background="#FFCDD2")  # Light red
            self.progress_tree.tag_configure("pending", background="#E1F5FE")  # Light blue
        except:
            pass
            
        # Update stats
        self.update_stats()
        
        # Store progress in dictionary for potential resume
        self.progress_data[device_name] = {
            "status": status, 
            "timestamp": timestamp,
            "filename": filename
        }
        self.save_progress()
        
    def update_stats(self):
        """Update statistics in the progress tab"""
        total = len(self.progress_tree.get_children())
        exported = 0
        downloaded = 0
        failed = 0
        
        for item_id in self.progress_tree.get_children():
            status = self.progress_tree.item(item_id, "values")[1]
            if status == "Exported":
                exported += 1
            elif status == "Downloaded":
                downloaded += 1
            elif status == "Failed":
                failed += 1
        
        pending = total - exported - downloaded - failed
        self.stats_var.set(f"Total: {total} | Exported: {exported} | Downloaded: {downloaded} | Failed: {failed} | Pending: {pending}")
        
    def validate_inputs(self):
        """Validate user inputs before starting automation"""
        if not self.email_var.get().strip():
            messagebox.showerror("Input Error", "Email cannot be empty")
            return False
            
        if not self.password_var.get().strip():
            messagebox.showerror("Input Error", "Password cannot be empty")
            return False
            
        # Validate date format
        date_format = "%Y/%m/%d"
        try:
            datetime.strptime(self.start_date_var.get(), date_format)
            datetime.strptime(self.end_date_var.get(), date_format)
        except ValueError:
            messagebox.showerror("Input Error", "Dates must be in YYYY/MM/DD format")
            return False
            
        # Validate page range
        if self.start_page_var.get() > self.end_page_var.get():
            messagebox.showerror("Input Error", "Start page cannot be greater than end page")
            return False
            
        return True
        
    def save_settings(self):
        """Save current settings to a file"""
        settings = {
            "email": self.email_var.get(),
            "password": self.password_var.get(),
            "start_date": self.start_date_var.get(),
            "end_date": self.end_date_var.get(),
            "mode": self.mode_var.get(),
            "start_page": self.start_page_var.get(),
            "end_page": self.end_page_var.get(),
            "batch_size": self.batch_size_var.get(),
            "headless": self.headless_var.get(),
            "auto_refresh": self.auto_refresh_var.get()
        }
        
        try:
            with open("qingping_settings.json", "w") as f:
                json.dump(settings, f)
            self.log_message("✅ Settings saved successfully", "success")
        except Exception as e:
            self.log_message(f"❌ Error saving settings: {e}", "error")
            
    def load_settings(self):
        """Load settings from file if it exists"""
        try:
            if os.path.exists("qingping_settings.json"):
                with open("qingping_settings.json", "r") as f:
                    settings = json.load(f)
                
                # Apply settings
                self.email_var.set(settings.get("email", ""))
                self.password_var.set(settings.get("password", ""))
                self.start_date_var.set(settings.get("start_date", datetime.now().strftime("%Y/%m/%d")))
                self.end_date_var.set(settings.get("end_date", datetime.now().strftime("%Y/%m/%d")))
                self.mode_var.set(settings.get("mode", "all"))
                self.start_page_var.set(settings.get("start_page", 1))
                self.end_page_var.set(settings.get("end_page", 12))
                self.batch_size_var.set(settings.get("batch_size", 20))
                self.headless_var.set(settings.get("headless", False))
                self.auto_refresh_var.set(settings.get("auto_refresh", True))
                
                self.log_message("✅ Settings loaded", "success")
                
                # Check if there's progress data to resume
                self.load_progress()
        except Exception as e:
            self.log_message(f"⚠️ Error loading settings: {e}", "warning")
            
    def save_progress(self):
        """Save current progress data to a file"""
        try:
            with open("qingping_progress.json", "w") as f:
                json.dump(self.progress_data, f)
        except Exception as e:
            self.log_message(f"⚠️ Error saving progress: {e}", "warning")
            
    def load_progress(self):
        """Load progress data from file if it exists"""
        try:
            if os.path.exists("qingping_progress.json"):
                with open("qingping_progress.json", "r") as f:
                    self.progress_data = json.load(f)
                
                # Populate the treeview
                for device_name, data in self.progress_data.items():
                    self.progress_tree.insert("", "end", values=(
                        device_name, data["status"], data["timestamp"]
                    ))
                
                # Update stats
                self.update_stats()
                
                # Enable resume button if there are pending devices
                for data in self.progress_data.values():
                    if data["status"] == "Pending":
                        self.resume_button.configure(state=tk.NORMAL)
                        break
                
                self.log_message(f"✅ Loaded progress data for {len(self.progress_data)} devices", "success")
        except Exception as e:
            self.log_message(f"⚠️ Error loading progress: {e}", "warning")
            
    def start_automation(self):
        """Start the automation process in a separate thread"""
        if not self.validate_inputs():
            return
            
        if self.running:
            messagebox.showinfo("Info", "Automation is already running")
            return
            
        # Ask if we should clear existing progress
        if len(self.progress_tree.get_children()) > 0:
            if messagebox.askyesno("Confirm", "Clear existing progress and start fresh?"):
                # Clear progress
                for item in self.progress_tree.get_children():
                    self.progress_tree.delete(item)
                self.progress_data = {}
                self.save_progress()
            else:
                self.log_message("📌 Starting with existing progress data", "info")
        
        # Update UI state
        self.running = True
        self.start_button.configure(state=tk.DISABLED)
        self.resume_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        
        # Clear log
        self.clear_log()
        
        # Start automation in a separate thread
        self.automation_thread = threading.Thread(target=self.run_automation)
        self.automation_thread.daemon = True
        self.automation_thread.start()
        
    def resume_automation(self):
        """Resume the automation process from where it left off"""
        if not self.validate_inputs():
            return
            
        if self.running:
            messagebox.showinfo("Info", "Automation is already running")
            return
            
        # Check if there's progress to resume
        pending_devices = []
        exported_devices = []
        downloaded_devices = []
        
        for device_name, data in self.progress_data.items():
            if data["status"] == "Pending":
                pending_devices.append(device_name)
            elif data["status"] == "Exported":
                exported_devices.append(device_name)
            elif data["status"] == "Downloaded":
                downloaded_devices.append(device_name)
        
        if not pending_devices and not exported_devices:
            messagebox.showinfo("Info", "No pending or exported devices found to resume")
            return
            
        self.log_message(f"🔄 Resuming automation with {len(pending_devices)} pending and {len(exported_devices)} exported devices", "info")
        self.log_message(f"ℹ️ Already downloaded: {len(downloaded_devices)} devices", "info")
        
        # Update UI state
        self.running = True
        self.start_button.configure(state=tk.DISABLED)
        self.resume_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        
        # Start automation in a separate thread
        self.automation_thread = threading.Thread(target=self.run_automation, kwargs={"resume": True})
        self.automation_thread.daemon = True
        self.automation_thread.start()
        
    def stop_automation(self):
        """Stop the automation process"""
        if not self.running:
            return
            
        self.running = False
        self.log_message("🛑 Stopping automation... (may take a moment)", "warning")
        
        # Close the driver if it exists
        if self.automation and self.automation.driver:
            try:
                self.automation.close_driver()
            except Exception as e:
                self.log_message(f"⚠️ Error closing browser: {e}", "warning")
        
        # Enable resume button
        self.resume_button.configure(state=tk.NORMAL)
        
    def reset_ui(self):
        """Reset UI state after automation completes"""
        self.running = False
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        
        # Check if there are pending devices
        for item_id in self.progress_tree.get_children():
            status = self.progress_tree.item(item_id, "values")[1]
            if status == "Pending" or status == "Exported":  # Allow resuming for both Pending and Exported
                self.resume_button.configure(state=tk.NORMAL)
                break
        else:
            self.resume_button.configure(state=tk.DISABLED)
        
    def run_automation(self, resume=False):
        """Run the automation process"""
        try:
            # Get input values
            email = self.email_var.get()
            password = self.password_var.get()
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
            mode = self.mode_var.get()
            start_page = self.start_page_var.get()
            end_page = self.end_page_var.get()
            batch_size = self.batch_size_var.get()
            headless = self.headless_var.get()
            
            # Initialize automation
            self.log_message("🚀 Starting QingPing IoT automation...", "info")
            self.automation = AutoQingPing(email, password)
            self.automation.set_log_callback(self.log_message)
            
            # Login
            self.log_message("🔐 Logging in to QingPing IoT...", "info")
            driver = self.automation.login(email, password, headless=headless)
            
            # Get devices or load from cached file if resuming
            devices_df = None
            csv_path = "devices_cache.csv"
            
            if resume and os.path.exists(csv_path):
                try:
                    self.log_message("📂 Loading devices from cache file...", "info")
                    devices_df = pd.read_csv(csv_path)
                    self.log_message(f"✅ Loaded {len(devices_df)} devices from cache", "success")
                except Exception as e:
                    self.log_message(f"⚠️ Error loading cached devices: {e}", "warning")
                    devices_df = None
            
            if devices_df is None:
                # Get devices fresh from the website
                self.log_message(f"📱 Getting devices (mode: {mode}, pages: {start_page}-{end_page})...", "info")
                devices_df = self.automation.get_devices(driver, mode=mode, start_page=start_page, end_page=end_page)
                
                # Save devices to CSV as cache
                devices_df.to_csv(csv_path, index=False)
                self.log_message(f"💾 Saved {len(devices_df)} devices to {csv_path}", "success")
            
            # Create a list of already processed devices if resuming
            completed_devices = []
            if resume:
                for device_name, data in self.progress_data.items():
                    if data["status"] == "Downloaded" or data["status"] == "Exported":
                        completed_devices.append(device_name)
                
                self.log_message(f"🔄 Skipping {len(completed_devices)} already processed devices", "info")
            
            # Process devices in batches
            total_devices = len(devices_df)
            total_batches = (total_devices + batch_size - 1) // batch_size  # Ceiling division
            
            # Track overall progress
            processed_count = 0
            success_count = 0
            failure_count = 0
            
            # Process each batch
            for batch_idx in range(total_batches):
                if not self.running:
                    self.log_message("🛑 Automation stopped by user", "warning")
                    break
                
                # Calculate batch range
                batch_start = batch_idx * batch_size
                batch_end = min(batch_start + batch_size, total_devices)
                batch_devices = devices_df.iloc[batch_start:batch_end]
                
                self.log_message(f"📦 Processing batch {batch_idx+1}/{total_batches} (devices {batch_start+1}-{batch_end}/{total_devices})", "info")
                
                # Process each device in the batch
                for i, (_, device) in enumerate(batch_devices.iterrows()):
                    if not self.running:
                        break
                    
                    device_name = device["device_name"]
                    device_index = batch_start + i + 1
                    
                    # Skip already completed devices
                    if device_name in completed_devices:
                        self.log_message(f"⏩ [{device_index}/{total_devices}] Skipping {device_name} (already completed)", "info")
                        continue
                    
                    # Update progress
                    self.update_progress(device_name, "Pending")
                    
                    self.log_message(f"🔍 [{device_index}/{total_devices}] Processing device: {device_name}", "info")
                    
                    try:
                        # Check if we need to refresh the page
                        if self.auto_refresh_var.get() and (i % 5 == 0):
                            self.automation.check_and_refresh_if_needed()
                        
                        # Process the device
                        success = self.automation.automate_date_selection(driver, device["link"], start_date, end_date)
                        
                        if success:
                            # Update progress
                            self.update_progress(device_name, "Exported")
                            
                            # Generate and store expected filename
                            device_name_clean = device_name.replace(" ", "_").replace("/", "_").lower()
                            filename_prefix = f"{device_name_clean}_{start_date.replace('/', '')}_{end_date.replace('/', '')}"
                            expected_filename = f"{filename_prefix}.csv"
                            
                            # Store in our device-to-filename mapping
                            self.device_filename_map[expected_filename] = device_name
                            self.log_message(f"🔗 Associated expected file: {expected_filename}", "info")
                            
                            success_count += 1
                        else:
                            self.log_message(f"❌ [{device_index}/{total_devices}] Failed to process: {device_name}", "error")
                            self.update_progress(device_name, "Failed")
                            failure_count += 1
                        
                    except Exception as e:
                        self.log_message(f"❌ [{device_index}/{total_devices}] Error processing {device_name}: {e}", "error")
                        self.update_progress(device_name, "Failed")
                        failure_count += 1
                    
                    processed_count += 1
                
                # Refresh the page between batches
                if self.running and batch_idx < total_batches - 1:
                    self.log_message("🔄 Refreshing page between batches...", "info")
                    try:
                        self.automation.check_and_refresh_if_needed(force=True)
                    except Exception as e:
                        self.log_message(f"⚠️ Error refreshing: {e}", "warning")
            
            if self.running:
                self.log_message(f"📊 Device processing completed. {success_count} devices exported, {failure_count} failed", "success")
                
                # Handle file downloads
                try:
                    self.log_message("🔍 Checking for downloadable files...", "info")
                    
                    # Calculate total pages for file downloads
                    files_per_page = 100  # QingPing typically shows 100 files per page
                    file_pages = (processed_count + files_per_page - 1) // files_per_page
                    
                    # Scrape files from each page
                    all_files = []
                    file_to_device_map = {}  # To keep track of which file belongs to which device
                    
                    for page_num in range(1, file_pages + 1):
                        if not self.running:
                            break
                            
                        self.log_message(f"📄 Scraping files from export page {page_num}/{file_pages}...", "info")
                        filenames = self.automation.scrape_filenames(driver, page_size=files_per_page, page_number=page_num)
                        
                        # Match filenames to devices using our mapping
                        for filename in filenames:
                            # Check for exact match
                            if filename in self.device_filename_map:
                                device_name = self.device_filename_map[filename]
                                file_to_device_map[filename] = device_name
                                self.log_message(f"🔍 Matched file {filename} to device {device_name}", "info")
                            else:
                                # Try partial matching for cases where the server modified the filename format
                                for expected_file, device in self.device_filename_map.items():
                                    # Extract device name part from expected filename (before first underscore)
                                    device_part = expected_file.split('_')[0]
                                    if device_part in filename:
                                        file_to_device_map[filename] = device
                                        self.log_message(f"🔍 Partially matched file {filename} to device {device}", "info")
                                        break
                                else:
                                    self.log_message(f"⚠️ Could not match file {filename} to any device", "warning")
                        
                        if filenames:
                            all_files.extend(filenames)
                            self.log_message(f"✅ Found {len(filenames)} files on page {page_num}", "success")
                            
                            # Click files to download them
                            self.log_message(f"⬇️ Downloading {len(filenames)} files from page {page_num}...", "info")
                            
                            # Track successful downloads
                            for i, filename in enumerate(filenames):
                                try:
                                    self.log_message(f"🔍 Looking for file: {filename}", "info")
                                    
                                    # Try to click the file (using the automation.click_files method for a single file)
                                    single_file_list = [filename]
                                    success = self.automation.click_files(driver, single_file_list, page_number=page_num)
                                    
                                    if success:
                                        self.log_message(f"⬇️ [{i+1}/{len(filenames)}] Downloaded: {filename}", "download")
                                        
                                        # Update the device status if we know which device this file belongs to
                                        if filename in file_to_device_map:
                                            device_name = file_to_device_map[filename]
                                            self.update_progress(device_name, "Downloaded", filename)
                                            self.log_message(f"✅ Updated status for {device_name} to Downloaded", "success")
                                    else:
                                        self.log_message(f"⚠️ Failed to download: {filename}", "warning")
                                    
                                    # Small delay between downloads
                                    time.sleep(1)
                                    
                                except Exception as e:
                                    self.log_message(f"❌ Error downloading file {filename}: {e}", "error")
                            
                        else:
                            self.log_message(f"⚠️ No files found on page {page_num}", "warning")
                    
                    # Summary of downloaded files
                    downloaded_count = sum(1 for item_id in self.progress_tree.get_children() 
                                          if self.progress_tree.item(item_id, "values")[1] == "Downloaded")
                    
                    self.log_message(f"📊 Download summary: {downloaded_count} files downloaded out of {len(all_files)} found", "success")
                    
                except Exception as e:
                    self.log_message(f"❌ Error during file download: {e}", "error")
                
                self.log_message("✅ Automation completed successfully!", "success")
            
            # Close the driver
            time.sleep(2)
            self.automation.close_driver()
            
        except Exception as e:
            self.log_message(f"❌ Error during automation: {e}", "error")
            import traceback
            self.log_message(traceback.format_exc(), "error")
        finally:
            # Update UI state
            self.root.after(0, self.reset_ui)
    
    def exit_app(self):
        """Exit the application"""
        if self.running:
            if messagebox.askyesno("Confirm Exit", "Automation is running. Are you sure you want to exit?"):
                self.stop_automation()
                self.root.destroy()
        else:
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = QingPingGUI(root)
    
    # Set a custom icon (optional)
    try:
        root.iconbitmap("icon.ico")  # Replace with your icon path
    except:
        pass
        
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()