import tkinter as tk
from tkinter import ttk, messagebox, StringVar, IntVar
import threading
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os

class AutoQingPing:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.driver = None
        self.log_callback = print  # Default to print

    def set_log_callback(self, callback):
        """Set a callback function for logging"""
        self.log_callback = callback

    def log(self, message):
        """Log a message using the callback"""
        self.log_callback(message)

    def login(self, email, password):
        options = Options()
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_experimental_option("detach", True)  # Keep browser open

        self.driver = webdriver.Chrome(options=options)
        self.driver.get("https://qingpingiot.com/user/login")
        wait = WebDriverWait(self.driver, 30)

        self.log("🔐 Opening login page...")

        # Click on the email login tab
        email_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-node-key='email']")))
        email_tab.click()

        # Fill in email
        email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Email']")))
        email_input.clear()
        email_input.send_keys(email)

        # Press Tab to move to password input
        email_input.send_keys(Keys.TAB)
        time.sleep(0.5)  # Small pause to ensure focus shifts

        # Type the password
        webdriver.ActionChains(self.driver).send_keys(password).perform()
        time.sleep(1)

        # Click the "Log In" button
        login_button = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[@type='button' and .//span[text()='Log In']]"
        )))
        login_button.click()

        self.log("✅ Login button clicked. Waiting for page load or verification (10s)...")
        time.sleep(10)

        return self.driver
    
    def get_devices(self, driver, mode="all", device_count=20):
        SCROLL_PAUSE_TIME = 1

        if device_count > 20:
            self.log(f"📜 Scrolling to load ~{device_count} devices...")

            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(SCROLL_PAUSE_TIME)

                new_height = driver.execute_script("return document.body.scrollHeight")
                cards = driver.find_elements(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-cardItem.antd-pro-pages-welcome-less-index-cardItemHiddenStatus")

                if len(cards) >= device_count or new_height == last_height:
                    break

                last_height = new_height

        # Grab all card elements after scroll
        cards = driver.find_elements(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-cardItem.antd-pro-pages-welcome-less-index-cardItemHiddenStatus")

        # Filter based on mode
        if mode == "online":
            filtered = [card for card in cards if "offlineCardItem" not in card.get_attribute("class")]
        elif mode == "offline":
            filtered = [card for card in cards if "offlineCardItem" in card.get_attribute("class")]
        else:
            filtered = cards

        self.log(f"📦 Found {len(filtered)} device(s) with mode = '{mode}'")

        data = []

        for i, card in enumerate(filtered[:device_count]):
            try:
                device_name = card.find_element(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-deviceNameText").text
                update_time = card.find_element(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-updateTime").text
                link = card.find_element(By.TAG_NAME, "a").get_attribute("href")

                titles = card.find_elements(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-title")
                values = card.find_elements(By.CSS_SELECTOR, ".antd-pro-pages-welcome-less-index-dataShow p")

                data_points = {}
                for title, value in zip(titles, values):
                    data_points[title.text] = value.text

                card_data = {
                    "device_name": device_name,
                    "update_time": update_time,
                    "link": link,
                    **data_points
                }

                data.append(card_data)
                self.log(f"[{i+1}] Scraped: {device_name}")

            except Exception as e:
                self.log(f"[CARD ERROR] {e}")
                continue

        df = pd.DataFrame(data)
        return df

    def automate_date_selection(self, driver, url, start_date, end_date):
        driver.get(url)
        wait = WebDriverWait(driver, 2)

        self.log("📄 Opening detail page...")

        try:
            # Step 1: Click Export button
            export_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span.antd-pro-pages-history-chart-chart-export_text"))
            )
            export_button.click()
            self.log("📦 Export button clicked.")

            # Step 2: Wait for modal to show
            modal_wrapper = wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ant-modal-content"))
            )
            self.log("🪟 Export modal appeared.")

            # Step 3: Open date range picker
            range_picker_wrapper = modal_wrapper.find_element(By.CLASS_NAME, "ant-picker-range")
            range_picker_wrapper.click()
            self.log("📆 Opened date range picker.")

            # Step 4: Type the start and end dates using send_keys
            formatted_start = start_date
            formatted_end = end_date

            inputs = modal_wrapper.find_elements(By.CSS_SELECTOR, "input[placeholder]")
            if len(inputs) >= 2:
                start_input, end_input = inputs[:2]

                # Clear and enter start date
                start_input.send_keys(Keys.CONTROL + "a")
                start_input.send_keys(Keys.BACKSPACE)
                start_input.send_keys(formatted_start)

                # Clear and enter end date
                end_input.send_keys(Keys.CONTROL + "a")
                end_input.send_keys(Keys.BACKSPACE)
                end_input.send_keys(formatted_end)

                # Press ENTER to confirm date range
                end_input.send_keys(Keys.ENTER)

                self.log(f"📅 Dates filled: {formatted_start} → {formatted_end}")
            else:
                raise Exception("⚠️ Couldn't locate both date input fields.")

            # Step 5: Click Confirm button
            confirm_button = modal_wrapper.find_element(By.CSS_SELECTOR, "button.ant-btn.ant-btn-primary")
            confirm_button.click()
            self.log("✅ Date selection confirmed.")

        except Exception as e:
            self.log(f"❌ Error during automation: {e}")

    def scrape_filenames(self, driver, device_count, page_number=1):
        url = "https://qingpingiot.com/export"
        driver.get(url)
        wait = WebDriverWait(driver, 10)

        if device_count >= 100:
            page_size_dropdown = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[@class='ant-select-selection-item' and contains(@title, '/ page')]")
            ))
            page_size_dropdown.click()

            # Wait for dropdown options to load
            option_100 = wait.until(EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@class, 'ant-select-item-option') and .//div[text()='100 / page']]")
            ))
            
            # Scroll into view and click the "100 / page" option
            driver.execute_script("arguments[0].scrollIntoView(true);", option_100)
            time.sleep(0.5)  # small delay to ensure it's interactable
            option_100.click()
            
            # Navigate to the desired page
            if page_number > 1:
                try:
                    # Wait for the pagination controls to load
                    wait.until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'ant-pagination')]")))

                    # Click on the desired page number
                    page_button = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, f"//a[@rel='nofollow' and text()='{page_number}']"))
                    )
                    self.log(f"➡️ Navigating to page {page_number}...")
                    page_button.click()

                    # Optional: wait for page to fully load after clicking
                    time.sleep(2)
                    self.log(f"✅ Now on page {page_number}.")

                except Exception as e:
                    self.log(f"❌ Failed to go to page {page_number}: {e}")

        try:
            # Wait until at least one relevant <span> appears
            wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '.csv') or contains(text(), '.xlsx')]")))

            # Find all .csv and .xlsx span elements
            file_elements = driver.find_elements(By.XPATH, "//span[contains(text(), '.csv') or contains(text(), '.xlsx')]")

            # Extract text
            filenames = [elem.text.strip() for elem in file_elements if elem.text.strip().endswith(('.csv', '.xlsx'))]

            self.log("✅ Found filenames:")
            for name in filenames:
                self.log(f"- {name}")

            return filenames

        except Exception as e:
            self.log(f"❌ Error while scraping filenames: {e}")
            return []
        
    def click_files(self, driver, filenames, page_number=1):
        url = "https://qingpingiot.com/export"
        driver.get(url)
        wait = WebDriverWait(driver, 30)
        # Navigate to the desired page
        if page_number > 1:
            try:
                # Wait for the pagination controls to load
                wait.until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class, 'ant-pagination')]")))

                # Click on the desired page number
                page_button = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//a[@rel='nofollow' and text()='{page_number}']"))
                )
                self.log(f"➡️ Navigating to page {page_number}...")
                page_button.click()

                # Optional: wait for page to fully load after clicking
                time.sleep(2)
                self.log(f"✅ Now on page {page_number}.")

            except Exception as e:
                self.log(f"❌ Failed to go to page {page_number}: {e}")
        time.sleep(15)  # Wait for the page to load
        for i in range (len(filenames)):
            filename = filenames[len(filenames)-1-i]
            try:
                # Wait for the specific filename element to appear and be clickable
                target = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//span[text()='{filename}']"))
                )
                self.log(f"🖱️ Clicking on file: {filename}")
                target.click()
                self.log(f"✅ Clicked: {filename}")

                # Optional: Wait between clicks (e.g., for download to start)
                time.sleep(2)

            except Exception as e:
                self.log(f"❌ Failed to click on {filename}: {e}")

    def close_driver(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.log("🔒 Browser closed.")


class QingPingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("QingPing IoT Automation")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        self.automation = None
        self.running = False
        
        # Create variables
        self.email_var = StringVar(value="")
        self.password_var = StringVar(value="")
        self.start_date_var = StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.end_date_var = StringVar(value=datetime.now().strftime("%Y/%m/%d"))
        self.mode_var = StringVar(value="all")
        self.device_count_var = IntVar(value=30)
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.tab_control = ttk.Notebook(self.main_frame)
        
        # Settings tab
        self.settings_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.settings_tab, text="Settings")
        
        # Log tab
        self.log_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.log_tab, text="Log")
        
        self.tab_control.pack(fill=tk.BOTH, expand=True)
        
        # Setup settings tab
        self.setup_settings_tab()
        
        # Setup log tab
        self.setup_log_tab()
        
        # Setup bottom buttons
        self.setup_buttons()
        
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
        
        ttk.Label(device_frame, text="Est. Device Count:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(device_frame, from_=1, to=1000, textvariable=self.device_count_var, width=10).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
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
        
        # Make the log text read-only
        self.log_text.configure(state=tk.DISABLED)
        
    def setup_buttons(self):
        button_frame = ttk.Frame(self.main_frame, padding=10)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.start_button = ttk.Button(button_frame, text="Start Automation", command=self.start_automation)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop Automation", command=self.stop_automation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
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
            
        return True
        
    def start_automation(self):
        """Start the automation process in a separate thread"""
        if not self.validate_inputs():
            return
            
        if self.running:
            messagebox.showinfo("Info", "Automation is already running")
            return
            
        # Update UI state
        self.running = True
        self.start_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        
        # Clear log
        self.clear_log()
        
        # Start automation in a separate thread
        self.automation_thread = threading.Thread(target=self.run_automation)
        self.automation_thread.daemon = True
        self.automation_thread.start()
        
    def stop_automation(self):
        """Stop the automation process"""
        if not self.running:
            return
            
        self.running = False
        self.log_message("Stopping automation...", "info")
        
        # Close the driver if it exists
        if self.automation and self.automation.driver:
            try:
                self.automation.close_driver()
            except Exception as e:
                self.log_message(f"Error closing browser: {e}", "error")
        
        # Update UI state
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        
    def run_automation(self):
        """Run the automation process"""
        try:
            # Get input values
            email = self.email_var.get()
            password = self.password_var.get()
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
            mode = self.mode_var.get()
            device_count = self.device_count_var.get()
            
            # Initialize automation
            self.log_message("Starting QingPing IoT automation...", "info")
            self.automation = AutoQingPing(email, password)
            self.automation.set_log_callback(self.log_message)
            
            # Login
            self.log_message("Logging in to QingPing IoT...", "info")
            driver = self.automation.login(email, password)
            
            # Get devices
            self.log_message(f"Getting devices (mode: {mode}, count: {device_count})...", "info")
            devices_df = self.automation.get_devices(driver, mode=mode, device_count=device_count)
            
            # Save devices to CSV
            csv_path = "devices.csv"
            devices_df.to_csv(csv_path, index=False)
            self.log_message(f"Saved device list to {csv_path}", "success")
            
            # Process each device
            for i in range(len(devices_df)):
                if not self.running:
                    break
                    
                device_name = devices_df.iloc[i]["device_name"]
                self.log_message(f"🔍 Processing device: {device_name}", "info")
                device_link = devices_df.iloc[i]["link"]
                time.sleep(1)  # Small delay to avoid overwhelming the server
                self.automation.automate_date_selection(driver, device_link, start_date, end_date)
            
            # Calculate pagination
            total_pages = len(devices_df) // 100 + 1
            remaining_devices = len(devices_df) % 100
            
            # Scrape filenames from each page
            file_names = {}
            for page_number in range(1, total_pages + 1):
                if not self.running:
                    break
                filenames = self.automation.scrape_filenames(driver, len(devices_df), page_number)
                file_names[page_number] = filenames
            
            # Click files to download them
            for page_number in range(1, total_pages + 1):
                if not self.running:
                    break
                    
                if page_number == total_pages:
                    filenames = file_names[page_number][:remaining_devices]
                else: 
                    filenames = file_names[page_number]
                
                self.automation.click_files(driver, filenames, page_number)
            
            self.log_message("✅ All files clicked. You can now check your downloads folder.", "success")
            
            # Close the driver
            time.sleep(2)
            self.automation.close_driver()
            
            self.log_message("Automation completed successfully!", "success")
            
        except Exception as e:
            self.log_message(f"Error during automation: {e}", "error")
            import traceback
            self.log_message(traceback.format_exc(), "error")
        finally:
            # Update UI state
            self.root.after(0, self.reset_ui)
    
    def reset_ui(self):
        """Reset UI state after automation completes"""
        self.running = False
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
    
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