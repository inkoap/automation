from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
from secrets import SECRET_VARS_ROSE
import subprocess
import psutil
import sys
import tkinter as tk
from tkinter import messagebox
import threading
import win32gui
import ctypes

def show_confirmation_dialog(flow_name="Google Meet Automation", wait_seconds=5):
    confirmed = {"status": False}

    def on_confirm():
        confirmed["status"] = True
        root.destroy()

    def on_cancel():
        root.destroy()
        sys.exit(0)

    def countdown():
        nonlocal wait_seconds
        def update():
            nonlocal wait_seconds
            if wait_seconds > 0 and not confirmed["status"]:
                message.set(f"⚡ Automation will run in {wait_seconds} seconds...\nClick Confirm to start now.")
                wait_seconds -= 1
                root.after(1000, update)
            elif not confirmed["status"]:
                confirmed["status"] = True
                root.destroy()
        update()

    # Create dialog window
    root = tk.Tk()
    root.title(f"Running {flow_name}")
    root.geometry("400x180")
    root.resizable(False, False)
    root.configure(bg="#f0f4fc")
    root.protocol("WM_DELETE_WINDOW", on_cancel)

    # Center the dialog
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (400 // 2)
    y = (root.winfo_screenheight() // 2) - (180 // 2)
    root.geometry(f"+{x}+{y}")
    print(x,y)

    # Header label
    header_label = tk.Label(root, text=flow_name, font=("Segoe UI", 14, "bold"), fg="#007bff", bg="#f0f4fc")
    header_label.pack(pady=(15, 5))

    # Countdown message label
    message = tk.StringVar()
    message.set("Preparing automation...")
    message_label = tk.Label(root, textvariable=message, font=("Segoe UI", 11), bg="#f0f4fc", justify="center")
    message_label.pack(pady=(0, 10))

    # Button frame
    button_frame = tk.Frame(root, bg="#f0f4fc")
    button_frame.pack(pady=10)

    # Confirm Button (Green)
    confirm_button = tk.Button(
        button_frame,
        text="✅ Confirm",
        font=("Segoe UI", 10, "bold"),
        bg="#28a745",
        fg="white",
        activebackground="#218838",
        activeforeground="white",
        relief=tk.FLAT,
        width=12,
        command=on_confirm
    )
    confirm_button.pack(side="right", padx=10)

    # Cancel Button (Gray)
    cancel_button = tk.Button(
        button_frame,
        text="❌ Cancel",
        font=("Segoe UI", 10, "bold"),
        bg="#6c757d",
        fg="white",
        activebackground="#5a6268",
        activeforeground="white",
        relief=tk.FLAT,
        width=12,
        command=on_cancel
    )
    cancel_button.pack(side="left", padx=10)

    # Start countdown thread
    threading.Thread(target=countdown, daemon=True).start()

    # Run dialog
    root.mainloop()
    return confirmed["status"]



class GoogleMeetAutomation:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.driver = self.setup_driver()
        # actions = ActionChains(self.driver)

    def setup_driver(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-infobars") # Remove pop-ups
            options.add_argument("--disable-dev-shm-usage") # Avoid crashes in headless mode
            options.add_argument("--ignore-certificate-errors") # Ignore SSL warnings
            options.add_argument("--disable-blink-features=AutomationControlled") # Prevent detection
            options.add_argument("--ignore-certificate-errors") # Bypass SSL warnings
            options.add_experimental_option("excludeSwitches", ["enable-automation"]) # Prevent detection
            options.add_argument("--disable-gpu")  # Fix GPU state error
            options.add_argument("--no-sandbox")  # Prevent security sandbox issues
            options.add_argument("--disable-software-rasterizer")  # 🔹 Fix WebGL GPU crash
            options.add_argument("--disable-features=IsolateOrigins,site-per-process")  # Improve performance
            options.add_experimental_option("detach", True)
            # Login default Chrome profile
            user_data_dir = os.path.join(os.environ["LOCALAPPDATA"], "Google", "Chrome", "User Data")
            options.add_argument(f"--user-data-dir={user_data_dir}")
            options.add_argument("--profile-directory=Default")

            return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        
        except Exception as e:
            print(f"Error setting up Chrome driver: {e}")
            # self.keep_browser_open()

    def open_google_meet(self):
        try:
            self.driver.get("https://meet.google.com/landing")
            time.sleep(1)
        except Exception as e:
            print(f"Error opening Google Meet: {e}")
            # self.keep_browser_open()

    def click_sign_in(self):
        try:
            sign_in_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@data-g-action='sign in']"))
            )
            sign_in_button.click()
        except Exception as e:
            print(f"Error: {e}")
            # self.keep_browser_open()

    def open_google_accounts(self):
        self.driver.get("https://accounts.google.com/")
        WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    def select_profile(self):
        try:
            profile_element = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, f"//div[@data-email='{self.email}']"))
            )
            profile_element.click()
        except Exception as e:
            print(f"Error selecting profile: {e}")
            try:
                no_account_element = self.driver.find_element(By.CSS_SELECTOR, "li.aZvCDf div.AsY17b")
                no_account_element.click()
            except:
                print(f"Error select another profile: {e}")
                pass
                

    def enter_password(self):
        try:
            try:
                # Wait for the 'Tampilkan sandi' parent element to appear
                show_password_div = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//div[contains(@class, 'sfqPrd') and .//div[text()='Tampilkan sandi']]"
                    ))
                )
                
                # Delete the element from the DOM
                self.driver.execute_script("arguments[0].remove();", show_password_div)
                print("✅ 'Tampilkan sandi' checkbox removed.")
            except Exception as e:
                print("⚠️ Failed to remove 'Tampilkan sandi':", e)

            """Enter password and sign in."""
            password_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "Passwd"))
            )
            password_field.send_keys(self.password)
            password_field.send_keys(Keys.RETURN)
        except Exception as e:
            pass

    def enter_credentials(self):
        try:
            email_field = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@type='email']"))
            )
            email_field.send_keys(self.email)
            email_field.send_keys(Keys.RETURN)

            time.sleep(1)

            try:
                # Wait for the 'Tampilkan sandi' parent element to appear
                show_password_div = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//div[contains(@class, 'sfqPrd') and .//div[text()='Show password']]"
                    ))
                )
                
                # Delete the element from the DOM
                self.driver.execute_script("arguments[0].remove();", show_password_div)
                print("✅ 'Tampilkan sandi' checkbox removed.")
            except Exception as e:
                print("⚠️ Failed to remove 'Tampilkan sandi':", e)
            
            password_field = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.NAME, "Passwd"))
            )
            password_field.send_keys(self.password)
            password_field.send_keys(Keys.RETURN)
        except Exception as e:
            print(f"Error: {e}")
            # self.keep_browser_open()

    def keep_browser_open(self):
        """Keep browser open until the user manually exits."""
        input("Press Enter to close the browser...")
        self.driver.quit()

    def check_state(self):
        try:
            elements = self.driver.find_elements(By.XPATH, "//span[@jsname='V67aGc' and contains(text(),'New meeting')]")
            if elements:
                state = 'logged_in'
            else:
                email_field = self.driver.find_element(By.XPATH, "//input[@type='email']")
                if email_field:
                    state = 'not_set'                    
                logo_element = self.driver.find_elements(By.XPATH, "//img[contains(@src, 'Google_Meet_Logo_x2.svg')]")
                if logo_element:
                    state = 'default'

            print(state)
            return state
        except Exception as e:
            print("Error in checking state", e)
            # self.keep_browser_open()

    def check_or_click_profile(self, expected_name=""):
        try:
            expected_name = self.email
            profile_element = self.driver.find_elements(By.XPATH, f"//a[contains(@aria-label, '{expected_name}')]")
            if profile_element:
                print(f"'{expected_name}' is already selected. No action needed.")
            else:
                print(f"'{expected_name}' not found. Clicking to select the correct profile...")
                profile_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//a[contains(@class, 'gb_B') and contains(@class, 'gb_Za')]")
                    )
                )
                profile_button.click()
                print("Profile selection menu opened!")

                time.sleep(1)
                logout_url = "https://accounts.google.com/Logout?"
                print("🔄 Redirecting to logout page to switch accounts...")
                state = 'logout'
                self.driver.get(logout_url)

            return state
        except Exception as e:
            print(f"Error: {e}")
            # self.keep_browser_open()

    def check_logged_in_accounts(self):
        try:
            print("Masuk")
            sign_out_element = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[@class='SedFmc']//span[contains(text(), 'Sign out')]"))
                )
            if sign_out_element:
                sign_out_element.click()
                print("Logged out.")
        except Exception as e:
            print(f"Error: {e}")
            # self.keep_browser_open()

def terminate_other_tv_automation():
    processes = []

    for process in psutil.process_iter(attrs=['pid', 'name', 'create_time']):
        try:
            if process.info['name'].lower() == 'gmeet-automation.exe':
                processes.append((process.info['pid'], process.info['create_time']))
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    if len(processes) > 2:
        processes.sort(key=lambda x: x[1]) 
        latest_pids = [p[0] for p in processes[-2:]] 

        for pid, _ in processes:
            if pid not in latest_pids: 
                try:
                    print(f"Terminating old instance of 'gmeet-automation.exe' (PID: {pid})...")
                    psutil.Process(pid).terminate()  
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
    else:
        print("No excess instances found. No action needed.")

def is_process_running(process_name):
    for process in psutil.process_iter(attrs=['name']):
        if process.info['name'].lower() == process_name.lower():
            return True
    return False

def terminate_if_running(process_name):
    if is_process_running(process_name):
        print(f"Terminating {process_name}...")
        subprocess.run(f"taskkill /F /IM {process_name} /T", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        print(f"{process_name} is not running.")

def disconnect_if_ssid_exist():
    try:
        result = subprocess.run("netsh wlan show interface", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        if result.returncode == 0:
            if "SSID" in result.stdout:
                print("SSID found. Disconnecting from Wi-Fi...")
                disconnect = subprocess.run("netsh wlan disconnect", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                if disconnect.returncode == 0:
                    print("Wi-Fi disconnected successfully.")
                else:
                    print("Failed to disconnect Wi-Fi.")
                    print(disconnect.stderr)
            else:
                print("No SSID found. Wi-Fi is not connected.")
        else:
            print("Failed to retrieve network interface details.")
            print(result.stderr)

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        result = subprocess.run("ping -n 3 google.com", shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if "Reply from" in result.stdout:
            pass
        else:
            print("No internet connection.")
            print(result.stderr)
            sys.exit(1)

# Run the Automation
if __name__ == "__main__":
    terminate_other_tv_automation()
    show_confirmation_dialog("Google Meet Automation", wait_seconds=5)
    terminate_if_running("chromedriver.exe")
    terminate_if_running("chrome.exe")
    disconnect_if_ssid_exist()

    EMAIL = SECRET_VARS_ROSE['EMAIL']
    PASSWORD = SECRET_VARS_ROSE['PASSWORD']

    gmeet_bot = GoogleMeetAutomation(EMAIL, PASSWORD)
    gmeet_bot.open_google_meet()
    while True:
        check_state = gmeet_bot.check_state()
        if check_state == 'logged_in':
            check_profile = gmeet_bot.check_or_click_profile()
            if check_profile == 'logout':
                time.sleep(1)
                gmeet_bot.open_google_meet()
                gmeet_bot.click_sign_in()
                gmeet_bot.select_profile()
                gmeet_bot.enter_password()
                # gmeet_bot.gmeet_bot.restore_chrome_window()
                break
            # gmeet_bot.gmeet_bot.restore_chrome_window()
            break

        elif check_state == 'not_set':
            gmeet_bot.enter_credentials()
            # gmeet_bot.gmeet_bot.restore_chrome_window()
            break
        
        else:
            gmeet_bot.click_sign_in()
            gmeet_bot.select_profile()
            gmeet_bot.enter_password()
            gmeet_bot.enter_credentials()
            # gmeet_bot.gmeet_bot.restore_chrome_window()
            break
