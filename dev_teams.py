from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from cryptography.fernet import Fernet
import os
import sys
import subprocess
import time
from secrets import SECRET_VARS_TEAMS
import psutil
import tkinter as tk
from tkinter import messagebox
import threading

def show_confirmation_dialog(flow_name="Teams Automation", wait_seconds=5):
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

class TeamsAutomation:
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

    def open_ms_teams(self):
        try:
            self.driver.get("https://teams.live.com/")
            time.sleep(1)
        except Exception as e:
            print(f"Error opening Teams: {e}")
            # self.keep_browser_open()

    def enter_email(self):
        try:
            time.sleep(1)
            print("input email")
            email_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "usernameEntry"))
            )
            print("email field is located")
            email_field.send_keys(self.email)
            email_field.send_keys(Keys.RETURN)
        except Exception as e:
            print(f"Error when input email {e}")

    def enter_password(self):
        try:
            try:
                # Wait until the button is present in DOM
                button = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Show password']"))
                )

                # Delete the element using JavaScript
                self.driver.execute_script("arguments[0].remove();", button)
                print("✅ 'Show password' button removed successfully.")

            except Exception as e:
                print("⚠️ Error removing 'Show password' button:", e)
            # Wait for password field first
            password_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "passwordEntry"))
            )
            # Enter password
            password_field.send_keys(self.password)

            password_field.send_keys(Keys.RETURN)

        except Exception as e:
            print(f"❌ Error when inputting password: {e}")

    
    def stay_sign_in(self):
        try:
            yes_button = WebDriverWait(self.driver, 3).until(
                EC.element_to_be_clickable((By.ID, "acceptButton"))
            )
            # Click the button
            yes_button.click()
            print("Clicked the 'Yes' button!")
        except Exception as e:
            print("Button not found or not clickable:", e)
            yes_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='primaryButton']"))
            )
            # Click the button
            yes_button.click()
            print("Clicked the 'Yes' button!")

    def click_sign_in(self):
        try:
            sign_in_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-tid='sign-in-button']"))
            )

            # Click the button
            sign_in_button.click()
            print("Clicked the 'Sign in' button!")

        except Exception as e:
            print("Button not found or not clickable:", e)

    def keep_browser_open(self):
        input("Press Enter to close the browser...")
        self.driver.quit()

    def check_state(self):
        try:
            elements = self.driver.find_elements(By.XPATH, "//button[@data-tid='sign-in-button']")
            if elements:
                state = 'default'
            else:
                email_element = self.driver.find_elements(By.ID, "usernameEntry")
                if email_element:
                    state = 'enter_email'
                else:
                    meet_element = self.driver.find_elements(By.ID, "40472f6e-248f-4599-842c-ff3ed8f0ae34")
                    if meet_element:
                        state = 'logged_in'
                    else:
                        state = 'unknown'
            print(state)
            return state
        except Exception as e:
            print("Error in checking state", e)
            # self.keep_browser_open()

    def check_or_click_profile(self):
        try:
            profile_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "idna-me-control-avatar-trigger"))
            )
            profile_button.click()

            try:
                email_lower = self.email.lower()
                email_span = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, f"//span[text()='{email_lower}']"))
                )
                if email_span:
                    print("✅ Email found, no action needed.")
                    meet_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "40472f6e-248f-4599-842c-ff3ed8f0ae34"))
                    )
                    meet_button.click()
                    end_state = True

            except Exception as e:
                sign_out_button = WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable((By.ID, "me-control-sign-out-button"))
                )
                sign_out_button.click()
                time.sleep(3)
                print("waiting for confirm pop up")
                sign_out_confirm_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@data-tid='sign-out-confirm']"))
                )
                sign_out_confirm_button.click()
                end_state = False
                print("Not found")
            return end_state
        except Exception as e:
            print(f"Error: {e}")
            # self.keep_browser_open()

    def check_logged_in_accounts(self):
        try:
            sign_out_element = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[@class='SedFmc']//span[contains(text(), 'Sign out')]"))
                )
            if sign_out_element:
                sign_out_element.click()
                print("Logged out.")
        except Exception as e:
            print(f"Error: {e}")
            # self.keep_browser_open()

    def actions(self, check_state):
        try:
            end_state = False
            if check_state == 'default':
                time.sleep(1)
                print("default")
                self.click_sign_in()
                time.sleep(1)
                print("input email")
                self.enter_email()
                print("enter password")
                self.enter_password()
                print("stay sign in")
                self.stay_sign_in()

            elif check_state == 'enter_email':
                time.sleep(1)
                self.enter_email()
                print("enter password")
                self.enter_password()
                print("stay sign in")
                self.stay_sign_in()

            elif check_state == 'logged_in':
                end_state = self.check_or_click_profile()

                print(end_state)
            return end_state
        except Exception as e:
            print("Error when to do the actions", e)
    
    def automation_teams(self):
        self.open_ms_teams()
        while True:
            time.sleep(1)
            check_state = self.check_state()
            print(check_state)
            end_state = self.actions(check_state)
            if end_state:
                break
            
def load_or_generate_key():
    key_path = "encryption_key_teams.key"
    if not os.path.exists(key_path):
        key = Fernet.generate_key()
        with open(key_path, "wb") as key_file:
            key_file.write(key)
        print("New encryption key generated. Save this securely!")
    else:
        with open(key_path, "rb") as key_file:
            key = key_file.read()
    return key

# Load encryption key
key = load_or_generate_key()
cipher = Fernet(key)

def encrypt_data(data):
    return cipher.encrypt(data.encode()).decode()

def decrypt_data(encrypted_text):
    try:
        return cipher.decrypt(encrypted_text.encode()).decode()
    except Exception as e:
        sys.exit(1)

with open("encrypted_credentials_teams.txt", "r") as f:
    encrypted_email, encrypted_password = f.read().split("\n")

def terminate_other_tv_automation():
    current_pid = os.getpid()

    processes = []

    for process in psutil.process_iter(attrs=['pid', 'name', 'create_time']):
        try:
            if process.info['name'].lower() == 'teams-automation.exe':
                processes.append((process.info['pid'], process.info['create_time']))
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    if len(processes) > 2:
        processes.sort(key=lambda x: x[1]) 
        latest_pids = [p[0] for p in processes[-2:]] 

        for pid, _ in processes:
            if pid not in latest_pids: 
                try:
                    print(f"Terminating old instance of 'teams-automation.exe' (PID: {pid})...")
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
    show_confirmation_dialog("Teams Automation", wait_seconds=5)
    terminate_if_running("chromedriver.exe")
    terminate_if_running("chrome.exe")
    disconnect_if_ssid_exist()
    EMAIL = SECRET_VARS_TEAMS['EMAIL']
    PASSWORD = SECRET_VARS_TEAMS['PASSWORD']

    encrypted_email = encrypt_data(EMAIL)
    encrypted_password = encrypt_data(PASSWORD)

    with open("encrypted_credentials_teams.txt", "w") as f:
        f.write(f"{encrypted_email}\n{encrypted_password}")

    teams_bot = TeamsAutomation(EMAIL, PASSWORD)
    # teams_bot.open_teams()
    teams_bot.automation_teams()