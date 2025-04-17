
import time
import os
import sys
from secrets import SECRET_VARS_ROSE
from cryptography.fernet import Fernet
import subprocess
import psutil
import ctypes
import tkinter as tk
from tkinter import messagebox
import threading
# Notification
# from PIL import Image, ImageTk

def show_confirmation_dialog(flow_name="Webex Automation", wait_seconds=5):
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
    # root.lift()
    # root.attributes("-topmost", True)
    root.focus_force()

    root.mainloop()
    return confirmed["status"]

def show_native_popup(title="Webex Automation", message="Automation will run now. Proceed?", wait_seconds=5):
    MB_OKCANCEL = 0x00000001
    IDOK = 1
    IDCANCEL = 2
    MB_SYSTEMMODAL = 0x00001000

    user32 = ctypes.windll.user32

    # Define the MessageBoxTimeoutW function
    MessageBoxTimeoutW = user32.MessageBoxTimeoutW
    MessageBoxTimeoutW.argtypes = [
        ctypes.wintypes.HWND,     # hWnd
        ctypes.wintypes.LPCWSTR,  # lpText
        ctypes.wintypes.LPCWSTR,  # lpCaption
        ctypes.wintypes.UINT,     # uType
        ctypes.wintypes.WORD,     # wLanguageId
        ctypes.wintypes.DWORD     # dwMilliseconds
    ]
    MessageBoxTimeoutW.restype = ctypes.wintypes.INT

    full_message = f"{message}\n\nAutomation will start in {wait_seconds} seconds..."

    # Show popup with timeout
    result = MessageBoxTimeoutW(
        0,
        full_message,
        title,
        MB_OKCANCEL | MB_SYSTEMMODAL,
        0,
        wait_seconds * 1000
    )

    if result == IDCANCEL:
        print("❌ User cancelled the automation.")
        sys.exit(0)
    
    # OK clicked or timeout (any other return)
    print("✅ Proceeding with automation...")

class WebexAutomation:
    def __init__(self):
        self.email_webex, self.password_webex = self.load_credentials()
        self.webex_path = os.path.join(os.getenv("APPDATA"), r"Microsoft\Windows\Start Menu\Programs\Webex\Webex.lnk")
        self.auto_ids = {
            "back_button_auto_id": "MainWindow.OnboardingView.onboardingScreen.backButton",
            "email_field_auto_id": "MainWindow.OnboardingView.onboardingScreen.stackedWidget.normalViewsPage.onboardingStackWidget.OnboardingSignInUpEmailWidget.dataStackedWidget.dataWidget.enterEmailField.textInput",
            "email_field_auto_id_2": "MainWindow.OnboardingView.onboardingScreen.stackedWidget.normalViewsPage.onboardingStackWidget.OnboardingSignInUpEmailWidget.dataStackedWidget.dataWidget.enterEmailField.UTTextField.textInput",
            "sign_in_button_auto_id": "MainWindow.OnboardingView.onboardingScreen.stackedWidget.normalViewsPage.onboardingStackWidget.OnboardingStartWidget.dataStackedWidget.dataWidget.nextButton",
            "profile_auto_id":'MainWindow.ConversationsForm.topLevelStack.mainAreasWidget.applicationHeaderView.applicationHeaderWidget.leftWidget.userInfoWidget.avatarButton',
            "configure_button_auto_id" : "MainWindow.ConversationsForm.topLevelStack.mainAreasWidget.applicationHeaderView.applicationHeaderWidget.leftWidget.userInfoWidget.configureButton",
            "sign_out_button_auto_id": "MainWindow.ConversationsForm.topLevelStack.mainAreasWidget.applicationHeaderView.applicationHeaderWidget.leftWidget.userInfoWidget.avatarButton.applicationMenu.signOutAction",
            "next_button_auto_id": "MainWindow.OnboardingView.onboardingScreen.stackedWidget.normalViewsPage.onboardingStackWidget.OnboardingSignInUpEmailWidget.dataStackedWidget.dataWidget.nextButton",
            "join_a_meeting_auto_id": "MainWindow.OnboardingView.onboardingScreen.stackedWidget.normalViewsPage.onboardingStackWidget.OnboardingStartWidget.dataStackedWidget.dataWidget.joinMeetingButton",
            "tab_selector_auto_id": "MainWindow.ConversationsForm.topLevelStack.mainAreasWidget.tabSelector.navigationWidget.meetTabSelector",
            "personal_link_auto_id": "MainWindow.ConversationsForm.topLevelStack.mainAreasWidget.sidePanelSplitter.rightSideStack.UnifiedMeetingViewWidget.mainStackedWidget.calendarMainWidget.meetingsHeadWidget.MeetingsHeadView.collapsedLine.copyButton",
            "sign_out_action_auto_id": "MainWindow.ConversationsForm.topLevelStack.mainAreasWidget.applicationHeaderView.applicationHeaderWidget.leftWidget.userInfoWidget.avatarButton.applicationMenu.signOutAction"
        }

    def load_or_generate_key(self):
        """Load encryption key or generate a new one if it doesn't exist."""
        key_path = "encryption_key.key"
        if not os.path.exists(key_path):
            key = Fernet.generate_key()
            with open(key_path, "wb") as key_file:
                key_file.write(key)
            print("🔑 New encryption key generated. Save this securely!")
        else:
            with open(key_path, "rb") as key_file:
                key = key_file.read()
        return key

    def load_credentials(self):
        """Load credentials securely from an encrypted file."""
        key = self.load_or_generate_key()
        self.cipher = Fernet(key)

        credentials_file = "encrypted_credentials.txt"
        if not os.path.exists(credentials_file):
            email = SECRET_VARS_ROSE["EMAIL"]
            password = SECRET_VARS_ROSE["PASSWORD"]

            encrypted_email = self.encrypt_data(email)
            encrypted_password = self.encrypt_data(password)

            with open(credentials_file, "w") as f:
                f.write(f"{encrypted_email}\n{encrypted_password}")

            print("🔐 Credentials encrypted and saved!")
            sys.exit(0)

        with open(credentials_file, "r") as f:
            encrypted_email, encrypted_password = f.read().split("\n")

        return self.decrypt_data(encrypted_email), self.decrypt_data(encrypted_password)

    
    def encrypt_data(self, data):
        """Encrypt sensitive data."""
        return self.cipher.encrypt(data.encode()).decode()

    def decrypt_data(self, encrypted_text):
        """Decrypt stored credentials."""
        try:
            return self.cipher.decrypt(encrypted_text.encode()).decode()
        except Exception as e:
            print(f"Decryption error: {e}")
            sys.exit(1)

    def set_window_resolution(self, window, width=1280, height=720):
        try:
            hwnd = window.handle

            # Get screen size
            user32 = ctypes.windll.user32
            screen_width = user32.GetSystemMetrics(0)
            screen_height = user32.GetSystemMetrics(1)

            x = (screen_width - width) // 2
            y = (screen_height - height) // 2

            SWP_NOZORDER = 0x0004
            SWP_NOACTIVATE = 0x0010

            result = ctypes.windll.user32.SetWindowPos(hwnd, 0, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE)

            if result:
                print(f"✅ Webex window moved to center with resolution {width}x{height}")
            else:
                print("[!] SetWindowPos failed")

        except Exception as e:
            print(f"[!] Failed to set window resolution: {e}")
    
    # def open_webex(self):
    #     try:
    #         if sys.platform.startswith('win'):
    #             if os.path.exists(self.webex_path):
    #                 os.system(f'"{self.webex_path}"')

    #                 while True:
    #                     webex_window, state = self.check_state()
    #                     print(webex_window)
                        
    #                     try:
    #                         if not webex_window.is_maximized():
    #                             try:
    #                                 webex_window.maximize()
    #                                 self.set_window_resolution(webex_window, width=9999, height=9999)
    #                             except Exception as e:
    #                                 print(e)
    #                     except:
    #                         pass
    #                     end_state, state = self.actions(webex_window, state)
    #                     if end_state:
    #                         break
    #     except Exception as e:
    #         print(f"Error: {e}")

    def open_webex(self):
        try:
            if sys.platform.startswith('win'):
                if os.path.exists(self.webex_path):
                    os.system(f'"{self.webex_path}"')
                    self.show_notification(message="Automation is running...")

                    while True:
                        webex_window, state = self.check_state()
                        try:
                            if state in ["default", "email_input", "password"]:
                                try:
                                    # webex_window.minimize()
                                    hwnd = webex_window.handle

                                    # Move window to x=0, y=0, width=1, height=1
                                    ctypes.windll.user32.MoveWindow(hwnd, -9000, -9000, 10, 10, True)
                                except Exception as e:
                                    print("Error minimizing window:", e)
                        except Exception as e:
                            print("Error brow pas minimize", e)

                        print(state)

                        end_state, state = self.actions(webex_window, state)
                        if end_state:
                            self.show_notification(message="Automation is completed", duration=1)
                            try:
                                if not webex_window.is_maximized():
                                    try:
                                        webex_window.maximize()
                                        self.set_window_resolution(webex_window, width=9999, height=9999)
                                    except Exception as e:
                                        print(e)
                            except:
                                pass
                            
                            break
        except Exception as e:
            print(f"Error: {e}")

    # def show_notification(self, title="Webex Automation", message="Automation is running...",duration=99):
    #     notification.notify(
    #         title=title,
    #         message=message,
    #         timeout=int(duration)  # in seconds
    #     )

    def show_notification(self, title="Webex Automation", message="Automation is running...", duration=90):
        def notifier():
            root = tk.Tk()
            root.overrideredirect(True)
            root.attributes("-topmost", True)
            root.configure(bg="#000000")
            root.attributes("-alpha", 0.95)  # Slight transparency

            # Position
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            width = 360
            height = 120
            x = screen_width - width - 20
            y = screen_height - height - 60
            root.geometry(f"{width}x{height}+{x}+{y}")

            # Shadow frame
            shadow = tk.Frame(root, bg="#444", padx=2, pady=2)
            shadow.pack(fill="both", expand=True)

            # Notification box
            box = tk.Frame(shadow, bg="#ffffff", padx=20, pady=15, bd=0, highlightthickness=0)
            box.pack(fill="both", expand=True)

            # Title
            tk.Label(
                box, text=title,
                font=("Segoe UI", 11, "bold"),
                fg="#111", bg="#ffffff",
                anchor="center", justify="center"
            ).pack(fill="x")

            # Message
            tk.Label(
                box, text=message,
                font=("Segoe UI", 10),
                fg="#333", bg="#ffffff",
                wraplength=300, justify="center"
            ).pack(fill="x")

            # Auto-close after `duration` seconds
            root.after(int(duration * 1000), root.destroy)
            root.mainloop()

        threading.Thread(target=notifier, daemon=True).start()

    # def show_notification(self, title="Automation Running", message="Webex automation is in progress...", duration=900):
    #     try:
    #         self.notifier = win10toast.ToastNotifier()
    #         self.notifier.show_toast(title, message, duration=int(duration), threaded=True)
    #     except Exception as e:
    #         print("Error when showing notification", e)

    def check_state(self):
        state = "unknown"  # Default state if no conditions match

        if Desktop(backend="uia").window(title="Webex").exists():
            webex_window = Desktop(backend="uia").window(title="Webex")
            if webex_window.child_window(title="Sign In - Webex", control_type="Pane").exists():
                state = 'password'
            elif webex_window.child_window(auto_id=self.auto_ids["profile_auto_id"], control_type='Custom').exists():
                state = 'logged_in'
            elif webex_window.child_window(title=" Join a meeting", auto_id=self.auto_ids["join_a_meeting_auto_id"], control_type="Button").exists():
                state = 'default'
            else:
                state = 'unknown'

        # Check if "Sign in - Webex" window exists
        elif Desktop(backend="uia").window(title="Sign in -  Webex").exists():
            webex_window = Desktop(backend="uia").window(title="Sign in -  Webex")
            state = 'email_input'

        elif Desktop(backend="uia").window(title="Join a meeting -  Webex").exists():
            webex_window = Desktop(backend="uia").window(title="Join a meeting -  Webex")
            state = "join_a_meeting"

        elif Desktop(backend="uia").window(title="Sign up -  Webex").exists():
            webex_window = Desktop(backend="uia").window(title="Sign up -  Webex")
            state = "sign_up"

        else:
            webex_window = None  # If no Webex window is found

        return webex_window, state

    def actions(self, webex_window, state):
        end_state = False
        try:
            if state == 'default':
                time.sleep(1)
                                
                sign_in_button = webex_window.child_window(auto_id=self.auto_ids["sign_in_button_auto_id"], control_type="Button")
                # sign_in_button.
                if sign_in_button.wait("exists", timeout=10):
                    sign_in_button.click()
            elif state == 'email_input':
                time.sleep(1)
                try:
                    webex_window = Desktop(backend="uia").window(title="Sign in -  Webex")
                    enter_email = webex_window.child_window(auto_id=self.auto_ids["email_field_auto_id"], control_type="Edit")
                    if enter_email.wait("exists", timeout=5):
                        enter_email.set_text(self.email_webex)
                        webex_window.child_window(title=" Next", auto_id=self.auto_ids["next_button_auto_id"], control_type="Button").click()
                        self.actions(webex_window = Desktop(backend="uia").window(title="Webex"), state='password')
                except Exception as e:
                    enter_email = webex_window.child_window(auto_id=self.auto_ids["email_field_auto_id_2"], control_type="Edit")
                    if enter_email.wait("exists", timeout=5):
                        enter_email.set_text(self.email_webex)
                        webex_window.child_window(title=" Next", auto_id=self.auto_ids["next_button_auto_id"], control_type="Button").click()
                        self.actions(webex_window = Desktop(backend="uia").window(title="Webex"), state='password')
                    
            elif state == 'password':
                time.sleep(1)
                        
                enter_password = webex_window.child_window(auto_id='IDToken2', control_type="Edit")
                if enter_password.wait("exists", timeout=10):
                    enter_password.set_text(self.password_webex) 
                    button_sign_in = webex_window.child_window(auto_id="Button1", control_type="Button").click()

            elif state == "join_a_meeting" or state == "sign_up":
                back_button = webex_window.child_window(auto_id=self.auto_ids["back_button_auto_id"], control_type="Button")
                if back_button:
                    back_button.click()

            elif state == 'logged_in':
                
                meeting_tab = webex_window.child_window(auto_id=self.auto_ids["tab_selector_auto_id"], control_type="Custom")
                if meeting_tab.wait("exists", timeout=10):
                    meeting_tab.click_input()
                    webex_window.wait("ready", timeout=5)

                    try:
                        username = self.email_webex.split('@')[0]
                        print(username)
                        copy_link_button = webex_window.child_window(
                            title=f" Copy my personal room link https://bercacloud.webex.com/meet/{username}",
                            auto_id=self.auto_ids["personal_link_auto_id"],
                            control_type="Button"
                        )

                        # copy_link_button = webex_window.child_window(title=f" Copy my personal room link https://bercacloud.webex.com/meet/rooms.rose", auto_id=self.auto_ids["personal_link_auto_id"], control_type="Button")

                        # Check if the button exists
                        if copy_link_button.wait("exists", timeout=2):
                            end_state = True
                    except:
                        webex_window.wait("ready", timeout=5)
                        profile_button=webex_window.child_window(auto_id=self.auto_ids["profile_auto_id"], control_type="Custom", found_index=0)
                        if not profile_button:
                            profile_button = webex_window.child_window(auto_id=self.auto_ids["profile_auto_id"], control_type="Custom")
                        if profile_button.wait("exists", timeout=10):
                            profile_button.click_input()
                            self.sign_out(webex_window)

        except Exception as e:
            self.open_webex()

        return end_state, state

    def sign_out(self, webex_window):
        try:    
            webex_window = Desktop(backend='uia').window(title_re=".*Profile and preferences.*")
            webex_window.wait("ready", timeout=5)
            sign_out_button = webex_window.child_window(auto_id=self.auto_ids["sign_out_action_auto_id"])
            if sign_out_button.wait("exists", timeout=10):
                sign_out_button.click_input()
                send_keys("{TAB}{ENTER}")
        except Exception as e:
            self.open_webex()

def terminate_other_tv_automation():
    processes = []

    for process in psutil.process_iter(attrs=['pid', 'name', 'create_time']):
        try:
            if process.info['name'].lower() == 'webex-automation.exe':
                processes.append((process.info['pid'], process.info['create_time']))
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    if len(processes) > 2:
        processes.sort(key=lambda x: x[1]) 
        latest_pids = [p[0] for p in processes[-2:]] 

        for pid, _ in processes:
            if pid not in latest_pids: 
                try:
                    print(f"Terminating old instance of 'webex-automation.exe' (PID: {pid})...")
                    psutil.Process(pid).terminate()
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
    else:
        print("No excess instances found. No action needed.")

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

if __name__ == "__main__":
    # show_native_popup()
    terminate_other_tv_automation()
    show_confirmation_dialog("Webex Automation", wait_seconds=5)
    disconnect_if_ssid_exist()

    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    bot = WebexAutomation()
    bot.open_webex()