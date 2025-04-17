# 💻 Meeting Room Automation (GMeet, Teams, Webex, Zoom)

This project automates login and session verification for various video conferencing platforms (**Google Meet**, **Microsoft Teams**, **Webex**, and **Zoom**) tailored for specific meeting room accounts.

---

## 🌐 Supported Platforms

- Google Meet  
- Microsoft Teams  
- Webex  
- Zoom  

> Each platform has its own executable (`.exe`) for automation.

---

## 🔧 Features

- Automated login to meeting accounts  
- Profile/account verification and switching  
- Session cleanup: terminates old or duplicate automation processes  
- Detects hotspot SSID and disconnects  
- Encrypted credential storage using `cryptography.fernet`  
- Uses default Chrome browser profile for seamless interaction  

---

## 📁 Folder Structure

automation-folder/ 
├── gmeet-automation.exe 
├── teams-automation.exe 
├── webex-automation.exe 
├── zoom-automation.exe 
├── encrypted_credentials.txt 🔒 Encrypted email & password 
├── encryption_key.key 🔑 Key for decryption


---

## 🔄 First-Time Setup

1. Download or place the `.exe` files in a folder.  
2. Ensure `encryption_key.key` and `encrypted_credentials.txt` are in the same directory.  
3. Run any executable (e.g., `gmeet-automation.exe`).  

> If `encrypted_credentials.txt` or `encryption_key.key` is missing, run `webex-automation.exe` and the script will:
- Automatically create `encrypted_credentials.txt` and `encryption_key.key`  
- On the next run, it will use the saved credentials  

---

## 🧪 Testing & Logging

Before running the automation, ensure:

- ✅ You have **logged in at least once manually** in the browser so Chrome saves the session.  
- ✅ You have **initialized Webex manually** to complete setup and store required data.  
- 🚫 **2FA or CAPTCHA is disabled**, or it may block automation.

---

## 📆 Google Meet Automation

- Opens Google Meet landing page  
- Detects if already signed in with the correct account  
- Logs out and re-authenticates if necessary  
- Supports switching between accounts  

### Checks:
- Not logged in (default state)  
- Needs email/password input  
- Logged in with the wrong account  

### Login Flow:
1. Open Google Meet  
2. Click **Sign in**  
3. Select saved account or re-enter credentials  
4. Complete password & login  

---

## 🚪 Microsoft Teams Automation

- Opens [teams.live.com](https://teams.live.com)  
- Enters credentials and manages session  
- Verifies logged-in account  
- Signs out and retries if wrong account is detected  

### Flow:
1. Open Teams Web  
2. Enter email & password  
3. Confirm "Stay signed in"  
4. Check profile match  
5. If mismatched → Sign out → Retry login  

---

## 📅 Webex Automation

- Uses `pywinauto` to control **Windows UI**  
- Launches Webex application from Start Menu  
- Detects login state  
- Signs in and navigates to copy personal room link  
- Logs out if the account is incorrect  

### Unique Flow:
- Operates at **Windows GUI level** (not browser-based)  
- Maximizes Webex window if not already maximized  
- Confirms account from profile section  

### Handles:
- Profile `auto_id` navigation  
- Email input via specific field `auto_id`  
- Personal link match (e.g., `https://bercacloud.webex.com/meet/rooms.sakura`)  

---

## 🌎 Zoom Automation

- Opens **Zoom Web Client**  
- Handles login and session validation  
- Navigates to profile to verify email  

### Actions:
- Logs in with email/password  
- Verifies correct account under avatar  
- Signs out and restarts flow if mismatched  

### Checks:
- "Join Meeting" button presence  
- Header presence for logged-in status  
- Email match inside `div.info-panel-detail`  

---

## 🔓 Security & Encryption

- Credentials are encrypted and stored **locally**.

### Files Used:
- `encryption_key.key` → symmetric key for Fernet encryption  
- `encrypted_credentials.txt` → stores encrypted email and password  

> ⚠️ If decryption fails, **delete both files** to reset the setup.

---

## 🚀 Automation Control Functions

Each script includes:

- `terminate_other_tv_automation()`  
  → Ensures only 1–2 instances of the same `.exe` are running  
  → Terminates older processes  

- `terminate_if_running(process_name)`  
  → Kills conflicting processes like `chrome.exe`, `chromedriver.exe`  

- `disconnect_if_ssid_exist()`  
  → Disconnects Wi-Fi if connected to hotspot or unnecessary SSID  
  → Pings Google to ensure internet connectivity  

---

## 🚧 Requirements

- 🪟 **Windows OS** (required for Webex automation)  
- 🌐 **Google Chrome** (installed and configured)

---

## 🎓 For Developers

To convert Python scripts to executable files:

```bash
pyinstaller --onefile --noconsole --clean --name=my_script my_script.py

---

## 🔄 First-Time Setup

1. Download or place the `.exe` files in a folder.  
2. Ensure `encryption_key.key` and `encrypted_credentials.txt` are in the same directory.  
3. Run any executable (e.g., `gmeet-automation.exe`).  

> If `encrypted_credentials.txt` or `encryption_key.key` is missing, run `webex-automation.exe` and the script will:
- Automatically create `encrypted_credentials.txt` and `encryption_key.key`  
- On the next run, it will use the saved credentials  

---

## 🧪 Testing & Logging

Before running the automation, ensure:

- ✅ You have **logged in at least once manually** in the browser so Chrome saves the session.  
- ✅ You have **initialized Webex manually** to complete setup and store required data.  
- 🚫 **2FA or CAPTCHA is disabled**, or it may block automation.

---

## 📆 Google Meet Automation

- Opens Google Meet landing page  
- Detects if already signed in with the correct account  
- Logs out and re-authenticates if necessary  
- Supports switching between accounts  

### Checks:
- Not logged in (default state)  
- Needs email/password input  
- Logged in with the wrong account  

### Login Flow:
1. Open Google Meet  
2. Click **Sign in**  
3. Select saved account or re-enter credentials  
4. Complete password & login  

---

## 🚪 Microsoft Teams Automation

- Opens [teams.live.com](https://teams.live.com)  
- Enters credentials and manages session  
- Verifies logged-in account  
- Signs out and retries if wrong account is detected  

### Flow:
1. Open Teams Web  
2. Enter email & password  
3. Confirm "Stay signed in"  
4. Check profile match  
5. If mismatched → Sign out → Retry login  

---

## 📅 Webex Automation

- Uses `pywinauto` to control **Windows UI**  
- Launches Webex application from Start Menu  
- Detects login state  
- Signs in and navigates to copy personal room link  
- Logs out if the account is incorrect  

### Unique Flow:
- Operates at **Windows GUI level** (not browser-based)  
- Maximizes Webex window if not already maximized  
- Confirms account from profile section  

### Handles:
- Profile `auto_id` navigation  
- Email input via specific field `auto_id`  
- Personal link match (e.g., `https://bercacloud.webex.com/meet/rooms.sakura`)  

---

## 🌎 Zoom Automation

- Opens **Zoom Web Client**  
- Handles login and session validation  
- Navigates to profile to verify email  

### Actions:
- Logs in with email/password  
- Verifies correct account under avatar  
- Signs out and restarts flow if mismatched  

### Checks:
- "Join Meeting" button presence  
- Header presence for logged-in status  
- Email match inside `div.info-panel-detail`  

---

## 🔓 Security & Encryption

- Credentials are encrypted and stored **locally**.

### Files Used:
- `encryption_key.key` → symmetric key for Fernet encryption  
- `encrypted_credentials.txt` → stores encrypted email and password  

> ⚠️ If decryption fails, **delete both files** to reset the setup.

---

## 🚀 Automation Control Functions

Each script includes:

- `terminate_other_tv_automation()`  
  → Ensures only 1–2 instances of the same `.exe` are running  
  → Terminates older processes  

- `terminate_if_running(process_name)`  
  → Kills conflicting processes like `chrome.exe`, `chromedriver.exe`  

- `disconnect_if_ssid_exist()`  
  → Disconnects Wi-Fi if connected to hotspot or unnecessary SSID  
  → Pings Google to ensure internet connectivity  

---

## 🚧 Requirements

- 🪟 **Windows OS** (required for Webex automation)  
- 🌐 **Google Chrome** (installed and configured)

---

## 🎓 For Developers

To convert Python scripts to executable files:

```bash
pyinstaller --onefile --noconsole --clean --name=my_script my_script.py
Optional: Add --upx-dir=C:\path\to\upx if you have UPX installed. (pyinstaller --onefile --noconsole --upx-dir=C:\path\to\upx --clean --name=my_script my_script.py)

## 🔧 Troubleshooting

- Wrong credentials? Delete both encryption_key.key and encrypted_credentials.txt.
- Stuck session? Force logout or clear Chrome profile session.
- Login blocked by security? Ensure CAPTCHA or MFA is disabled.
