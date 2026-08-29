# ============================================================
# AUTO-INSTALL DEPENDENCIES
# ============================================================
import subprocess
import sys

def _ensure_dependencies():
    packages = {
        "dotenv": "python-dotenv",
        "pyotp": "pyotp",
        "requests": "requests",
        "selenium": "selenium"
    }
    for mod, pkg in packages.items():
        try:
            __import__(mod)
        except ImportError:
            print(f"[*] Installing missing package: {pkg}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

_ensure_dependencies()
# ============================================================

from urllib.parse import parse_qs, urlparse, quote
import pyotp
import requests
import dotenv
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ============================================================
# LOAD .ENV
# ============================================================
dotenv.load_dotenv()

# ============================================================
# UPSTOX CONFIGURATION
# ============================================================
API_KEY = os.getenv("API_KEY")
SECRET_KEY = os.getenv("SECRET_KEY")
RURL = os.getenv("RURL")
MOBILE_NO = os.getenv("MOBILE_NO")
TOTP_KEY = os.getenv("TOTP_KEY")
PIN = os.getenv("PIN")

# ============================================================
# TOKEN FILE
# ============================================================
TOKEN_FILE = os.getenv("TOKEN_FILE", "access_token.txt")

# ============================================================
# CHECK CONFIGURATION
# ============================================================
def check_config():
    print("\n" + "=" * 70)
    print("CHECKING UPSTOX CONFIGURATION")
    print("=" * 70)
    required = {
        "API_KEY": API_KEY,
        "SECRET_KEY": SECRET_KEY,
        "RURL": RURL,
        "MOBILE_NO": MOBILE_NO,
        "TOTP_KEY": TOTP_KEY,
        "PIN": PIN
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        print("\nMissing .env values:")
        for item in missing:
            print(" -", item)
        raise RuntimeError("Please check your .env file.")
    print("Configuration OK.")

# ============================================================
# CREATE AUTH URL
# ============================================================
rurlEncode = quote(RURL, safe="")
AUTH_URL = (
    f"https://api.upstox.com/v2/login/authorization/dialog"
    f"?response_type=code"
    f"&client_id={API_KEY}"
    f"&redirect_uri={rurlEncode}"
)

# ============================================================
# GET ACCESS TOKEN
# ============================================================
def getAccessToken(code):
    print("\n" + "=" * 70)
    print("EXCHANGING AUTH CODE FOR ACCESS TOKEN")
    print("=" * 70)
    url = "https://api.upstox.com/v2/login/authorization/token"
    headers = {
        "accept": "application/json",
        "Api-Version": "2.0",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    data = {
        "code": code,
        "client_id": API_KEY,
        "client_secret": SECRET_KEY,
        "redirect_uri": RURL,
        "grant_type": "authorization_code"
    }
    print("\nSending authorization code to Upstox...")
    response = requests.post(url, headers=headers, data=data, timeout=30)
    print("HTTP STATUS:", response.status_code)
    
    try:
        json_response = response.json()
    except Exception:
        print("\nRAW RESPONSE:\n", response.text)
        raise RuntimeError("Upstox returned invalid response.")
    
    if response.status_code != 200:
        print("\nTOKEN ERROR:\n", json_response)
        raise RuntimeError("Failed to get access token.")
    
    access_token = json_response.get("access_token")
    if not access_token:
        print("\naccess_token missing:\n", json_response)
        raise RuntimeError("Access token not received.")
    
    print("\nAccess token received successfully.")
    return access_token

# ============================================================
# CHECK TOKEN VALIDITY
# ============================================================
def check_token_validity(access_token):
    print("\n" + "=" * 70)
    print("CHECKING EXISTING ACCESS TOKEN")
    print("=" * 70)
    url = "https://api.upstox.com/v2/user/get-funds-and-margin"
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }
    try:
        response = requests.get(url, headers=headers, timeout=30)
    except Exception as exc:
        print("Token validation request failed:", exc)
        return False
    
    print("HTTP STATUS:", response.status_code)
    if response.status_code != 200:
        print("Token validation failed.\n", response.text)
        return False
    
    try:
        json_response = response.json()
    except Exception:
        print("Invalid JSON response.")
        return False
    
    print("\nToken validation successful.")
    return True

# ============================================================
# LOAD CACHED TOKEN
# ============================================================
def load_cached_token():
    if not os.path.exists(TOKEN_FILE):
        print("\nNo cached token found.")
        return None
    try:
        with open(TOKEN_FILE, "r", encoding="utf-8") as token_file:
            token = token_file.read().strip()
            return token if token else None
    except OSError as exc:
        print("Failed to read cached token:", exc)
        return None

# ============================================================
# SAVE CACHED TOKEN
# ============================================================
def save_cached_token(access_token):
    try:
        with open(TOKEN_FILE, "w", encoding="utf-8") as token_file:
            token_file.write(access_token)
        print("\nAccess token saved:\n", os.path.abspath(TOKEN_FILE))
    except OSError as exc:
        print("Failed to save access token:", exc)

# ============================================================
# FIND PIN FIELD
# ============================================================
def find_pin_field(driver):
    print("\nSearching for PIN field...")
    selectors = [
        (By.XPATH, "//input[@aria-label='Enter 6-digit PIN']"),
        (By.XPATH, "//input[contains(@placeholder,'PIN')]"),
        (By.XPATH, "//input[@type='password']"),
        (By.XPATH, "//input[contains(translate(@name,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'PIN')]"),
        (By.XPATH, "//input[contains(translate(@id,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'PIN')]")
    ]
    for number, selector in enumerate(selectors, start=1):
        try:
            print(f"Trying PIN selector {number}...")
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located(selector))
            print(f"PIN field found using selector {number}.")
            return element
        except Exception:
            print(f"Selector {number} not found.")
    return None

# ============================================================
# FIND CONTINUE BUTTON
# ============================================================
def find_continue_button(driver):
    print("\nSearching for Continue button...")
    selectors = [
        (By.XPATH, "//button[normalize-space()='Continue']"),
        (By.XPATH, "//button[contains(.,'Continue')]"),
        (By.XPATH, "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'continue')]")
    ]
    for number, selector in enumerate(selectors, start=1):
        try:
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(selector))
            print(f"Continue button found using selector {number}.")
            return button
        except Exception:
            print(f"Continue selector {number} not found.")
    return None

# ============================================================
# SELENIUM CHROME LOGIN
# ============================================================
def run():
    print("\n" + "=" * 70)
    print("STARTING UPSTOX SELENIUM CHROME LOGIN")
    print("=" * 70)
    
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 60)
    
    try:
        print("\nOpening Upstox login...")
        driver.get(AUTH_URL)
        time.sleep(5)
        
        print("\nWaiting for mobile number field...")
        mobile = wait.until(EC.visibility_of_element_located((By.ID, "mobileNum")))
        time.sleep(1)
        mobile.clear()
        mobile.send_keys(MOBILE_NO)
        print("Mobile number entered.")
        time.sleep(2)
        
        print("\nClicking Get OTP...")
        get_otp = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Get OTP']")))
        time.sleep(1)
        get_otp.click()
        print("Get OTP clicked.")
        time.sleep(5)
        
        print("\nGenerating TOTP...")
        otp = pyotp.TOTP(TOTP_KEY).now()
        print("TOTP generated.")
        time.sleep(1)
        
        print("\nWaiting for OTP field...")
        otp_field = wait.until(EC.visibility_of_element_located((By.ID, "otpNum")))
        time.sleep(1)
        otp_field.clear()
        otp_field.send_keys(otp)
        print("TOTP entered.")
        time.sleep(2)
        
        print("\nSubmitting OTP...")
        continue_button = find_continue_button(driver)
        if continue_button is None:
            raise RuntimeError("OTP Continue button not found.")
        time.sleep(1)
        continue_button.click()
        print("OTP submitted.")
        
        print("\n" + "=" * 70)
        print("WAITING FOR PIN PAGE")
        print("=" * 70)
        time.sleep(10)
        
        pin_field = find_pin_field(driver)
        if pin_field is None:
            print("\n" + "=" * 70)
            print("PIN FIELD NOT FOUND")
            print("=" * 70)
            print("\nCurrent URL:\n", driver.current_url)
            print("\nPage title:\n", driver.title)
            print("\nChrome will remain open for 30 seconds.")
            time.sleep(30)
            raise RuntimeError("PIN field not found.")
        
        print("\nPIN field detected.")
        time.sleep(2)
        pin_field.click()
        time.sleep(1)
        pin_field.clear()
        pin_field.send_keys(PIN)
        print("PIN entered.")
        time.sleep(2)
        
        print("\nSubmitting PIN...")
        continue_button = find_continue_button(driver)
        if continue_button is None:
            raise RuntimeError("PIN Continue button not found.")
        time.sleep(2)
        continue_button.click()
        print("PIN submitted.")
        
        print("\n" + "=" * 70)
        print("WAITING FOR UPSTOX REDIRECT")
        print("=" * 70)
        time.sleep(5)
        start_time = time.time()
        code = None
        while time.time() - start_time < 90:
            current_url = driver.current_url
            if current_url.startswith(RURL):
                print("\nRedirect detected:\n", current_url)
                parsed = urlparse(current_url)
                query = parse_qs(parsed.query)
                if "code" in query:
                    code = query["code"][0]
                    break
            time.sleep(1)
        
        if not code:
            print("\nFinal browser URL:\n", driver.current_url)
            raise RuntimeError("Authorization code not found.")
        
        print("\n" + "=" * 70)
        print("AUTH CODE RECEIVED")
        print("=" * 70)
        print("Authorization code received successfully.")
        time.sleep(2)
        return code
    finally:
        print("\nClosing Chrome...")
        time.sleep(2)
        driver.quit()

# ============================================================
# GET TOKEN
# ============================================================
def getToken():
    cached_token = load_cached_token()
    if cached_token:
        print("\nFound cached token. Validating token...")
        if check_token_validity(cached_token):
            print("\nCached token is valid.")
            return cached_token
        print("\nCached token is expired. Generating new token...")
    
    code = run()
    token = getAccessToken(code)
    save_cached_token(token)
    return token

# ============================================================
# PUBLIC FUNCTION
# ============================================================
def get_access_token():
    return getToken()

# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    try:
        check_config()
        token = getToken()
        print("\n" + "=" * 70)
        print("UPSTOX LOGIN SUCCESS")
        print("=" * 70)
        print("\nAccess Token received successfully.")
        print("Token saved in:", TOKEN_FILE)
    except KeyboardInterrupt:
        print("\nProgram stopped by user.")
    except Exception as e:
        print("\n" + "=" * 70)
        print("ERROR")
        print("=" * 70)
        print(type(e).__name__, ":", str(e))
