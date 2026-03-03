import logging, gc, psutil, time, subprocess, os, traceback, signal
from difflib import SequenceMatcher
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
# Add after the selenium imports
from pyvirtualdisplay import Display

SPEED_MULTIPLIER = 0.3
fast_sleep = lambda d: time.sleep(d * SPEED_MULTIPLIER)

# Add these lines:
USE_VIRTUAL_DISPLAY = False  # Set to False if you want to see the browser
VIRTUAL_DISPLAY_SIZE = (1920, 1080)

logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(os.path.expanduser("~/Documents/cfms_automation.log")), logging.StreamHandler()]
)

TO_BE_FILLED = os.path.expanduser("~/Documents/to_be_filled.txt")
FILLED_ENTRIES = os.path.expanduser("~/Documents/filled_entries.txt")
ERROR_LOG = os.path.expanduser("~/Documents/cfms_errors.log")
INVALID_CASES = os.path.expanduser("~/Documents/invalidcases.txt")
NO_FINAL_ORDER_CASES = os.path.expanduser("~/Documents/no_final_order_cases.txt")

MAX_RETRIES_PER_CASE = 5
MAX_ELEMENT_RETRIES = 3
ELEMENT_RETRY_DELAY = 2
BROWSER_RESTART_DELAY = 5
WATCHDOG_TIMEOUT = 300

CASE_FIR_NUMBER = "1345"
CASE_FIR_YEAR = "2024"
CASE_POLICE_STATION = "Islam Pura"
CASE_ACT = "PPC"
CASE_OFFENCE = "155C"
DECISION_DATE = "2025-11-21"
DECISION_TYPE = "Acquitted Due to Compromised"

DECISION_MAPPING = {
    "Agreed": ("Cancellation accepted by court", "Yes"),
    "Convicted": ("Conviction", "Other imprisonment"),
    "Fined": ("Conviction", "Fine only"),
    "Acquitted": ("Acquittal", "Poor Investigation"),
    "Acquitted Due to Compromised": ("Acquittal", "Due to Compromise"),
    "u/s 249-A": ("Acquittal", "Poor Investigation"),
    "u/s 345 Cr.P.C": ("Acquittal", "Due to Compromise"),
    "Consign to Record u/s 512": ("Consign to Record", "512 Cr.P.C"),
    "Consign to Record Room": ("Consign to Record", "512 Cr.P.C"),
    "داخل دفتر زیردفعہ512": ("Consign to Record", "512 Cr.P.C"),  # ← ADD THIS
    "داخل دفتر": ("Consign to Record", "512 Cr.P.C"),  # ← ADD THIS
    "منظور شد": ("Cancellation accepted by court", "Yes"),
    "فیصلہ شد": ("Cancellation accepted by court", "Yes")
}

POLICE_STATION_MAPPING = {
    "شفیق‌آباد": "Shafiq Abad", "شفیقآباد": "Shafiq Abad",
    "گلشن‌راوی": "Gulshan-e-Ravi", "گلشنراوی": "Gulshan-e-Ravi",
    "اسلام‌پورہ": "Islam Pura", "اسلامپورہ": "Islam Pura", 
    "مزنگ": "Mozang",
    "موچی گیٹ": "Mochi Gate", "موچیگیٹ": "Mochi Gate",  # ← Mochi Gate added
    "شاہدرہ": "Shahdara", "شاہدره": "Shahdara",
    "Gulshan Ravi": "Gulshan-e-Ravi", "Gulshan ravi": "Gulshan-e-Ravi",
    "gulshan ravi": "Gulshan-e-Ravi", "GULSHAN RAVI": "Gulshan-e-Ravi",
    "GulshanRavi": "Gulshan-e-Ravi", 
    "Shafiq Abad": "Shafiq Abad", "ShafiqAbad": "Shafiq Abad",
    "Islam Pura": "Islam Pura", "IslamPura": "Islam Pura", 
    "Mozang": "Mozang",
    "Mochi Gate": "Mochi Gate", "MochiGate": "Mochi Gate", "mochi gate": "Mochi Gate",
    "Shahdara": "Shahdara", "shahdara": "Shahdara", "SHAHDARA": "Shahdara"
} 

def translate_station(station):
    s = station.strip()
    if s in POLICE_STATION_MAPPING: 
        return POLICE_STATION_MAPPING[s]
    for k, v in POLICE_STATION_MAPPING.items():
        if s.lower() == k.lower(): 
            return v
    norm = s.replace(" ", "").replace("‌", "").replace("-", "").lower()
    for k, v in POLICE_STATION_MAPPING.items():
        if norm == k.replace(" ", "").replace("‌", "").replace("-", "").lower(): 
            return v
    best_m, best_s = None, 0.0
    for k, v in POLICE_STATION_MAPPING.items():
        score = SequenceMatcher(None, norm, k.replace(" ", "").replace("‌", "").replace("-", "").lower()).ratio()
        if score > best_s: 
            best_s, best_m = score, v
    if best_m and best_s >= 0.80:
        logging.info(f"Station matched: '{station}' -> '{best_m}'")
        return best_m
    logging.warning(f"No translation: {station}")
    return station

error_counts = {}
last_error_time = {}

def log_error(error_type, message, exception=None):
    error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {error_type}: {message}"
    if exception: 
        error_msg += f"\n{traceback.format_exc()}"
    logging.error(error_msg)
    try:
        with open(ERROR_LOG, 'a', encoding='utf-8') as f:
            f.write(error_msg + '\n' + '='*80 + '\n')
    except: 
        pass
    error_counts[error_type] = error_counts.get(error_type, 0) + 1
    last_error_time[error_type] = time.time()

def timeout_handler(signum, frame):
    raise TimeoutError("Operation exceeded watchdog timeout")

def sync_to_gdrive():
    try:
        local_file = os.path.expanduser("~/GoogleDrive/STATUS.txt")
        if not os.path.exists(local_file):
            logging.error("STATUS.txt not found")
            return False
        os.sync()
        time.sleep(0.3)
        result = subprocess.run(["rclone", "copy", local_file, "gdrive:GoogleDrive", "-v"],
                              capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            logging.info("STATUS.txt synced")
            return True
        logging.error(f"Sync failed: {result.stderr}")
        return False
    except subprocess.TimeoutExpired:
        logging.error("Sync timeout")
        return False
    except Exception as e:
        logging.error(f"Sync error: {e}")
        return False

class BrowserManager:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.restart_count = 0
        self.display = None  # Add this line
        
    def start_display(self):
        """Start virtual display"""
        if USE_VIRTUAL_DISPLAY:
            try:
                logging.info("Starting virtual display...")
                self.display = Display(visible=False, size=VIRTUAL_DISPLAY_SIZE)
                self.display.start()
                logging.info(f"Virtual display started: {VIRTUAL_DISPLAY_SIZE[0]}x{VIRTUAL_DISPLAY_SIZE[1]}")
                return True
            except Exception as e:
                logging.error(f"Failed to start virtual display: {e}")
                logging.error("Continuing without virtual display...")
                return False
        return True
    
    def stop_display(self):
        """Stop virtual display"""
        if self.display:
            try:
                self.display.stop()
                logging.info("Virtual display stopped")
            except Exception as e:
                logging.error(f"Error stopping display: {e}")
        
   
class BrowserManager:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.restart_count = 0
        self.display = None
        
    def start_display(self):
        """Start virtual display"""
        if USE_VIRTUAL_DISPLAY:
            try:
                logging.info("Starting virtual display...")
                self.display = Display(visible=False, size=VIRTUAL_DISPLAY_SIZE)
                self.display.start()
                logging.info(f"Virtual display started: {VIRTUAL_DISPLAY_SIZE[0]}x{VIRTUAL_DISPLAY_SIZE[1]}")
                return True
            except Exception as e:
                logging.error(f"Failed to start virtual display: {e}")
                logging.error("Continuing without virtual head...")
                return False
        return True
    
    def stop_display(self):
        """Stop virtual display"""
        if self.display:
            try:
                self.display.stop()
                logging.info("Virtual display stopped")
            except Exception as e:
                logging.error(f"Error stopping display: {e}")
    
    def start_browser(self):
        for attempt in range(3):
            try:
                logging.info(f"Starting browser (attempt {attempt + 1}/3)")
                fast_sleep(1)
                
                opts = Options()
                
                # Use a SEPARATE dedicated profile for automation
                # This won't touch your regular Chrome
                opts.add_argument("--user-data-dir=/home/control/.selenium-cfms-profile")
                
                # ALWAYS run in headless mode - completely invisible
                #opts.add_argument("--headless=new")
                #opts.add_argument("--window-size=1920,1080")
                opts.add_argument("--disable-gpu")
                opts.add_argument("--no-sandbox")
                opts.add_argument("--disable-dev-shm-usage")
                opts.add_argument("--disable-software-rasterizer")
                opts.add_argument("--disable-extensions")
                opts.add_argument("--disable-dev-tools")
                opts.add_argument("--remote-debugging-port=0")
                opts.add_experimental_option('excludeSwitches', ['enable-logging'])
                opts.add_experimental_option("detach", False)
                
                logging.info("🚀 Starting HEADLESS browser - completely invisible!")
                
                self.driver = webdriver.Chrome(options=opts)
                self.wait = WebDriverWait(self.driver, 20)
                self.driver.get("https://cfms.prosecution.punjab.gov.pk/#/")
                fast_sleep(3)
                
                logging.info("✅ Browser started successfully in HEADLESS mode")
                self.restart_count += 1
                return True
                
            except Exception as e:
                log_error("BROWSER_START", f"Failed: {e}", e)
                if self.driver:
                    try: 
                        self.driver.quit()
                    except: 
                        pass
                if attempt < 2:
                    fast_sleep(BROWSER_RESTART_DELAY * (attempt + 1))
        return False  
    
    def cleanup(self):
        """Cleanup browser and display"""
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass
        self.stop_display()
    
    def restart_browser(self):
        logging.info("="*60 + "\nRESTARTING BROWSER\n" + "="*60)
        try:
            if self.driver: 
                self.driver.quit()
        except: 
            pass
        fast_sleep(BROWSER_RESTART_DELAY)
        gc.collect()
        return self.start_browser()
    
    def is_browser_alive(self):
        try:
            _ = self.driver.current_url
            return True
        except:
            return False

def safe_find_element(driver, wait, by, value, description="element", timeout=10):
    for attempt in range(MAX_ELEMENT_RETRIES):
        try:
            return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        except TimeoutException:
            if attempt < MAX_ELEMENT_RETRIES - 1:
                logging.warning(f"{description} not found, retry {attempt + 1}")
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                log_error("ELEMENT_NOT_FOUND", f"{description}: {value}")
                return None
        except Exception as e:
            log_error("ELEMENT_FIND_ERROR", f"{description}: {e}", e)
            return None
    return None

def safe_click(driver, wait, element_or_locator, description="element", use_js=False):
    for attempt in range(MAX_ELEMENT_RETRIES):
        try:
            if isinstance(element_or_locator, tuple):
                element = safe_find_element(driver, wait, *element_or_locator, description)
                if not element: 
                    return False
            else:
                element = element_or_locator
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            fast_sleep(0.3)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(element))
            if use_js:
                driver.execute_script("arguments[0].click();", element)
            else:
                element.click()
            logging.info(f"Clicked {description}")
            return True
        except StaleElementReferenceException:
            if attempt < MAX_ELEMENT_RETRIES - 1:
                logging.warning(f"Stale {description}, retrying")
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                log_error("STALE_ELEMENT", description)
                return False
        except Exception as e:
            if attempt < MAX_ELEMENT_RETRIES - 1:
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                log_error("CLICK_ERROR", f"{description}: {e}", e)
                return False
    return False

norm = lambda txt: txt.lower().replace(" ", "").replace("-", "").replace("_", "").replace(".", "") if txt else ""
def select_dropdown_robust(driver, wait, label, val):
    for attempt in range(MAX_ELEMENT_RETRIES):
        try:
            xpath = f"//label[contains(text(),'{label}')]/following-sibling::div//span[contains(@class,'select2-selection')]"
            if not safe_click(driver, wait, (By.XPATH, xpath), f"{label} dropdown"):
                if attempt < MAX_ELEMENT_RETRIES - 1: 
                    continue
                return False
            fast_sleep(0.2)
            search_field = safe_find_element(driver, wait, By.CSS_SELECTOR, ".select2-search__field", "search", 5)
            if not search_field:
                if attempt < MAX_ELEMENT_RETRIES - 1: 
                    continue
                return False
            search_field.clear()
            search_field.send_keys(val)
            fast_sleep(0.3)
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__options")))
            fast_sleep(0.1)
            options = driver.find_elements(By.XPATH, "//li[contains(@class,'select2-results__option')]")
            for opt in options:
                try:
                    opt_text = opt.text.strip()
                    if norm(val) == norm(opt_text):
                        opt.click()
                        logging.info(f"Selected '{opt_text}'")
                        return True
                except: 
                    continue
            best_match, best_score = None, 0.0
            for opt in options:
                try:
                    opt_text = opt.text.strip()
                    score = SequenceMatcher(None, norm(val), norm(opt_text)).ratio()
                    if score > best_score: 
                        best_score, best_match = score, opt
                except: 
                    continue
            if best_match and best_score >= 0.70:
                best_match.click()
                logging.info(f"Selected '{best_match.text}' (fuzzy: {best_score*100:.1f}%)")
                return True
            logging.warning(f"No match for '{val}'")
            if attempt < MAX_ELEMENT_RETRIES - 1:
                fast_sleep(ELEMENT_RETRY_DELAY)
                continue
            return False
        except Exception as e:
            log_error("DROPDOWN_ERROR", f"{label}: {e}", e)
            if attempt < MAX_ELEMENT_RETRIES - 1:
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                return False
    return False

def verify_station_selected(driver, expected_station):
    try:
        selected_span = driver.find_element(By.XPATH, 
            "//label[contains(text(),'Police Station')]/following-sibling::div//span[@class='select2-selection__rendered']")
        selected_text = selected_span.text.strip()
        if selected_text == "Select Police Station" or not selected_text:
            logging.error("Station NOT selected")
            return False
        score = SequenceMatcher(None, norm(expected_station), norm(selected_text)).ratio()
        if score >= 0.70:
            logging.info(f"Station verified: '{selected_text}'")
            return True
        logging.warning(f"Station mismatch: expected '{expected_station}', got '{selected_text}'")
        return False
    except Exception as e:
        logging.error(f"Verify failed: {e}")
        return False

def select_police_station_verified(driver, wait, station_name):
    for attempt in range(3):
        try:
            if attempt > 0: 
                logging.info(f"Retry {attempt+1}/3")
            if not safe_click(driver, wait, 
                (By.XPATH, "//label[contains(text(),'Police Station')]/following::span[contains(@class,'select2-selection')]"), 
                "Police Station"):
                if attempt < 2:
                    fast_sleep(1)
                    continue
                return False
            fast_sleep(0.5)
            search = safe_find_element(driver, wait, By.XPATH, "//input[@class='select2-search__field']", "search", timeout=5)
            if not search:
                if attempt < 2:
                    fast_sleep(1)
                    continue
                return False
            search.clear()
            fast_sleep(0.2)
            search_text = station_name[:5]
            for char in search_text:
                search.send_keys(char)
                fast_sleep(0.05)
            logging.info(f"Searching: '{search_text}'")
            fast_sleep(1.5)
            try:
                WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__options")))
            except TimeoutException:
                logging.error("No results")
                if attempt < 2:
                    try: 
                        driver.find_element(By.TAG_NAME, "body").click()
                        fast_sleep(0.5)
                    except: 
                        pass
                    continue
                return False
            fast_sleep(0.5)
            options = []
            for wait_attempt in range(3):
                options = driver.find_elements(By.XPATH, 
                    "//li[contains(@class,'select2-results__option') and not(contains(@class,'loading'))]")
                if options: 
                    break
                fast_sleep(0.3)
            valid_options = []
            for opt in options:
                try:
                    opt_text = opt.text.strip()
                    if opt_text and opt_text.lower() not in ['loading...', 'searching...', 'no results']:
                        valid_options.append((opt, opt_text))
                except: 
                    continue
            logging.info(f"Found {len(valid_options)} options")
            if not valid_options:
                logging.error("No valid options")
                if attempt < 2:
                    try: 
                        driver.find_element(By.TAG_NAME, "body").click()
                        fast_sleep(0.5)
                    except: 
                        pass
                    continue
                return False
            for opt, opt_text in valid_options:
                if norm(station_name) == norm(opt_text):
                    logging.info(f"Exact match: '{opt_text}'")
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", opt)
                        fast_sleep(0.2)
                        opt.click()
                        fast_sleep(0.5)
                        if verify_station_selected(driver, station_name): 
                            return True
                        if attempt < 2: 
                            continue
                    except Exception as e:
                        logging.error(f"Click error: {e}")
                        if attempt < 2: 
                            continue
            best_match, best_opt, best_score = None, None, 0.0
            for opt, opt_text in valid_options:
                score = SequenceMatcher(None, norm(station_name), norm(opt_text)).ratio()
                if score > best_score: 
                    best_score, best_match, best_opt = score, opt_text, opt
            if best_opt and best_score >= 0.70:
                logging.info(f"Fuzzy match: '{best_match}' ({best_score*100:.1f}%)")
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", best_opt)
                    fast_sleep(0.2)
                    best_opt.click()
                    fast_sleep(0.5)
                    if verify_station_selected(driver, station_name): 
                        return True
                    if attempt < 2: 
                        continue
                except Exception as e:
                    logging.error(f"Click error: {e}")
                    if attempt < 2: 
                        continue
            logging.error(f"No good match")
            if attempt < 2:
                try: 
                    driver.find_element(By.TAG_NAME, "body").click()
                    fast_sleep(0.5)
                except: 
                    pass
                continue
            return False
        except Exception as e:
            log_error("STATION_SELECT", f"Attempt {attempt+1}: {e}", e)
            if attempt < 2:
                try: 
                    driver.find_element(By.TAG_NAME, "body").click()
                    fast_sleep(0.5)
                except: 
                    pass
                continue
            return False
    return False

def fuzzy_match_decision(raw_text):
    if not raw_text: 
        return None, 0.0
    cleaned = raw_text.lower().strip()
    cleaned = ''.join(c for c in cleaned if c.isalnum() or c.isspace() or ord(c) > 127)
    cleaned = ' '.join(cleaned.split())
    best_match, best_score = None, 0.0
    decision_variations = {
        "Agreed": ["agreed", "agree", "منظور", "منظورشد", "منظورہ", "منظورهشد"],
        "Convicted": ["convicted", "convict", "conviction", "سزا", "سزاشد", "سزائے", "بریشد"],
        "Fined": ["fined", "fine", "جرمانہ", "جرمانے"],
        "Acquitted": ["acquitted", "acquit", "acquittal", "بری", "بریشد", "بریٔت"],
        "Acquitted Due to Compromised": ["compromised", "compromise", "سمجھوتہ", "acquitted due to compromised"],
        "u/s 249-A": ["249", "249a", "249-a", "قزیردفعہ249", "us249a", "us249-a"],
        "u/s 345 Cr.P.C": ["345", "345crpc", "us345", "دفعہ345", "us345crpc"],
        "داخل دفتر زیردفعہ512": ["داخلدفترزیردفعہ512", "داخل دفتر زیردفعہ512", "512"],
        "داخل دفتر": ["داخلدفتر", "داخل دفتر", "دفتر", "dakhil"]
    }
    for decision, variations in decision_variations.items():
        for variant in variations:
            if variant in cleaned or cleaned in variant:
                score = SequenceMatcher(None, cleaned, variant).ratio()
                if score > best_score: 
                    best_score, best_match = score, decision
            score = SequenceMatcher(None, cleaned, variant).ratio()
            if score > best_score: 
                best_score, best_match = score, decision
    for key in DECISION_MAPPING.keys():
        key_clean = key.lower().strip()
        score = SequenceMatcher(None, cleaned, key_clean).ratio()
        if score > best_score: 
            best_score, best_match = score, key
    words = cleaned.split()
    if len(words) > 1:
        for decision, variations in decision_variations.items():
            for variant in variations:
                for word in words:
                    if len(word) > 2:
                        score = SequenceMatcher(None, word, variant).ratio()
                        if score > best_score: 
                            best_score, best_match = score, decision
    return best_match, best_score
def get_next_case():
    try:
        if not os.path.exists(TO_BE_FILLED):
            logging.warning(f"File not found: {TO_BE_FILLED}")
            return None
        with open(TO_BE_FILLED, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        if len(lines) <= 1:
            logging.info("No more cases")
            return None
        case_line = lines[1].strip()
        if not case_line:
            logging.warning("Empty line - skipping")
            with open(TO_BE_FILLED, 'w', encoding='utf-8') as f:
                f.write(lines[0])
                if len(lines) > 2: 
                    f.writelines(lines[2:])
            return get_next_case()
        return case_line
    except Exception as e:
        log_error("FILE_READ", f"Error: {e}", e)
        return None

def check_fir_not_found(driver, wait, case_data):
    try:
        popup = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'swal2-popup')]")))
        txt = popup.text.lower()
        if 'fir data not found' in txt or 'record not found' in txt:
            logging.info("FIR NOT FOUND - Logging to invalid cases")
            try:
                with open(INVALID_CASES, 'a', encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\t{case_data}\n")
            except Exception as e:
                log_error("INVALID_CASE_LOG", f"Failed: {e}", e)
            try:
                driver.find_element(By.XPATH, "//button[normalize-space()='Cancel']").click()
                fast_sleep(0.5)
            except: 
                pass
            return True
        return False
    except:
        return False

def mark_as_filled(case_data):
    try:
        with open(FILLED_ENTRIES, 'a', encoding='utf-8') as f:
            f.write(case_data + '\n')
        logging.info("Added to filled")
    except Exception as e:
        log_error("FILE_WRITE", f"Error: {e}", e)
    try:
        with open(TO_BE_FILLED, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        with open(TO_BE_FILLED, 'w', encoding='utf-8') as f:
            f.write(lines[0])
            if len(lines) > 2: 
                f.writelines(lines[2:])
        logging.info("Removed from queue")
    except Exception as e:
        log_error("FILE_UPDATE", f"Error: {e}", e)

def parse_case_data(row_data):
    global CASE_FIR_NUMBER, CASE_FIR_YEAR, CASE_POLICE_STATION, CASE_ACT, CASE_OFFENCE, DECISION_DATE, DECISION_TYPE
    try:
        # Split by tabs and remove ONLY trailing empty fields
        all_fields = row_data.split('\t')
        # Remove trailing empty strings
        while all_fields and not all_fields[-1].strip():
            all_fields.pop()
        
        # Now strip whitespace from remaining fields
        fields = [f.strip() for f in all_fields if f.strip()]
        
        logging.info(f"Parsed {len(fields)} fields from {len(all_fields)} columns")
        
        # Extract the 7 core fields (ignore extra columns like case numbers, duplicates)
        if len(fields) >= 7:
            # Always take first 7 meaningful fields
            name, CASE_FIR_NUMBER, CASE_FIR_YEAR, CASE_OFFENCE, urdu_station, date_str, raw_decision = fields[0:7]
            CASE_OFFENCE = CASE_OFFENCE.replace(" PPC", "").replace("PPC", "").strip()
            logging.info(f"Using 7-field format (ignored {len(fields)-7} extra fields)")
        
        elif len(fields) == 6:
            last_field = fields[-1].strip()
            if ':' in last_field:
                date_decision = last_field.split(':', 1)
                date_str = date_decision[0].strip()
                raw_decision = date_decision[1].strip() if len(date_decision) > 1 else ""
                name, CASE_FIR_NUMBER, CASE_FIR_YEAR, CASE_OFFENCE, urdu_station = fields[:5]
                CASE_OFFENCE = CASE_OFFENCE.replace(" PPC", "").replace("PPC", "").strip()
            else:
                raise ValueError("Expected date:decision format in 6-field data")
        
        else:
            raise ValueError(f"Invalid format - expected at least 6 fields, got {len(fields)}")
        
        # Translate station
        CASE_POLICE_STATION = translate_station(urdu_station)
        if not CASE_FIR_NUMBER or not CASE_FIR_YEAR or not CASE_POLICE_STATION:
            raise ValueError("Empty required field")
        
        # Parse date (handle both - and . separators)
        date_parts = date_str.replace(':', '').replace('.', '-').split('-')
        if len(date_parts) == 3:
            day, month, year = date_parts[0].zfill(2), date_parts[1].zfill(2), date_parts[2]
            if len(year) == 2: 
                year = "20" + year
            DECISION_DATE = f"{year}-{month}-{day}"
        else:
            raise ValueError(f"Invalid date: {date_str}")
        
        # Match decision
        matched_decision, match_score = fuzzy_match_decision(raw_decision)
        if matched_decision and match_score >= 0.70:
            DECISION_TYPE = matched_decision
            logging.info(f"Decision: '{raw_decision}' -> '{matched_decision}'")
        else:
            logging.warning(f"Decision '{raw_decision}' NO MATCH")
            if matched_decision and match_score >= 0.50:
                DECISION_TYPE = matched_decision
                logging.warning(f"WEAK MATCH: '{matched_decision}'")
            else:
                raise ValueError(f"Decision '{raw_decision}' cannot be matched")
        
        CASE_ACT = "PPC"
        logging.info(f"Case: FIR {CASE_FIR_NUMBER}/{CASE_FIR_YEAR} | {CASE_POLICE_STATION} | {DECISION_TYPE}")
        return True
        
    except Exception as e:
        log_error("PARSE_ERROR", f"Failed: {e}", e)
        logging.error(f"Raw: {row_data}")
        return False

def check_framing_required(driver, wait, case_data):
    try:
        popup = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'swal2-popup')]")))
        txt = popup.text.lower()
        if 'framing' in txt and 'charge' in txt:
            logging.info("FRAMING REQUIRED - Logging to framing errors")
            try:
                with open(os.path.expanduser("~/Documents/framing_errors.txt"), 'a', encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\t{case_data}\n")
            except Exception as e:
                log_error("FRAMING_LOG", f"Failed: {e}", e)
            try:
                driver.find_element(By.XPATH, "//button[normalize-space()='OK']").click()
                fast_sleep(0.5)
            except: 
                pass
            return True
        return False
    except:
        return False

def fill_decision_fields(driver, wait, decision, detail, max_retries=3):
    """
    Fills Decision and Detail fields for ALL sections in the modal.
    
    Args:
        decision: First value from DECISION_MAPPING (e.g., "Acquittal", "Conviction")
        detail: Second value from DECISION_MAPPING (e.g., "Due to Compromise", "Fine only")
    """
    def select2_fill(span, value):
        """Helper to fill a select2 dropdown"""
        span.click()
        search = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".select2-search__field")))
        search.clear()
        search.send_keys(value)
        fast_sleep(0.3)
        search.send_keys(Keys.ENTER)
        fast_sleep(0.5)
    
    try:
        # Find all sections with Decision label (one per offense: PPC 420, PPC 468, etc.)
        decision_labels = driver.find_elements(By.XPATH, "//div[@id='IdCrudModal']//label[normalize-space()='Decision']")
        total = len(decision_labels)
        
        logging.info(f"Found {total} Decision section(s) - Filling with Decision='{decision}', Detail='{detail}'")
        
        for index in range(total):
            for attempt in range(1, max_retries + 1):
                try:
                    # Re-find elements to avoid stale references
                    decision_labels = driver.find_elements(By.XPATH, "//div[@id='IdCrudModal']//label[normalize-space()='Decision']")
                    decision_label = decision_labels[index]
                    
                    # Scroll to section
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", decision_label)
                    fast_sleep(0.3)
                    
                    # ===== Fill Decision Field =====
                    decision_span = decision_label.find_element(By.XPATH, 
                        "./following-sibling::div//span[@class='select2-selection select2-selection--single']")
                    
                    logging.info(f"Filling Decision [{index+1}/{total}]: '{decision}'")
                    select2_fill(decision_span, decision)
                    
                    # Verify Decision was set
                    current_decision = decision_span.find_element(By.CLASS_NAME, "select2-selection__rendered").text.strip()
                    if norm(decision) not in norm(current_decision):
                        raise Exception(f"Decision verify failed: expected '{decision}', got '{current_decision}'")
                    
                    logging.info(f"✓ Decision set [{index+1}/{total}]: '{current_decision}'")
                    
                    # ===== Fill Detail Field =====
                    try:
                        # Wait for Detail dropdown to populate (it depends on Decision selection)
                        fast_sleep(0.5)
                        
                        detail_label = decision_label.find_element(By.XPATH,
                            "../following-sibling::div[.//label[contains(text(),'Detail of Decision')]]//label[contains(text(),'Detail of Decision')]")
                        
                        detail_span = detail_label.find_element(By.XPATH,
                            "./following-sibling::div//span[@class='select2-selection select2-selection--single']")
                        
                        # Wait for Detail options to load
                        detail_select = detail_label.find_element(By.XPATH, "./following-sibling::div//select")
                        WebDriverWait(driver, 10).until(lambda d: len(detail_select.find_elements(By.TAG_NAME, "option")) > 1)
                        
                        logging.info(f"Filling Detail [{index+1}/{total}]: '{detail}'")
                        select2_fill(detail_span, detail)
                        
                        # Verify Detail was set
                        current_detail = detail_span.find_element(By.CLASS_NAME, "select2-selection__rendered").text.strip()
                        if norm(detail) not in norm(current_detail):
                            raise Exception(f"Detail verify failed: expected '{detail}', got '{current_detail}'")
                        
                        logging.info(f"✓ Detail set [{index+1}/{total}]: '{current_detail}'")
                        
                    except NoSuchElementException:
                        logging.warning(f"No Detail field found [{index+1}/{total}]")
                    
                    break  # Success, move to next section
                    
                except Exception as e:
                    logging.warning(f"Retry {attempt}/{max_retries} failed for section [{index+1}]: {e}")
                    if attempt == max_retries:
                        log_error("DECISION_SECTION_FAILED", f"Section {index+1} failed after {max_retries} attempts", e)
        
        logging.info(f"✓ Completed filling {total} Decision section(s)")
        return True
        
    except Exception as e:
        log_error("FILL_DECISION_FATAL", f"Fatal error: {e}", e)
        return False

def automate_final_order(driver, wait, date_val, dec_type, case_data):
    logging.info(f"Final Order - Type: {dec_type}")
    if dec_type not in DECISION_MAPPING:
        log_error("INVALID_DECISION", f"Type '{dec_type}' not in mapping")
        return "ERROR"
    decision, detail = DECISION_MAPPING[dec_type]
    failed_attempts = {}
    MAX_ATTEMPTS_PER_INDEX = 3
    is_cancellation = (decision == "Cancellation accepted by court" and detail == "Yes")
    try:
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr")))
            fast_sleep(0.3)
        except TimeoutException:
            logging.warning("No table")
        btns = driver.find_elements(By.XPATH, "//tbody/tr/td[last()]//a[@id='roleActions']")
        total_entries = len(btns)
        if total_entries == 0:
            logging.info("NO FINAL ORDER ENTRIES - Logging to no_final_order file")
            try:
                with open(NO_FINAL_ORDER_CASES, 'a', encoding='utf-8') as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\t{case_data}\n")
            except Exception as e:
                log_error("NO_FINAL_ORDER_LOG", f"Failed: {e}", e)
            return "SKIP"
        logging.info(f"Found {total_entries} entries")
        if is_cancellation:
            logging.info("CANCELLATION - Only first entry")
        processed_count, i = 0, 0
        while i < total_entries:
            if is_cancellation and i > 0:
                logging.info(f"Skip entry {i+1}")
                i += 1
                continue
            logging.info(f"\n{'='*60}\nEntry {i+1}/{total_entries}\n{'='*60}")
            try:
                btns = driver.find_elements(By.XPATH, "//tbody/tr/td[last()]//a[@id='roleActions']")
                if i >= len(btns):
                    logging.warning(f"Entry {i+1} not found")
                    break
                if failed_attempts.get(i, 0) >= MAX_ATTEMPTS_PER_INDEX:
                    logging.error(f"Skip entry {i+1} - max attempts")
                    i += 1
                    continue
                if not safe_click(driver, wait, btns[i], f"three-dot {i+1}"):
                    failed_attempts[i] = failed_attempts.get(i, 0) + 1
                    continue
                fast_sleep(0.3)
            except Exception as e:
                log_error("THREE_DOT_CLICK", f"Entry {i+1}: {e}", e)
                failed_attempts[i] = failed_attempts.get(i, 0) + 1
                i += 1
                continue
            edit_clicked = False
            try:
                edit_btn = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'dropdown-menu') and contains(@class,'show')]//a[contains(text(),'Edit')]"))
                )
                edit_btn.click()
                edit_clicked = True
                logging.info("Clicked Edit dropdown")
            except TimeoutException:
                try:
                    edit_btn = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Edit' and contains(@class,'btn')]"))
                    )
                    edit_btn.click()
                    edit_clicked = True
                    logging.info("Clicked Edit floating")
                except TimeoutException:
                    logging.warning(f"No Edit [{i+1}]")
                    try: 
                        driver.find_element(By.TAG_NAME, "body").click()
                    except: 
                        pass
                    i += 1
                    continue
            fast_sleep(0.3)
            if check_framing_required(driver, wait, case_data): 
                return "SKIP"
            try:
                wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class,'modal') and contains(@class,'show')]")))
            except TimeoutException:
                log_error("MODAL_TIMEOUT", f"Modal not appeared [{i+1}]")
                failed_attempts[i] = failed_attempts.get(i, 0) + 1
                continue
            fast_sleep(0.3)
            try:
                date_inp = driver.find_element(By.XPATH, "//div[contains(@class,'modal')]//input[@type='date']")
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].blur();
                """, date_inp, date_val)
                fast_sleep(0.3)
                logging.info(f"Date: {date_val}")
            except Exception as e:
                log_error("DATE_FILL", f"Error: {e}", e)
            fill_decision_fields(driver, wait, decision, detail)
            try:
                upd_btn = driver.find_element(By.XPATH, "//div[contains(@class,'modal')]//button[normalize-space()='Update']")
                driver.execute_script("arguments[0].scrollIntoView(true);", upd_btn)
                fast_sleep(0.2)
                upd_btn.click()
                logging.info("Clicked Update")
            except Exception as e:
                log_error("UPDATE_CLICK", f"Error: {e}", e)
                failed_attempts[i] = failed_attempts.get(i, 0) + 1
                continue
            try:
                wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class,'modal') and contains(@class,'show')]")))
                fast_sleep(0.3)
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "modal-backdrop")))
                logging.info("Modal closed")
            except: 
                pass
            fast_sleep(0.3)
            try:
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='OK']"))).click()
                logging.info("Popup handled")
            except: 
                pass
            try:
                wait.until(EC.invisibility_of_element_located((By.ID, "IdCrudModal")))
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "modal-backdrop")))
                fast_sleep(1)
            except:
                fast_sleep(1)
            processed_count += 1
            failed_attempts[i] = 0
            logging.info(f"Entry {i+1} completed")
            if is_cancellation:
                logging.info("Cancellation done")
                break
            i += 1
            fast_sleep(0.3)
        logging.info(f"\nProcessed {processed_count} entries")
        if is_cancellation:
            logging.info("CANCELLATION COMPLETE")
            return "COMPLETE"
        try:
            driver.refresh()
            fast_sleep(0.3)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Final Order']"))).click()
            fast_sleep(0.5)
        except Exception as e:
            log_error("REFRESH_ERROR", f"Error: {e}", e)
        try:
            remaining_btns = driver.find_elements(By.XPATH, "//tbody/tr/td[last()]//a[@id='roleActions']")
            remaining_count = len(remaining_btns)
        except:
            remaining_count = 0
        logging.info(f"After refresh: {remaining_count} entries")
        if remaining_count == 0:
            logging.info("NO ENTRIES - Complete")
            return "COMPLETE"
        elif remaining_count < total_entries:
            logging.info("Entries reduced")
            return "RESTART"
        else:
            logging.warning("Count unchanged - force complete")
            return "COMPLETE"
    except Exception as e:
        log_error("FINAL_ORDER_ERROR", f"Failed: {e}", e)
        return "ERROR"

def handle_judicial_proceedings(driver, wait):
    logging.info("\n" + "="*60 + "\nJUDICIAL PROCEEDINGS\n" + "="*60)
    try:
        if not safe_click(driver, wait, (By.XPATH, "//a[normalize-space()='Judicial Proceedings']"), "Judicial tab"):
            return False
        fast_sleep(1)
        rows = driver.find_elements(By.CSS_SELECTOR, ".laravel-vue-datatable-tbody tr")
        if rows and len(rows) > 0:
            logging.info("Existing record")
            try:
                three_dots = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "td:last-child .feather-more-horizontal"))
                )
                if three_dots.is_displayed() and three_dots.is_enabled():
                    three_dots.click()
                    fast_sleep(0.3)
                    edit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'dropdown-menu')]//a[contains(text(),'Edit')]")))
                    edit_btn.click()
                    fast_sleep(0.5)
                    return True
                return False
            except Exception as e:
                log_error("JUDICIAL_EDIT", f"Error: {e}", e)
                return False
        else:
            logging.info("Adding new")
            try:
                add_btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Add New')]")))
                add_btn.click()
                fast_sleep(0.5)
                return True
            except Exception as e:
                log_error("JUDICIAL_ADD", f"No button: {e}", e)
                return False
    except Exception as e:
        log_error("JUDICIAL_PROCEEDINGS", f"Error: {e}", e)
        return False

def fill_court_modal(driver, wait):
    logging.info("Filling Court modal")
    try:
        for attempt in range(3):
            if select_dropdown_robust(driver, wait, "Category", "Magistrate"):
                fast_sleep(0.3)
                try:
                    cat_span = driver.find_element(By.XPATH, "//label[contains(text(),'Category')]/following-sibling::div//span[contains(@class,'select2-selection')]")
                    if "magistrate" in cat_span.text.lower():
                        logging.info("Category verified")
                        break
                except: 
                    pass
            if attempt == 2:
                log_error("CATEGORY_FILL", "Failed after 3 attempts")
                return False
            fast_sleep(0.5)
        fast_sleep(0.3)
        for attempt in range(3):
            if select_dropdown_robust(driver, wait, "Court Type", "Section 30 Magistrate"):
                fast_sleep(0.3)
                try:
                    ct_span = driver.find_element(By.XPATH, "//label[contains(text(),'Court Type')]/following-sibling::div//span[contains(@class,'select2-selection')]")
                    if "section 30" in ct_span.text.lower():
                        logging.info("Court Type verified")
                        break
                except: 
                    pass
            if attempt == 2:
                log_error("COURT_TYPE_FILL", "Failed after 3 attempts")
                return False
            fast_sleep(0.5)
        fast_sleep(0.3)
        button_clicked = False
        for btn_text in ["Save", "Create"]:
            if button_clicked: 
                break
            try:
                btn = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, f"//button[@type='submit' and normalize-space()='{btn_text}']")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                fast_sleep(0.1)
                btn.click()
                logging.info(f"Clicked {btn_text}")
                button_clicked = True
                break
            except: 
                pass
        if not button_clicked:
            log_error("SAVE_BUTTON", "Failed to click")
            return False
        fast_sleep(0.5)
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='OK']"))).click()
        except: 
            pass
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Close']"))).click()
        except: 
            pass
        fast_sleep(0.5)
        return True
    except Exception as e:
        log_error("COURT_MODAL", f"Error: {e}", e)
        return False

def process_single_case(browser_mgr, case_data):
    driver, wait = browser_mgr.driver, browser_mgr.wait
    try:
        if not browser_mgr.is_browser_alive():
            log_error("BROWSER_DEAD", "Unresponsive")
            return "ERROR"
        logging.info("Navigating to Cases")
        try:
            cases_menu = wait.until(EC.presence_of_element_located((By.XPATH, "//span[normalize-space()='Cases']")))
            driver.execute_script("arguments[0].click();", cases_menu)
            fast_sleep(1)
        except Exception as e:
            log_error("CASES_NAV", f"Error: {e}", e)
            return "ERROR"
        if not safe_click(driver, wait, (By.XPATH, "//button[normalize-space()='OK']"), "OK"): 
            return "ERROR"
        if not safe_click(driver, wait, (By.XPATH, "//button[contains(@class,'btn-success')]"), "green"): 
            return "ERROR"
        fir = safe_find_element(driver, wait, By.XPATH, "//input[@placeholder='FIR Number']", "FIR")
        if not fir: 
            return "ERROR"
        fir.clear()
        fir.send_keys(CASE_FIR_NUMBER)
        if not safe_click(driver, wait, (By.XPATH, "//label[contains(text(),'FIR Year')]/following::span[contains(@class,'select2-selection')]"), "FIR Year"): 
            return "ERROR"
        if not safe_click(driver, wait, (By.XPATH, f"//li[normalize-space()='{CASE_FIR_YEAR}']"), "Year"): 
            return "ERROR"
        if not select_police_station_verified(driver, wait, CASE_POLICE_STATION):
            logging.error("Station selection failed")
            return "ERROR"
        if not safe_click(driver, wait, (By.XPATH, "//button[contains(text(),'Fetch FIR Data')]"), "Fetch"): 
            return "ERROR"
        fast_sleep(1)
        if check_fir_not_found(driver, wait, case_data): 
            return "SKIP"
        if not safe_click(driver, wait, (By.XPATH, "//button[normalize-space()='Edit Case']"), "Edit Case"): 
            return "ERROR"
        fast_sleep(2)
        driver.refresh()
        fast_sleep(3)
        logging.info("\n" + "="*60 + "\nPROSECUTION\n" + "="*60)
        if not safe_click(driver, wait, (By.XPATH, "//a[normalize-space()='Prosecution']"), "Prosecution"): 
            return "ERROR"
        fast_sleep(0.5)
        if not safe_click(driver, wait, (By.XPATH, "//button[contains(@class,'btn-success')]"), "green"): 
            return "ERROR"
        fast_sleep(0.5)
        role = safe_find_element(driver, wait, By.XPATH, "//select[.//option[contains(text(),'Conduct Trial')]]", "role")
        if not role: 
            return "ERROR"
        Select(role).select_by_visible_text("Conduct Trial")
        fast_sleep(2.5)
        dates = [d for d in driver.find_elements(By.XPATH, "//input[@type='date']") if d.is_displayed()]
        if len(dates) >= 2:
            for idx, d in enumerate(dates[:2]):
                d.click()
                fast_sleep(0.2)
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].blur();
                """, d, DECISION_DATE)
                fast_sleep(0.5)
                logging.info(f"Date {idx+1} filled")
        fast_sleep(0.2)
        if not safe_click(driver, wait, (By.XPATH, "//button[normalize-space()='Create']"), "Create"): 
            return "ERROR"
        fast_sleep(1)
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='OK']"))).click()
        except: 
            pass
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Close']"))).click()
        except: 
            pass
        fast_sleep(1)
        should_fill = handle_judicial_proceedings(driver, wait)
        if should_fill:
            if not fill_court_modal(driver, wait):
                log_error("COURT_MODAL_FAIL", "Failed")
        logging.info("\n" + "="*60 + "\nFINAL ORDER\n" + "="*60)
        if not safe_click(driver, wait, (By.XPATH, "//a[normalize-space()='Final Order']"), "Final Order"): 
            return "ERROR"
        fast_sleep(0.3)
        status = automate_final_order(driver, wait, DECISION_DATE, DECISION_TYPE, case_data)
        return status
    except Exception as e:
        log_error("CASE_PROCESSING", f"Error: {e}", e)
        return "ERROR"

def write_status_file(session_start_time, session_cases):
    sf = os.path.expanduser("~/GoogleDrive/STATUS.txt")
    try:
        with open(TO_BE_FILLED, 'r', encoding='utf-8') as f:
            remaining = len([l for l in f.readlines()[1:] if l.strip()])
    except:
        remaining = 0
    try:
        with open(FILLED_ENTRIES, 'r', encoding='utf-8') as f:
            filled = len([l for l in f.readlines() if l.strip()])
    except:
        filled = 0
    total = filled + remaining
    progress_pct = (filled / total * 100) if total > 0 else 0
    session_elapsed = time.time() - session_start_time
    cases_this_session = len(session_cases)
    if cases_this_session > 0:
        avg_case_time = sum(c['duration'] for c in session_cases) / cases_this_session
        cases_per_hour = (cases_this_session / session_elapsed * 3600) if session_elapsed > 0 else 0
    else:
        avg_case_time = 0
        cases_per_hour = 0
    if cases_this_session >= 2:
        eta_seconds = int(avg_case_time * remaining)
        eta_hours = eta_seconds // 3600
        eta_mins = (eta_seconds % 3600) // 60
        completion_time = datetime.now() + timedelta(seconds=eta_seconds)
        eta_str = f"{eta_hours}h {eta_mins}m"
        completion_str = completion_time.strftime('%I:%M %p on %b %d')
    else:
        eta_str = "Calculating..."
        completion_str = "Calculating..."
    bar_length = 40
    filled_length = int(bar_length * progress_pct / 100)
    progress_bar = '█' * filled_length + '░' * (bar_length - filled_length)
    try:
        with open(sf, 'w', encoding='utf-8') as f:
            f.write("=" * 70 + "\n")
            f.write("CFMS AUTOMATION - STATUS REPORT\n")
            f.write("=" * 70 + "\n\n")
            f.write(f"Updated: {datetime.now().strftime('%I:%M:%S %p - %b %d, %Y')}\n")
            f.write(f"Status: 🟢 ACTIVE\n\n")
            f.write("-" * 70 + "\n")
            f.write("PROGRESS\n")
            f.write("-" * 70 + "\n")
            f.write(f"Total:      {total:>6} cases\n")
            f.write(f"Completed:  {filled:>6} cases ({progress_pct:>5.1f}%)\n")
            f.write(f"Remaining:  {remaining:>6} cases\n\n")
            f.write(f"[{progress_bar}] {progress_pct:>5.1f}%\n\n")
            f.write("-" * 70 + "\n")
            f.write("PERFORMANCE\n")
            f.write("-" * 70 + "\n")
            f.write(f"This Session:       {cases_this_session:>4} cases\n")
            f.write(f"Processing Speed:   {cases_per_hour:>5.1f} cases/hour\n")
            f.write(f"Avg Time/Case:      {avg_case_time:>5.1f} seconds\n\n")
            f.write("-" * 70 + "\n")
            f.write("FORECAST\n")
            f.write("-" * 70 + "\n")
            f.write(f"Time Remaining:  {eta_str}\n")
            f.write(f"Completion ETA:  {completion_str}\n\n")
            f.write("-" * 70 + "\n")
            f.write("LAST CASE\n")
            f.write("-" * 70 + "\n")
            f.write(f"FIR:      {CASE_FIR_NUMBER}/{CASE_FIR_YEAR}\n")
            f.write(f"Station:  {CASE_POLICE_STATION}\n")
            f.write(f"Decision: {DECISION_TYPE}\n\n")
            f.write("=" * 70 + "\n")
            f.write(f"Note: '{filled}' entries = Successfully edited cases\n")
            f.write("=" * 70 + "\n")
        logging.info(f"STATUS.txt updated - Progress: {progress_pct:.1f}%")
    except Exception as e:
        log_error("STATUS_UPDATE", f"Failed: {e}", e)

def main():
    browser_mgr = BrowserManager()
    if not browser_mgr.start_browser():
        logging.error("Browser start failed")
        return
    loop_count, max_loops, consecutive_errors, MAX_CONSECUTIVE_ERRORS = 0, 1000, 0, 3
    session_start_time = time.time()
    session_cases = []
    try:
        while loop_count < max_loops:
            loop_count += 1
            logging.info("\n" + "#"*60 + f"\n### LOOP {loop_count} ###\n" + "#"*60)
            case_line = get_next_case()
            if not case_line:
                logging.info("NO MORE CASES")
                break
            logging.info(f"Processing: {case_line}")
            if not parse_case_data(case_line):
                logging.error("Parse failed - skipping")
                mark_as_filled(case_line)
                consecutive_errors = 0
                continue
            case_retry_count, case_success = 0, False
            case_start_time = time.time()
            while case_retry_count < MAX_RETRIES_PER_CASE and not case_success:
                case_retry_count += 1
                if case_retry_count > 1: 
                    logging.info(f"Retry {case_retry_count}/{MAX_RETRIES_PER_CASE}")
                if not browser_mgr.is_browser_alive():
                    logging.warning("Browser died - restarting")
                    if not browser_mgr.restart_browser():
                        log_error("BROWSER_RESTART_FAIL", "Failed")
                        break
                try:
                    signal.signal(signal.SIGALRM, timeout_handler)
                    signal.alarm(WATCHDOG_TIMEOUT)
                    status = process_single_case(browser_mgr, case_line)
                    signal.alarm(0)
                    if status == "COMPLETE":
                        mark_as_filled(case_line)
                        case_duration = time.time() - case_start_time
                        session_cases.append({
                            'fir': f"{CASE_FIR_NUMBER}/{CASE_FIR_YEAR}",
                            'station': CASE_POLICE_STATION,
                            'decision': DECISION_TYPE,
                            'duration': case_duration,
                            'timestamp': datetime.now()
                        })
                        logging.info("CASE SUCCESSFULLY COMPLETED")
                        write_status_file(session_start_time, session_cases)
                        sync_to_gdrive()
                        case_success, consecutive_errors = True, 0
                    elif status == "SKIP":
                        logging.info("Case skipped - logged to error file")
                        mark_as_filled(case_line)
                        case_success, consecutive_errors = True, 0
                    elif status == "RESTART":
                        logging.info("Restarting")
                        continue
                    elif status == "ERROR":
                        consecutive_errors += 1
                        logging.error(f"Error (consecutive: {consecutive_errors})")
                        if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                            logging.error("Too many errors - restarting browser")
                            if not browser_mgr.restart_browser():
                                log_error("CRITICAL", "Cannot restart")
                                return
                            consecutive_errors = 0
                        if case_retry_count >= MAX_RETRIES_PER_CASE:
                            logging.error("Max retries - skipping")
                            mark_as_filled(case_line)
                            case_success = True
                except TimeoutError:
                    log_error("WATCHDOG_TIMEOUT", f"Timeout {WATCHDOG_TIMEOUT}s")
                    consecutive_errors += 1
                    logging.info("Timeout - restarting")
                    if not browser_mgr.restart_browser():
                        log_error("CRITICAL", "Cannot restart after timeout")
                        return
                except Exception as e:
                    log_error("PROCESSING_ERROR", f"Unexpected: {e}", e)
                    consecutive_errors += 1
                    if case_retry_count >= MAX_RETRIES_PER_CASE:
                        mark_as_filled(case_line)
                        case_success = True
            logging.info("="*60 + "\nRESTARTING BROWSER\n" + "="*60)
            if not browser_mgr.restart_browser():
                log_error("CRITICAL", "Restart failed")
                break
            logging.info(f"RAM: {psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024:.1f} MB")
        if loop_count >= max_loops: 
            logging.warning(f"Max loops ({max_loops})")
        print("\n" + "="*60 + "\nAUTOMATION COMPLETE\n" + "="*60)
        print(f"\nTotal errors: {sum(error_counts.values())}")
        print(f"Browser restarts: {browser_mgr.restart_count}")
    except KeyboardInterrupt:
        logging.info("\nInterrupted")
    except Exception as e:
        log_error("CRITICAL", f"Fatal: {e}", e)
    finally:
        browser_mgr.cleanup()  # Changed from just browser_mgr.driver.quit()
        logging.info("\nCleanup complete")

if __name__ == "__main__":
    main()