import logging
import gc
import psutil
import time
import subprocess
from difflib import SequenceMatcher
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
import os
from datetime import datetime
import traceback
import signal
import sys

# SPEED OPTIMIZATION
SPEED_MULTIPLIER = 0.3

def fast_sleep(duration):
    """Optimized sleep function for faster navigation"""
    time.sleep(duration * SPEED_MULTIPLIER)

# Enhanced logging with error tracking
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.expanduser("~/Documents/cfms_automation.log")),
        logging.StreamHandler()
    ]
)

# File paths
TO_BE_FILLED = os.path.expanduser("~/Documents/to_be_filled.txt")
FILLED_ENTRIES = os.path.expanduser("~/Documents/filled_entries.txt")
ERROR_LOG = os.path.expanduser("~/Documents/cfms_errors.log")
INVALID_CASES = os.path.expanduser("~/Documents/invalidcases.txt")  # 🔥 NEW
# Global retry configuration
MAX_RETRIES_PER_CASE = 5
MAX_ELEMENT_RETRIES = 3
ELEMENT_RETRY_DELAY = 2
BROWSER_RESTART_DELAY = 5
WATCHDOG_TIMEOUT = 300

# Global state
CASE_FIR_NUMBER = "1345"
CASE_FIR_YEAR = "2024"
CASE_POLICE_STATION = "Islam Pura"
CASE_ACT = "PPC"
CASE_OFFENCE = "155C"
DECISION_DATE = "2025-11-21"
DECISION_TYPE = "Acquitted Due to Compromised"
DECISION_MAPPING = {
    "Agreed": ("Cancellation accepted by court", "No"),
    "Convicted": ("Conviction", "Other imprisonment"),
    "Fined": ("Conviction", "Fine only"),
    "Acquitted": ("Acquittal", "Poor Investigation"),
    "Acquitted Due to Compromised": ("Acquittal", "Due to Compromise"),
    "u/s 249-A": ("Acquittal", "Poor Investigation"),
    "u/s 345 Cr.P.C": ("Acquittal", "Due to Compromise"),
    "Consign to Record u/s 512": ("Consign to Record", "512 Cr.P.C"),
    "Consign to Record Room": ("Consign to Record", "512 Cr.P.C")
}

# 🔥 FIXED: Urdu + English Police Station mapping
POLICE_STATION_MAPPING = {
    # Urdu → English
    "شفیق‌آباد": "Shafiq Abad",
    "شفیقآباد": "Shafiq Abad",
    "گلشن‌راوی": "Gulshan-e-Ravi",
    "گلشنراوی": "Gulshan-e-Ravi",
    "اسلام‌پورہ": "Islam Pura",
    "اسلامپورہ": "Islam Pura",
    
    # 🔥 NEW: English → English normalization
    "Gulshan Ravi": "Gulshan-e-Ravi",
    "Gulshan ravi": "Gulshan-e-Ravi",
    "gulshan ravi": "Gulshan-e-Ravi",
    "GULSHAN RAVI": "Gulshan-e-Ravi",
    "GulshanRavi": "Gulshan-e-Ravi",
    "Shafiq Abad": "Shafiq Abad",
    "ShafiqAbad": "Shafiq Abad",
    "Islam Pura": "Islam Pura",
    "IslamPura": "Islam Pura",
}

def translate_station(urdu_or_english_station):
    """Convert Urdu/English station name to normalized English"""
    station = urdu_or_english_station.strip()
    
    # Try exact match first
    if station in POLICE_STATION_MAPPING:
        return POLICE_STATION_MAPPING[station]
    
    # Try case-insensitive match
    for key, value in POLICE_STATION_MAPPING.items():
        if station.lower() == key.lower():
            return value
    
    # Try normalized match (remove spaces/dashes)
    normalized = station.replace(" ", "").replace("‌", "").replace("-", "").lower()
    for key, value in POLICE_STATION_MAPPING.items():
        key_norm = key.replace(" ", "").replace("‌", "").replace("-", "").lower()
        if normalized == key_norm:
            return value
    
    # 🔥 NEW: Fuzzy match for variations
    best_match = None
    best_score = 0.0
    
    for key, value in POLICE_STATION_MAPPING.items():
        key_norm = key.replace(" ", "").replace("‌", "").replace("-", "").lower()
        score = SequenceMatcher(None, normalized, key_norm).ratio()
        if score > best_score:
            best_score = score
            best_match = value
    
    # If 80%+ match, use it
    if best_match and best_score >= 0.80:
        logging.info(f"✓ Station fuzzy-matched: '{station}' → '{best_match}' ({best_score*100:.1f}%)")
        return best_match
    
    # Return as-is if no match
    logging.warning(f"⚠ No translation for: {station} (best: {best_score*100:.1f}%)")
    return station

# Error tracking
error_counts = {}
last_error_time = {}

def log_error(error_type, message, exception=None):
    """Log errors to both console and file"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    error_msg = f"[{timestamp}] {error_type}: {message}"
    
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
    """Handle operation timeout"""
    raise TimeoutError("Operation exceeded watchdog timeout")

def sync_to_gdrive():
    """Robust STATUS.txt sync to Google Drive"""
    try:
        local_file = os.path.expanduser("~/GoogleDrive/STATUS.txt")
        
        if not os.path.exists(local_file):
            logging.error("❌ STATUS.txt not found")
            return False
        
        os.sync()
        time.sleep(0.3)
        
        result = subprocess.run(
            ["rclone", "copy", local_file, "gdrive:GoogleDrive", "-v"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            logging.info("✅ STATUS.txt synced to Google Drive")
            return True
        else:
            logging.error(f"❌ Sync failed: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        logging.error("❌ Sync timeout (30s)")
        return False
    except Exception as e:
        logging.error(f"❌ Sync error: {e}")
        return False

class BrowserManager:
    """Manages browser lifecycle with automatic recovery"""
    
    def __init__(self):
        self.driver = None
        self.wait = None
        self.restart_count = 0
        
    def start_browser(self):
        """Start browser with error handling"""
        max_attempts = 3
        
        for attempt in range(max_attempts):
            try:
                logging.info(f"Starting browser (attempt {attempt + 1}/{max_attempts})...")
                
                # Kill any existing chrome processes
                os.system("pkill -f chrome >/dev/null 2>&1")
                fast_sleep(2)
                
                opts = Options()
                opts.add_argument("--start-maximized")
                opts.add_argument("--user-data-dir=/home/control/selenium-profile")
                opts.add_argument("--disable-dev-shm-usage")
                opts.add_argument("--no-sandbox")
                opts.add_argument("--disable-gpu")
                opts.add_experimental_option('excludeSwitches', ['enable-logging'])
                
                self.driver = webdriver.Chrome(options=opts)
                self.wait = WebDriverWait(self.driver, 20)
                
                # Test browser
                self.driver.get("https://cfms.prosecution.punjab.gov.pk/#/")
                fast_sleep(3)
                
                logging.info("✓ Browser started successfully")
                self.restart_count += 1
                return True
                
            except Exception as e:
                log_error("BROWSER_START", f"Failed to start browser: {e}", e)
                
                if self.driver:
                    try:
                        self.driver.quit()
                    except:
                        pass
                
                if attempt < max_attempts - 1:
                    fast_sleep(BROWSER_RESTART_DELAY * (attempt + 1))
                else:
                    return False
        
        return False
    
    def restart_browser(self):
        """Restart browser completely"""
        logging.info("="*60)
        logging.info("🔄 RESTARTING BROWSER...")
        logging.info("="*60)
        
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass
        
        fast_sleep(BROWSER_RESTART_DELAY)
        gc.collect()
        
        return self.start_browser()
    
    def is_browser_alive(self):
        """Check if browser is responsive"""
        try:
            _ = self.driver.current_url
            return True
        except:
            return False

def safe_find_element(driver, wait, by, value, description="element", timeout=10):
    """Find element with retry logic"""
    for attempt in range(MAX_ELEMENT_RETRIES):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            if attempt < MAX_ELEMENT_RETRIES - 1:
                logging.warning(f"⚠ {description} not found, retry {attempt + 1}/{MAX_ELEMENT_RETRIES}")
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                log_error("ELEMENT_NOT_FOUND", f"Could not find {description}: {value}")
                return None
        except Exception as e:
            log_error("ELEMENT_FIND_ERROR", f"Error finding {description}: {e}", e)
            return None
    
    return None

def safe_click(driver, wait, element_or_locator, description="element", use_js=False):
    """Click element with comprehensive retry logic"""
    for attempt in range(MAX_ELEMENT_RETRIES):
        try:
            # Get element if locator provided
            if isinstance(element_or_locator, tuple):
                element = safe_find_element(driver, wait, *element_or_locator, description)
                if not element:
                    return False
            else:
                element = element_or_locator
            
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            fast_sleep(0.3)
            
            # Wait for clickable
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(element))
            
            # Click method
            if use_js:
                driver.execute_script("arguments[0].click();", element)
            else:
                element.click()
            
            logging.info(f"✓ Clicked {description}")
            return True
            
        except StaleElementReferenceException:
            if attempt < MAX_ELEMENT_RETRIES - 1:
                logging.warning(f"Stale element on {description}, retrying...")
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                log_error("STALE_ELEMENT", f"Stale element after retries: {description}")
                return False
                
        except Exception as e:
            if attempt < MAX_ELEMENT_RETRIES - 1:
                logging.warning(f"Click failed on {description}, retry {attempt + 1}")
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                log_error("CLICK_ERROR", f"Failed to click {description}: {e}", e)
                return False
    
    return False

def select_dropdown_robust(driver, wait, label, val):
    """Enhanced dropdown selector with better fuzzy matching"""
    for attempt in range(MAX_ELEMENT_RETRIES):
        try:
            logging.info(f"Selecting '{val}' from '{label}' (attempt {attempt + 1})...")
            
            xpath = f"//label[contains(text(),'{label}')]/following-sibling::div//span[contains(@class,'select2-selection')]"
            
            if not safe_click(driver, wait, (By.XPATH, xpath), f"{label} dropdown"):
                if attempt < MAX_ELEMENT_RETRIES - 1:
                    continue
                return False
            
            fast_sleep(0.2)
            
            # Search field
            search_field = safe_find_element(driver, wait, By.CSS_SELECTOR, ".select2-search__field", "search field", 5)
            if not search_field:
                if attempt < MAX_ELEMENT_RETRIES - 1:
                    continue
                return False
            
            search_field.clear()
            search_field.send_keys(val)
            fast_sleep(0.3)
            
            # Wait for results
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__options"))
            )
            fast_sleep(0.1)
            
            options = driver.find_elements(By.XPATH, "//li[contains(@class,'select2-results__option')]")
            
            # Try exact match first
            for opt in options:
                try:
                    opt_text = opt.text.strip()
                    if norm(val) == norm(opt_text):
                        opt.click()
                        logging.info(f"✓ Selected '{opt_text}' (exact match)")
                        return True
                except:
                    continue
            
            # Try fuzzy match with SequenceMatcher
            best_match = None
            best_score = 0.0
            
            for opt in options:
                try:
                    opt_text = opt.text.strip()
                    score = SequenceMatcher(None, norm(val), norm(opt_text)).ratio()
                    
                    if score > best_score:
                        best_score = score
                        best_match = opt
                except:
                    continue
            
            # Click best match if confidence > 70%
            if best_match and best_score >= 0.70:
                best_match.click()
                logging.info(f"✓ Selected '{best_match.text}' (fuzzy match: {best_score*100:.1f}%)")
                return True
            
            logging.warning(f"No good match for '{val}' (best: {best_score*100:.1f}%)")
            if attempt < MAX_ELEMENT_RETRIES - 1:
                fast_sleep(ELEMENT_RETRY_DELAY)
                continue
            
            return False
            
        except Exception as e:
            log_error("DROPDOWN_ERROR", f"Error selecting {label}: {e}", e)
            if attempt < MAX_ELEMENT_RETRIES - 1:
                fast_sleep(ELEMENT_RETRY_DELAY)
            else:
                return False
    
    return False

def norm(txt):
    return txt.lower().replace(" ", "").replace("-", "").replace("_", "").replace(".", "") if txt else ""

def verify_station_selected(driver, expected_station):
    """Verify police station was actually selected"""
    try:
        selected_span = driver.find_element(By.XPATH, 
            "//label[contains(text(),'Police Station')]/following-sibling::div//span[@class='select2-selection__rendered']")
        
        selected_text = selected_span.text.strip()
        
        if selected_text == "Select Police Station" or not selected_text:
            logging.error(f"❌ Station NOT selected (still shows placeholder)")
            return False
        
        # Check if it's a reasonable match
        score = SequenceMatcher(None, norm(expected_station), norm(selected_text)).ratio()
        
        if score >= 0.70:
            logging.info(f"✓ Station verified: '{selected_text}'")
            return True
        else:
            logging.warning(f"⚠ Selected station mismatch: expected '{expected_station}', got '{selected_text}'")
            return False
            
    except Exception as e:
        logging.error(f"❌ Could not verify station selection: {e}")
        return False

def select_police_station_verified(driver, wait, station_name):
    """Select police station with post-selection verification"""
    max_attempts = 3
    
    for attempt in range(max_attempts):
        try:
            if attempt > 0:
                logging.info(f"Retry {attempt+1}/{max_attempts} for station selection")
            
            # Click dropdown
            if not safe_click(driver, wait, 
                (By.XPATH, "//label[contains(text(),'Police Station')]/following::span[contains(@class,'select2-selection')]"), 
                "Police Station"):
                if attempt < max_attempts - 1:
                    fast_sleep(1)
                    continue
                return False
            
            fast_sleep(0.5)
            
            # Search
            search = safe_find_element(driver, wait, 
                By.XPATH, "//input[@class='select2-search__field']", 
                "search field", timeout=5)
            if not search:
                if attempt < max_attempts - 1:
                    fast_sleep(1)
                    continue
                return False
            
            # Clear and type search query
            search.clear()
            fast_sleep(0.2)
            
            # Type character by character (more reliable)
            search_text = station_name[:5]
            for char in search_text:
                search.send_keys(char)
                fast_sleep(0.05)
            
            logging.info(f"Searching for: '{search_text}'")
            fast_sleep(1.5)  # 🔥 Increased wait for results
            
            # Wait for results to load
            try:
                WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__options"))
                )
            except TimeoutException:
                logging.error("❌ No search results appeared")
                if attempt < max_attempts - 1:
                    # Close dropdown and retry
                    try:
                        driver.find_element(By.TAG_NAME, "body").click()
                        fast_sleep(0.5)
                    except:
                        pass
                    continue
                return False
            
            fast_sleep(0.5)
            
            # Get all options - try multiple times
            options = []
            for wait_attempt in range(3):
                options = driver.find_elements(By.XPATH, 
                    "//li[contains(@class,'select2-results__option') and not(contains(@class,'loading'))]")
                
                if options:
                    break
                
                fast_sleep(0.3)
            
            # Filter out empty/loading options
            valid_options = []
            for opt in options:
                try:
                    opt_text = opt.text.strip()
                    if opt_text and opt_text.lower() not in ['loading...', 'searching...', 'no results']:
                        valid_options.append((opt, opt_text))
                except:
                    continue
            
            # 🔥 DEBUG: Log what we found
            logging.info(f"🔍 Found {len(valid_options)} valid options")
            for i, (opt, opt_text) in enumerate(valid_options):
                logging.info(f"  [{i+1}] '{opt_text}'")
            
            if not valid_options:
                logging.error("❌ No valid options found")
                if attempt < max_attempts - 1:
                    try:
                        driver.find_element(By.TAG_NAME, "body").click()
                        fast_sleep(0.5)
                    except:
                        pass
                    continue
                return False
            
            # Try exact match first
            for opt, opt_text in valid_options:
                if norm(station_name) == norm(opt_text):
                    logging.info(f"✓ Exact match found: '{opt_text}'")
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", opt)
                        fast_sleep(0.2)
                        opt.click()
                        fast_sleep(0.5)
                        logging.info(f"✓ Clicked station: {opt_text}")
                        
                        # Verify selection
                        if verify_station_selected(driver, station_name):
                            return True
                        else:
                            logging.warning("⚠ Verification failed after exact match")
                            if attempt < max_attempts - 1:
                                continue
                    except Exception as e:
                        logging.error(f"Error clicking exact match: {e}")
                        if attempt < max_attempts - 1:
                            continue
            
            # Try fuzzy match
            best_match = None
            best_opt = None
            best_score = 0.0
            
            for opt, opt_text in valid_options:
                score = SequenceMatcher(None, norm(station_name), norm(opt_text)).ratio()
                logging.info(f"  Fuzzy: '{opt_text}' = {score*100:.1f}%")
                
                if score > best_score:
                    best_score = score
                    best_match = opt_text
                    best_opt = opt
            
            if best_opt and best_score >= 0.70:
                logging.info(f"✓ Fuzzy match found: '{best_match}' ({best_score*100:.1f}%)")
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", best_opt)
                    fast_sleep(0.2)
                    best_opt.click()
                    fast_sleep(0.5)
                    logging.info(f"✓ Clicked station: {best_match}")
                    
                    # Verify selection
                    if verify_station_selected(driver, station_name):
                        return True
                    else:
                        logging.warning("⚠ Verification failed after fuzzy match")
                        if attempt < max_attempts - 1:
                            continue
                except Exception as e:
                    logging.error(f"Error clicking fuzzy match: {e}")
                    if attempt < max_attempts - 1:
                        continue
            
            logging.error(f"❌ No good match (best: '{best_match}' = {best_score*100:.1f}%)")
            
            if attempt < max_attempts - 1:
                # Close dropdown and retry
                try:
                    driver.find_element(By.TAG_NAME, "body").click()
                    fast_sleep(0.5)
                except:
                    pass
                continue
            
            return False
            
        except Exception as e:
            log_error("STATION_SELECT", f"Error selecting station (attempt {attempt+1}): {e}", e)
            if attempt < max_attempts - 1:
                try:
                    driver.find_element(By.TAG_NAME, "body").click()
                    fast_sleep(0.5)
                except:
                    pass
                continue
            return False
    
    return False

def fuzzy_match_decision(raw_text):
    """
    SUPER AGGRESSIVE fuzzy matcher for terrible formats
    Handles: Urdu, English, typos, mixed formats, partial text
    """
    if not raw_text:
        return None, 0.0
    
    # Clean input AGGRESSIVELY
    cleaned = raw_text.lower().strip()
    cleaned = ''.join(c for c in cleaned if c.isalnum() or c.isspace() or ord(c) > 127)
    cleaned = ' '.join(cleaned.split())
    
    best_match = None
    best_score = 0.0
    
    # All possible variations to check against
    decision_variations = {
        "Agreed": ["agreed", "agree", "منظور", "منظورشد", "منظورہ", "منظورهشد"],
        "Convicted": ["convicted", "convict", "conviction", "سزا", "سزاشد", "سزائے", "بریشد"],
        "Fined": ["fined", "fine", "جرمانہ", "جرمانے"],
        "Acquitted": ["acquitted", "acquit", "acquittal", "بری", "بریشد", "بریٔت"],
        "Acquitted Due to Compromised": ["compromised", "compromise", "سمجھوتہ", "acquitted due to compromised"],
        "u/s 249-A": ["249", "249a", "249-a", "قزیردفعہ249", "us249a", "us249-a"],
        "u/s 345 Cr.P.C": ["345", "345crpc", "us345", "دفعہ345", "us345crpc"],
        "Consign to Record u/s 512": ["512", "consign 512", "قزیردفعہ512", "داخلدفترزیردفعہ512", "us512"],
        "Consign to Record Room": ["consign", "record room", "داخلدفتر", "فیصلہشد", "dakhil", "دفتر"]
    }
    
    # Try exact match first
    for decision, variations in decision_variations.items():
        for variant in variations:
            # Direct substring match
            if variant in cleaned or cleaned in variant:
                score = SequenceMatcher(None, cleaned, variant).ratio()
                if score > best_score:
                    best_score = score
                    best_match = decision
            
            # Fuzzy match
            score = SequenceMatcher(None, cleaned, variant).ratio()
            if score > best_score:
                best_score = score
                best_match = decision
    
    # Try matching with DECISION_MAPPING keys too
    for key in DECISION_MAPPING.keys():
        key_clean = key.lower().strip()
        score = SequenceMatcher(None, cleaned, key_clean).ratio()
        if score > best_score:
            best_score = score
            best_match = key
    
    # Word-level matching (for multi-word inputs)
    words = cleaned.split()
    if len(words) > 1:
        for decision, variations in decision_variations.items():
            for variant in variations:
                for word in words:
                    if len(word) > 2:
                        score = SequenceMatcher(None, word, variant).ratio()
                        if score > best_score:
                            best_score = score
                            best_match = decision
    
    return best_match, best_score

def get_next_case():
    """Get next case with error handling"""
    try:
        if not os.path.exists(TO_BE_FILLED):
            logging.warning(f"File not found: {TO_BE_FILLED}")
            return None
        
        with open(TO_BE_FILLED, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        if len(lines) <= 1:
            logging.info("Only header remaining - no more cases")
            return None
        
        case_line = lines[1].strip()
        
        if not case_line:
            logging.warning("Empty line found - skipping")
            with open(TO_BE_FILLED, 'w', encoding='utf-8') as f:
                f.write(lines[0])
                if len(lines) > 2:
                    f.writelines(lines[2:])
            return get_next_case()
        
        return case_line
        
    except Exception as e:
        log_error("FILE_READ", f"Error reading to_be_filled.txt: {e}", e)
        return None
def check_fir_not_found(driver, wait, case_data):
    """Check for 'FIR not found' error popup"""
    try:
        popup = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'swal2-popup')]"))
        )
        txt = popup.text.lower()
        
        if 'fir data not found' in txt or 'record not found' in txt:
            logging.info("⚠ FIR NOT FOUND IN SYSTEM!")
            
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            try:
                with open(INVALID_CASES, 'a', encoding='utf-8') as f:
                    f.write(f"{current_time}\t{case_data}\n")
                logging.info(f"✓ Data appended to: {INVALID_CASES}")
            except Exception as e:
                log_error("INVALID_CASE_LOG", f"Failed to log invalid case: {e}", e)
            
            mark_as_filled(case_data)
            
            try:
                cancel_btn = driver.find_element(By.XPATH, "//button[normalize-space()='Cancel']")
                cancel_btn.click()
                fast_sleep(0.5)
            except:
                pass
            
            return True
        
        return False
    except:
        return False
def mark_as_filled(case_data):
    """Mark case as complete with error handling"""
    try:
        with open(FILLED_ENTRIES, 'a', encoding='utf-8') as f:
            f.write(case_data + '\n')
        logging.info(f"✓ Added to filled_entries.txt")
    except Exception as e:
        log_error("FILE_WRITE", f"Error writing to filled_entries.txt: {e}", e)
    
    try:
        with open(TO_BE_FILLED, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        with open(TO_BE_FILLED, 'w', encoding='utf-8') as f:
            f.write(lines[0])
            if len(lines) > 2:
                f.writelines(lines[2:])
        logging.info(f"✓ Removed from to_be_filled.txt")
    except Exception as e:
        log_error("FILE_UPDATE", f"Error updating to_be_filled.txt: {e}", e)

def parse_case_data(row_data):
    """Parse case data with SUPER AGGRESSIVE fuzzy matching - HANDLES 6, 7, OR 8 FIELDS"""
    global CASE_FIR_NUMBER, CASE_FIR_YEAR, CASE_POLICE_STATION
    global CASE_ACT, CASE_OFFENCE, DECISION_DATE, DECISION_TYPE
    
    try:
        # Split by tab
        fields = row_data.split('\t')
        
        # Clean up fields (remove extra spaces)
        fields = [f.strip() for f in fields if f.strip()]
        
        logging.info(f"Parsed {len(fields)} fields from row")
        
        urdu_station = ""
        
        # Handle different formats
        if len(fields) == 7:
            # Format: name | FIR | year | offence | station | date | decision
            logging.info("Using 7-field format: name | FIR | year | offence | station | date | decision")
            
            name = fields[0].strip()
            CASE_FIR_NUMBER = fields[1].strip()
            CASE_FIR_YEAR = fields[2].strip()
            CASE_OFFENCE = fields[3].strip().replace(" PPC", "").replace("PPC", "").strip()
            urdu_station = fields[4].strip()
            date_str = fields[5].strip()
            raw_decision = fields[6].strip()
            
        elif len(fields) == 8:
            # Format: case | name | FIR | year | offence | station | date | decision
            logging.info("Using 8-field format: case | name | FIR | year | offence | station | date | decision")
            
            case_number = fields[0].strip()
            name = fields[1].strip()
            CASE_FIR_NUMBER = fields[2].strip()
            CASE_FIR_YEAR = fields[3].strip()
            CASE_OFFENCE = fields[4].strip().replace(" PPC", "").replace("PPC", "").strip()
            urdu_station = fields[5].strip()
            date_str = fields[6].strip()
            raw_decision = fields[7].strip()
            
        elif len(fields) == 6:
            # Format: name | case_num | year | offence | police_station | date:decision
            logging.info("Using 6-field format with date:decision")
            
            last_field = fields[-1].strip()
            
            if ':' in last_field:
                date_decision = last_field.split(':', 1)
                date_str = date_decision[0].strip()
                raw_decision = date_decision[1].strip() if len(date_decision) > 1 else ""
                
                name = fields[0].strip()
                CASE_FIR_NUMBER = fields[1].strip()
                CASE_FIR_YEAR = fields[2].strip()
                CASE_OFFENCE = fields[3].strip().replace(" PPC", "").replace("PPC", "").strip()
                urdu_station = fields[4].strip()
            else:
                raise ValueError("Expected date:decision format in last field")
        else:
            raise ValueError(f"Invalid row format - expected 6, 7, or 8 fields, got {len(fields)}")
        
        # 🔥 TRANSLATE STATION (handles both Urdu and English)
        CASE_POLICE_STATION = translate_station(urdu_station)
        
        if urdu_station != CASE_POLICE_STATION:
            logging.info(f"✓ Station translated: '{urdu_station}' → '{CASE_POLICE_STATION}'")
        else:
            logging.info(f"✓ Station: '{CASE_POLICE_STATION}'")
        
        # Validate required fields
        if not CASE_FIR_NUMBER or not CASE_FIR_YEAR or not CASE_POLICE_STATION:
            raise ValueError("FIR Number, Year, or Police Station is empty")
        
        # Parse date - handle formats like "30-10-2025" or "30-10-25"
        date_parts = date_str.replace(':', '').split('-')
        
        if len(date_parts) == 3:
            day = date_parts[0].zfill(2)
            month = date_parts[1].zfill(2)
            year = date_parts[2]
            
            # If year is 2-digit, convert to 4-digit
            if len(year) == 2:
                year = "20" + year
            
            DECISION_DATE = f"{year}-{month}-{day}"
        else:
            raise ValueError(f"Invalid date format: {date_str}")
        
        # 🔥 SUPER AGGRESSIVE FUZZY MATCH FOR DECISION
        matched_decision, match_score = fuzzy_match_decision(raw_decision)
        
        if matched_decision and match_score >= 0.70:
            DECISION_TYPE = matched_decision
            logging.info(f"✓ Decision matched: '{raw_decision}' → '{matched_decision}' ({match_score*100:.1f}% match)")
        else:
            # If no match, log warning but try to continue with raw text
            logging.warning(f"⚠ Decision '{raw_decision}' NO GOOD MATCH (best: {match_score*100:.1f}% - {matched_decision})")
            
            # Try one more time with even more aggressive matching
            if matched_decision and match_score >= 0.50:
                DECISION_TYPE = matched_decision
                logging.warning(f"⚠ WEAK MATCH ACCEPTED: '{raw_decision}' → '{matched_decision}' ({match_score*100:.1f}%)")
            else:
                raise ValueError(f"Decision '{raw_decision}' cannot be matched (best: {match_score*100:.1f}%)")
        
        CASE_ACT = "PPC"
        
        logging.info(f"✓ Parsed Case: FIR {CASE_FIR_NUMBER}/{CASE_FIR_YEAR} | {CASE_POLICE_STATION} | {DECISION_TYPE}")
        return True
        
    except Exception as e:
        log_error("PARSE_ERROR", f"Failed to parse row: {e}", e)
        logging.error(f"Raw data: {row_data}")
        try:
            logging.error(f"Split into {len(fields)} fields: {fields}")
        except:
            pass
        return False

def check_framing_required(driver, wait, case_data):
    """Check for framing error popup"""
    try:
        popup = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'swal2-popup')]"))
        )
        txt = popup.text.lower()
        
        if 'framing' in txt and 'charge' in txt:
            logging.info("⚠ FRAMING OF CHARGE REQUIRED!")
            
            txt_path = os.path.expanduser("~/Documents/framing_errors.txt")
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            try:
                with open(txt_path, 'a', encoding='utf-8') as f:
                    f.write(f"{current_time}\t{case_data}\n")
                logging.info(f"✓ Data appended to: {txt_path}")
            except Exception as e:
                log_error("FRAMING_LOG", f"Failed to log framing error: {e}", e)
            
            mark_as_filled(case_data)
            
            try:
                ok_btn = driver.find_element(By.XPATH, "//button[normalize-space()='OK']")
                ok_btn.click()
                fast_sleep(0.5)
            except:
                pass
            
            return True
        
        return False
    except:
        return False


def fill_decision_fields(driver, wait, decision, detail, max_retries=3):
    """
    Enterprise-grade Decision/Detail filler
    - Bottom-to-top
    - Re-finds DOM every step
    - Verifies after fill
    - Retries per section
    - ONLY FILLS EMPTY FIELDS
    """

    def get_sections():
        return driver.find_elements(
            By.XPATH,
            "//div[@id='IdCrudModal']//div[.//label[contains(text(),'Decision')] and .//select]"
        )

    def select2_fill(span, value):
        span.click()
        search = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".select2-search__field"))
        )
        search.clear()
        search.send_keys(value)
        fast_sleep(0.3)
        search.send_keys(Keys.ENTER)
        fast_sleep(0.5)

    try:
        sections = get_sections()
        total = len(sections)
        logging.info(f"Found {total} Decision section(s)")

        # 🔥 BOTTOM → TOP (prevents last-row skip)
        for index in reversed(range(total)):
            for attempt in range(1, max_retries + 1):
                try:
                    sections = get_sections()
                    section = sections[index]

                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", section
                    )
                    fast_sleep(0.3)

                    # ================= DECISION =================
                    decision_span = section.find_element(
                        By.XPATH,
                        ".//label[contains(text(),'Decision')]/following-sibling::div//span[contains(@class,'select2-selection')]"
                    )

                    current_decision = decision_span.text.strip()
                    
                    # ✅ ONLY FILL IF EMPTY
                    if not current_decision or current_decision == "-- Select --" or "select" in current_decision.lower():
                        select2_fill(decision_span, decision)

                        # ✅ VERIFY
                        if norm(decision) not in norm(decision_span.text):
                            raise Exception("Decision verification failed")

                        logging.info(f"✓ Decision set [{index+1}/{total}]")
                    else:
                        logging.info(f"⊘ Decision already filled [{index+1}/{total}]: {current_decision}")

                    # ================= DETAIL =================
                    try:
                        detail_span = section.find_element(
                            By.XPATH,
                            ".//label[contains(text(),'Detail')]/following-sibling::div//span[contains(@class,'select2-selection')]"
                        )

                        detail_select = section.find_element(
                            By.XPATH,
                            ".//label[contains(text(),'Detail')]/following-sibling::div//select"
                        )

                        WebDriverWait(driver, 10).until(
                            lambda d: len(detail_select.find_elements(By.TAG_NAME, "option")) > 1
                        )

                        current_detail = detail_span.text.strip()
                        
                        # ✅ ONLY FILL IF EMPTY
                        if not current_detail or current_detail == "-- Select --" or "select" in current_detail.lower():
                            select2_fill(detail_span, detail)

                            # ✅ VERIFY
                            if norm(detail) not in norm(detail_span.text):
                                raise Exception("Detail verification failed")

                            logging.info(f"✓ Detail set [{index+1}/{total}]")
                        else:
                            logging.info(f"⊘ Detail already filled [{index+1}/{total}]: {current_detail}")

                    except NoSuchElementException:
                        logging.warning(f"⚠ No Detail field in section {index+1}")

                    # ✅ SUCCESS → break retry loop
                    break

                except Exception as e:
                    logging.warning(
                        f"Retry {attempt}/{max_retries} failed for section {index+1}: {e}"
                    )
                    if attempt == max_retries:
                        log_error(
                            "DECISION_SECTION_FAILED",
                            f"Section {index+1} failed after retries",
                            e,
                        )

        return True

    except Exception as e:
        log_error("FILL_DECISION_FATAL", f"Fatal error: {e}", e)
        return False

def automate_final_order(driver, wait, date_val, dec_type, case_data):
    """Process final order with comprehensive error handling"""
    logging.info(f"Starting Final Order - Type: {dec_type}")
    
    if dec_type not in DECISION_MAPPING:
        log_error("INVALID_DECISION", f"Decision type '{dec_type}' not in mapping")
        return "ERROR"
    
    decision, detail = DECISION_MAPPING[dec_type]
    failed_attempts = {}
    MAX_ATTEMPTS_PER_INDEX = 3
    
    # 🔥 CHECK IF CANCELLATION - ONLY FILL FIRST ENTRY
    is_cancellation = (decision == "Cancellation accepted by court" and detail == "No")
    
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr")))
        fast_sleep(0.5)
        
        btns = driver.find_elements(By.XPATH, "//tbody/tr/td[last()]//a[@id='roleActions']")
        total_entries = len(btns)
        logging.info(f"Found {total_entries} entry/entries")
        
        if is_cancellation:
            logging.info("⚠ CANCELLATION TYPE - Only filling FIRST entry (auto-fills all)")
        
        processed_count = 0
        i = 0
        
        while i < total_entries:
            # 🔥 SKIP ALL EXCEPT FIRST IF CANCELLATION
            if is_cancellation and i > 0:
                logging.info(f"⊘ Skipping entry {i+1} (Cancellation auto-fill)")
                i += 1
                continue
            
            logging.info(f"\n{'='*60}\nEntry {i+1}/{total_entries}\n{'='*60}")
            
            try:
                btns = driver.find_elements(By.XPATH, "//tbody/tr/td[last()]//a[@id='roleActions']")
                
                if i >= len(btns):
                    logging.warning(f"Entry {i+1} not found")
                    break
                
                if failed_attempts.get(i, 0) >= MAX_ATTEMPTS_PER_INDEX:
                    logging.error(f"Skipping entry {i+1} after max attempts")
                    i += 1
                    continue
                
                if not safe_click(driver, wait, btns[i], f"three-dot menu {i+1}"):
                    failed_attempts[i] = failed_attempts.get(i, 0) + 1
                    continue
                
                fast_sleep(0.5)
                
            except Exception as e:
                log_error("THREE_DOT_CLICK", f"Error at entry {i+1}: {e}", e)
                failed_attempts[i] = failed_attempts.get(i, 0) + 1
                i += 1
                continue
            
            # Click Edit
            edit_clicked = False
            
            try:
                edit_btn = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, 
                        "//div[contains(@class,'dropdown-menu') and contains(@class,'show')]//a[contains(text(),'Edit')]"))
                )
                edit_btn.click()
                edit_clicked = True
                logging.info("✓ Clicked Edit from dropdown")
            except TimeoutException:
                try:
                    edit_btn = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH, 
                            "//a[normalize-space()='Edit' and contains(@class,'btn')]"))
                    )
                    edit_btn.click()
                    edit_clicked = True
                    logging.info("✓ Clicked floating Edit")
                except TimeoutException:
                    logging.warning(f"No Edit for entry {i+1}")
                    try:
                        driver.find_element(By.TAG_NAME, "body").click()
                    except:
                        pass
                    i += 1
                    continue
            
            fast_sleep(0.5)
            
            # Check framing error
            if check_framing_required(driver, wait, case_data):
                return "NEXT_CASE"
            
            # Wait for modal
            try:
                wait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//div[contains(@class,'modal') and contains(@class,'show')]")
                ))
            except TimeoutException:
                log_error("MODAL_TIMEOUT", f"Modal didn't appear for entry {i+1}")
                failed_attempts[i] = failed_attempts.get(i, 0) + 1
                continue
            
            fast_sleep(0.5)
            
            # Fill date
            try:
                date_inp = driver.find_element(By.XPATH, "//div[contains(@class,'modal')]//input[@type='date']")
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].blur();
                """, date_inp, date_val)
                fast_sleep(0.5)
                logging.info(f"✓ Date filled: {date_val}")
            except Exception as e:
                log_error("DATE_FILL", f"Error filling date: {e}", e)
            
            # Fill decision fields
            fill_decision_fields(driver, wait, decision, detail)
            
            # Click Update
            try:
                upd_btn = driver.find_element(By.XPATH, 
                    "//div[contains(@class,'modal')]//button[normalize-space()='Update']")
                driver.execute_script("arguments[0].scrollIntoView(true);", upd_btn)
                fast_sleep(0.2)
                upd_btn.click()
                logging.info("✓ Clicked Update")
            except Exception as e:
                log_error("UPDATE_CLICK", f"Error clicking Update: {e}", e)
                failed_attempts[i] = failed_attempts.get(i, 0) + 1
                continue
            
            # Wait for modal close
            try:
                wait.until(EC.invisibility_of_element_located(
                    (By.XPATH, "//div[contains(@class,'modal') and contains(@class,'show')]")
                ))
                fast_sleep(0.3)
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "modal-backdrop")))
                logging.info("✓ Modal closed")
            except:
                pass
            
            fast_sleep(0.5)
            
            # Handle success popup
            try:
                ok_btn = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='OK']"))
                )
                ok_btn.click()
                logging.info("✓ Success popup handled")
            except:
                pass
            
            # Final cleanup
            try:
                wait.until(EC.invisibility_of_element_located((By.ID, "IdCrudModal")))
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "modal-backdrop")))
                fast_sleep(1)
            except:
                fast_sleep(1)
            
            processed_count += 1
            failed_attempts[i] = 0
            logging.info(f"✓✓✓ Entry {i+1} completed")
            
            # 🔥 IF CANCELLATION - STOP AFTER FIRST ENTRY
            if is_cancellation:
                logging.info("✓ Cancellation first entry filled - stopping (auto-fills rest)")
                break
            
            i += 1
            fast_sleep(1)
        
        logging.info(f"\n✓ Processed {processed_count} entry/entries")
        
        # 🔥 IF CANCELLATION - SKIP REFRESH CHECK
        if is_cancellation:
            logging.info("✓✓✓ CANCELLATION COMPLETE - No refresh needed")
            return "COMPLETE"
        
        # Refresh and check
        try:
            driver.refresh()
            fast_sleep(0.5)
            
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[normalize-space()='Final Order']")
            )).click()
            fast_sleep(0.5)
        except Exception as e:
            log_error("REFRESH_ERROR", f"Could not refresh: {e}", e)
        
        # Check remaining
        try:
            remaining_btns = driver.find_elements(By.XPATH, "//tbody/tr/td[last()]//a[@id='roleActions']")
            remaining_count = len(remaining_btns)
        except:
            remaining_count = 0
        
        logging.info(f"After refresh: {remaining_count} entries found")
        
        if remaining_count == 0:
            logging.info("✓✓✓ NO ENTRIES REMAINING - Complete!")
            return "COMPLETE"
        elif remaining_count < total_entries:
            logging.info(f"Entries reduced - continuing")
            return "RESTART"
        else:
            logging.warning("Entry count unchanged - forcing complete")
            return "COMPLETE"
        
    except Exception as e:
        log_error("FINAL_ORDER_ERROR", f"Final Order Failed: {e}", e)
        return "ERROR"

def handle_judicial_proceedings(driver, wait):
    """Handle Judicial Proceedings with error recovery"""
    logging.info("\n" + "="*60)
    logging.info("JUDICIAL PROCEEDINGS")
    logging.info("="*60)
    
    try:
        if not safe_click(driver, wait, (By.XPATH, "//a[normalize-space()='Judicial Proceedings']"), "Judicial Proceedings tab"):
            return False
        
        fast_sleep(1)
        
        rows = driver.find_elements(By.CSS_SELECTOR, ".laravel-vue-datatable-tbody tr")
        
        if rows and len(rows) > 0:
            logging.info("✓ Found existing record")
            
            try:
                three_dots = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "td:last-child .feather-more-horizontal")
                    )
                )
                
                if three_dots.is_displayed() and three_dots.is_enabled():
                    three_dots.click()
                    fast_sleep(0.3)
                    
                    edit_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//div[contains(@class,'dropdown-menu')]//a[contains(text(),'Edit')]")
                    ))
                    edit_btn.click()
                    fast_sleep(0.5)
                    return True
                else:
                    return False
                    
            except Exception as e:
                log_error("JUDICIAL_EDIT", f"Error editing judicial: {e}", e)
                return False
        else:
            logging.info("Adding new record")
            
            try:
                add_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Add New')]"))
                )
                add_btn.click()
                fast_sleep(0.5)
                return True
            except Exception as e:
                log_error("JUDICIAL_ADD", f"No Add New button: {e}", e)
                return False
                
    except Exception as e:
        log_error("JUDICIAL_PROCEEDINGS", f"Error in Judicial Proceedings: {e}", e)
        return False

def fill_court_modal(driver, wait):
    """Fill court modal with error recovery"""
    logging.info("Filling Court modal")
    
    try:
        # Category with retries
        for attempt in range(3):
            if select_dropdown_robust(driver, wait, "Category", "Magistrate"):
                fast_sleep(0.3)
                try:
                    cat_span = driver.find_element(By.XPATH, 
                        "//label[contains(text(),'Category')]/following-sibling::div//span[contains(@class,'select2-selection')]")
                    if "magistrate" in cat_span.text.lower():
                        logging.info("✓ Category verified")
                        break
                except:
                    pass
            
            if attempt == 2:
                log_error("CATEGORY_FILL", "Failed to select Category after 3 attempts")
                return False
            
            fast_sleep(0.5)
        
        fast_sleep(0.3)
        
        # Court Type with retries
        for attempt in range(3):
            if select_dropdown_robust(driver, wait, "Court Type", "Section 30 Magistrate"):
                fast_sleep(0.3)
                try:
                    ct_span = driver.find_element(By.XPATH, 
                        "//label[contains(text(),'Court Type')]/following-sibling::div//span[contains(@class,'select2-selection')]")
                    if "section 30" in ct_span.text.lower():
                        logging.info("✓ Court Type verified")
                        break
                except:
                    pass
            
            if attempt == 2:
                log_error("COURT_TYPE_FILL", "Failed to select Court Type after 3 attempts")
                return False
            
            fast_sleep(0.5)
        
        fast_sleep(0.3)
        
        # Click Save/Create
        button_clicked = False
        
        for btn_text in ["Save", "Create"]:
            if button_clicked:
                break
            
            try:
                btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, 
                        f"//button[@type='submit' and normalize-space()='{btn_text}']"))
                )
                
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                fast_sleep(0.1)
                
                btn.click()
                logging.info(f"✓ Clicked {btn_text}")
                button_clicked = True
                break
            except:
                pass
        
        if not button_clicked:
            log_error("SAVE_BUTTON", "Failed to click Save/Create")
            return False
        
        fast_sleep(0.5)
        
        # Handle popups
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='OK']")
            )).click()
        except:
            pass
        
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='Close']")
            )).click()
        except:
            pass
        
        fast_sleep(0.5)
        return True
        
    except Exception as e:
        log_error("COURT_MODAL", f"Error filling court modal: {e}", e)
        return False

def process_single_case(browser_mgr, case_data):
    """Process single case with comprehensive error handling"""
    driver = browser_mgr.driver
    wait = browser_mgr.wait
    
    try:
        # Check browser health
        if not browser_mgr.is_browser_alive():
            log_error("BROWSER_DEAD", "Browser unresponsive")
            return "ERROR"
        
        # Navigate to Cases
        logging.info("Navigating to Cases...")
        try:
            cases_menu = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[normalize-space()='Cases']")
            ))
            driver.execute_script("arguments[0].click();", cases_menu)
            fast_sleep(1)
        except Exception as e:
            log_error("CASES_NAV", f"Error navigating to Cases: {e}", e)
            return "ERROR"
        
        # Click OK
        if not safe_click(driver, wait, (By.XPATH, "//button[normalize-space()='OK']"), "OK button"):
            return "ERROR"
        
        # Click green button
        if not safe_click(driver, wait, (By.XPATH, "//button[contains(@class,'btn-success')]"), "green button"):
            return "ERROR"
        
        # Fill FIR
        fir = safe_find_element(driver, wait, By.XPATH, "//input[@placeholder='FIR Number']", "FIR field")
        if not fir:
            return "ERROR"
        
        fir.clear()
        fir.send_keys(CASE_FIR_NUMBER)
        
        # FIR Year
        if not safe_click(driver, wait, (By.XPATH, "//label[contains(text(),'FIR Year')]/following::span[contains(@class,'select2-selection')]"), "FIR Year"):
            return "ERROR"
        
        if not safe_click(driver, wait, (By.XPATH, f"//li[normalize-space()='{CASE_FIR_YEAR}']"), "Year option"):
            return "ERROR"
        
        # 🔥 FIXED: Police Station with verification
        if not select_police_station_verified(driver, wait, CASE_POLICE_STATION):
            logging.error("❌ Failed to select police station - aborting")
            return "ERROR"
        
        # Fetch FIR
       
        if not safe_click(driver, wait, (By.XPATH, "//button[contains(text(),'Fetch FIR Data')]"), "Fetch FIR"):
            return "ERROR"
        
        fast_sleep(1)  # 🔥 Wait for potential error popup
        
        # 🔥 Check if FIR not found
        if check_fir_not_found(driver, wait, case_data):
            return "NEXT_CASE"
        
        if not safe_click(driver, wait, (By.XPATH, "//button[normalize-space()='Edit Case']"), "Edit Case"):
            return "ERROR"
        
        fast_sleep(2)
        driver.refresh()
        fast_sleep(3)
        
        # Prosecution
        logging.info("\n" + "="*60)
        logging.info("PROSECUTION")
        logging.info("="*60)
        
        if not safe_click(driver, wait, (By.XPATH, "//a[normalize-space()='Prosecution']"), "Prosecution tab"):
            return "ERROR"
        
        fast_sleep(0.5)
        
        if not safe_click(driver, wait, (By.XPATH, "//button[contains(@class,'btn-success')]"), "green button"):
            return "ERROR"
        
        fast_sleep(0.5)
        
        role = safe_find_element(driver, wait, By.XPATH, "//select[.//option[contains(text(),'Conduct Trial')]]", "role dropdown")
        if not role:
            return "ERROR"
        
        Select(role).select_by_visible_text("Conduct Trial")
        fast_sleep(2.5)
        
        # Fill dates
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
                logging.info(f"✓ Date {idx+1} filled")
        
        fast_sleep(0.2)
        
        if not safe_click(driver, wait, (By.XPATH, "//button[normalize-space()='Create']"), "Create button"):
            return "ERROR"
        
        fast_sleep(1)
        
        # Handle popups
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='OK']")
            )).click()
        except:
            pass
        
        try:
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='Close']")
            )).click()
        except:
            pass
        
        fast_sleep(1)
        
        # Judicial Proceedings
        should_fill = handle_judicial_proceedings(driver, wait)
        
        if should_fill:
            if not fill_court_modal(driver, wait):
                log_error("COURT_MODAL_FAIL", "Failed to fill court modal")
        
        # Final Order
        logging.info("\n" + "="*60)
        logging.info("FINAL ORDER")
        logging.info("="*60)
        
        if not safe_click(driver, wait, (By.XPATH, "//a[normalize-space()='Final Order']"), "Final Order tab"):
            return "ERROR"
        
        fast_sleep(0.3)
        
        status = automate_final_order(driver, wait, DECISION_DATE, DECISION_TYPE, case_data)
        
        return status
        
    except Exception as e:
        log_error("CASE_PROCESSING", f"Error processing case: {e}", e)
        return "ERROR"

def main():
    """Main execution with self-healing"""
    browser_mgr = BrowserManager()
    
    # Start browser
    if not browser_mgr.start_browser():
        logging.error("❌ Failed to start browser after multiple attempts")
        return
    
    loop_count = 0
    max_loops = 1000
    consecutive_errors = 0
    MAX_CONSECUTIVE_ERRORS = 3
    
    try:
        while loop_count < max_loops:
            loop_count += 1
            logging.info("\n" + "#"*60)
            logging.info(f"### LOOP {loop_count} ###")
            logging.info("#"*60 + "\n")
            
            # Get next case
            case_line = get_next_case()
            if not case_line:
                logging.info("✓✓✓ NO MORE CASES!")
                break
            
            logging.info(f"Processing: {case_line}")
            
            # Parse case
            if not parse_case_data(case_line):
                logging.error("❌ Parse failed - skipping")
                mark_as_filled(case_line)
                consecutive_errors = 0
                continue
            
            # Try processing with retries
            case_retry_count = 0
            case_success = False
            
            while case_retry_count < MAX_RETRIES_PER_CASE and not case_success:
                case_retry_count += 1
                
                if case_retry_count > 1:
                    logging.info(f"⟳ Case retry {case_retry_count}/{MAX_RETRIES_PER_CASE}")
                
                # Check browser health before processing
                if not browser_mgr.is_browser_alive():
                    logging.warning("Browser died - restarting...")
                    if not browser_mgr.restart_browser():
                        log_error("BROWSER_RESTART_FAIL", "Failed to restart browser")
                        break
                
                # Process case with timeout
                try:
                    # Set watchdog timer
                    signal.signal(signal.SIGALRM, timeout_handler)
                    signal.alarm(WATCHDOG_TIMEOUT)
                    
                    status = process_single_case(browser_mgr, case_line)
                    
                    # Cancel watchdog
                    signal.alarm(0)
                    
                    if status == "COMPLETE":
                        mark_as_filled(case_line)
                        logging.info("✓✓✓ CASE COMPLETE")
                        
                        # 🔥 FIXED: Update STATUS.txt with proper counting
                        sf = os.path.expanduser("~/GoogleDrive/STATUS.txt")
                        
                        try:
                            # Count remaining
                            try:
                                with open(TO_BE_FILLED, 'r') as f:
                                    remaining = len([l for l in f.readlines()[1:] if l.strip()])
                            except:
                                remaining = 0
                            
                            # Count filled
                            try:
                                with open(FILLED_ENTRIES, 'r') as f:
                                    filled = len([l for l in f.readlines() if l.strip()])
                            except:
                                filled = 0
                            
                            # Read existing STATUS entries (only [xxx] lines)
                            existing_entries = []
                            try:
                                if os.path.exists(sf) and os.path.getsize(sf) > 0:
                                    with open(sf, 'r') as f:
                                        for line in f:
                                            if line.strip() and line.strip().startswith('['):
                                                existing_entries.append(line.strip())
                            except:
                                existing_entries = []
                            
                            # Calculate ETA
                            prediction = "ETA: Calculating..."
                            if len(existing_entries) >= 2:
                                times = []
                                for line in existing_entries:
                                    try:
                                        # Extract time: [001] HH:MM:SS | ...
                                        time_str = line.split('|')[0].split(']')[1].strip()
                                        h, m, s = map(int, time_str.split(':'))
                                        times.append(h * 3600 + m * 60 + s)
                                    except:
                                        pass
                                
                                if len(times) >= 2:
                                    time_diffs = [times[i] - times[i-1] for i in range(1, len(times))]
                                    avg_seconds = sum(time_diffs) / len(time_diffs)
                                    eta_seconds = int(avg_seconds * remaining)
                                    eta_hours = eta_seconds // 3600
                                    eta_mins = (eta_seconds % 3600) // 60
                                    prediction = f"ETA: {eta_hours}h {eta_mins}m ({avg_seconds:.1f}s/case avg)"
                            
                            # Write STATUS.txt
                            with open(sf, 'w') as f:
                                f.write("="*80 + '\n')
                                f.write("CFMS AUTOMATION LOG\n")
                                f.write("="*80 + '\n')
                                f.write(f"📊 Filled: {filled} | Remaining: {remaining} | {prediction}\n")
                                f.write("="*80 + '\n\n')
                                
                                # Write all existing entries
                                for entry in existing_entries:
                                    f.write(entry + '\n')
                                
                                # Add new entry
                                entry_num = len(existing_entries) + 1
                                new_entry = f"[{entry_num:03d}] {datetime.now().strftime('%H:%M:%S')} | FIR {CASE_FIR_NUMBER}/{CASE_FIR_YEAR} | {DECISION_TYPE} ✅"
                                f.write(new_entry + '\n')
                                
                                logging.info(f"✓ STATUS Entry #{entry_num} written")
                        
                        except Exception as e:
                            log_error("STATUS_UPDATE", f"Failed to update STATUS.txt: {e}", e)
                        
                        # Sync to Google Drive
                        sync_to_gdrive()
                        
                        case_success = True
                        consecutive_errors = 0
                        
                    elif status == "NEXT_CASE":
                        logging.info("✓ Framing error - moving on")
                        case_success = True
                        consecutive_errors = 0
                        
                    elif status == "RESTART":
                        logging.info("↻ Restarting same case")
                        continue
                        
                    elif status == "ERROR":
                        consecutive_errors += 1
                        logging.error(f"❌ Case error (consecutive: {consecutive_errors})")
                        
                        if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                            logging.error("❌ Too many consecutive errors - restarting browser")
                            if not browser_mgr.restart_browser():
                                log_error("CRITICAL", "Cannot restart browser")
                                return
                            consecutive_errors = 0
                        
                        if case_retry_count >= MAX_RETRIES_PER_CASE:
                            logging.error(f"❌ Max retries reached - skipping case")
                            mark_as_filled(case_line)
                            case_success = True
                    
                except TimeoutError:
                    log_error("WATCHDOG_TIMEOUT", f"Operation timed out after {WATCHDOG_TIMEOUT}s")
                    consecutive_errors += 1
                    
                    # Force browser restart on timeout
                    logging.info("⟳ Timeout - restarting browser")
                    if not browser_mgr.restart_browser():
                        log_error("CRITICAL", "Cannot restart browser after timeout")
                        return
                    
                except Exception as e:
                    log_error("PROCESSING_ERROR", f"Unexpected error: {e}", e)
                    consecutive_errors += 1
                    
                    if case_retry_count >= MAX_RETRIES_PER_CASE:
                        mark_as_filled(case_line)
                        case_success = True
            
            # Browser restart after each case
            logging.info("="*60)
            logging.info("🔄 RESTARTING BROWSER")
            logging.info("="*60)
            
            if not browser_mgr.restart_browser():
                log_error("CRITICAL", "Failed to restart browser")
                break
            
            logging.info(f"💾 RAM: {psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024:.1f} MB")
        
        if loop_count >= max_loops:
            logging.warning(f"⚠ Max loops ({max_loops}) reached")
        
        print("\n" + "="*60)
        print("✓✓✓ AUTOMATION COMPLETE!")
        print("="*60)
        print(f"\nTotal errors logged: {sum(error_counts.values())}")
        print(f"Browser restarts: {browser_mgr.restart_count}")
        
    except KeyboardInterrupt:
        logging.info("\n⚠ Interrupted by user")
    except Exception as e:
        log_error("CRITICAL", f"Fatal error: {e}", e)
    finally:
        try:
            browser_mgr.driver.quit()
        except:
            pass
        
        logging.info("\n✓ Cleanup complete")

if __name__ == "__main__":
    main()