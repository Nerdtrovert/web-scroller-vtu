"""
Web scraper module for VTU Result Automation
Handles Selenium operations to scrape student results from the VTU website
Includes automatic CAPTCHA solving with multi-engine OCR and retry mechanism.
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    NoAlertPresentException,
    UnexpectedAlertPresentException
)
import time
import os
import logging
from logger import setup_logger
from captcha_solver import CaptchaSolver
from config import (
    XPATH_USN_INPUT,
    XPATH_USN_INPUT_FALLBACK,
    XPATH_CAPTCHA_INPUT,
    XPATH_CAPTCHA_INPUT_FALLBACK,
    XPATH_CAPTCHA_IMAGE,
    XPATH_SUBMIT_BUTTON,
    XPATH_STUDENT_USN,
    XPATH_STUDENT_NAME,
    XPATH_SUBJECT_BASE,
    SUBJECT_NAME_INDEX,
    SUBJECT_IA_INDEX,
    SUBJECT_SEE_INDEX,
    SUBJECT_TOTAL_INDEX,
    SUBJECT_RESULT_INDEX,
    MAX_SUBJECTS,
    MAX_CAPTCHA_RETRIES,
    WAIT_AFTER_STARTUP,
    WAIT_BEFORE_CAPTCHA,
    WAIT_AFTER_INPUT,
    WAIT_TIMEOUT
)

logger = setup_logger(__name__)


class ResultScraper:
    """Handles web scraping operations for VTU result portal"""
    
    def __init__(self, driver_path, website_url):
        """
        Initialize the scraper with driver path and website URL
        
        Args:
            driver_path: Path to ChromeDriver executable
            website_url: URL of the VTU result portal
        """
        self.driver_path = driver_path
        self.website_url = website_url
        self.driver = None
        self.actions = None
        self.main_window = None
        self.wait = None
        
        # Page elements
        self.usn_box = None
        self.captcha_box = None
        self.captcha_img = None
        self.submit_btn = None
    
    def setup_driver(self):
        """Initialize and configure the Chrome WebDriver"""
        service = Service(self.driver_path) if self.driver_path else None
        opts = Options()
        opts.add_experimental_option("detach", True)
        
        if service:
            self.driver = webdriver.Chrome(service=service, options=opts)
        else:
            self.driver = webdriver.Chrome(options=opts)

        self.actions = ActionChains(self.driver)
        self.wait = WebDriverWait(self.driver, WAIT_TIMEOUT)
        self.driver.maximize_window()
        self.driver.get(self.website_url)
        
        self.main_window = self.driver.window_handles[0]
        logger.info(f"Driver setup complete. Opened {self.website_url}")
        time.sleep(WAIT_AFTER_STARTUP)
    
    def locate_page_elements(self):
        """
        Locate the main page elements (USN input, captcha input, captcha image, submit button)
        
        Returns:
            True if mandatory elements found, False otherwise
        """
        try:
            # Locate USN input with fallback
            try:
                self.usn_box = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_USN_INPUT)))
            except TimeoutException:
                self.usn_box = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_USN_INPUT_FALLBACK)))

            # Locate Captcha input with fallback
            try:
                self.captcha_box = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_CAPTCHA_INPUT)))
            except TimeoutException:
                self.captcha_box = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_CAPTCHA_INPUT_FALLBACK)))

            # Locate Submit button
            self.submit_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_SUBMIT_BUTTON)))

            # Locate Captcha image element for auto-solving
            try:
                self.captcha_img = self.driver.find_element(By.XPATH, XPATH_CAPTCHA_IMAGE)
                logger.info("Captcha image element located.")
            except NoSuchElementException:
                logger.warning("Captcha image element not found via default XPath. Auto-solve may need refresh.")
                self.captcha_img = None

            logger.info("Page elements located successfully.")
            return True
        except (NoSuchElementException, TimeoutException) as e:
            logger.error(f"Page elements not found: {e}")
            return False
            
    def check_alert(self):
        """
        Check if a JavaScript alert is present, dismiss it, and return its text.
        
        Returns:
            Alert text if present, None otherwise.
        """
        try:
            alert = self.driver.switch_to.alert
            text = alert.text
            alert.accept()
            logger.info(f"Browser alert handled: {text}")
            return text
        except NoAlertPresentException:
            return None
        except Exception as e:
            logger.debug(f"Error checking alert: {e}")
            return None

    def solve_current_captcha(self, save_debug_path="temp_captcha.png"):
        """
        Capture the current captcha element screenshot and automatically solve it.
        
        Args:
            save_debug_path: File path to save captcha screenshot for debugging.
            
        Returns:
            Solved 6-character captcha string, or None if failed.
        """
        try:
            if not self.captcha_img:
                try:
                    self.captcha_img = self.driver.find_element(By.XPATH, XPATH_CAPTCHA_IMAGE)
                except NoSuchElementException:
                    logger.error("Cannot solve captcha: image element not found.")
                    return None

            img_bytes = self.captcha_img.screenshot_as_png
            
            # Save to disk for user visibility and debugging
            if save_debug_path:
                try:
                    with open(save_debug_path, "wb") as f:
                        f.write(img_bytes)
                except Exception as e:
                    logger.debug(f"Failed to save temp captcha image: {e}")

            predicted = CaptchaSolver.solve(img_bytes)
            logger.info(f"Auto-solved captcha: '{predicted}'")
            return predicted
        except Exception as e:
            logger.error(f"Error capturing or solving captcha: {e}")
            return None
    
    def enter_usn_and_captcha(self, usn, captcha):
        """
        Enter USN and captcha values into the form
        
        Args:
            usn: Student USN to enter
            captcha: Captcha value to enter
        """
        self.usn_box.clear()
        self.usn_box.send_keys(usn)
        time.sleep(WAIT_AFTER_INPUT)
        
        self.captcha_box.clear()
        self.captcha_box.send_keys(captcha)
        time.sleep(WAIT_AFTER_INPUT)
        logger.debug(f"Entered USN: {usn} and Captcha: {captcha}")
    
    def submit_and_switch_to_result(self):
        """
        Submit the form and switch to the result window
        
        Returns:
            True if result window opened successfully, False otherwise
        """
        try:
            # Open result in new tab using CTRL+Click
            self.actions.key_down(Keys.LEFT_CONTROL).perform()
            self.submit_btn.click()
            self.actions.key_up(Keys.LEFT_CONTROL).perform()
        except UnexpectedAlertPresentException:
            alert_text = self.check_alert()
            logger.warning(f"Submission failed due to alert: {alert_text}")
            return False
        
        # Check for immediate alert
        alert_text = self.check_alert()
        if alert_text:
            logger.warning(f"Alert upon form submit: {alert_text}")
            return False

        # Wait for new window
        try:
            self.wait.until(EC.number_of_windows_to_be(2))
            result_window = self.driver.window_handles[1]
            self.driver.switch_to.window(result_window)
            logger.info("Switched to result window.")
            return True
        except TimeoutException:
            logger.warning("Result window did not open within timeout.")
            # Check if alert popped up during wait
            self.check_alert()
            self.usn_box.clear()
            return False

    def submit_with_auto_captcha(self, usn, max_retries=MAX_CAPTCHA_RETRIES):
        """
        High-level automation: types USN, automatically captures & solves captcha,
        submits, and handles retries if VTU alerts an invalid captcha.
        
        Args:
            usn: Student USN
            max_retries: Maximum attempts to solve and submit captcha
            
        Returns:
            True if successfully opened result window, False otherwise
        """
        for attempt in range(1, max_retries + 1):
            logger.info(f"Processing {usn} - Captcha attempt {attempt}/{max_retries}")
            
            # Ensure elements are fresh
            self.locate_page_elements()
            
            captcha = self.solve_current_captcha()
            if not captcha:
                logger.warning(f"Could not solve captcha on attempt {attempt}.")
                time.sleep(1)
                continue
                
            self.enter_usn_and_captcha(usn, captcha)
            
            # Attempt submission
            opened = self.submit_and_switch_to_result()
            if opened:
                return True
                
            # If not opened, check if there was an alert and pause briefly
            alert = self.check_alert()
            if alert:
                logger.info(f"VTU portal message: {alert}")
                if "invalid" in alert.lower() and "usn" in alert.lower():
                    logger.warning(f"USN {usn} is invalid or result not found.")
                    return False
            
            time.sleep(1)

        logger.error(f"Exhausted {max_retries} captcha attempts for {usn}.")
        return False
    
    def scrape_student_info(self):
        """
        Scrape student USN and name from the result page
        
        Returns:
            Tuple of (usn, name) or (None, None) on error
        """
        try:
            usn_element = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_STUDENT_USN)))
            name_element = self.driver.find_element(By.XPATH, XPATH_STUDENT_NAME)
            
            usn = usn_element.text
            name = name_element.text
            logger.info(f"Scraped info for USN: {usn}, Name: {name}")
            return usn, name
        except (NoSuchElementException, TimeoutException) as e:
            logger.error(f"Error extracting student info: {e}")
            return None, None
            
    def scrape_semester(self):
        """Scrape semester information from the result page"""
        try:
            body_text = self.driver.find_element(By.ID, "dataPrint").text
            for line in body_text.splitlines():
                if "Semester" in line:
                    return line.strip()
        except Exception:
            pass
        return ""
    
    def scrape_subjects(self):
        """
        Scrape all subject details from the result page
        
        Returns:
            List of dictionaries containing subject data (name, ia, see, total, res)
        """
        subjects = []
        
        for k in range(1, MAX_SUBJECTS + 1):
            try:
                # Build XPath for current subject row
                base_xpath = XPATH_SUBJECT_BASE.format(k + 1)
                
                # Extract subject information
                sub_name = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_NAME_INDEX}]"
                ).text
                ia = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_IA_INDEX}]"
                ).text
                see = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_SEE_INDEX}]"
                ).text
                total = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_TOTAL_INDEX}]"
                ).text
                res = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_RESULT_INDEX}]"
                ).text
                
                subjects.append({
                    "name": sub_name,
                    "ia": ia,
                    "see": see,
                    "total": total,
                    "res": res
                })
            except NoSuchElementException:
                # No more subjects found
                break
        
        logger.info(f"Scraped {len(subjects)} subjects.")
        return subjects
    
    def close_result_and_return_to_main(self):
        """Close the result window and switch back to main window"""
        try:
            self.driver.close()
            self.driver.switch_to.window(self.main_window)
            self.usn_box.clear()
            # Wait for main window to be active/interactable again
            try:
                self.wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_USN_INPUT)))
            except TimeoutException:
                self.wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_USN_INPUT_FALLBACK)))
            logger.debug("Closed result window and returned to main.")
        except Exception as e:
             logger.error(f"Error returning to main window: {e}")
    
    def cleanup(self):
        """Clean up driver resources"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("Driver exited successfully.")
            except Exception as e:
                logger.error(f"Error closing driver: {e}")
