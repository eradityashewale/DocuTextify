import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from configparser import ConfigParser
import logging
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
)
from selenium.webdriver.common.action_chains import ActionChains

# --- 1. Configure Logging ---
logging.basicConfig(
    level=logging.INFO,
    filename="selenium_log.log",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# --- 2. Load Vocabulary Data ---
file_path = "idioms_definitions.csv"
try:
    vocab_df = pd.read_csv(file_path)
    logging.info(f"Successfully loaded CSV file: {file_path}")
except FileNotFoundError:
    logging.error(f"CSV file not found: {file_path}")
    raise
except Exception as e:
    logging.error(f"Error loading CSV file: {e}")
    raise

# --- 3. Load Configuration ---
config = ConfigParser()
config_file = "config.ini"
try:
    config.read(config_file)
    email = config["TARUN_GROVER"]["email"]
    password = config["TARUN_GROVER"]["password"]
    logging.info(f"Successfully loaded configuration from {config_file}")
except KeyError as e:
    logging.error(f"Missing key in configuration file: {e}")
    raise
except Exception as e:
    logging.error(f"Error reading configuration file: {e}")
    raise

# --- 4. Select Idioms to Add ---
# Define the numbers you want to select (ensure these exist in your CSV)
selected_numbers = [405, 415, 425, 435, 445, 455, 465, 475, 485, 495]

# Filter the DataFrame for the selected numbers
selected_vocab = vocab_df[vocab_df["number"].isin(selected_numbers)]

# Verify that exactly 10 idioms are selected
if len(selected_vocab) != 10:
    logging.error(
        f"Expected 10 idioms, but found {len(selected_vocab)}. Please check the selected numbers."
    )
    raise ValueError("Number of selected idioms does not equal 10.")
else:
    logging.info("Successfully selected 10 idioms based on the provided numbers.")

# --- 5. Set Up Selenium WebDriver ---
service = Service(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
# Optional: Run in headless mode
# chrome_options.add_argument("--headless")
driver = webdriver.Chrome(service=service, options=chrome_options)

# --- 6. Define Helper Functions ---

def retry_on_exception(func, retries=3, delay=2):
    """
    Retries a function if an exception occurs.

    Args:
        func (callable): The function to execute.
        retries (int): Number of retry attempts.
        delay (int): Delay in seconds between retries.

    Returns:
        The return value of the function if successful.

    Raises:
        The exception from the last failed attempt.
    """
    for attempt in range(retries):
        try:
            return func()
        except Exception as e:
            logging.warning(f"Attempt {attempt + 1} failed with exception: {e}")
            time.sleep(delay)
    logging.error(f"All {retries} attempts failed for function {func.__name__}.")
    raise

def navigate_to_add_idioms(driver, wait):
    """Navigates to the 'Add Idioms' page."""
    try:
        # Click on 'Vocabs' link
        vocabs_link_xpath = "//nav//a[.//h6[text()='Vocabs']]"
        vocabs_link = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, vocabs_link_xpath))
            )
        )
        vocabs_link.click()
        logging.info("Clicked on 'Vocabs' link.")
    except TimeoutException:
        logging.error("'Vocabs' link not found or not clickable.")
        driver.save_screenshot("vocabs_link_not_clickable.png")
        raise

    try:
        # Click on 'Idioms' section
        idioms_section_xpath = "//header//a[.//div[text()='Idioms']]"
        idioms_section = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, idioms_section_xpath))
            )
        )
        idioms_section.click()
        logging.info("Navigated to 'Idioms' section.")
    except TimeoutException:
        logging.error("'Idioms' section not found or not clickable.")
        driver.save_screenshot("idioms_section_not_clickable.png")
        raise

    try:
        # Click on 'Add Idioms' button
        add_idioms_button_xpath = "//a[contains(@href, 'add-idiom') or contains(text(), 'Add Idioms')]"
        add_idioms_button = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, add_idioms_button_xpath))
            )
        )
        add_idioms_button.click()
        logging.info("Clicked on 'Add Idioms' button.")
    except TimeoutException:
        logging.error("'Add Idioms' button not found or not clickable.")
        driver.save_screenshot("add_idioms_button_not_clickable.png")
        raise

def add_idiom(driver, wait, idiom):
    """Adds a single idiom to the platform."""
    idiom_number = idiom['number']
    idiom_phrase = idiom['idiom']
    idiom_definition = idiom['definition']
    idiom_example = idiom['example']
    hard_coded_date = "25/12/2024"  # Correct 'YYYY-MM-DD' format
    logging.info(f"Adding idiom number {idiom_number}: {idiom_phrase}")

    try:
        # Phrase Name Input
        phrase_input_xpath = "//input[@name='phrase']"
        phrase_input = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, phrase_input_xpath))
            )
        )
        phrase_input.clear()
        phrase_input.send_keys(idiom_phrase)
        logging.info(f"Entered idiom phrase: {idiom_phrase}")
    except TimeoutException:
        logging.error(f"Phrase input field not found for idiom {idiom_number}.")
        driver.save_screenshot(f"idiom_{idiom_number}_phrase_input_not_found.png")
        return False

    try:
        # Definition Input
        definition_input_xpath = "//input[@name='definition']"
        try:
            definition_input = retry_on_exception(
                lambda: wait.until(
                    EC.element_to_be_clickable((By.XPATH, definition_input_xpath))
                )
            )
            definition_input.clear()
            definition_input.send_keys(idiom_definition)
            logging.info(f"Entered definition: {idiom_definition}")
        except TimeoutException:
            # If input field not found, try contenteditable div
            definition_input_xpath = "//div[contains(@class, 'definition') and @contenteditable='true']"
            definition_input = retry_on_exception(
                lambda: wait.until(
                    EC.element_to_be_clickable((By.XPATH, definition_input_xpath))
                )
            )
            definition_input.click()
            definition_input.send_keys(idiom_definition)
            logging.info(f"Entered definition via contenteditable div: {idiom_definition}")
    except (TimeoutException, ElementNotInteractableException) as e:
        logging.error(
            f"Definition input field not found or not interactable for idiom {idiom_number}. Exception: {e}"
        )
        driver.save_screenshot(f"idiom_{idiom_number}_definition_input_error.png")
        return False

    try:
        # Example Input
        example_input_xpath = "//input[@name='example']"
        example_input = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, example_input_xpath))
            )
        )
        example_input.clear()
        example_input.send_keys(idiom_example)
        logging.info(f"Entered example: {idiom_example}")
    except TimeoutException:
        logging.error(f"Example input field not found for idiom {idiom_number}.")
        driver.save_screenshot(f"idiom_{idiom_number}_example_input_not_found.png")
        return False

    try:
        # Date Input
        date_input_xpath = "//input[@type='date' or contains(@placeholder, 'Date')]"
        date_input = retry_on_exception(
            lambda: wait.until(EC.element_to_be_clickable((By.XPATH, date_input_xpath)))
        )
        date_input.clear()
        date_input.send_keys(hard_coded_date)
        logging.info(f"Entered hard-coded date: {hard_coded_date}")
    except TimeoutException:
        logging.error("Date input field not found.")
        driver.save_screenshot("date_input_not_found.png")
        return False

    try:
        # Submit the Form
        submit_button_xpath = "//button[contains(text(), 'Create') or contains(text(), 'Submit')]"
        submit_button = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, submit_button_xpath))
            )
        )
        submit_button.click()
        logging.info("Clicked 'Submit' button.")
    except TimeoutException:
        logging.error(f"Submit button not found or not clickable for idiom {idiom_number}.")
        driver.save_screenshot(f"idiom_{idiom_number}_submit_button_not_clickable.png")
        return False

    try:
        # Click on 'Vocabs' link
        vocabs_link_xpath = "//nav//a[.//h6[text()='Vocabs']]"
        vocabs_link = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, vocabs_link_xpath))
            )
        )
        vocabs_link.click()
        logging.info("Clicked on 'Vocabs' link.")
    except TimeoutException:
        logging.error("'Vocabs' link not found or not clickable.")
        driver.save_screenshot("vocabs_link_not_clickable.png")
        raise

    try:
        # Click on 'Idioms' section
        idioms_section_xpath = "//header//a[.//div[text()='Idioms']]"
        idioms_section = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, idioms_section_xpath))
            )
        )
        idioms_section.click()
        logging.info("Navigated to 'Idioms' section.")
    except TimeoutException:
        logging.error("'Idioms' section not found or not clickable.")
        driver.save_screenshot("idioms_section_not_clickable.png")
        raise

    try:
        # Click on 'Add Idioms' button
        add_idioms_button_xpath = "//a[contains(@href, 'add-idiom') or contains(text(), 'Add Idioms')]"
        add_idioms_button = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, add_idioms_button_xpath))
            )
        )
        add_idioms_button.click()
        logging.info("Clicked on 'Add Idioms' button.")
    except TimeoutException:
        logging.error("'Add Idioms' button not found or not clickable.")
        driver.save_screenshot("add_idioms_button_not_clickable.png")
        raise
    # --- Ensure the Form is Ready for Next Idiom ---
    try:
        # Option 1: Wait until the form fields are cleared
        wait.until(lambda d: d.find_element(By.XPATH, "//input[@name='phrase']").get_attribute('value') == "")
        wait.until(lambda d: d.find_element(By.XPATH, "//input[@name='definition']").get_attribute('value') == "")
        wait.until(lambda d: d.find_element(By.XPATH, "//input[@name='example']").get_attribute('value') == "")
        wait.until(lambda d: d.find_element(By.XPATH, "//input[@type='date']").get_attribute('value') == hard_coded_date)
        logging.info(f"Form is ready for the next idiom {idiom_number}.")
    except TimeoutException:
        logging.warning(f"Form fields not cleared or not ready for idiom {idiom_number}. Attempting to refresh the form.")
        try:
            # Option 2: Refresh the form or page
            driver.refresh()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='phrase']")))
            logging.info("Form refreshed successfully.")
        except Exception as e:
            logging.error(f"Failed to refresh the form for idiom {idiom_number}. Exception: {e}")
            driver.save_screenshot(f"idiom_{idiom_number}_form_refresh_failed.png")
            return False

    return True

# --- 7. Main Execution Block ---
try:
    # Open the login page
    login_url = "https://admin.tarungroverenglish.com/app/"
    driver.get(login_url)
    logging.info(f"Navigated to login page: {login_url}")

    # Initialize WebDriverWait
    wait = WebDriverWait(driver, 15)

    # --- Login Process ---
    logging.info("Starting login process.")

    # Wait for email field and enter email
    try:
        email_input = retry_on_exception(
            lambda: wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='email']"))
            )
        )
        email_input.clear()
        email_input.send_keys(email)
        logging.info("Entered email.")
    except TimeoutException:
        logging.error("Email input field not found.")
        driver.save_screenshot("email_input_not_found.png")
        raise

    # Wait for password field and enter password
    try:
        password_input = retry_on_exception(
            lambda: wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
            )
        )
        password_input.clear()
        password_input.send_keys(password)
        logging.info("Entered password.")
    except TimeoutException:
        logging.error("Password input field not found.")
        driver.save_screenshot("password_input_not_found.png")
        raise

    # Click the login button
    try:
        login_button = retry_on_exception(
            lambda: wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Login')]"))
            )
        )
        login_button.click()
        logging.info("Clicked login button.")
    except TimeoutException:
        logging.error("Login button not found or not clickable.")
        driver.save_screenshot("login_button_not_clickable.png")
        raise

    # Wait for the dashboard page to load
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//h6[contains(text(), 'Dashboard')]")
            )
        )
        logging.info("Logged in successfully. Dashboard loaded.")
    except TimeoutException:
        logging.error("Dashboard did not load successfully.")
        driver.save_screenshot("dashboard_not_loaded.png")
        raise

    # --- Navigate to "Add Idioms" Page ---
    navigate_to_add_idioms(driver, wait)

    # --- Iterate Over Selected Idioms and Add Them ---
    for index, row in selected_vocab.iterrows():
        idiom = {
            "number": row["number"],
            "idiom": row["idiom"],
            "definition": row["definition"],
            "example": row["example"],
        }
        success = add_idiom(driver, wait, idiom)
        if not success:
            logging.warning(f"Skipping idiom number {idiom['number']} due to previous errors.")
            continue  # Proceed to the next idiom

        # Optional: Short wait before adding the next idiom
        time.sleep(1)

    logging.info("All selected idioms have been added successfully.")
    print("Idioms added successfully.")

except Exception as e:
    logging.error("An error occurred during the Selenium script execution.", exc_info=True)
    print(f"An error occurred: {e}")

finally:
    # Ensure the browser is closed
    driver.quit()
    logging.info("Browser closed.")
