from datetime import datetime, timedelta
import os
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException)
from webdriver_manager.chrome import ChromeDriverManager
from configparser import ConfigParser
import logging
import time
import sys
import traceback

# Setup Logging
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Create handlers
file_handler = logging.FileHandler('automation.log')
file_handler.setLevel(logging.INFO)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)

# Create formatters and add to handlers
formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

def load_config(config_path='config.ini'):
    """Load configuration from the config.ini file."""
    config = ConfigParser()
    config.read(config_path)
    try:
        email = config["TARUN_GROVER"]["email"]
        password = config["TARUN_GROVER"]["password"]
        logger.info("Configuration loaded successfully.")
        return email, password
    except KeyError as e:
        logger.error(f"Missing configuration for {e}. Please check config.ini.")
        raise

def setup_webdriver():
    """Set up the Chrome WebDriver with desired options."""
    logger.info("Setting up the WebDriver.")
    try:
        service = Service(ChromeDriverManager().install())
    except Exception as e:
        logger.error(f"Failed to install ChromeDriver: {e}")
        raise

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    # Optional: Run Chrome in headless mode
    # chrome_options.add_argument("--headless")
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()
        logger.info("WebDriver setup complete.")
        return driver
    except Exception as e:
        logger.error(f"Failed to initialize WebDriver: {e}")
        raise

def save_cookies(driver, file_path):
    """Save cookies to a JSON file."""
    try:
        with open(file_path, "w") as file:
            json.dump(driver.get_cookies(), file)
        logger.info(f"Cookies saved to {file_path}.")
    except Exception as e:
        logger.error(f"Failed to save cookies: {e}")
        raise

def load_cookies(driver, file_path, domain=".tarungroverenglish.com"):
    """Load cookies from a JSON file."""
    if os.path.exists(file_path):
        try:
            with open(file_path, "r") as file:
                cookies = json.load(file)
                for cookie in cookies:
                    cookie["domain"] = domain
                    try:
                        driver.add_cookie(cookie)
                    except Exception as e:
                        logger.warning(f"Failed to add cookie: {cookie} due to {e}")
            logger.info(f"Cookies loaded from {file_path}.")
        except Exception as e:
            logger.error(f"Error loading cookies from {file_path}: {e}")
            raise
    else:
        logger.info(f"No cookies file found at {file_path}.")

def is_logged_in(driver):
    """Check if the user is already logged in by verifying the presence of the Dashboard."""
    try:
        driver.find_element(By.XPATH, "//h6[contains(text(), 'Dashboard')]")
        return True
    except NoSuchElementException:
        return False

def login(driver, wait, email, password, login_url, cookies_file):
    """Handle the login process, using cookies if available."""
    logger.info("Starting login process.")
    driver.get(login_url)
    logger.info(f"Navigated to {login_url}.")

    # Load cookies and refresh
    load_cookies(driver, cookies_file)
    driver.refresh()
    time.sleep(2)  # Wait for potential redirection after loading cookies

    # Check if already logged in
    if is_logged_in(driver):
        logger.info("Successfully logged in using cookies.")
        return
    else:
        logger.info("Cookies did not work, proceeding with login.")

    try:
        email_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='email']")))
        email_input.clear()
        email_input.send_keys(email)
        logger.info("Entered email.")

        password_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
        password_input.clear()
        password_input.send_keys(password)
        logger.info("Entered password.")

        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Login')]")))
        login_button.click()
        logger.info("Clicked the Login button.")

        # Wait until Dashboard is present
        wait.until(EC.presence_of_element_located((By.XPATH, "//h6[contains(text(), 'Dashboard')]")))
        logger.info("Login successful, saving cookies.")
        save_cookies(driver, cookies_file)

    except TimeoutException as e:
        driver.save_screenshot("login_timeout.png")
        logger.error("Login elements not found. Screenshot saved as login_timeout.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise
    except Exception as e:
        driver.save_screenshot("login_exception.png")
        logger.error("An unexpected error occurred during login. Screenshot saved as login_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def navigate_to_section(driver, wait, section_name):
    """Navigate to a specified section by its name."""
    logger.info(f"Navigating to '{section_name}' section.")
    try:
        # Adjusted XPath to account for nested <h6> within <a>
        section = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//a[.//h6[contains(text(), '{section_name}')]]")))
        
        # Scroll into view
        driver.execute_script("arguments[0].scrollIntoView();", section)
        
        # Click the section
        section.click()
        logger.info(f"Navigated to '{section_name}' section.")
    except TimeoutException:
        logger.warning(f"Primary locator failed for '{section_name}'. Trying alternative locators.")
        
        try:
            # Alternative Locator 1: Using href attribute (Example for 'Quizzes')
            href_fragment = '/quizzes' if section_name.lower() == 'quizzes' else '/other_section'
            section = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, f"a[href*='{href_fragment}']")))
            driver.execute_script("arguments[0].scrollIntoView();", section)
            section.click()
            logger.info(f"Navigated to '{section_name}' section using alternative locator 1.")
        except TimeoutException:
            logger.warning(f"Alternative locator 1 failed for '{section_name}'. Trying alternative locator 2.")
            
            try:
                # Alternative Locator 2: Using aria-label attribute if available
                aria_label = 'Quizzes' if section_name.lower() == 'quizzes' else 'OtherSection'
                section = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//a[@aria-label='{aria_label}']")))
                driver.execute_script("arguments[0].scrollIntoView();", section)
                section.click()
                logger.info(f"Navigated to '{section_name}' section using alternative locator 2.")
            except TimeoutException as e:
                # Save page source and screenshot for debugging
                with open("page_source.html", "w", encoding='utf-8') as f:
                    f.write(driver.page_source)
                
                driver.save_screenshot(f"{section_name}_section_timeout.png")
                logger.error(f"Failed to locate the '{section_name}' section with all locator strategies. Screenshot and page source saved for debugging.")
                logger.error(f"Exception: {e}")
                logger.error(f"Stacktrace: {traceback.format_exc()}")
                driver.quit()
                raise
            except Exception as e:
                driver.save_screenshot(f"{section_name}_section_exception.png")
                logger.error(f"An unexpected error occurred while navigating to '{section_name}'. Screenshot saved as {section_name}_section_exception.png.")
                logger.error(f"Exception: {e}")
                logger.error(f"Stacktrace: {traceback.format_exc()}")
                driver.quit()
                raise
    except Exception as e:
        driver.save_screenshot(f"{section_name}_section_exception.png")
        logger.error(f"An unexpected error occurred while navigating to '{section_name}'. Screenshot saved as {section_name}_section_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def click_add_quiz(driver, wait):
    """Click the 'Add Quiz' button using a reliable locator."""
    logger.info("Clicking the 'Add Quiz' button.")
    try:
        add_quiz_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@href, '/app/quizzes/add?preselectedType=daily_free')]")))
        add_quiz_button.click()
        logger.info("Clicked the 'Add Quiz' button.")
    except TimeoutException:
        driver.save_screenshot("add_quiz_timeout.png")
        logger.error("Add Quiz button not found. Screenshot saved as add_quiz_timeout.png.")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise
    except Exception as e:
        driver.save_screenshot("add_quiz_exception.png")
        logger.error("An unexpected error occurred while clicking 'Add Quiz'. Screenshot saved as add_quiz_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def set_quiz_details(driver, wait, quiz_date, points):
    """Set the quiz date and points."""
    logger.info("Setting quiz details.")
    try:
        date_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@name='forDate']")))
        date_input.clear()
        date_input.send_keys(quiz_date)
        logger.info(f"Set quiz date to {quiz_date}.")

        points_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@name='points']")))
        points_input.clear()
        points_input.send_keys(str(points))
        logger.info(f"Set quiz points to {points}.")

        create_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(), 'Create')]")))
        create_button.click()
        logger.info("Clicked the 'Create' button to create the quiz.")

        # Log current page title and URL
        logger.info(f"Current Page Title: {driver.title}")
        logger.info(f"Current URL: {driver.current_url}")

        # Wait until the quiz is created and redirected appropriately
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/header/div[2]/div[2]/div/a[2]/div")))
        logger.info("Quiz created successfully.")

    except TimeoutException as e:
        driver.save_screenshot("set_quiz_details_timeout.png")
        logger.error("Failed to set quiz details. Screenshot saved as set_quiz_details_timeout.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise
    except Exception as e:
        driver.save_screenshot("set_quiz_details_exception.png")
        logger.error("An unexpected error occurred while setting quiz details. Screenshot saved as set_quiz_details_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def add_first(driver, wait):
    """Add the first question by navigating through the UI elements."""
    logger.info("Adding the first question.")
    try:
        # Click on the "Questions" section
        questions_section = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class, 'MuiStack-root') and contains(text(), 'Questions')]")))
        questions_section.click()
        logger.info("Clicked on the 'Questions' section.")

        # Click on the "Plus" button to add a new question
        plus_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div")))
        plus_button.click()
        # The above XPath targets the SVG path of the plus icon.
        # It's generally better to click the parent button if possible.
        # Alternatively, you can use a more reliable locator if available.
        # plus_button_parent = plus_button.find_element(By.XPATH, "./ancestor::button")
        # plus_button_parent.click()
        # logger.info("Clicked on the 'Plus' button to add a new question.")

        # Click on "One" to select the type or first option
        one_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class, 'MuiChip-root') and .//div[text()='1']]")))
        one_button.click()
        logger.info("Clicked on the 'One' button to select the first option.")

    except TimeoutException as e:
        driver.save_screenshot("add_first_timeout.png")
        logger.error("Timeout while adding the first question. Screenshot saved as add_first_timeout.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise
    except Exception as e:
        driver.save_screenshot("add_first_exception.png")
        logger.error("An unexpected error occurred while adding the first question. Screenshot saved as add_first_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def read_questions(file_path):
    """Read questions and options from a CSV file."""
    logger.info(f"Reading questions from {file_path}.")
    try:
        data = pd.read_csv(file_path)
        
        # Strip whitespace from column names to avoid mismatches
        data.columns = data.columns.str.strip()
        
        required_columns = ['Question', 'Option_A', 'Option_B', 'Option_C', 'Option_D']
        missing_columns = [col for col in required_columns if col not in data.columns]
        
        if missing_columns:
            logger.error(f"Missing columns in {file_path}: {', '.join(missing_columns)}.")
            raise KeyError(f"Missing columns: {', '.join(missing_columns)}.")
        
        # Drop rows with any missing values in required columns
        data = data.dropna(subset=required_columns)
        
        # Convert to list of dictionaries
        questions = data.to_dict(orient='records')
        
        logger.info(f"Loaded {len(questions)} questions from {file_path}.")
        return questions
    except FileNotFoundError:
        logger.error(f"Quiz data file {file_path} not found.")
        raise
    except Exception as e:
        logger.error(f"Error reading quiz data file: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        raise


def click_add_question_button(driver, wait):
    """Click the 'Add Question' button using a reliable locator."""
    logger.info("Clicking the 'Add Question' button.")
    try:
        # Updated locator: Adjust based on actual HTML structure
        plus_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@aria-label='Add Question']")))
        plus_button.click()
        logger.info("Clicked the 'Add Question' button.")
    except TimeoutException:
        logger.warning("Primary locator for 'Add Question' button failed. Trying alternative locators.")
        try:
            # Example alternative: using a CSS selector based on class
            plus_button = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.add-question-button")))
            plus_button.click()
            logger.info("Clicked the 'Add Question' button using alternative locator.")
        except TimeoutException:
            driver.save_screenshot("add_question_button_timeout.png")
            logger.error("Failed to locate the 'Add Question' button. Screenshot saved as add_question_button_timeout.png.")
            logger.error(f"Stacktrace: {traceback.format_exc()}")
            driver.quit()
            raise
        except Exception as e:
            driver.save_screenshot("add_question_button_exception.png")
            logger.error("An unexpected error occurred while clicking 'Add Question'. Screenshot saved as add_question_button_exception.png.")
            logger.error(f"Exception: {e}")
            logger.error(f"Stacktrace: {traceback.format_exc()}")
            driver.quit()
            raise
    except Exception as e:
        driver.save_screenshot("add_question_button_exception.png")
        logger.error("An unexpected error occurred while clicking 'Add Question'. Screenshot saved as add_question_button_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def add_question(driver, wait, question_text, question_number, option_A_text, option_B_text, option_C_text, option_D_text):
    """Add a single question to the quiz."""
    logger.info(f"Adding question {question_number}.")
    try:
        # Click the 'Add Question' button
        # click_add_question_button(driver, wait)

        # Locate the textarea for question description
        question_textarea = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//textarea[@name='description']")))
        question_textarea.clear()
        question_textarea.send_keys(question_text)
        logger.info(f"Entered text for question {question_number}.")

        Add_choice_A = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div/button")))
        Add_choice_A.click()
        logger.info(f"Add choice for A")

        Add_choice_B = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div/button")))
        Add_choice_B.click()
        logger.info(f"Add choice for B")

        Add_choice_C = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div/button")))
        Add_choice_C.click()
        logger.info(f"Add choice for C")

        Add_choice_D = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div/button")))
        Add_choice_D.click()
        logger.info(f"Add choice for D")

        # Locate the textarea for question description
        option_A_textarea = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[2]/div/textarea[1]")))
        option_A_textarea.clear()
        option_A_textarea.send_keys(option_A_text)
        logger.info(f"option A text added for question {question_number}.")

        # Locate the textarea for question description
        option_B_textarea = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[3]/div/div/div/div[1]/div[2]/div/textarea[1]")))
        option_B_textarea.clear()
        option_B_textarea.send_keys(option_B_text)
        logger.info(f"option B text added for question {question_number}.")

        # Locate the textarea for question description
        option_C_textarea = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[3]/div/div/div/div[1]/div[2]/div/textarea[1]")))
        option_C_textarea.clear()
        option_C_textarea.send_keys(option_C_text)
        logger.info(f"option C text added for question {question_number}.")

        # Locate the textarea for question description
        option_D_textarea = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[2]/div/textarea[1]")))
        option_D_textarea.clear()
        option_D_textarea.send_keys(option_D_text)
        logger.info(f"option D text added for question {question_number}.")

        # Click the 'Save' button to add the question
        save_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[2]/button")))
        save_button.click()
        logger.info(f"Description for on ")

        # Optional: Wait for confirmation or the question to appear in the list
        wait.until(EC.presence_of_element_located(
            (By.XPATH, f"//div[contains(text(), '{question_text[:20]}')]")))  # Adjust as needed
        logger.info(f"Question {question_number} is now listed in the quiz.")

    except TimeoutException as e:
        driver.save_screenshot(f"add_question_{question_number}_timeout.png")
        logger.error(f"Failed to add question {question_number}. Screenshot saved as add_question_{question_number}_timeout.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise
    except Exception as e:
        driver.save_screenshot(f"add_question_{question_number}_exception.png")
        logger.error(f"An unexpected error occurred while adding question {question_number}. Screenshot saved as add_question_{question_number}_exception.png.")
        logger.error(f"Exception: {e}")
        logger.error(f"Stacktrace: {traceback.format_exc()}")
        driver.quit()
        raise

def add_all_questions(driver, wait, questions):
    """Iterate through all questions and add them to the quiz."""
    logger.info("Starting to add all questions.")
    for i, question in enumerate(questions, start=1):
        required_keys = ['Question', 'Option_A', 'Option_B', 'Option_C', 'Option_D']
        present_keys = question.keys()
        missing_keys = [key for key in required_keys if key not in present_keys]
        
        if missing_keys:
            logger.error(f"Question {i} is missing required fields: {', '.join(missing_keys)}.")
            continue  # Skip this question or handle as needed
        
        logger.info(f"Adding question {i}: {question['Question']}")
        add_question(
            driver,
            wait,
            question_text=question['Question'],
            question_number=i,
            option_A_text=question['Option_A'],
            option_B_text=question['Option_B'],
            option_C_text=question['Option_C'],
            option_D_text=question['Option_D']
        )
    logger.info("All questions added successfully.")


def main():
    logger.info("Script started.")
    # Configuration
    config_path = 'config.ini'
    cookies_file = "cookies.json"
    login_url = "https://admin.tarungroverenglish.com/app/"
    quiz_data_file = "quiz_data.csv"
    points = 10
    # Dynamic Quiz Date: Uncomment the following line to set the quiz date to tomorrow
    # quiz_date = (datetime.now() + timedelta(days=1)).strftime('%d-%m-%Y')
    quiz_date = '02-01-2025'  # Use dynamic date or set manually as needed
    
    # Load credentials
    try:
        email, password = load_config(config_path)
    except Exception as e:
        logger.error("Failed to load configuration. Exiting script.")
        sys.exit(1)

    # Set up WebDriver
    try:
        driver = setup_webdriver()
    except Exception as e:
        logger.error("Failed to set up WebDriver. Exiting script.")
        sys.exit(1)
    
    wait = WebDriverWait(driver, 60)  # Increased timeout to 60 seconds

    try:
        # Perform login
        login(driver, wait, email, password, login_url, cookies_file)

        # Navigate to "Quizzes" section
        navigate_to_section(driver, wait, "Quizzes")

        # Click the "Add Quiz" button
        click_add_quiz(driver, wait)

        # Set quiz details
        set_quiz_details(driver, wait, quiz_date, points)

        # Add the first question setup (if needed)
        add_first(driver, wait)

        # Read questions from CSV
        questions = read_questions(quiz_data_file)

        # Navigate to "Questions" section (if not already there)
        # navigate_to_section(driver, wait, "Questions")

        # Add all questions
        add_all_questions(driver, wait, questions)

    except Exception as e:
        logger.error(f"An error occurred during automation: {e}")
    finally:
        # Optionally, close the browser after completion
        try:
            driver.quit()
            logger.info("Browser closed.")
        except Exception as e:
            logger.warning(f"Failed to close the browser gracefully: {e}")
        logger.info("Script finished.")

if __name__ == "__main__":
    main()
