from datetime import datetime, timedelta
import time
import os
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from configparser import ConfigParser

# Load configuration
config = ConfigParser()
config.read('config.ini')
email = config["TARUN_GROVER"]["email"]
password = config["TARUN_GROVER"]["password"]

# File to store cookies
cookies_file = "cookies.json"

# Set up the Chrome WebDriver
service = Service(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 30)

# Login URL
login_url = "https://admin.tarungroverenglish.com/app/"

# Function to save cookies
def save_cookies(driver, file_path):
    with open(file_path, "w") as file:
        json.dump(driver.get_cookies(), file)

# Function to load cookies
def load_cookies(driver, file_path):
    if os.path.exists(file_path):
        with open(file_path, "r") as file:
            cookies = json.load(file)
            for cookie in cookies:
                cookie["domain"] = ".tarungroverenglish.com"
                driver.add_cookie(cookie)

# Function to check if logged in
def is_logged_in(driver):
    try:
        driver.find_element(By.XPATH, "//h6[contains(text(), 'Dashboard')]")
        return True
    except:
        return False

# Open the login page
driver.get(login_url)
time.sleep(3)

# Load cookies and refresh
load_cookies(driver, cookies_file)
driver.refresh()
time.sleep(3)

# Check if already logged in
if not is_logged_in(driver):
    print("Cookies did not work, proceeding with login.")
    
    email_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='email']")))
    email_input.clear()
    email_input.send_keys(email)

    password_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
    password_input.clear()
    password_input.send_keys(password)

    login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]")
    login_button.click()

    wait.until(EC.presence_of_element_located((By.XPATH, "//h6[contains(text(), 'Dashboard')]") ))
    print("Login successful, saving cookies.")
    save_cookies(driver, cookies_file)
else:
    print("Successfully logged in using cookies.")

# Navigate to "Quizzes" section
quizzes_section = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//h6[contains(text(), 'Quizzes')]")))
quizzes_section.click()
time.sleep(3)

# Wait for the "Add Quiz" button and click it
add_quiz_button = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//a[contains(@href, '/app/quizzes/add?preselectedType=daily_free')]")))
add_quiz_button.click()
time.sleep(3)

# Set the quiz date
tomorrow_date = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
current_date = '08-12-2024'

date_input = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//input[@name='forDate']")))
date_input.send_keys(current_date)
print(f"Tomorrow's date {tomorrow_date} selected for the quiz.")

# Add points to the quiz
points_input = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//input[@name='points']")))
points_input.send_keys("10")
time.sleep(1)

create_button = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//button[contains(text(), 'Create')]")))
create_button.click()
time.sleep(3)
print("Quiz created successfully.")

# Function to read questions from Excel file
def read_questions(file_path):
    data = pd.read_excel(file_path)
    questions = data['Question'].tolist()
    return questions

# File path for quiz data
quiz_data_file = "quiz_data.xlsx"

# Read questions
questions = read_questions(quiz_data_file)
print(f"Loaded {len(questions)} questions from the Excel file.")

# Navigate to "Questions" section
questions_section = wait.until(EC.presence_of_element_located(
    (By.XPATH, "//div[contains(text(), 'Questions')]")))
questions_section.click()
time.sleep(3)

# Add questions
for i, question_text in enumerate(questions, start=1):
    print(f"Adding question {i}: {question_text}")

    add_question_button = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//div[contains(@class, 'MuiChip-root') and contains(@class, 'MuiChip-filled') and contains(@class, 'MuiChip-clickable')]")
    ))
    add_question_button.click()
    time.sleep(1)

    question_input = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//textarea[@name='description']")))
    question_input.clear()
    question_input.send_keys(question_text)
    time.sleep(1)

    save_button = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//button[contains(text(), 'Save')]")))
    save_button.click()
    time.sleep(2)

print("All questions added successfully.")
