import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from configparser import ConfigParser
from datetime import datetime

# Load vocab data from Excel
file_path = "Extracted_Vocabulary.xlsx"
vocab_df = pd.read_excel(file_path)

# Load configuration
config = ConfigParser()
config.read('config.ini')
email = config["TARUN_GROVER"]["email"]
password = config["TARUN_GROVER"]["password"]

# Set up the Chrome WebDriver
service = Service(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open the login page
login_url = "https://admin.tarungroverenglish.com/app/"
driver.get(login_url)

# Wait for the page to load and log in if necessary (customize the login process if required)
wait = WebDriverWait(driver, 30)

# Wait for email field and enter email
email_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='email']")))
email_input.clear()
email_input.send_keys(email)

# Wait for password field and enter password
password_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
password_input.clear()
password_input.send_keys(password)

# Click the login button
login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]")
login_button.click()

# Wait for the dashboard page to load
wait.until(EC.presence_of_element_located((By.XPATH, "//h6[contains(text(), 'Dashboard')]")))

# Click on the "Vocabs" link to navigate to the vocab page
vocabs_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div/div/nav/a[2]/div[2]/h6")))
vocabs_link.click()

# Wait for the vocab page to load
wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/app/vocabs/edit?preselectedType=normal')]")))

# Loop through the vocabulary data and add each entry to the website
for index, row in vocab_df.iterrows():
    # Click on the "Add Normal" button
    add_normal_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/app/vocabs/edit?preselectedType=normal')]")))
    add_normal_button.click()
    
    # Fill in the vocab details
    time.sleep(2)  # Wait for the modal to open
    
    # Fill in the "Word" field
    word_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[1]/div/div/input")))
    word_input.clear()
    word_input.send_keys(row["name"])
    
    # Fill in the "Part of Speech" dropdown
    dropdown_element = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/div/div/div")))
    dropdown_element.click()
    part_of_speech_mapping = {
        'n': 'Noun',
        'v': 'Verb',
        'adj': 'Adjective',
        'adv': 'Adverb'
    }
    part_of_speech_value = row['type']
    if isinstance(part_of_speech_value, str):
        part_of_speech = part_of_speech_mapping.get(part_of_speech_value.lower(), 'Noun')
    else:
        part_of_speech = 'Noun'
    
    if part_of_speech:
        option = driver.find_element(By.XPATH, f"//li[contains(text(), '{part_of_speech}')]")
        option.click()
    
    # Fill in the "Date" field with the current date
    current_date = '01-01-2025'
    date_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[3]/div/div/input")))
    date_input.clear()  # Clear existing input
    date_input.send_keys(current_date)

    # Fill in the "Definition" field
    definition_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[3]/div/input")))
    definition_input.clear()
    definition_input.send_keys(row["meaning"])
    
    # Fill in the "Examples" field if available
    if pd.notna(row["examples"]):
        example_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[4]/div/input")))
        example_input.clear()
        example_input.send_keys(row["examples"])
    
    # Fill in the "Synonyms" field if available
    if pd.notna(row["synonyms"]):
        synonym_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[5]/div/input")))
        synonym_input.clear()
        synonym_input.send_keys(row["synonyms"])

    # Fill in the "Trick" field if available
    if pd.notna(row["hint"]):
        trick_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[8]/div/input")))
        trick_input.clear()
        trick_input.send_keys(row["hint"])
    
    # Click on the "Create" button
    create_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div/div/div/div/div[10]/button")))
    create_button.click()

    # Wait for the vocab page to load again
    time.sleep(3)

    # Navigate back to the vocab page to add the next word
    vocabs_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div/div/nav/a[2]/div[2]/h6")))
    vocabs_link.click()

    # Wait for the vocab page to fully load
    wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/app/vocabs/edit?preselectedType=normal')]")))

# Close the browser
driver.quit()
