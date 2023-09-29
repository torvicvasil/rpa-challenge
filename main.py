from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
ChromeDriverManager().install()
# Initialize the WebDriver
driver = webdriver.Chrome()
from selenium.webdriver.common.by import By
import requests
import pandas as pd
import time

# URL of the page you want to visit
url = "https://rpachallenge.com/"

def download_file():
    file_url = "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"
    r = requests.get(file_url, allow_redirects=True)
    df = pd.read_excel(r.content)
    return df

def fill_in_fields(first_name, last_name, email, role, company_name, phone_number, address):
    first_name_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelFirstName']")
    first_name_element.send_keys(first_name)

    last_name_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelLastName']")
    last_name_element.send_keys(last_name)

    email_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelEmail']")
    email_element.send_keys(email)

    role_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelRole']")
    role_element.send_keys(role)

    company_name_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelCompanyName']")
    company_name_element.send_keys(company_name)

    phone_number_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelPhone']")
    phone_number_element.send_keys(phone_number)

    address_element = driver.find_element(By.CSS_SELECTOR, "[ng-reflect-name='labelAddress']")
    address_element.send_keys(address)

try:
    # Open the URL
    driver.get(url)

    # Get the title of the page
    page_title = driver.title

    # Print the title
    print("Page title:", page_title)

    file = download_file()

    print(file)

    start_element = driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")
    start_element.click()

    # using a itertuples()
    for i in file.itertuples():
        first_name = i[1]
        last_name = i[2]
        email = i[6]
        role = i[4]
        company_name = i[3]
        phone_number = i[7]
        address = i[5]
        print(f"({first_name}, {last_name}, {email}, {role}, {company_name}, {phone_number}, {address})")
        fill_in_fields(first_name, last_name, email, role, company_name, phone_number, address)
        submit_element = driver.find_element(By.CSS_SELECTOR, "[type='submit']")
        submit_element.click()




finally:
    # Close the browser window
    time.sleep(10)
    driver.quit()