import selenium.common.exceptions
import time
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv


user_id = input("Enter User ID: ")
user_password = input("Enter Password: ")
file_name = input("Enter CSV File Name (with extension): ")

chrome_options = Options() 
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
driver.get("http://172.16.0.6:81/ords/f?p=700:740:14062566633627.:::740:hide_unhide_region:1&cs=15D3B1DF29B0283C9835B2D43941D190D")


user_name = driver.find_element(by=By.ID, value="P101_USERNAME")
user_name.send_keys(user_id)

password = driver.find_element(by=By.ID, value="P101_PASSWORD")
password.send_keys(user_password)

general_key = driver.find_element(by=By.ID, value="P101_LOGIN_CODE")
general_key.send_keys("123456")

sign_in = driver.find_element(by=By.ID, value="P101_LOGIN")
sign_in.click()


with open(file_name, "r") as csv_file:
    csv_reader = csv.reader(csv_file)
    line_count = 0
    for voucher in csv_reader:
        line_count += 1           
        voucher_no = driver.find_element(by=By.ID, value="P740_VOUCHER_NO")
        count = 0
        # wait for element "P740_VOUCHER_NO_TO_SEARCH"
        while count < 3:
            count += 1
            try:         
                voucher_no.clear()
                voucher_no.send_keys(voucher[1])
                voucher_no.send_keys(Keys.ENTER)
                break
            except Exception as e:
                time.sleep(1)
driver.close()
