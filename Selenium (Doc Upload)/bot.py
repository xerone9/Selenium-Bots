import selenium.common.exceptions
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
# from selenium.common.exceptions import NoSuchElementException
# from selenium.common.exceptions import StaleElementReferenceException
# from selenium.webdriver.support import expected_conditions
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

import openpyxl as xl
import os.path, sys


JS_DROP_FILE = """
    var target = arguments[0],
        offsetX = arguments[1],
        offsetY = arguments[2],
        document = target.ownerDocument || document,
        window = document.defaultView || window;

    var input = document.createElement('INPUT');
    input.type = 'file';
    input.onchange = function () {
      var rect = target.getBoundingClientRect(),
          x = rect.left + (offsetX || (rect.width >> 1)),
          y = rect.top + (offsetY || (rect.height >> 1)),
          dataTransfer = { files: this.files };

      ['dragenter', 'dragover', 'drop'].forEach(function (name) {
        var evt = document.createEvent('MouseEvent');
        evt.initMouseEvent(name, !0, !0, window, 0, 0, 0, x, y, !1, !1, !1, !1, 0, null);
        evt.dataTransfer = dataTransfer;
        target.dispatchEvent(evt);
      });

      setTimeout(function () { document.body.removeChild(input); }, 25);
    };
    document.body.appendChild(input);
    return input;
"""

doc_types = []
with open("Doc Type.ini", "r") as file_types:
    for line in file_types.readlines():
        if line[0] == "#" or line[0] == "\n":
            pass
        else:
            doc_types.append(line.strip())

folder_path = ""
with open("location.ini", "r") as folder_root:
    folder_path = folder_root.read() + "\\"

user_id = input("Enter User ID: ")
user_password = input("Enter Password: ")  

PATH = "C:\Program Files (x86)\chromedriver.exe"
chrome_options = Options() 
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=chrome_options)
driver.execute_script("window.scroll(0, 1250)")

driver.get("http://172.16.0.6:81/ords/r/erasoft/a500/student-documents?hide_unhide_region=1&clear=838&session=1619755178786.&cs=1y5xey-06OpND47wjMWq3OydyHWA")

user_name = driver.find_element(by=By.ID, value="P101_USERNAME")
user_name.send_keys(user_id)

password = driver.find_element(by=By.ID, value="P101_PASSWORD")
password.send_keys(user_password)

general_key = driver.find_element(by=By.ID, value="P101_LOGIN_CODE")
general_key.send_keys("123456")

sign_in = driver.find_element(by=By.ID, value="LOGIN")
sign_in.click() 

for directory in os.listdir(folder_path):
    if os.path.isdir(folder_path + "\\" + directory):               
        count = 0
        while count < 3:
            count += 1
            try:
                stu_id = driver.find_element(by=By.ID, value="P0_V_DIRECT_STUDENT_ID")        
                stu_id.clear()
                stu_id.send_keys(directory)
                stu_id.send_keys(Keys.ENTER)  
                time.sleep(1)
                count = 0    
                break
            except Exception as e:
                time.sleep(1)   
        for filename in os.listdir(folder_path + "\\" + directory + "\\"):
            # print(directory + "-->" + filename)
            try:
                document_type = filename.split(" - ")[0]  
                multiple_files = ""
                if filename.split(" - ")[1].__contains__("("):
                    multiple_files = " (" + filename.split("(")[1].split(".")[0]      
                while count < 3:
                    count += 1
                    try:     
                        select_item = Select(driver.find_element(by=By.ID, value="P838_DOC_TYPE_ID"))
                        if document_type not in doc_types:
                            select_item.select_by_visible_text("Others") 
                        else:                     
                            select_item.select_by_visible_text(document_type) 
                        description = driver.find_element(by=By.ID, value="P838_DESCRIPTION") 
                        description.clear()               
                        description.send_keys(document_type + multiple_files)
                        upload_file = driver.find_element(by=By.CLASS_NAME, value="apex-item-filedrop-body").parent
                        file_input = driver.execute_script(JS_DROP_FILE, driver.find_element(by=By.CLASS_NAME, value="apex-item-filedrop-body"), 0, 0)
                        file_input.send_keys(folder_path + "\\" + directory + "\\" + filename)
                        push_doc = driver.find_element(by=By.ID, value="B4877606197040446")
                        push_doc.click()                         
                        count = 0
                        break
                    except Exception as e:                        
                        time.sleep(1)
                time.sleep(1)
            except IndexError:
                print("no id found in file name. File not uploaded")

input("Review Files If Required. Press Enter To Exit...")

