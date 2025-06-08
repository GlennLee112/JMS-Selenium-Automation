import JMS

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

from datetime import datetime

import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException


driver = JMS.Driver()

# 688575345125

Delivery_Sign_xpath = (f"//div[@class='el-table__body-wrapper is-scrolling-none']//tbody//"
                    f"tr[.//td[contains(., 'Delivery Signature')]and .//td[contains(., 'ECP GAMBIR 366')]]")

Delivery_Sign_elements = driver.find_elements(By.XPATH, Delivery_Sign_xpath)

print(len(Delivery_Sign_elements))

time_xpath = "./td[3]//span"

Delivery_Sign = datetime.strptime((Delivery_Sign_elements[0].find_element(By.XPATH, time_xpath).text),
                               '%Y-%m-%d %H:%M:%S')

print(Delivery_Sign)