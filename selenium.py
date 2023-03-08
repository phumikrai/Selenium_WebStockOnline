import os
import time
import json
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from functions import dataslicing

# load parameter from parameters.json

with open("parameters.json", "r") as openfile:
    parameters = json.load(openfile)

print(parameters)

# load excel data

filepath = os.getcwd()+"\{}".format(parameters["file name"])
df = pd.read_excel(filepath, sheet_name=parameters["sheet name"])
df = df.drop(index=0).dropna(how="all").set_index("I.ITEM_NO")

# separate data group by 50

def dataseparation(df):
    """
    df = dataframe
    -----
    return = 
    """

    datarange = np.arange(1, n_data, 50)
n_range = len(datarange)

for i in np.arange(n_range-1):
    print(datarange[i], datarange[i+1])

# set pageLoadStrategy

caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "eager"

# load chrome driver

driverpath = os.getcwd()+"\chromedriver.exe"
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path=driverpath, options=options, desired_capabilities=caps)

# set url for web stock online "http://gcgplbiis/webstock/"

homeurl = "http://gcgplbiis/webstock/"

# navigate to url

driver.get(homeurl)

# click material management button

driver.find_element(By.LINK_TEXT, "Material management").click()
time.sleep(1)

# click new material button

driver.find_element(By.LINK_TEXT, "New material").click()
time.sleep(1)

# show maintanance plant list

driver.find_element(
    By.CSS_SELECTOR, 
    '#frm > div.wrapper > div.content-wrapper > section.content > div > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > span > span.selection > span > span.select2-selection__arrow'
    ).click()

# select plant

plantname = parameters["plant name"]
plantlist = driver.find_elements(By.XPATH, "/html/body/span/span/span[2]/ul/li")

for plant in plantlist:
    if plant.text == plantname:
        plant.click()
        break
time.sleep(1)

# show MRP Controller list

driver.find_element(
    By.CSS_SELECTOR, 
    '#frm > div.wrapper > div.content-wrapper > section.content > div > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody > tr:nth-child(3) > td:nth-child(2) > div > span > span.selection > span > span.select2-selection__arrow'
    ).click()

# select plant

mrpcname = parameters["mrpc name"]
mrpclist = driver.find_elements(By.XPATH, "/html/body/span/span/span[2]/ul/li")

for mrpc in mrpclist:
    if mrpc.text == mrpcname:
        mrpc.click()
        break

# click "add item" Button

driver.find_element(
    By.CSS_SELECTOR
    , "#MainContent_btnAdd"
    ).click()

# switch driver to add item window

driver.switch_to.window(driver.window_handles[1])
current_url = driver.current_url
print(current_url)

time.sleep(5)
driver.quit()