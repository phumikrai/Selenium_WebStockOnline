import os
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from functions import dataslicing, loaddata

# load parameter from parameters.json

with open("parameters.json", "r") as openfile:
    parameters = json.load(openfile)

# load excel data

filepath = os.getcwd()+"\{}".format(parameters["file name"])

df = loaddata(filepath, parameters["sheet name"])
df = df.drop(index=0).dropna(how="all").set_index("I.ITEM_NO")

# count row

n_row = len(df.index)
print(n_row)

# separate data group by 50

sliced_df = dataslicing(df)

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

# loop through sliced dataframe

for dataframe in sliced_df:
    
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

    # select mrpc

    mrpcname = parameters["mrpc name"]
    mrpclist = driver.find_elements(By.XPATH, "/html/body/span/span/span[2]/ul/li")

    for mrpc in mrpclist:
        if mrpc.text == mrpcname:
            mrpc.click()
            break

    # loop through each row of dataframe

    for row in dataframe.iterrows():

        # # click "add item" Button

        # driver.find_element(
        #     By.CSS_SELECTOR
        #     , "#MainContent_btnAdd"
        #     ).click()
        # time.sleep(1)

        # # switch driver to add item window

        # driver.switch_to.window(driver.window_handles[1])
        # current_url = driver.current_url

        # time.sleep(5)
        # driver.quit()

        completion = (row[0]/n_row)*100
        partname = row[1]["S.PRT_NAME"]
        eqtag = row[1]["S.EQ_TAG"]
        print("{}% {} is uploaded for {}.".format(completion, partname, eqtag))