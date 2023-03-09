# standard library

import os
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

# local library

from functions import dataslicing, loaddata, greeting, progressreport, dropdown_selection
from material_spec import autodump_mat_spec

# load parameter from parameters.json

with open("parameters.json", "r") as openfile:
    parameters = json.load(openfile)

# load excel data

filepath = os.getcwd()+"\{}".format(parameters["file name"])

df = loaddata(filepath, parameters["sheet name"])
df = df.drop(index=0).dropna(how="all")

# count row

n_row = len(df.index)

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

try:
    manage_button = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Material management")))
    manage_button.click()

except TimeoutException:
    print("Loading took too much time!")

# print text within console

greeting(parameters["plant name"], parameters["mrpc name"], parameters["file name"])

# error collection

errorlist = []

# loop through sliced dataframe

for dataframe in sliced_df:
    
    # click new material button
    
    try:
        mat_button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.LINK_TEXT, "New material")))
        mat_button.click()
    
    except TimeoutException:
        print("Loading took too much time!")

    # select maintanance plant from dropdown list

    plant_selector = """#frm > div.wrapper > div.content-wrapper > section.content > div
                        > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody
                        > tr:nth-child(2) > td:nth-child(2) > div > span > span.selection 
                        > span > span.select2-selection__arrow"""
    plantname = parameters["plant name"]
    dropdown_selection(driver=driver, button_css=plant_selector, selectname=plantname)

    # select MRP Controller from dropdown list

    mrpc_selector = """#frm > div.wrapper > div.content-wrapper > section.content > div 
                        > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody 
                        > tr:nth-child(3) > td:nth-child(2) > div > span > span.selection
                        > span > span.select2-selection__arrow"""
    mrpcname = parameters["mrpc name"]
    dropdown_selection(driver=driver, button_css=mrpc_selector, selectname=mrpcname)

    # loop through each row of dataframe

    for row in dataframe.iterrows():

        # click "add item" Button

        driver.find_element(
            By.CSS_SELECTOR, 
            "#MainContent_btnAdd"
            ).click()

        # check page redirect and set windows

        try:
            WebDriverWait(driver, 10).until(
                EC.new_window_is_opened(driver.window_handles))
            before_window = driver.window_handles[0]
            after_window = driver.window_handles[-1]
        
        except TimeoutException:
            print("Loading took too much time!")

        # switch driver to new one

        driver.switch_to.window(after_window)

        """
        1. Material Specification
        """

        # set parameters

        groupalias = parameters["group alias"]
        matcodealias = parameters["matcode alias"]
        groupitem = row[1]["I.GROUP_TYPE"]
        matcodeitem = str(row[1]["I.MATGRP_CODE"])

        # dump data into material specification session

        check_error = autodump_mat_spec(driver, groupalias, matcodealias, groupitem, matcodeitem)
        
        # if error is still raising, print error as file and break this loop

        if check_error:
            set_format = (row[0], matcodeitem, row[1]["S.PRT_NAME"])
            error_text = "Item No. %s, Material Group %s for %s is not found." %(set_format)
            errorlist.append(error_text)
            
            break

            driver.close()
        else:
            pass

        # report status

        progressreport(indexnumber=row[0], totalrow=n_row)

        break

        # switch driver back to material request page  

        driver.switch_to.window(before_window)
    
    # this break will be removed after done

    break

print("\nDone, Sed Lew Jaa~")
# driver.quit()

# print error report if found

if len(errorlist) > 0:
    with open("error_report.txt", "w") as errorfile:
        for erroritem in errorlist:
            errorfile.write('%s\n' %erroritem)
    errorfile.close()
else:
    pass