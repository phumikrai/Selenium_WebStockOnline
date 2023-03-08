import os
import sys
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from functions import dataslicing, loaddata

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
    manage_button = WebDriverWait(driver, parameters["time delay"]).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Material management")))
    manage_button.click()

except TimeoutException:
    print("Loading took too much time!")

# print text within console

texts = (parameters["plant name"], parameters["mrpc name"])

sys.stdout.write("Stock code is being created for plant \"%s\"\n" %parameters["plant name"])
sys.stdout.write("through MRP Controller \"%s\"\n" %parameters["mrpc name"])
sys.stdout.write("using \"%s\" file.\n\n" %parameters["file name"])
sys.stdout.write("%Progress\n")

# loop through sliced dataframe

for dataframe in sliced_df:
    
    # click new material button
    
    try:
        mat_button = WebDriverWait(driver, parameters["time delay"]).until(
            EC.presence_of_element_located((By.LINK_TEXT, "New material")))
        mat_button.click()
    
    except TimeoutException:
        print("Loading took too much time!")

    # show maintanance plant list

    try:
        plant_button = WebDriverWait(driver, parameters["time delay"]).until(
            EC.presence_of_element_located((
            By.CSS_SELECTOR, 
             '#frm > div.wrapper > div.content-wrapper > section.content > div > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody > tr:nth-child(2) > td:nth-child(2) > div > span > span.selection > span > span.select2-selection__arrow'
             )))
        plant_button.click()
    
    except TimeoutException:
        print("Loading took too much time!")

    # select plant

    plantname = parameters["plant name"]
    plantlist = driver.find_elements(By.XPATH, "/html/body/span/span/span[2]/ul/li")

    for plant in plantlist:
        if plant.text == plantname:
            plant.click()
            break

    # show MRP Controller list

    try:
        mrpc_button = WebDriverWait(driver, parameters["time delay"]).until(
            EC.presence_of_element_located((
            By.CSS_SELECTOR, 
             '#frm > div.wrapper > div.content-wrapper > section.content > div > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody > tr:nth-child(3) > td:nth-child(2) > div > span > span.selection > span > span.select2-selection__arrow'
             )))
        mrpc_button.click()
    
    except TimeoutException:
        print("Loading took too much time!")

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

        # report status

        sys.stdout.write('\r')
        progress = round((row[0]/n_row)*100, 2)
        i = int(progress//5)
        sys.stdout.write("[%-20s] %.2f%%" %('='*i, progress))
        sys.stdout.flush()
        # time.sleep(0.1)

sys.stdout.write("\n")
driver.quit()
