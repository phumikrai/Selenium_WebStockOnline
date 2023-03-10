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

from functions import dataslicing, loaddata, greeting, progressreport, dropdown_selection, dumpinput
from material_spec import select_mat_spec, dump_mat_spec, matspec
from apply_equip import select_equipment
from mat_master import matmaster

# load parameter from parameters.json

with open("aliasfile.json", "r") as openfile:
    parameters = json.load(openfile)

# load col to css input data

mat_spec_input = matspec().col_to_css
mat_master_input = matmaster().col_to_css

# load excel data

filepath = os.getcwd()+"\{}".format(parameters["file name"])

df = loaddata(filepath, parameters["sheet name"])
df = df.drop(index=0).dropna(how="all")

# count row

n_row = len(df.index)

# separate data group by 50

sliced_df = dataslicing(df)

# get column name

columnname = df.columns

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
        EC.presence_of_element_located((By.LINK_TEXT, "Material management"))
        )
    manage_button.click()

except TimeoutException:
    print("Loading took too much time!")


# print text within console

greeting(parameters["plant name"], parameters["mrpc name"], parameters["file name"], n_row)

# error collection

errorlist = []

# loop through sliced dataframe

for dataframe in sliced_df:
    
    # click new material button
    
    try:
        mat_button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.LINK_TEXT, "New material"))
            )
        mat_button.click()
    
    except TimeoutException:
        print("Loading took too much time!")

    # set base selector for maintance plant and mrp controller

    plant_mrpc_base = """#frm > div.wrapper > div.content-wrapper > section.content > div
                        > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody
                        > tr:nth-child({index}) > td:nth-child(2) > div > span > span.selection 
                        > span > span.select2-selection__arrow"""
    
    # select maintanance plant from dropdown list

    plant_selector = plant_mrpc_base.format(index="2")
    plantname = parameters["plant name"]
    load_error = dropdown_selection(driver=driver, button_css=plant_selector, selectname=plantname)

    # select mrp controller from dropdown list

    mrpc_selector = plant_mrpc_base.format(index="3")
    mrpcname = parameters["mrpc name"]
    dropdown_selection(driver=driver, button_css=mrpc_selector, selectname=mrpcname)

    # loop through each row of dataframe

    for row in dataframe.iterrows():

        # click "add item" Button

        driver.find_element(By.CSS_SELECTOR, "#MainContent_btnAdd").click()

        # check page redirect and set windows

        try:
            WebDriverWait(driver, 10).until(
                EC.new_window_is_opened(driver.window_handles)
                )
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

        # select material specification session

        check_error = select_mat_spec(driver, groupalias, matcodealias, groupitem, matcodeitem)
        
        # if error is still raising, print error as file and break this loop

        if check_error:
            set_format = (row[0], matcodeitem, row[1]["S.PRT_NAME"])
            error_text = "Item No. %s, Material Group %s for %s is not found." %(set_format)
            errorlist.append(error_text)
            
            break

            driver.close()
        else:
            pass

        # set base selector for all dropdown lists

        matspec_base = """#MainContent_dvR_{index} > table > tbody > tr > td:nth-child(2) > div 
                        > span > span.selection > span > span.select2-selection__arrow"""
        
        # select all dropdown lists

        obg = ("001", "obgalias", "S.MATSPEC_01")
        store = ("049", "storealias", "S.STORING")
        detail = ("050", "detailalias", "S.STORING_DETAIL")
        weight = ("052", "weightalias", "S.WEIGHT")
        unit = ("054", "unitalias", "S.UNIT")

        all_dropdowns = [obg, store, detail, weight, unit]

        for dropdown in all_dropdowns:
            selector = matspec_base.format(index=dropdown[0])
            selectname = parameters[dropdown[1]][row[1][dropdown[2]]]
            dropdown_selection(driver=driver, button_css=selector, selectname=selectname)

        # dump data into inputs of material specification session

        dump_mat_spec(driver, row[1], mat_spec_input, columnname)

        # dump "set of" data if available 

        if (row[1]["S.UNIT"] == "SET"):
            dumpinput(driver, str(row[1]["S.SET_CONSIST"]), "#MainContent_txtR_SET_CONSIST")

        # ***** leave default values for receiving method ***** 

        """
        2. Apply Equipment (BOM)
        """

        # click "search eq.bom" Button

        driver.find_element(By.CSS_SELECTOR, "#MainContent_btnBOM").click()

        # select equipment

        check_error = select_equipment(driver, row[1]["S.EQ_TAG"])

        # if error is raising, print error as file and break this loop

        if check_error:
            set_format = (row[0], row[1]["S.EQ_TAG"], row[1]["S.PRT_NAME"])
            error_text = "Item No. %s, Equipment Tag %s for %s is not found." %(set_format)
            errorlist.append(error_text)
            break

            driver.close()
        else:
            pass


        """
        3. Material Master
        """

        # dump data into inputs of material master session

        dump_mat_master(driver, row[1], mat_master_input, columnname)

        """
        4. Material Return Stock
        """

        

        """
        5. Attach Document
        """


        """
        6. Request Reason
        """

        # report status

        progressreport(indexnumber=row[0], totalrow=n_row)        

        # switch driver back to material request page  

        driver.switch_to.window(before_window)

        break
    
    # this break will be removed after done

    break

print("\nSed Lew Jaa~")
# driver.quit()

# print error report if found

if len(errorlist) > 0:
    with open("error_report.txt", "w") as errorfile:
        for erroritem in errorlist:
            errorfile.write('%s\n' %erroritem)
    errorfile.close()
else:
    pass