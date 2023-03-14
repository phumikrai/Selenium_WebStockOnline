# standard library

import os
import sys
import json
import time 
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from webdriver_manager.chrome import ChromeDriverManager # remote webdriver

# local library

from functions import dataslicing, loaddata, greeting, progressreport, dropdown_selection, dumpinput, loopdump
from material_spec import select_mat_spec, matspec
from apply_equip import select_equipment
from mat_master import matmaster

# determine if application is a script file or exe file

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

# load parameter from parameters.json

jsonpath = os.path.join(application_path, "aliasfile.json")
with open(jsonpath, "r") as openfile:
    parameters = json.load(openfile)

# print text within console

greeting(parameters["plant name"], parameters["mrpc name"], parameters["file name"])

# load excel data

excelpath = os.path.join(application_path, parameters["file name"])
df = loaddata(excelpath, parameters["sheet name"])
df = df.drop(index=0).dropna(how="all")

# separate data group by 50

sliced_df = dataslicing(df)

# get general parameters

homeurl = "http://gcgplbiis/webstock/" # web stock online url
mat_spec_col_to_field = matspec().col_to_css # mapping column name with css field
mat_master_col_to_field = matmaster().col_to_css # mapping column name with css field
n_row = len(df.index) # count row
columnname = df.columns # column name
errorlist = [] # error collection
timeout_error = False # default timeout error
item_error = False # default item error

# get selenium parameters

caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "eager" # action without full loading

# print initial progress

text = "Press Enter to start..."
sys.stdout.write(input(text))
sys.stdout.write("\r% Progress\n" + " "*len(text))
sys.stdout.flush()
progressreport(indexnumber=0, totalrow=n_row)

# start program

starttime = time.time()

# load remote chrome driver

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), 
                          options=options, desired_capabilities=caps)

# set loaded driver

driver.implicitly_wait(1) # implicitly wait = 1 sec

# navigate to url

driver.get(homeurl)

# click material management button

try:
    manage_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Material management"))
        )
    manage_button.click()

except (TimeoutException, ConnectionResetError):
    timeout_error = True

# program is runnning if no timeout error

if not(timeout_error):

    # loop through sliced dataframe

    for dataframe in sliced_df:
        
        # click new material button
        
        try:
            mat_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.LINK_TEXT, "New material"))
                )
            mat_button.click()
            
            # set base selector for maintance plant and mrp controller

            plant_mrpc_base = """#frm > div.wrapper > div.content-wrapper > section.content > div
                                > div > div > div:nth-child(2) > div:nth-child(2) > table > tbody
                                > tr:nth-child({index}) > td:nth-child(2) > div > span > span.selection 
                                > span > span.select2-selection__arrow"""
            
            # select maintanance plant from dropdown list

            plant_selector = plant_mrpc_base.format(index="2")
            plantname = parameters["plant name"]
            dropdown_selection(driver=driver, button_css=plant_selector, selectname=plantname)

            # select mrp controller from dropdown list

            mrpc_selector = plant_mrpc_base.format(index="3")
            mrpcname = parameters["mrpc name"]
            dropdown_selection(driver=driver, button_css=mrpc_selector, selectname=mrpcname)

        except (TimeoutException, ConnectionResetError):
            timeout_error = True
            break

        # loop through each row of dataframe

        for row in dataframe.iterrows():
            
            try:
                # click "add item" Button

                driver.find_element(By.CSS_SELECTOR, "#MainContent_btnAdd").click()

                # check page redirect 

                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                before_window = driver.window_handles[0]
                after_window = driver.window_handles[-1]

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

                select_mat_spec(driver, groupalias, matcodealias, groupitem, matcodeitem)

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

                loopdump(driver, row[1], mat_spec_col_to_field, columnname)

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

                """
                3. Material Master
                """

                # dump data into inputs of material master session

                loopdump(driver, row[1], mat_master_col_to_field, columnname)

                """
                4. Material Return Stock
                """

                # ***** leave default values for this ***** 

                """
                5. Attach Document
                """

                # ***** assume that there are no attachments ***** 

                """
                6. Request Reason
                """

                # dump reason from aliasfile

                dumpinput(driver, parameters["request reason"], "#MainContent_txtREASON")
                
                """
                save record 
                """

                # click save button

                driver.find_element(By.CSS_SELECTOR, "#MainContent_btnSave2").click()

                # check completion

                message = driver.find_element(By.CSS_SELECTOR, "#MsgDetail").text

                if message == "Save completed.":
                    driver.find_element(By.CSS_SELECTOR, "#Msg-OK").click() 
                else:
                    driver.find_element(By.CSS_SELECTOR, "#Msg-OK").click()
                    raise Exception
                
                # report progress status 

                progressreport(indexnumber=row[0], totalrow=n_row)        

                # close driver and switch back to material request page  

                driver.close()
                driver.switch_to.window(before_window)
                driver.find_element(By.CSS_SELECTOR, "#MainContent_btnRefresh").click()
                    
            except (TimeoutException, KeyError, NoSuchElementException, Exception, ConnectionResetError):
                
                # set text components
                
                item_error = True
                item = row[0]
                part = row[1]["S.PRT_NAME"]
                mat = str(row[1]["I.MATGRP_CODE"])
                tag = row[1]["S.EQ_TAG"]
                error_text = "Item No. {item}, {part}, got an error, maybe mat.group {mat} or eq.tag {tag} is wrong."
                
                # combine and collect an error
                
                format_text = error_text.format(item=item, part=part, mat=mat, tag=tag)
                errorlist.append(format_text)
                
                # report progress status

                progressreport(indexnumber=row[0], totalrow=n_row)

                # close driver and switch back to material request page
                
                driver.close()
                driver.switch_to.window(before_window)
                continue 

        # save draft if loop through each 50 items

        driver.find_element(By.CSS_SELECTOR, "#MainContent_btnSaveDraft").click()

else:
    pass

# print result and error (if available)

if (timeout_error == True) and (item_error == False):
    print("\n\nSome web element is not found. Please check your internet, maybe it is slow.")
elif (timeout_error == True) and (item_error == True):
    print("\n\nSome web element is not found. Please check your internet and errorlog.txt file.")
elif (timeout_error == False) and (item_error == True):
    print("\n\nSed Lew Jaa~, but there are some error na~, please check errorlog.txt file.")
else:
    print("\n\nSed Lew Jaa~")

# finish progream and print out time spending

endtime = time.time()
timediff = (endtime-starttime)/60
print("You are just spending only {timediff:.2f} mins for this automation".format(timediff=timediff))
driver.quit()

# print error report (if available)

if len(errorlist) > 0:
    with open("errorlog.txt", "w") as errorfile:
        for erroritem in errorlist:
            errorfile.write('%s\n' %erroritem)
    errorfile.close()
else:
    pass

# quit program

text = "\nPress Enter to quit..."
sys.stdout.write(input(text))