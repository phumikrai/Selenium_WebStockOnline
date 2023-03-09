# this file is for function collection

def dataslicing(df):
    """\
    This function is to slice dataframe by 50 
    df: dataframe
    -----
    return: list of sliced dataframe
    """
    # import numpy

    import numpy as np

    # get parameters

    n_data = len(df.index)
    datarange = np.arange(0, n_data, 50)
    n_range = len(datarange)
    
    # extract data range 
    
    range_set = [(datarange[i], datarange[i+1]) for i in np.arange(n_range-1)]
    
    # append last range if available
    
    if n_data != datarange[-1]:
        range_set.append((datarange[-1], n_data))
    
    # return list of sliced dataframe

    return [df.iloc[a:b] for a, b in range_set]

def loaddata(filepath: str, sheetname: str):
    """
    Load data from an Excel file without warning
    filepath: file path of excel file
    sheetname: sheet name will be focused on
    -----
    return: excel data within dataframe
    """
    # import pandas and warning

    import pandas as pd
    import warnings

    warnings.simplefilter(action='ignore', category=UserWarning)
    
    # return read excel function

    return pd.read_excel(filepath, sheet_name=sheetname)

def greeting(plantname: str, mrpc: str, filename: str):
    """
    Greeting text using parameters.json file
    plantname: name of plant
    mrpc: name of mrpc
    filename: name of excel file
    -----
    return: greeting text on console
    """
    import sys

    sys.stdout.write("Bot Jaa~ - Automation for dumping data to Web Stock Online\n\n")
    sys.stdout.write("Stock code is being created for plant \n\t\"%s\"\n" %plantname)
    sys.stdout.write("through MRP Controller \n\t\"%s\"\n" %mrpc)
    sys.stdout.write("using \n\t\"%s\" file.\n\n" %filename)
    sys.stdout.write("% Progress\n")

def progressreport(indexnumber: int, totalrow: int):
    """
    print progress report on console
    indexnumber: indexnumber
    totalrow: total row of data as data count
    -----
    return: progress text and bar on console
    """
    import sys

    # get parameters
    
    strprogress = "{}/{}".format(indexnumber,totalrow)
    progress = round((indexnumber/totalrow)*100, 2)
    step = int(progress//5)

    # print out progress

    sys.stdout.write('\r')
    sys.stdout.write("%9s [%-20s] %.2f%%" %(strprogress, '='*step, progress))
    sys.stdout.flush()

def dropdown_selection(driver, button_css, selectname):
    """
    This function is for dropdown selection
    driver: browser driver class created from webdriver
    button_css: css selector for dropdown button
    selectname: name of item to be selected
    -----
    return: progress text and bar on console
    """
    # import relevant libraries

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.support import expected_conditions as EC

    # click dropdown button

    try:
        button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
            By.CSS_SELECTOR, button_css
            )))
        button.click()
    
    except TimeoutException:
        print("Loading took too much time!")

    # select item

    itemlist = driver.find_elements(By.XPATH, "/html/body/span/span/span[2]/ul/li")

    for item in itemlist:
        if item.text == selectname:
            item.click()
            break