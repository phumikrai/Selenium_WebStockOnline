"""
1. Material Specification
"""

def select_mat_spec(driver, groupalias, materialalias, groupitem, matcodeitem):
    """
    This function is for automatically dumping data into material specification session
    driver: browser driver class created from webdriver
    groupalias: group alias from parameters.json
    matcodealias: material code alias from parameters.json
    groupitem: item from group type column
    matcodeitem: item from material group code column
    -----
    return: error checker
    """
    # import library

    from selenium.webdriver.common.by import By
    from functions import dropdown_selection

    # select group type from dropdown list

    group_selector = """#frmDoc > div:nth-child(5) > section > div:nth-child(2) 
                        > div.box-body > div:nth-child(1) > div.col-md-5 > table 
                        > tbody > tr > td:nth-child(2) > span > span.selection 
                        > span > span.select2-selection__arrow"""
    groupname = groupalias[groupitem]
    dropdown_selection(driver=driver, button_css=group_selector, selectname=groupname)

    # select material group from dropdown list

    material_selector = """#frmDoc > div:nth-child(5) > section > div:nth-child(2) 
                            > div.box-body > div:nth-child(1) > div.col-md-7 > table 
                            > tbody > tr > td:nth-child(2) > div > span > span.selection 
                            > span > span.select2-selection__arrow"""

    try:
        materialname = materialalias[matcodeitem]
        dropdown_selection(driver=driver, button_css=material_selector, selectname=materialname)

        # click load spec and check error

        check_error = loadspec(driver)

        # if error then loop through the remaining type for checking

        if check_error:
            for key, item in groupalias.items():
                if key != groupitem:

                    # select group type from dropdown list again

                    dropdown_selection(driver=driver, button_css=group_selector, selectname=item)

                    # select material group from dropdown list again

                    dropdown_selection(driver=driver, button_css=material_selector, selectname=materialname)

                    # click load spec and check error

                    check_error = loadspec(driver)

                    if check_error:
                        continue
                    else:
                        break
        else:
            pass

    except KeyError:
        check_error = True
        
    return check_error

def loadspec(driver):
    """
    This fucntion is for load spec button and return error checker
    driver: browser driver class created from webdriver
    """
    # import library

    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException

    # click "Load spec" Button

    driver.find_element(
        By.CSS_SELECTOR
        , "#MainContent_btnSPEC"
        ).click()
    
    # check error
    
    try:
        check_error = driver.find_element(
                            By.CSS_SELECTOR, 
                            "#MainContent_ddlMATGRP-error"
                            ).is_displayed()
    
    except NoSuchElementException:
        check_error = False

    return check_error

def dump_mat_spec(driver, selectedrow, column_to_field, mat_spec_input):
    """
    This function is for automatically dumping data into material specification session
    driver: browser driver class created from webdriver
    selectedrow: seleced row for data input
    mat_spec_input: column and field related as dict
    columnname : name of column within dataframe
    """
    # import library

    from functions import dumpinput

    # loop through each column to dump available field

    for col, value in zip(columnname, selectedrow):
        if col in column_to_field:
            dumpinput(driver, str(value), column_to_field[col])
        else:
            pass