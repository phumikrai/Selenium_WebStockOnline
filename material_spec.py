"""
1. Material Specification
"""

class matspec:
    def __init__(self):
        self.col_to_css = {
        "S.MAF_PRT":"#MainContent_txtR_MAF_PRT",
        "S.PRT_NAME":"#MainContent_txtR_PRT_NAME",
        "S.MAF_PRT_NO":"#MainContent_txtR_MAF_PRT_NO",
        "S.SET_POSIT":"#MainContent_txtR_SET_POSIT",
        "S.MAT_PRT":"#MainContent_txtR_MAT_PRT",
        "S.SIZE":"#MainContent_txtR_SIZE",
        "S.PRT_MODEL":"#MainContent_txtR_PRT_MODEL",
        "S.PRT_SERL":"#MainContent_txtR_PRT_SERL",
        "S.DRAWING":"#MainContent_txtR_DRAWING",
        "S.VENDOR":"#MainContent_txtR_VENDOR",
        "S.OTHER":"#MainContent_txtR_OTHER",
        "S.SET_INSUR":"#MainContent_txtR_SET_INSUR",
        "S.RECOMMEND":"#MainContent_txtR_RECOMMEND",
        "S.SUBTITUT":"#MainContent_txtR_SUBTITUT",
        "S.OLD_MATERIAL":"#MainContent_txtR_OLD_MATERIAL",
        "S.TRNS_TO_MAT":"#MainContent_txtR_TRNS_TO_MAT",
        "S.MAT_TRANSFER":"#MainContent_txtR_MAT_TRANSFER",
        "S.PROJ_NO":"#MainContent_txtR_PROJ_NO",
        "S.MATSPEC_03":"#MainContent_txtR_MATSPEC_03",
        "S.SET_PRESER":"#MainContent_txtR_SET_PRESER",
        "S.WEB_REQUEST":"#MainContent_txtR_WEB_REQUEST",
        "S.SET_TECH":"#MainContent_txtR_SET_TECH"
        }

def select_mat_spec(driver, groupalias, materialalias, groupitem, matcodeitem):
    """
    This function is for data selection into material specification session
    driver: browser driver class created from webdriver
    groupalias: group alias from parameters.json
    matcodealias: material code alias from parameters.json
    groupitem: item from group type column
    matcodeitem: item from material group code column
    """
    
    # import library

    from functions import dropdown_selection

    # set base selector for group type and material group

    base_selector = """#frmDoc > div:nth-child(5) > section > div:nth-child(2) 
                        > div.box-body > div:nth-child(1) > div.col-md-{index} > table 
                        > tbody > tr > td:nth-child(2) {tag}> span > span.selection 
                        > span > span.select2-selection__arrow"""

    # select group type from dropdown list

    group_selector = base_selector.format(index="5",tag="")
    groupname = groupalias[groupitem]
    dropdown_selection(driver=driver, button_css=group_selector, selectname=groupname)

    # select material group from dropdown list

    material_selector = base_selector.format(index="7",tag="> div ")
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
                continue
    else:
        pass

def loadspec(driver):
    """
    This fucntion is for load spec button and return error checker
    driver: browser driver class created from webdriver
    -----
    return: error checker
    """
    # import library

    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException

    # click "Load spec" Button

    driver.find_element(By.CSS_SELECTOR, "#MainContent_btnSPEC").click()
    
    # check error
    try:
        check_error = driver.find_element(
            By.CSS_SELECTOR, "#MainContent_ddlMATGRP-error"
            ).is_displayed()
    
    except NoSuchElementException:
        check_error = False
    
    return check_error

def dump_mat_spec(driver, selectedrow, mat_spec_input, columnname):
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
        if (col in mat_spec_input) and (str(value) != "nan"):
            dumpinput(driver, str(value), mat_spec_input[col])
        else:
            pass



