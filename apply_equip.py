"""
2. Apply Equipment (BOM)
"""

def select_equipment(driver, equipitem):
    """
    This function is for searching and selecting equipment if available.
    driver: browser driver class created from webdriver
    equipitem: item from equipment tag column
    -----
    return: error checker
    """

    # import libraries

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from functions import dumpinput

    # set default output

    check_error = False

    # insert eq.tag for searching

    eqtag_input = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "#MainContent_txtEqTagNo")
            )
        )
    eqtag_input.send_keys(equipitem)

    # click "search" Button

    driver.find_element(By.CSS_SELECTOR, "#btnMdlSearch").click()

    # check search results

    checkbox = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located(
        (By.CSS_SELECTOR, "#tblEQList > tbody > tr > td > input")
        )
    )
    checkbox.click()
    
    # click "add selected to equipment list" Button

    driver.find_element(By.CSS_SELECTOR, "#btnMdlAdd").click()

    # click radio button for equipment selection

    div_css = "#MainContent_gvBOM > tbody > tr > td > div"
    input_css = div_css + " > input"
    javascript = (
        """
            var checkbox = document.querySelector("{input_css}");
            checkbox.checked = true; 
            var setInput = document.querySelector("{div_css}");
            setInput.setAttribute("class", "iradio_square-blue checked");
        """
    ).format(input_css=input_css, div_css=div_css)
    driver.execute_script(javascript)

    # set quantity equal one

    dumpinput(driver, "1", 'input[type="text"][name^="txtBOMQTY"]')

    return check_error