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
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.support import expected_conditions as EC

    # set default output

    check_error = False

    try:
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

        select_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#rdoBOMSEL0"))
            )
        select_button.click()

    except TimeoutException:
        check_error = True

    return check_error