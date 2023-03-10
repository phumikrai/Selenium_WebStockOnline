"""
3. Material Master
"""

class matmaster:
    def __init__(self):
        self.col_to_css = {
        "M.SAFETY_STOCK":"#MainContent_tn0SAFETY_STOCK_N",
        "M.MAX_STOCK":"#MainContent_tn0MAX_STOCK_N",
        "M.MOVING_PRICE":"#MainContent_tn0MOVING_PRICE_N"
        }

def dump_mat_master(driver, selectedrow, mat_master_input, columnname):
    """
    This function is for automatically dumping data into material specification session
    driver: browser driver class created from webdriver
    selectedrow: seleced row for data input
    mat_master_input: column and field related as dict
    columnname : name of column within dataframe
    """
    # import library

    from functions import dumpinput

    # loop through each column to dump available field

    for col, value in zip(columnname, selectedrow):
        if (col in mat_master_input) and (str(value) != "nan"):
            dumpinput(driver, str(value), mat_master_input[col])
        else:
            pass