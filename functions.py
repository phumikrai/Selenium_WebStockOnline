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
    """
    # import pandas and warning

    import pandas as pd
    import warnings

    warnings.simplefilter(action='ignore', category=UserWarning)
    
    # return read excel function

    return pd.read_excel(filepath, sheet_name=sheetname)