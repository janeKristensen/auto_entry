import pandas as pd
import openpyxl

def get_customer_data(filenames: list):
    """ Skip rows and retrieve nrows, 
    adjusted to fit the header of order form
    Return a pandas dataframe with customer information"""

    for file in filenames: 
        df_customer = pd.read_excel(
            file,
            #sheet_name='Sheet1',
            header=None,
            skiprows=8,
            nrows=11,
            usecols="A:B",
            index_col=0
            )
        df_customer = df_customer.fillna('')
        break
          
    return df_customer


def get_order_data(filenames: list):
    """ Skip rows of order form header and retrieve all other rows
    Call split method to divide order into requisitions 
    based on shipping temp. and DGR status
    Returns a list of dataframes containing order"""
    df_order = pd.DataFrame()
    
    for file in filenames:
        df_data = pd.read_excel(
            file,
            #sheet_name='Sheet1',
            header=1,
            skiprows=19,
            usecols="A:I"
            )
        df_data = df_data.dropna(subset=['Ordered QTY'])
        
        # If no country of origin is listed in order form,
        # the program defaults to DK
        if 'Origin' not in df_data:
            df_data['Origin'] = 'DK'

        if df_order.empty:
            df_order = df_data
        else:
            df_order = pd.concat([df_data, df_order])
    
            
    df_order = df_order.groupby(['Products'], as_index=False).agg(
        {'Ordered QTY':'sum', 
         'HS Code': 'first', 
         'Dangerous goods': 'first', 
         'Price/Unit in USD only for customs purpose': 'max', 
         'Shipping temp.': 'first', 
         'Shipping Instructions': 'first',
         'Origin': 'first'
         }).reset_index()       
    split_order_arr = split_order(df_order)
    print(split_order_arr)
    return split_order_arr
    


def sort_items(order_data_arr, items):
    room = items[items['Shipping temp.']  == '15-25°C']
    ambient = items[items['Shipping temp.']  == 'Ambient']
    cool = items[items['Shipping temp.']  == '2-8°C']
    freeze = items[items['Shipping temp.']  == 'Below -15°C']
    
    # If no items are at room temp. ship at ambient
    if room.empty:
        if not ambient.empty:
            order_data_arr.append(ambient)      
    else:
        # If all items are at room temp  
        if ambient.empty:
            order_data_arr.append(room)    
        else:
            # Ambient and 15-25DC is shipped together
            room = pd.concat([room, ambient], ignore_index=True)
            order_data_arr.append(room) 
        
    if not cool.empty:
        order_data_arr.append(cool)
    
    if not freeze.empty:
        order_data_arr.append(freeze)


def split_order(df_order):

    """ All items with DGR status is put into a new dataframe.
    Order dataframe is copied to Non-DGR dataframe.
    Items in DGR dataframe are removed from Non-DGR """

    dg_items = df_order.dropna(subset=['Dangerous goods'])
    dg_items.fillna('')
   
    n_dg_items = df_order
    cond = n_dg_items['Dangerous goods'].isin(dg_items['Dangerous goods'])
    n_dg_items.drop(n_dg_items[cond].index, inplace = True)
    n_dg_items = n_dg_items.fillna('')

    order_data_arr = list()
    sort_items(order_data_arr, dg_items)
    sort_items(order_data_arr, n_dg_items)

    return order_data_arr

