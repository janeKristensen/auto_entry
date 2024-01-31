import string
import dialog
import pandas as pd
import functions as fc
import data_entry as de
import shipment_request as request
from tkinter import messagebox as MB
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC




def fill_customer(driver, wait: WebDriverWait, df_customer):
    """ Retrieve shipping data from customer dataframe
        and send to functions that locate and enter data into form"""

    try:
        # Collect information from dataframe.
        country = df_customer.loc["COUNTRY:"].item()
        customer = df_customer.loc["RECEIVER:"].item()
        city = df_customer.loc["CITY:"].item()
        post = df_customer.loc["POSTAL CODE:"].item()
        street = df_customer.loc["STREET:"].item()
        contact = df_customer.loc["CONTACT PERSON:"].item()
        email = df_customer.loc["E-MAIL:"].item()
        phone = str(df_customer.loc["PHONE NO.:"].item())

        # Remove spaces in phone number 
        phone = phone.replace(' ', '') 
        p_str = string.punctuation
        # Remove all punctuation except '+'
        p_str = p_str.replace('+', '')
        phone = phone.translate(str.maketrans({key: None for key in p_str}))  
    except Exception as e:
        err = MB.showerror(
            "A critical error occured",
            "A critical error occured while collecting data from file. "
            "Click OK to terminate the program."
            )
        print({str(e)})
        if(err):
            fc.kill_driver(driver)

    # Fill data into form  
    de.enter_country(driver, wait, country)     
    de.enter_customer(wait, customer)
    de.enter_city(wait, city)   
    de.enter_post(wait, post)    
    de.enter_street(wait, street)     
    de.enter_contact(wait, contact)     
    de.enter_email(wait, email)
    de.enter_phone(wait, phone)
    de.enter_eori(wait, "EORI") # EORI number - always the same
        
   

def fill_order(driver, wait: WebDriverWait, df_order):
    """ Retrieve item data from customer order form dataframe
    and send to functions that locate and enter data into form.
    Orders has previously been divided into dataframes.
    This function loops through all items in a dataframe."""
    
    # Resetting row index of the dataframe to start from 0
    df_order_i = df_order.reset_index()

    # Get shipping temp. and DGR status and input to form
    ship_temp = str(df_order_i.loc[0, "Shipping temp."])
    dgr = str(df_order_i.loc[0, "Dangerous goods"])
    de.temp_input(wait, ship_temp)
    de.dgr_input(driver, wait, dgr)

    # Run through all items in list and enter into requisition
    for row in range(len(df_order.index)):
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@data-control-name='icoSaveSec']"))
            ).click()
        
        item = df_order_i.loc[:, "Products"]
        amount = df_order_i.loc[:, "Ordered QTY"]
        hscode = df_order_i.loc[:, "HS Code"]  
        inst = df_order_i.loc[:, "Shipping Instructions"]
        price = df_order_i.loc[:, "Price/Unit in USD only for customs purpose"]
        origin = df_order_i.loc[:, "Origin"] 
        # If price value is missing or not a number, price is set to empty value
        try:   
            i_price = int(price[row])
        except ValueError:
            i_price = 1
        
        de.item_input(driver, row, item[row])
        
        try:
            de.qty_input(driver, row, int(amount[row]))
        except ValueError:
            pass
        
        de.unit_input(driver, wait, row)
        de.price_input(driver, row, i_price)         
        de.hs_input(driver, row, int(hscode[row]))
        de.origin_input(driver, row, str(origin[row]))
       
        # Scroll into view before clicking
        element = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@data-control-name='DataCardValue53']"))
            )
        
        driver.execute_script(
            "arguments[0].scrollIntoView();",
            element
            )
        de.enter_inst(wait, str(inst[row]))

    # Outside for loop, only filled in once
    de.enter_currency(driver, wait)

   

def fill_other(driver, wait: WebDriverWait, df_customer):
    """ Fills form data not found on the customer order form,
    and purpose of shipment, which is entered in the same section on the form """

    purpose = df_customer.loc["Purpose of order:"].item()

    de.enter_sendercountry(driver, wait)
    de.enter_costcenter(driver, wait)
    de.enter_ref_email(wait, "@")
    de.enter_purpose(driver, wait, purpose)



def process_order(driver, wait: WebDriverWait):
    """ Create dataframes from customer requisition excel file,
        containing shipping information and items ordered.
        Items ordered are split into dataframes depending on shipping temp.,
        and dangerouse goods, and stored in an array.
        process_order loops trough the array of orders and
        creates a new requisition for each order """   
      
    try:
        files = ()
        files = dialog.get_filename()
        df_customer = request.get_customer_data(files)
        df_order_arr = request.get_order_data(files)     
    except (FileNotFoundError, ValueError, TypeError) as e:
        print({str(e)})
        MB.showinfo(
            "No file selected",
            "Program will be terminated as no file was selected."
            )
        fc.kill_driver(driver)
            
        
    # Loop through all dataframes in the list of orders
    # and create a requesition for each dataframe
    i = 0
    while i < len(df_order_arr):
                
        order_data = pd.DataFrame(df_order_arr[i])
                
        # Bring up browser window, switch to iframe
        # and click button for new requisition
        driver.switch_to.window(
            driver.window_handles[-1]
            )
        fc.get_iframe(wait)
        fc.new_requesition(wait)
        
        # Enter values from files into new request form
        fill_customer(driver, wait, df_customer)
        fill_other(driver, wait, df_customer)
        fill_order(driver, wait, order_data)
       
   
        # Pause the program to allow time for manual entry
        # of attachments and missing data
        msg = MB.askyesnocancel(
            "Program paused",
            "Click YES to continue with next requisition. "
            "Click NO to close the browser."
            )

        # On 'Yes' continue with next dataframe
        # Add 1 to array position
        if(msg):
            i += 1 
            continue
        # On 'Cancel' end script but keep browser open   
        elif msg is None: 
            break
        # On 'No' end script and close browser
        else: 
            fc.kill_driver(driver)

    if i >= len(df_order_arr):
        # If there are no more dataframes in array  
        MB.showinfo(
            "Script completed",
            "All requisitions has been entered."
            )  
    else:
        MB.showwarning(
            "Script terminated",
            "Script ended before all requisitions had been entered."
            )

