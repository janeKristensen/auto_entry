import country_converter as coco
import functions as fc
from phonenumbers import phonenumberutil 
from tkinter import messagebox as MB
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



def enter_customer(wait:WebDriverWait, text: str):
    """ Locates the input field for customer on the form
    and enters data retrieved from customer order """

    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue28textbox']",
            text
            )       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter customer, "
            "please correct the entry manually."
            )    


def enter_city(wait:WebDriverWait, text: str):
    """ Locates the input field for city on the form
    and enters data retrieved from customer order """

    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue30textbox']",
            text
            )      
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter city, "
            "please correct the entry manually."
            )    
    

def enter_post(wait:WebDriverWait, text: str):
    """ Locates the input field for postal code on the form
    and enters data retrieved from customer order """

    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue54textbox']",
            text
            )  
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter postal code, "
            "please correct the entry manually."
            )    
    

def enter_street(wait:WebDriverWait, text: str):
    """ Locates the input field for address on the form
    and enters data retrieved from customer order """

    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue24textbox']",
            text
            )       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter street, "
            "please correct the entry manually."
            )    


def enter_contact(wait:WebDriverWait, text: str):
    """ Locates the input field for contact person on the form
    and enters data retrieved from customer order """
    
    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue32textbox']",
            text
            )        
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter contact person, "
            "please correct the entry manually."
            )    


def enter_email(wait:WebDriverWait, text: str):
    """ Locates the input field for contact e-mail on the form
    and enters data retrieved from customer order """

    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue33textbox']",
            text
            )       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter e-mail, "
            "please correct the entry manually."
            )    


def enter_phone(wait:WebDriverWait, text: str):
    """ Locates the input field for phone number on the form
    and enters data retrieved from customer order """
    
    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue31textbox']",
            text
            )       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter phone number, "
            "please correct the entry manually."
            )    


def enter_eori(wait:WebDriverWait, text: str):
    """ Locates the input field for eori number on the form
    and enters data retrieved from customer order """
    
    try:
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue2textbox']",
            text
            )       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter EORI number, "
            "please correct the entry manually."
            )    
                                                 
    
def enter_country(driver, wait: WebDriverWait, country: str):
    """ Takes the country specified in order form and converts it, 
    to iso2 standard, and inputs in country datafield.
    Enters phone extension code based on specified country.
    Checks if country is inside EU and enters result in datafield. """ 
    
    # Xpath of label elements, used in all 3 dropdown combo-boxes
    xpath_label = "//span[@class='itemTemplateLabel_dqr75c']"

    cc = coco.CountryConverter()    
    c_iso2 = cc.convert(names=country, to='iso2', not_found=None)
        
    try: 
        # If not recognized in country converter enter the string from order form
        # and return error messagebox
        if(len(c_iso2) > 2 or c_iso2 == ""):
            raise Exception
        
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@data-control-name='ComboBox1']")
                )
            ).click()
        
        # after update of courier app it's not possible to write in country field
        # Locate input field and enter country       
        xpath_country = (
            "//input[@class='labelText_awd9vl-o_O-searchInput_10muldr"
            "-o_O-labelText_vuv5ov-o_O-label_fn3faz-o_O-"
            "inputTop_11b8lle mousetrap block-undo-redo']"
            )
        # send_keys_to_xpath(wait, xpath_country, c_iso2)

        # Click on the country in fly-out combobox to select it
        xpath_country_iso2 = "//span[text()='" + c_iso2 + "']"

        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, xpath_label)
                )
            )
        
        elem_s = driver.find_element(By.XPATH, xpath_country_iso2)
        driver.execute_script(
            "arguments[0].scrollIntoView();",
            elem_s
        )
         
        elem_c = driver.find_element(By.XPATH, xpath_label) 
        elem_c.find_element(
            By.XPATH, xpath_country_iso2
            ).click()              
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter country, calling code and EU-country. "
            "Please correct the entry manually."
            )

    try:
        # Find country code based on the selected country
        ccphone = phonenumberutil.country_code_for_region(c_iso2)
    
        # Locate input field for phone extension and input data
        str_phone = "+" + str(ccphone)
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue14textbox']",
            str_phone
            )
    except:
       pass
        
    try:
        # Locate input field "Outside EU" and use country converter
        # to check if country is inside EU
        xpath_eu = "//div[@data-control-name='DataCardValue45']"
        
        iso3 = cc.convert(
            names=country, to='iso3',
            not_found='not found'
            )

        if iso3 != 'not found':
            df_eu = cc.EU27as('ISO3')
            
            if iso3 in df_eu.values:
                # Set input field to "No" for country inside EU
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, xpath_eu)
                        )
                    ).click()
                
                elem_eu = driver.find_element(By.XPATH, xpath_label)
                elem_eu.find_element(
                    By.XPATH, "//span[text()='No']"
                    ).click()
                
            else:
                # Set input to "Yes" for country outside EU
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, xpath_eu)
                        )
                    ).click()
                
                elem = driver.find_element(By.XPATH, xpath_label)
                elem.find_element(
                    By.XPATH,
                    "//span[text()='Yes']"
                    ).click()
                
        else:
            raise Exception        
    except:
        pass

                       
def enter_purpose(driver, wait: WebDriverWait, purpose: str):
    """ Locates the input field for shipment purpose on the form
    and enters data retrieved from customer order """
    #Currently hardcoded to enter 'Analysis'
    purpose = "Analysis"
    
    try:
        xpath_p = "//div[@data-control-name='DataCardValue41']"
        
        # Scroll into view
        elem_p = driver.find_element(By.XPATH, xpath_p)
        
        driver.execute_script(
            "arguments[0].scrollIntoView();",
            elem_p
            )
          
        # Open combobox and choose "others" to be able to enter text.
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, xpath_p)
                )
            ).click()

        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[text()='Others']")
                )
            ).click()
       
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@appmagic-control='txtOtherstextbox']")
                )
            ).send_keys(purpose) 
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter purpose of shipment, "
            "please correct the entry manually."
            )    
   
   
def enter_sendercountry(driver, wait: WebDriverWait):
    """ Locates the input field for sender_country on the form
    and enters 'DK' as default"""
    
    try:
        xpath_s = "//div[@data-control-name='DataCardValue51']"
        #Locate the combobox and scroll into view.  
        elem_s = driver.find_element(By.XPATH, xpath_s)
        
        driver.execute_script("arguments[0].scrollIntoView();", elem_s)
 
        #Click and select DK from the pop up menu.
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, xpath_s)
                )
            ).click()

        elem_s.find_element(By.XPATH, "//li[@data-index='0']").click()
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter sender country, "
            "please correct the entry manually."
            )    
          

def enter_ref_email(wait:WebDriverWait, text: str):
    """ Locates the input field for sender e-amil on the form
    and enters the reference substances shared e-mail """
    
    try:
        # Locate input field and send data from form
        fc.send_keys_to_xpath(
            wait,
            "//input[@appmagic-control='DataCardValue34textbox']",
            text
            )
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter sender group e-mail, "
            "please correct the entry manually."
            )    


def temp_input(wait: WebDriverWait, temp: str):
    """ Locates the input field for shipment temperature on the form
    and enters data retrieved from customer order """
    
    try:
        # Locate drop down menu and click to open       
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@data-control-name='DataCardValue25']")
                )
            ).click()

        # Choose shipping temperature from list
        if(temp == "Ambient"):
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[text()='Ambient (No Requirement)']")
                    )
                ).click()            
        elif(temp == "2-8°C"):
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[text()='Cold +2C to +8C']")
                    )
                ).click()            
        elif(temp == "Below -15°C"):
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[text()='Frozen -25C to -15C']")
                    )
                ).click()           
        elif(temp == "15-25°C"):
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[text()='Room Temp. +15C to +25C']")
                    )
                ).click()  
        else:
            MB.showwarning(
                "An error occured",
                "Program was unable to enter shipping temperature, "
                "please correct the entry manually."
                )
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter shipping temperature, "
            "please correct the entry manually."
            )
        

def dgr_input(driver, wait: WebDriverWait, dgr: str):
    """ Locates the input field for DGR on the form
    and enters data retrieved from customer order """
    
    try:
        # Locate drop down menu and click to open
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@data-control-name='DataCardValue1']"))
            ).click()

        # Make a list of all comboboxes
        # Select span from index 5 in this list
        flyout_list = driver.find_elements(
            By.XPATH,
            "//div[@class='powerapps-flyout-content']"
            )

        # Click "Yes" or "No" depending status
        if(dgr == "X"):
            flyout_list[0].find_element(
                By.XPATH,
                "//li[@data-index='0']"
                ).click()  
        else:
            flyout_list[0].find_element(
                By.XPATH,
                "//li[@data-index='1']"
                ).click()     
    except:
        MB.showwarning(
            "An error occured",
            "An error occured while entering DGR classification, "
            "please select manually."
            )
        

def enter_costcenter(driver, wait: WebDriverWait):
    """ Locates the input field for costcenter on the form
    and enters 'cost center' """
    
    try:
        # Locate the textbox and scroll into view before sending keys.
        elem_cc = driver.find_element(
            By.XPATH,
            "//input[@appmagic-control='DataCardValue40textbox']"
            )
        
        driver.execute_script(
            "arguments[0].scrollIntoView();",
            elem_cc
            )
        
        elem_cc.send_keys("cost center")        
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter cost center, please correct the entry manually."
            )    

    
def enter_currency(driver, wait: WebDriverWait):
    """ Locates the input field for currency on the form
    and enters 'USD' which is used as default"""
    
    try:
        # Locate the combobox and scroll into view.
        xpath_currency = "//div[@data-control-name='DataCardValue3']"       
        element = driver.find_element(
            By.XPATH,
            xpath_currency
            )
        
        driver.execute_script(
            "arguments[0].scrollIntoView();",
            element
            )

        # Click and select USD from the pop up menu.
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, xpath_currency))
            ).click()
        
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[text()='USD']"))
            ).click()   
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter currency, "
            "please correct the entry manually."
            )
        

def enter_inst(wait: WebDriverWait, inst: str):
    """ Locates the input field for shipping instructions on the form
    and enters data from customer order form """    

    try:
        # Send whitespace if str is blank
        if inst == "nan" or inst == 'None':
            inst = ''

        fc.send_keys_to_xpath(
            wait,
            "//textarea[@appmagic-control='DataCardValue53textarea']",
            inst
            )        
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter instructions, "
            "please correct the entry manually."
            )
   

def item_input(driver, x, item: str):
    """ Locates the input field for shipment items on the form
    and enters data from customer order dataframe one line at a time """

    try:
        element = fc.get_gallery_row(driver, x)
       
        # Get input field
        prods = element.find_elements(
            By.XPATH,
            "//input[@appmagic-control='txtItemUsagetextbox']"
            )

        # Input data from form 
        fc.action_chain(driver, prods[x], item)    
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter item name, "
            "please correct the entry manually."
            )


def qty_input(driver, x, amount: int):
    """ Locates the input field for shipment items on the form
    and enters quantity data from customer order dataframe one line at a time """

    try:
        element = fc.get_gallery_row(driver, x)

        # Get input field
        qtys = element.find_elements(
            By.XPATH,
            "//input[@appmagic-control='txtItemQtytextbox']"
            )

        # Input data from form
        fc.action_chain(driver, qtys[x], amount)       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter quantity, "
            "please correct the entry manually."
            )


def origin_input(driver, x, origin: str):
    """ Locates the input field for shipment items on the form
    and enters origin data from customer order dataframe one line at a time.
    Only Segrate samples have a different country of origin """

    try:
        element = fc.get_gallery_row(driver, x)

        # Get input field
        origin_field = element.find_elements(
            By.XPATH,
            "//input[@appmagic-control='countryoforigintextbox']"
            )

        # Input data from form
        fc.action_chain(driver, origin_field[x], origin)       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter country of origin, "
            "please correct the entry manually.")


def unit_input(driver, wait: WebDriverWait, x):
    """ Locates the input field for units on the form
    and enters 'EA' as default """

    try:
        element = fc.get_gallery_row(driver, x)

        # Click the drop down menu and select "EA"
        units = element.find_elements(
            By.XPATH,
            "//div[@data-control-name='Dropdown5']"
            )
        
        units[x].click()
        
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@class='item3 appmagic-dropdownListItem']")
                )
            ).click()        
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter unit, "
            "please correct the entry manually."
            )
        

def price_input(driver, x, price: int):
    """ Locates the input field for price on the form
    and enters data from customer order dataframe one line at a time.
    Enters the value 1 if nothing else specified """

    try:
        element = fc.get_gallery_row(driver, x)

        # Get input field
        prices = element.find_elements(
            By.XPATH,
            "//input[@appmagic-control='txtPricePerItemtextbox']"
            )

        # Input data from form
        fc.action_chain(driver, prices[x], price)       
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter price, "
            "please correct the entry manually."
            )


def hs_input(driver, x, hscode: int):
    """ Locates the input field for HScode on the form
    and enters data from customer order dataframe one line at a time """

    try:
        element = fc.get_gallery_row(driver, x)

        # Get input field
        hs = element.find_elements(
            By.XPATH,
            "//input[@appmagic-control='txtHSCodetextbox']"
            )

        # Input data from form
        fc.action_chain(driver, hs[x], hscode)    
    except:
        MB.showwarning(
            "An error occured",
            "Program was unable to enter HS code, "
            "please correct the entry manually."
            )
