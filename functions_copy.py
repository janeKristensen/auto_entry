import os
import sys
from tkinter import messagebox as MB
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_app_browser():
    """ Start the browser driver and load the requesition webpage.
    Return the driver to main script """

    user = os.path.expanduser('~')
    user_dir = "user-data-dir=" + user + "\\AppData\\Local\\Microsoft\\Edge\\User Data\\Selenium"
    
    # Open a driver and detach to prevent closing at script end
    dr_options = webdriver.EdgeOptions()
    dr_options.add_experimental_option("detach", True)
    dr_options.add_argument(user_dir)
    driver = webdriver.Edge(options=dr_options)

    # Open CourierRequesition app in browser and maximize the window
    url = ("url")
    driver.get(url)
    driver.maximize_window()
    driver.switch_to.default_content()
    
    return driver


def kill_driver(driver):
    driver.quit()
    sys.exit()

    
def get_iframe(wait: WebDriverWait):
    """ CourierRequisition app is inside an iframe which needs to be selected
    before elements of the webpage can be located """
    
    wait.until(
        lambda d: d.execute_script(
            "return document.readyState"
            ) == "complete"
        )
    
    iframe = wait.until(
        lambda d: d.find_element(
            By.XPATH,"//iframe[@id='fullscreen-app-host']"
            ).get_attribute("id")
        )
   
    wait.until(
        EC.frame_to_be_available_and_switch_to_it(
            (By.ID, iframe))
        )


def new_requesition(wait: WebDriverWait):
    """ Clicks on the "New Requisition" button to start a new shipment request """
    
    # Throws exception if switch to frame didn't happen
    try:
        xpath_i = "//div[@data-control-name='icoAddCR']"
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, xpath_i))
            )
        
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, xpath_i))
            ).click()        
    except:
        error = MB.showerror(
            "An error occured",
            "Browser was not able to locate iframe."
            "Click OK to terminate the program."
            )
        
        return error


def send_keys_to_xpath(wait:WebDriverWait, xpath: str, text: str):
    """ Utility function for sending text to elements on the form """
    
    # Locate input field by xpath and send a string value 
    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, xpath)
            )
        ).send_keys(text)


def get_gallery_row(driver, x: int):
 # Move into the item gallery control
    gallery = driver.find_element(
        By.XPATH,
        "//div[@class='virtualized-gallery']"
        )
        
    gallery.click()
    
    # Identify the row number    
    row_str = str(int(x + 1))
    element = gallery.find_element(
        By.XPATH,
        "//div[@aria-posinset='" + row_str + "']"
        )
    return element


def action_chain(driver, row, input):
    # Input data from form into row field
    actions = ActionChains(driver)
    actions.double_click(row)
    actions.send_keys(input)
    actions.perform()
    actions.reset_actions()  




