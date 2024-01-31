import functions as fc
import process_order as po
from tkinter import messagebox as MB
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

    
def main():
    """ Main loop of the program.
    Choose an excel file with customer order.
    Program retrieves order data into two dataframes, containing order and shipment information.
    Data is automatically entered into shipment request forms in the CousierRequisition app """   

    try:
        driver = fc.get_app_browser()

        # Wait is defined as fluent wait - will keep polling for 10 seconds
        wait = WebDriverWait(
            driver,
            timeout=10,
            poll_frequency=2,
            ignored_exceptions=[
                NoSuchElementException,
                StaleElementReferenceException
                ]
            )

        ##Order processing
        next_order = True

        while next_order:
            po.process_order(driver, wait)
            next_order = MB.askyesnocancel(
                "Load next file",
                "Do you want to load another file? "
                "Click 'No' to end program. "
                "Click 'Cancel' to leave the browser open."
                )
            if next_order:
                continue
            elif next_order is None:
                pass
            else:
                fc.kill_driver(driver)  
            
    except Exception as e:
        err = MB.showwarning(
            "An unexpected error occured",
            "An unexpected error occured. "
            "Click OK to terminate the program."
            )
        print({str(e)})
        if(err):
            fc.kill_driver(driver)


                                    

if __name__ == '__main__':
    main()
    


  
