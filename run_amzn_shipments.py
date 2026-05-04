import os
import time
import shutil
import ctypes
import traceback
import xlwings as xw
from rich import print
from dotenv import load_dotenv
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import chrome, accounts, custom_functions, outlook
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException

###############################################################################################################################################
#Get the user and working directory
directory: str = os.getcwd()
win_user: str = os.getlogin()

#Get Seller Central credentials from environment
load_dotenv()
username: str = os.getenv("AMZN_email")
password: str = os.getenv("AMZN_pass")

#Set Chrome User Data Directory
user_data_dir: str = f"C:/ChromeAutomationProfile"

#Get all the Amazon accounts
Accounts = accounts.Amazon()

##################################################################################################################################################
# Create the body of the email
body = """
Good morning,<br><br>
Please find attached the Shipments report updated for today.<br><br>
If any questions, please let me know.<br><br>
Thanks,<br><br>
"""

##################################################################################################################################################
def seconds_until_target(TargetTime: str):
    #Calculate the number of seconds until the target time
    now = datetime.now()
    TargetTime = datetime.strptime(TargetTime, "%H:%M:%S").replace(year=now.year, month=now.month, day=now.day)

    if TargetTime < now:
        TargetTime += timedelta(days=1)

    return (TargetTime - now).total_seconds()

##################################################################################################################################################
def ShouldRun() -> bool:
    #Check if today is Tuesday
    today: str = datetime.now().strftime("%A")

    return today in ["Tuesday"]

###############################################################################################################################################
#Ask the user if they want to start the process now
BtnPressed = ctypes.windll.user32.MessageBoxW(
    0,
    "Do you want to start the script now?",
    "Amazon Shipments",
    4 | 0x20
)

while True:
    #Time to start
    StartTime = "09:00:00"
    StartHour = int(StartTime.split(":")[0])
    StartMin = StartTime.split(":")[1]

    if ShouldRun():
        SleepTime = seconds_until_target(StartTime)
        nowHour = int(datetime.now().strftime("%H"))

        #If the user pressed "Yes", then start the process
        if BtnPressed == 7:
            if nowHour >= StartHour:
                print(f"'[INFO]' Shipment updates will be worked on next Tuesday at {StartHour}:{StartMin} AM.")
            else:
                print(f"'[INFO]' Shipment updates will be worked today at {StartHour}:{StartMin} AM.")

            #Sleep until just before the Start time
            time.sleep(max(SleepTime - 1, 0))

            #Loop to ensure that we catch the exact time
            while datetime.now().strftime("%H:%M:%S") != StartTime:
                time.sleep(0.5)

            #Get today's name
            today: str = datetime.now().strftime("%A")

            if today != "Tuesday":
                continue

        #Reset the value of the button
        BtnPressed = 7

        ###############################################################################################################################################
        #Initialize Chrome
        opening_browser = True
        while opening_browser:
            try:
                driver: object = chrome.start_browser(
                    user_data_dir,
                    "Default",
                    headless=True
                )
                opening_browser = False

            except (SessionNotCreatedException, RuntimeError):
                print("'[ERROR]' Failed to open the Chrome. It seems Chrome was already open. Killing the application and retrying.")
                custom_functions.kill_app("chrome")
                time.sleep(5)

            except PermissionError:
                print("'[ERROR]' Failed to open the Chrome. It seems Chrome was already open. Killing the application and retrying.")
                custom_functions.kill_app("uc_driver")
                time.sleep(5)

        ###############################################################################################################################################
        for account, url in Accounts.items():
            match account:
                case "FocusCam":
                    root = "SellerOrg Corp"
                case "LifeS":
                    root = "Lifestyle By Focus"
                case "XtraB":
                    root = "XtraBargains"
                case "KnoxGear":
                    root = "Knox Gear"
                case "Apple":
                    root = "Apple Renewed Focus"
                case "FocusHome":
                    root = "Focus Home"

            #Delete all files in the root directory for each account
            print(f"'[INFO]' Removing all files in the root directory for '{root}'.")
            for file in os.listdir(f"{directory}/Amazon/Shipments/{root}/"):
                os.remove(os.path.join(f"{directory}/Amazon/Shipments/{root}/", file))

            print(f"'[INFO]' Navigating to '{root}' account.")
            driver.get(url)
            driver.switch_to_window(0)

            try:
                code = None
                while not code:
                    code = accounts.Amazon_login(driver, username, password)

                    if not code:
                        print("'[ERROR]' Failed to log in to Amazon. Trying again.")
                        driver.get(url)
                        driver.switch_to_window(0)

            except TimeoutException:
                pass

            print("'[INFO]' Getting shipments.")
            driver.get("https://sellercentral.amazon.com/gp/ssof/shipping-queue.html/ref=xx_fbashipq_favb_xx#fbashipment")
            time.sleep(2)
            driver.switch_to_window(0)

            #If "Feedback" banner pops-up, close it
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "#vibes-close-button"
                ))).click()
            except TimeoutException:
                pass

            #Increase the range of results per page from 25 to 100
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "#pagination-dropdown"
                ))).click()

                custom_functions.shadow_element(
                    driver,
                    "#pagination-dropdown > kat-dropdown",
                    "div > div:nth-child(3) > div > div > div > slot:nth-child(2) > kat-option:nth-child(4)"
                )

                time.sleep(3)
            except TimeoutException:
                print(f"'[INFO]' No results on '{root}' account. Moving to next account.")
                continue

            #TODO: USE THE NUMBER FOR THE DESIRED RANGE
            #* For "All" = 1
            #* For "Within 24 hours" = 2
            #* For "Within 1 week" = 3
            #* For "Within 30 days" = 4
            #* For "Within 90 days" = 5
            #* For "Within 1 year" = 6
            #* For "Custom date range" = 7

            #Set the value for the desired range
            within_range = 6
            match within_range:
                case 1:
                    range_value = "ALL" #All
                case 2:
                    range_value = "1" #Within 24 hours
                case 3:
                    range_value = "7" #Within 1 week
                case 4:
                    range_value = "30" #Within 30 days
                case 5:
                    range_value = "90" #Within 90 days
                case 6:
                    range_value = "365" #Within 1 year
                case 7:
                    range_value = "CUSTOM" #Custom date range

            try:
                #Select the desired range
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "input[type='text'][value='Last updated']"
                ))).click()

                #Get the parent element of the radio button
                parent = driver.find_element(
                    By.CLASS_NAME,
                    "date-filter-radio-inline"
                )

                #Find the input element for the radio button by its value or label
                radio_input = parent.find_element(
                    By.CSS_SELECTOR,
                    f"input[type='radio'][value='{range_value}']"
                )

                #Find the clickable icon *after* the input - sibling span
                #and click it to select the radio button
                radio_input.find_element(
                    By.XPATH,
                    "./following-sibling::span[contains(@class, 'kat-radiobutton-icon')]"
                ).click()
                time.sleep(2)

            except TimeoutException:
                #Take a screenshot if the element is not found
                print("'[TimeoutException]' Element not found.")
                driver.save_screenshot(f"Shipments error.png")
                driver.quit()
                quit()

            try:
                #Select only orders with "Closed" status
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR,
                    "input[type='text'][value='Status']"
                ))).click()

                parent = driver.find_element(
                    By.CLASS_NAME,
                    "shipment-status-filter-inline"
                )

                parent.find_element(
                    By.CSS_SELECTOR,
                    "kat-checkbox[label='Closed'][value='CLOSED']"
                ).click()
                time.sleep(2)

            except TimeoutException:
                #Take a screenshot if the element is not found
                print("'[TimeoutException]' Element not found.")
                driver.save_screenshot(f"Shipments error.png")
                driver.quit()
                quit()

            page = 1
            getting_shipments = True
            while getting_shipments:
                #Download the table data
                print("'[INFO]' Downloading file.")
                driver.find_element(
                    By.CSS_SELECTOR,
                    "#export-link-container > a"
                ).click()

                #Wait for the file to download
                time.sleep(5)

                #Extract the downloaded file and move it to the corresponding location
                for file in os.listdir(f"{directory}/downloaded_files/"):
                    if file.endswith("Z.csv"):
                        print(f"'[INFO]' File '{file}' downloaded.")
                        shutil.move(
                            f"{directory}/downloaded_files/{file}",
                            f"{directory}/Amazon/Shipments/{root}/Shipments - Page {page}.csv"
                        )
                        break

                #Move to the next page
                try:
                    NextPageBtn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                        By.CSS_SELECTOR,
                        "#tab-view > div:nth-child(3) > div > div.paging.flex-row > div:nth-child(4) > a"
                    )))

                    #Use Javascript to scroll down to the next page button
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", NextPageBtn)
                    NextPageBtn.click()

                    time.sleep(2)
                    page += 1

                except TimeoutException:
                    print("'[INFO]' No more pages to download.")
                    getting_shipments = False

        #Close Firefox
        driver.quit()

        ###############################################################################################################################################
        #Get today's date
        Date = datetime.now().strftime("%m/%d/%Y")
        
        #Update the queries in the "Shipment" workbook
        print("'[INFO]' Updating queries in the 'Shipment' workbook.")
        ShipmentWbPath = f"{directory}/Amazon/Reports/Shipments.xlsm"
        ShipmentWb = xw.Book(ShipmentWbPath)
        DataValSh = ShipmentWb.sheets("DataVal")
        custom_functions.update_directory(ShipmentWb)

        #Get the latest workbook Status
        status = DataValSh.range("B2").value
        send = False if status == "Sent" else True

        #Update the status in the workbook
        if send:
            DataValSh.range("B2").value = "Sent"
        else:
            DataValSh.range("B2").value = "Not Sent"

        #Refresh the queries, save and close the workbook
        Refresh = ShipmentWb.macro("Module1.Refresh")
        Refresh()
        time.sleep(30)

        print("'[INFO]' Saving and closing the workbook.")
        ShipmentWb.save()
        ShipmentWb.close()
        time.sleep(60)

        if send:
            #Send email notification
            print("'[INFO]' Sending email.")
            outlook.send_email(
                account="user@example.com",
                subject=f"Shipments Report - {Date}",
                body=body,
                to=["user@example.com", "user@example.com"],
                cc=["user@example.com"],
                attachments=[ShipmentWbPath],
                show=True,
                send=True
            )

            print("'[INFO]' Email has been sent.")

    #Sleep 60 seconds before starting over
    time.sleep(60)