import os
import time
import shutil
import traceback
import xlwings as xw
from rich import print
from dotenv import load_dotenv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fc_utils import chrome, accounts, custom_functions, outlook, alert_utils
from fc_utils.config_utils import get_env
from fc_utils.schedule_utils import run_on_schedule
from fc_utils.accounts import AMAZON_ACCOUNT_NAMES
from selenium.common.exceptions import TimeoutException

directory: str = os.getcwd()

load_dotenv()
username: str = os.getenv("AMZN_email")
password: str = os.getenv("AMZN_pass")
sender_email: str = os.getenv("SENDER_EMAIL", "")
to_email: list[str] = [e.strip() for e in os.getenv("TO_EMAIL", "").split(",") if e.strip()]
cc_email: list[str] = [e.strip() for e in os.getenv("CC_EMAIL", "").split(",") if e.strip()]
user_data_dir: str = get_env("CHROME_USER_DATA_DIR", required=True)

amzn_accounts = accounts.Amazon()
shipment_wb_path: str = f"{directory}/Amazon/Reports/Shipments.xlsm"

# within_range selects the date filter applied to shipments:
# 1=All  2=24h  3=1 week  4=30 days  5=90 days  6=1 year  7=Custom
within_range: int = 6
_range_map = {1: "ALL", 2: "1", 3: "7", 4: "30", 5: "90", 6: "365", 7: "CUSTOM"}
range_value: str = _range_map[within_range]

body = """
Good morning,<br><br>
Please find attached the Shipments report updated for today.<br><br>
If any questions, please let me know.<br><br>
Thanks,<br><br>
"""


def main() -> None:
    """Download Amazon shipment CSVs for each account and email the refreshed report.

    Scrapes closed shipments within the configured date range, saves each page
    as a CSV, refreshes the Shipments workbook, and emails the report if it
    has not already been sent today.
    """
    driver = None
    try:
        driver = chrome.start_browser(user_data_dir, "Default", headless=True)

        for account, url in amzn_accounts.items():
            root = AMAZON_ACCOUNT_NAMES[account]

            print(f"[cyan][INFO][/cyan] Removing all files in the shipments folder for [cyan]{root}[/cyan].")
            for file in os.listdir(f"{directory}/Amazon/Shipments/{root}/"):
                os.remove(os.path.join(f"{directory}/Amazon/Shipments/{root}/", file))

            print(f"[cyan][INFO][/cyan] Navigating to [cyan]{root}[/cyan] account.")
            driver.get(url)
            driver.switch_to_window(0)

            try:
                code = None
                while not code:
                    code = accounts.amazon_login(driver, username, username, password)
                    if not code:
                        print("[bold red][ERROR][/bold red] Failed to log in to Amazon. Trying again.")
                        driver.get(url)
                        driver.switch_to_window(0)
            except TimeoutException:
                pass

            print("[cyan][INFO][/cyan] Getting shipments.")
            driver.get("https://sellercentral.amazon.com/gp/ssof/shipping-queue.html/ref=xx_fbashipq_favb_xx#fbashipment")
            time.sleep(2)
            driver.switch_to_window(0)

            # Dismiss the feedback banner if present
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "#vibes-close-button"
                ))).click()
            except TimeoutException:
                pass

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "#pagination-dropdown"
                ))).click()
                custom_functions.shadow_element(
                    driver,
                    "#pagination-dropdown > kat-dropdown",
                    "div > div:nth-child(3) > div > div > div > slot:nth-child(2) > kat-option:nth-child(4)"
                )
                time.sleep(3)
            except TimeoutException:
                print(f"[cyan][INFO][/cyan] No results on [cyan]{root}[/cyan] account. Moving to next account.")
                continue

            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "input[type='text'][value='Last updated']"
                ))).click()

                parent = driver.find_element(By.CLASS_NAME, "date-filter-radio-inline")
                radio_input = parent.find_element(
                    By.CSS_SELECTOR,
                    f"input[type='radio'][value='{range_value}']"
                )
                # Click the sibling span — the input itself is not directly clickable
                radio_input.find_element(
                    By.XPATH,
                    "./following-sibling::span[contains(@class, 'kat-radiobutton-icon')]"
                ).click()
                time.sleep(2)

            except TimeoutException:
                print("[bold red][TimeoutException][/bold red] Date filter element not found.")
                driver.save_screenshot("Shipments error.png")
                raise RuntimeError("Critical element not found; aborting.")

            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "input[type='text'][value='Status']"
                ))).click()

                parent = driver.find_element(By.CLASS_NAME, "shipment-status-filter-inline")
                parent.find_element(
                    By.CSS_SELECTOR,
                    "kat-checkbox[label='Closed'][value='CLOSED']"
                ).click()
                time.sleep(2)

            except TimeoutException:
                print("[bold red][TimeoutException][/bold red] Status filter element not found.")
                driver.save_screenshot("Shipments error.png")
                raise RuntimeError("Critical element not found; aborting.")

            page = 1
            getting_shipments = True
            while getting_shipments:
                print("[cyan][INFO][/cyan] Downloading file.")
                driver.find_element(By.CSS_SELECTOR, "#export-link-container > a").click()
                time.sleep(5)

                for file in os.listdir(f"{directory}/downloaded_files/"):
                    if file.endswith("Z.csv"):
                        print(f"[cyan][INFO][/cyan] File [cyan]{file}[/cyan] downloaded.")
                        shutil.move(
                            f"{directory}/downloaded_files/{file}",
                            f"{directory}/Amazon/Shipments/{root}/Shipments - Page {page}.csv"
                        )
                        break

                try:
                    next_page_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                        By.CSS_SELECTOR,
                        "#tab-view > div:nth-child(3) > div > div.paging.flex-row > div:nth-child(4) > a"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_page_btn)
                    next_page_btn.click()
                    time.sleep(2)
                    page += 1
                except TimeoutException:
                    print("[cyan][INFO][/cyan] No more pages to download.")
                    getting_shipments = False

        driver.quit()

        date_str: str = datetime.now().strftime("%m/%d/%Y")

        print("[cyan][INFO][/cyan] Updating queries in the [cyan]Shipment[/cyan] workbook.")
        shipment_wb = xw.Book(shipment_wb_path)
        data_val_sh = shipment_wb.sheets("DataVal")
        custom_functions.update_directory(shipment_wb)

        status = data_val_sh.range("B2").value
        send = status != "Sent"

        data_val_sh.range("B2").value = "Sent" if send else "Not Sent"

        refresh = shipment_wb.macro("Module1.Refresh")
        refresh()
        time.sleep(30)

        print("[cyan][INFO][/cyan] Saving and closing the workbook.")
        shipment_wb.save()
        shipment_wb.close()
        time.sleep(60)

        if send:
            print("[cyan][INFO][/cyan] Sending email.")
            outlook.send_email(
                account=sender_email,
                subject=f"Shipments Report - {date_str}",
                body=body,
                to=to_email,
                cc=cc_email,
                attachments=[shipment_wb_path],
                show=True,
                send=True
            )
            print("[cyan][INFO][/cyan] Email has been sent.")

    except (KeyboardInterrupt, SystemExit):
        print("[yellow][WARNING][/yellow] Script interrupted by user.")
        raise SystemExit(0)

    except Exception:
        alert_utils.handle_crash(driver, traceback.format_exc(), "Shipments")
        raise SystemExit(1)

    finally:
        try:
            driver.quit()
        except Exception:
            pass


run_on_schedule(main, hour=9, minute=0, day_of_week="tue")
