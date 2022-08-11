import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from tempfile import gettempdir
from pathlib import Path
from random import random
from datetime import datetime, date
import sheets


def query_sheet(path):
    """
    Wait for a sheet to be downloaded
    """

    while True:
        files = os.listdir(path)

        if len(files) == 1 and files[0].endswith(".xlsx"):
            return os.path.join(path, files[0])
        elif len(files) > 1:
            raise ValueError


def clear_files(path):
    """
    Remove all files from a directory
    """
    files = os.listdir(path)

    for f in files:
        os.remove(os.path.join(path, f))


def get_warehouse(driver):
    """
    Determine the warehouse
    """

    wh_strip = lambda s: "".join(c for c in s if c.isalpha())
    for elem in driver.find_elements(By.CSS_SELECTOR, "b"):
        if "green" in elem.get_attribute("style"):
            warehouse = wh_strip(elem.get_attribute("innerText").strip())
            if any(wh in warehouse for wh in sheets.WAREHOUSE_IDS):
                return warehouse
    warehouse = wh_strip(driver.find_elements(By.CSS_SELECTOR, "tbody")[2].find_elements(By.CSS_SELECTOR, "td")[1].get_attribute("innerText"))
    if any(wh in warehouse for wh in sheets.WAREHOUSE_IDS):
        return warehouse
    raise ValueError("Could not find warehouse")


def get_carrier(driver, carrier_i):
    """
    Determine the carrier
    """

    carrier = driver.find_elements(By.CSS_SELECTOR, "tbody")[2].find_elements(By.CSS_SELECTOR, "td")[carrier_i].get_attribute("innerText")

    try:
        return carrier[carrier.index("[") + 1:carrier.index("]")]
    except ValueError:
        return "UNKWN"


def scrape(driver, data, wait, site, cmd_options, temp_dir, username, password):
    """
    Scrape one website and add all the data to `data`
    """

    driver.get(f"https://www2.order-fulfillment.bz/{site}/reports")

    wait.until(expected_conditions.presence_of_element_located((By.ID, "btnLogin")))
    driver.find_element(By.NAME, "LoginId").send_keys(username)
    driver.find_element(By.NAME, "Password").send_keys(password)
    driver.find_element(By.ID, "btnLogin").click()

    wait.until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, "#btnPendingShipment"))).click()
    clear_files(temp_dir)
    sheet_path = query_sheet(temp_dir)
    order_nums = sheets.extract_order_nums(sheet_path)

    total = len(order_nums)

    for i, order_num in enumerate(order_nums):
        if cmd_options["show_progress"]:
            print(f"{i + 1}/{total}")

        url = f"https://www2.order-fulfillment.bz/{site}/orders"
        order_url = f"{url}/{order_num}/manage"

        driver.get(order_url)
        wait.until(lambda driver: driver.current_url == order_url)

        tbodys = driver.find_elements(By.CSS_SELECTOR, "tbody")
        while len(tbodys) < 4:
            tbodys = driver.find_elements(By.CSS_SELECTOR, "tbody")
        tbody = tbodys[3]

        items = []

        order_time = None
        po_num = None

        for elem in driver.find_elements(By.CSS_SELECTOR, "strong"):
            inner = elem.get_attribute("innerText")

            if order_time is None and inner == "Order Date":
                order_time = driver.execute_script("return arguments[0].nextSibling", elem)
                order_time = datetime.strptime(order_time["textContent"], " - %m/%d/%Y")
            elif po_num is None and inner == "PO #":
                po_num = driver.execute_script("return arguments[0].nextSibling", elem)

            if po_num is not None and order_time is not None:
                break
        else:
            raise ValueError

        warehouse = get_warehouse(driver)
        po = po_num["textContent"][3:]
        status = "Not Shipped"
        ship_status = sheets.get_ship_status(order_time, status)

        headings = driver.find_elements(By.CSS_SELECTOR, "thead")[3].find_element(By.CSS_SELECTOR, "tr").find_elements(By.CSS_SELECTOR, "th")
        headings2 = driver.find_elements(By.CSS_SELECTOR, "thead")[2].find_element(By.CSS_SELECTOR, "tr").find_elements(By.CSS_SELECTOR, "th")

        num_index = None
        qty_index = None
        carrier_i = None

        for i, heading in enumerate(headings):
            text = heading.get_attribute("innerText")

            if text == "Item Name":
                num_index = i
            elif text == "Quanity":
                qty_index = i

            if num_index is not None and qty_index is not None:
                break
        else:
            raise ValueError

        for i, heading in enumerate(headings2):
            text = heading.get_attribute("innerText")

            if text == "Tracking #":
                carrier_i = i

            if carrier_i is not None:
                break
        else:
            raise ValueError

        carrier = get_carrier(driver, carrier_i)

        for trow in tbody.find_elements(By.CSS_SELECTOR, "tr"):
            item_data = trow.find_elements(By.CSS_SELECTOR, "td")

            if not item_data:
                break

            num_td = item_data[num_index]
            qty_td = item_data[qty_index]
            num = num_td.get_attribute("innerText")
            qty = qty_td.get_attribute("innerText")

            items.append((num, qty))

        data.append((
            po,
            carrier,
            status,
            warehouse,
            ship_status,
            items,
        ))

        if cmd_options["debug"]:
            print(data[-1])

        if cmd_options["pause"]:
            while True: pass


def parse_cmd_options():
    """
    Parse options passed from the command line
    """

    opts = {
        "show_progress": False, # show how much total progress has been made by printing to stdout
        "non_headless": False,  # don't operate in headless mode (show browser window)
        "pause": False,         # pause on the first order (for inspecting)
        "debug": False,         # debug the data
    }

    for opt in sys.argv:
        if opt in opts:
            opts[opt] = True

    return opts


def main():
    cmd_options = parse_cmd_options()
    temp_dir = os.path.join(gettempdir(), f"scraped_report{random()}")
    Path(temp_dir).mkdir(parents=True, exist_ok=False)

    capabilities = DesiredCapabilities().CHROME
    capabilities["pageLoadStrategy"] = "none"

    options = webdriver.ChromeOptions()
    options.add_argument("--log-level=3")

    if not cmd_options["non_headless"]:
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--headless")

    options.add_experimental_option("prefs", { "download.default_directory": temp_dir })

    driver = webdriver.Chrome("C:/Selenium/chromedriver.exe", chrome_options=options, desired_capabilities=capabilities)
    wait = WebDriverWait(driver, 10)
    data = []

    username = input("Username:\n")
    password = input("Password:\n")

    scrape(driver, data, wait, "homebeyond", cmd_options, temp_dir, username, password)
    scrape(driver, data, wait, "vanityart",  cmd_options, temp_dir, username, password)

    class_lookup = sheets.load_class_lookup("class_lookup.xlsx")
    combo_lookup = sheets.load_combo_lookup("combo_lookup.xlsx")
    warehouses = sheets.input_warehouses()
    output_data = sheets.parse_data(data, warehouses, class_lookup, combo_lookup)

    sheets.write_data(output_data, f"scraped_{date.today()}.xlsx")


if __name__ == "__main__":
    main()
