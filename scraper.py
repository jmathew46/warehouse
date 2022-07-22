import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from tempfile import gettempdir
from pathlib import Path
from random import random
from datetime import datetime
import sheets


def query_sheet(path):
    while True:
        files = os.listdir(path)

        if len(files) == 1 and files[0].endswith(".xlsx"):
            return os.path.join(path, files[0])
        elif len(files) > 1:
            raise ValueError


def main():
    progress = len(sys.argv) > 1 and sys.argv[1] == "prog"
    temp_dir = os.path.join(gettempdir(), f"scraped_report{random()}")
    Path(temp_dir).mkdir(parents=True, exist_ok=False)

    options = webdriver.ChromeOptions()
    options.add_argument("--log-level=3")
    options.add_experimental_option("prefs", { "download.default_directory": temp_dir })

    driver = webdriver.Chrome("C:/Selenium/chromedriver.exe", chrome_options=options)
    wait = WebDriverWait(driver, 10)
    username = input("Username:\n")
    password = input("Password:\n")

    driver.get("https://www2.order-fulfillment.bz/vanityart/reports")
    driver.find_element(By.NAME, "LoginId").send_keys(username)
    driver.find_element(By.NAME, "Password").send_keys(password)
    driver.find_element(By.ID, "btnLogin").click()

    wait.until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, "#btnPendingShipment"))).click()
    sheet_path = query_sheet(temp_dir)
    order_nums = sheets.extract_order_nums(sheet_path)

    data = []

    total = len(order_nums)

    for i, order_num in enumerate(order_nums):
        if progress:
            print(f"{i + 1}/{total}")

        url = "https://www2.order-fulfillment.bz/vanityart/orders"

        driver.get(f"{url}/{order_num}/manage")

        tbody = driver.find_element(By.CSS_SELECTOR, "tbody")
        items = []
        carrier = driver.find_element(By.CSS_SELECTOR, "address").get_attribute("innerHTML")

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

        warehouse = driver.find_elements(By.CSS_SELECTOR, "tbody")[2].find_elements(By.CSS_SELECTOR, "td")[1].get_attribute("innerText")

        po = po_num["textContent"][3:]
        print(warehouse)
        warehouse = warehouse[warehouse.index("(") + 1:warehouse.index(")")]
        carrier = carrier[carrier.index("Shipping Method -") + 18:].strip()
        status = "Not Shipped"
        ship_status = sheets.get_ship_status(order_time, status)

        headings = driver.find_element(By.CSS_SELECTOR, "thead").find_element(By.CSS_SELECTOR, "tr").find_elements(By.CSS_SELECTOR, "th")

        num_index = None
        qty_index = None

        for i, heading in enumerate(headings):
            text = heading.get_attribute("innerText")

            if text == "Item #":
                num_index = i
            elif text == "Quantity":
                qty_index = i

            if num_index is not None and qty_index is not None:
                break
        else:
            raise ValueError

        for trow in tbody.find_elements(By.CSS_SELECTOR, "tr"):
            item_data = trow.find_elements(By.CSS_SELECTOR, "td")

            if not item_data:
                break

            num_td = item_data[num_index]
            qty_td = item_data[qty_index]
            num = num_td.get_attribute("title")
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

    class_lookup = sheets.load_class_lookup("class_lookup.xlsx")
    combo_lookup = sheets.load_combo_lookup("combo_lookup.xlsx")
    warehouses = sheets.input_warehouses()
    output_data = sheets.parse_data(data, warehouses, class_lookup, combo_lookup)

    sheets.write_data(output_data, "scraped.xlsx")


if __name__ == "__main__":
    main()
