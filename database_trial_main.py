import asyncio
import os
import pandas as pd
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
from datetime import datetime, timedelta
import re


# Define constants
INPUT_DIR = '../Task_12/'
BASE_URL = "https://theautotrender.com"
PREVIOUS_DATA_FILE = 'previous_data.json'  # File to store previous data

# Selenium setup
def setup_driver():
    chrome_options = Options()
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    return driver

def login(driver):
    driver.get(BASE_URL)
    wait = WebDriverWait(driver, 10)

    # Enter username
    username_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mui-1"]')))
    username_field.send_keys("patelsuchit@gmail.com")

    # Enter password
    password_field = driver.find_element(By.XPATH, '//*[@id="mui-2"]')
    password_field.send_keys("Pat@27811")

    # Check checkbox
    checkbox = driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/div[2]/section[2]/div/form/div[3]/label/span/input')
    checkbox.click()

    # Click submit button
    submit_button = driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/div[2]/section[2]/div/form/button')
    submit_button.click()

def sanitize_filename(url, output_dir):
    match = re.search(r'=(W?)(.*?)&', url)
    if match:
        sheet_name = match.group(2).strip().upper()
    else:
        sheet_name = re.sub(r'[\/:*?"<>|]', '_', url)
    
    if sheet_name.endswith('W'):
        sheet_name = sheet_name[:-1] + ' '
    elif sheet_name.endswith('w'):
        sheet_name = sheet_name[:-1] + ' '

    # Truncate the sheet name to ensure it doesn't exceed 31 characters
    if len(sheet_name) >= 31:
        sheet_name = sheet_name[:31]

    base_sheet_name = sheet_name  # Store the original sheet name
    count = 1
    base_filename = os.path.join(output_dir, f"{base_sheet_name}_{count}.xlsx")

    while os.path.exists(base_filename):
        count += 1
        base_filename = os.path.join(output_dir, f"{base_sheet_name}_{count}.xlsx")

    sheet_name = f"{base_sheet_name}_{count}"

    return sheet_name

def handle_popups(driver):
    try:
        close_button = driver.find_element(By.XPATH, "//button[contains(@class, 'close-button-class')]")
        close_button.click()
        print("Closed a popup.")
    except (NoSuchElementException, TimeoutException):
        pass

async def scrape_table(driver, table_index=0):
    wait = WebDriverWait(driver, 10)
    tables = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//table')))
    
    if len(tables) > table_index:
        table = tables[table_index]
        headers = table.find_elements(By.TAG_NAME, "th")
        header_names = [header.text for header in headers]

        rows = table.find_elements(By.TAG_NAME, "tr")
        data = []

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            cols = [col.text if col.text.strip() != "" else "" for col in cols]
            data.append(cols)

        if header_names:
            df = pd.DataFrame(data, columns=header_names)
        else:
            df = pd.DataFrame(data)

        return df
    else:
        return None

def load_previous_data():
    if os.path.exists(PREVIOUS_DATA_FILE):
        with open(PREVIOUS_DATA_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_current_data(url, data):
    previous_data = load_previous_data()
    previous_data[url] = data[:4]  # Save first 4 rows for comparison
    with open(PREVIOUS_DATA_FILE, 'w') as f:
        json.dump(previous_data, f)


def compare_data(url, new_data):
    previous_data = load_previous_data()
    if url in previous_data:
        for prev, new in zip(previous_data[url], new_data):
            if prev != new:
                return False
        return True
    return False


async def scrape_and_process_tab(driver, url):
    retry_count = 3
    while retry_count > 0:
        try:
            # Start timing
            start_time = time.time()

            driver.get(url)
            wait = WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.XPATH, '//table')))
            handle_popups(driver)

            sheet_name = sanitize_filename(url, INPUT_DIR)
            excel_path = os.path.join(INPUT_DIR, f"{sheet_name}.xlsx")

            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                data_frame = await scrape_table(driver, table_index=0)
                if data_frame is not None:
                    first_rows_data = data_frame.iloc[:4].values.tolist()  # Get the first 4 rows of data

                    if compare_data(url, first_rows_data):
                        print(f"Skipping scraping for {url} as the data hasn't changed.")
                        os.remove(excel_path)  # Remove the file if the data hasn't changed
                        return None, None
                    
                    else:
                        data_frame.to_excel(writer, sheet_name=sheet_name + "Option", index=False)
                        save_current_data(url, first_rows_data)
                        print(f"Scraped and saved data from {url}")
                    
                btn_3min_xpath = '//*[@id="__next"]/main/div[2]/div[2]/main/div[3]/div/div/div[2]/div[1]/div/button[1]'
                driver.find_element(By.XPATH, btn_3min_xpath).click()
                time.sleep(5)
                data_frame = await scrape_table(driver, table_index=1)
                if data_frame is not None:
                    data_frame.to_excel(writer, sheet_name='3 Min', index=False)

                btn_5min_xpath = '//*[@id="__next"]/main/div[2]/div[2]/main/div[3]/div/div/div[2]/div[1]/div/button[2]'
                driver.find_element(By.XPATH, btn_5min_xpath).click()
                time.sleep(5)
                data_frame = await scrape_table(driver, table_index=1)
                if data_frame is not None:
                    data_frame.to_excel(writer, sheet_name='5 Min', index=False)

                data_frame = await scrape_table(driver, table_index=1)
                if data_frame is not None:
                    data_frame.to_excel(writer, sheet_name='15 Min', index=False)
            
           
            
            # Check if the Excel file is empty or contains fewer than 3 non-empty sheets
            if os.path.exists(excel_path):
                excel_file = pd.ExcelFile(excel_path)
                sheet_names = excel_file.sheet_names
                if len(sheet_names) < 3 or all([pd.read_excel(excel_file, sheet).empty for sheet in sheet_names]):
                    os.remove(excel_path)
                    print(f"Removed file due to insufficient data or being empty: {excel_path}")
                else:
                    print(f"File saved with sufficient tables: {excel_path}")

             # End timing
            end_time = time.time()
            time_taken = end_time - start_time
            

            print(f"Time taken for {url}: {time_taken:.2f} seconds")

            return excel_path, sheet_name
        except StaleElementReferenceException:
            retry_count -= 1
            print(f"Retrying scraping {url} due to stale element reference ({3 - retry_count}/3)")
        except Exception as e:
            print(f"Error scraping {url}: {e}")
            break
    return None, None

async def process_tabs(driver, urls):
    tasks = []
    
    for url in urls:
        driver.execute_script(f"window.open('{url}', '_blank');")
        await asyncio.sleep(2)  # Introduce a small delay between opening tabs

    await asyncio.sleep(10)  # Wait for all tabs to load

    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        try:
            if driver.current_url in urls:
                tasks.append(asyncio.create_task(scrape_and_process_tab(driver, driver.current_url)))
        except Exception as e:
            print(f"Error processing tab: {e}")

    await asyncio.gather(*tasks)


async def main():
    driver = setup_driver()
    end_time = datetime.now() + timedelta(hours=5)

    start_time = datetime.now()  # Record start time

    try:
        login(driver)
        while datetime.now() < end_time:
            urls = [
                "https://theautotrender.com/derivative?category=niftyW&id=OIData",
                "https://theautotrender.com/derivative?category=bankNiftyW&id=OIDataW",
                "https://theautotrender.com/derivative?category=FINNiftyW&id=OIData",
            ]
            await process_tabs(driver, urls)
            await asyncio.sleep(60)

    finally:
        driver.quit()
    
    end_time = datetime.now()  # Record end time
    total_time = end_time - start_time  # Calculate elapsed time
    print(f"Total time taken: {total_time}")

if __name__ == "__main__":
    asyncio.run(main())

