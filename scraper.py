import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Path to your ChromeDriver
chrome_driver_path = r"chromedriver-win64/chromedriver.exe"

# Create a new instance of the Chrome driver
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)
url = 'https://edge.pse.com.ph/otherReports/form.do'
driver.get(url)

# Capture the handle for the main window
main_page = driver.current_window_handle

# Wait for the page to load
time.sleep(5)

# Function to filter by "Public Ownership"
def filter_public_ownership():
    print("Locating the 'Template Name' textbox...")
    template_name_box = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'tmplNm'))
    )
    print("Textbox located. Clearing and entering 'Public Ownership'...")
    template_name_box.clear()
    template_name_box.send_keys('Public Ownership')

    print("Locating the 'Search' button...")
    search_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'btnSearch'))
    )
    print("Search button located. Clicking the button...")
    search_button.click()

    print("Waiting for the results table to load...")
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//table[contains(@class, "list")]'))
    )
    print("Results table loaded.")

# Function to click the "Public Ownership Report" link for the first row
def click_first_public_ownership_report():
    print("Locating the results table...")
    table = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//table[contains(@class, "list")]'))
    )
    print("Table located. Finding the first 'Public Ownership Report' link...")
    first_row_link = table.find_element(By.XPATH, './/tbody/tr[1]/td[2]/a[text()="Public Ownership Report"]')
    
    onclick_attribute = first_row_link.get_attribute('onclick')
    print(f"Found 'onclick' attribute: {onclick_attribute}. Executing script...")
    driver.execute_script(onclick_attribute)

    print("Waiting for the popup to load...")
    time.sleep(10)  # Wait for the popup to appear

    # Switch to the new window
    for handle in driver.window_handles:
        if handle != main_page:
            popup_page = handle
            break
    driver.switch_to.window(popup_page)

    print("Popup opened successfully. Waiting for the content to load completely...")
    form_element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//iframe[@id="viewContents"]'))
    )
    print("Iframe element found. Switching to iframe...")
    driver.switch_to.frame(form_element)
    
    return True

# Function to extract and save table data as CSV files
def extract_and_save_table_data():
    print("Extracting data from tables...")

    def parse_table_by_id(table_id):
        try:
            print(f"Attempting to locate table with ID: {table_id}")
            table_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, f'//table[@id="{table_id}"]'))
            )
            print(f"Table with ID: {table_id} found.")
            rows = table_element.find_elements(By.TAG_NAME, "tr")
            headers = [th.text for th in rows[0].find_elements(By.TAG_NAME, "th")]
            data = []
            for row in rows[1:]:
                cells = row.find_elements(By.TAG_NAME, "td")
                data.append([cell.text for cell in cells])
            return headers, data
        except Exception as e:
            print(f"Error parsing table with ID {table_id}: {e}")
            return [], []

    # Get the company name
    company_name = driver.find_element(By.ID, 'companyName').text.strip()
    print(f"Company Name: {company_name}")

    # Create a directory for the company
    if not os.path.exists(company_name):
        os.makedirs(company_name)

    # Parse and save tables
    table_ids = ["Directors", "Officers"]
    for table_id in table_ids:
        print(f"Parsing table with ID {table_id}...")
        headers, data = parse_table_by_id(table_id)
        if headers and data:
            df = pd.DataFrame(data, columns=headers)
            file_path = os.path.join(company_name, f"{table_id}.csv")
            df.to_csv(file_path, index=False)
            print(f"Saved {table_id} data to {file_path}")

# Filter the results by "Public Ownership"
filter_public_ownership()

# Add a delay to ensure the filter operation completes
print("Waiting for the filter operation to complete...")
time.sleep(10)  # Increased delay to ensure the page fully loads

# Click the "Public Ownership Report" link for the first row
if click_first_public_ownership_report():
    # Extract and save table data
    extract_and_save_table_data()

# Close the popup and switch back to the main window
driver.close()
driver.switch_to.window(main_page)

# Close the driver
driver.quit()
