"""
Automation script for this process: https://youtu.be/xWQro7ZHbao
Remember to include a XLS file in the same directory with IDs to look up
Run pip install -r requirements.txt and set password in environment or below before running
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import random
import time

# Initialize an empty list to store log entries
log_entries = []

# setup random time interval 0.5 to 2.5 seconds to avoid bot detection
random_time = random.uniform(0.5, 2.5)

# Initialize Selenium
print("Initializing Selenium...")
driver = webdriver.Chrome()

# First, go to the login page
print("Navigating to login page...")
login_url = "https://4395.east.3cx.us/#/login"
driver.get(login_url)

# Wait for the login page to load
print("Waiting for login page to load...")
time.sleep(random_time)

# Perform login
print("Locating login fields...")
username_field = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="User name or extension number"]')
password_field = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="Password"]')
login_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')

print("Entering credentials and clicking login...")
username_field.send_keys("admin")
password_field.send_keys("") # update with password before running
login_button.click()

# Wait for the login to complete
print("Waiting for login to complete...")
time.sleep(random_time)

# Navigate to the admin panel
print("Navigating to admin panel...")
admin_panel_url = "https://4395.east.3cx.us/#/app/extension_editor/"
driver.get(admin_panel_url)

# Wait for the admin panel to load
print("Waiting for admin panel to load...")
time.sleep(random_time)

# Read the spreadsheet to get IDs
print("Reading data from spreadsheet...")
df = pd.read_excel('cn_automation.xlsx')  # Replace with your actual path
ids = df['ID'].tolist()

# Loop through each ID to perform actions
for id in ids:
    # pause on every 25th ID to avoid rate limiting
    if id % 25 == 0:
        print("Pausing for 10 seconds to avoid rate limiting ...")
        time.sleep(10)

    # Initialize a log entry dictionary
    log_entry = {'ID': id}
    
    try:
        print(f"Processing ID: {id}")
        time.sleep(random_time)
        
        # Find the search bar and enter the ID
        print("Locating search bar...")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'inputSearch')))
        search_bar = driver.find_element(By.ID, 'inputSearch')

        print(f"Entering ID {id} into search bar...")
        search_bar.clear()
        search_bar.send_keys(str(id))

        # Wait for search results to load
        print("Waiting for search results...")
        time.sleep(random_time)

        print("Clicking the first search result...")
        first_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//tbody/tr[1]")))
        first_row.click()

        # Wait for the new page to load
        print("Waiting for the new page to load...")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.LINK_TEXT, 'Phone Provisioning')))
        time.sleep(random_time)

        # Click the "Phone Provisioning" link
        print("Navigating to Phone Provisioning...")
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Phone Provisioning'))).click()
        time.sleep(random_time)

        # Check if the panel exists. If not, click the "Cancel" button and skip to the next ID
        ip_phone_panel = driver.find_elements(By.XPATH, '//panel[@mc-title="EXTENSIONS.EDITOR.PROVISIONING.IP_PHONE"]')
        if not ip_phone_panel:
            print("IP Phone panel not found. Skipping to the next ID.")
            cancel_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'btnCancel')))
            cancel_button.click()
            log_entry['Status'] = 'Skipped'
            log_entries.append(log_entry)
            continue
        else:
            # Change the "Provisioning Method" dropdown to "3CX SBC (remote)"
            try:
                print("Changing Provisioning Method...")
                time.sleep(random_time)
                provisioning_dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//label[text()='Provisioning Method']/following-sibling::select")))
                select_provisioning = Select(provisioning_dropdown)
                select_provisioning.select_by_visible_text("3CX SBC (remote)")
            except Exception as e:
                print(f"Did not find the Provisioning Method dropdown. Error: {e}")

            # Change the "3CX Session Border Controller" dropdown to "SBC1"
            try:
                print("Changing 3CX Session Border Controller...")
                time.sleep(random_time)
                sbc_dropdown = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//label[text()='3CX Session Border Controller']/following-sibling::div/select")))
                select_sbc = Select(sbc_dropdown)
                select_sbc.select_by_visible_text("SBC1") # Update this to whatever the target dropdown option is
            except Exception as e:
                print(f"Did not find the 3CX Session Border Controller dropdown. Error: {e}")

        # Save the changes
        print("Saving changes...")
        save_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'btnSave')))
        save_button.click()

        log_entry['Status'] = 'Success'
        
    except Exception as e:
        print(f"An error occurred for ID {id}: {e}")
        log_entry['Status'] = f"Failure: {e}"
        
    # Append the log entry to the list
    log_entries.append(log_entry)

# Create a DataFrame from the list of log entries
log_df = pd.DataFrame(log_entries)

# Save the DataFrame to a CSV file with timestamp in the filename
log_df.to_csv('cn_automation_log_' + time.strftime("%Y%m%d-%H%M%S") + '.csv', index=False)

# Close the driver
print("Closing driver...")
driver.quit()
