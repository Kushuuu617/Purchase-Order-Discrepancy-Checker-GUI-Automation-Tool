from selenium import webdriver # type: ignore
from selenium.webdriver.common.by import By # type: ignore
from selenium.webdriver.chrome.options import Options # type: ignore
from selenium.webdriver.support.ui import WebDriverWait # type: ignore
from selenium.webdriver.support import expected_conditions as EC # type: ignore
from selenium.webdriver.support.ui import Select # type: ignore
from openpyxl import Workbook # type: ignore

import time

Username = input("Username : ")
Password = input("Password : ")

options = Options()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)
driver.get("https://induserp.industowers.com/OA_HTML/AppsLocalLogin.jsp")

driver.find_element(By.ID, 'usernameField').send_keys(Username) 
driver.find_element(By.ID, 'passwordField').send_keys(Password)
  
driver.find_element(By.CLASS_NAME, "OraButton").click()

WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, "//div[text()='India Local iSupplier']"))
).click()

WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, "//div[text()='Home Page']"))
).click()

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ILS_POS_HOME"))
).click()

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "PosHpOrdersTable:PosHpoPoNum:0"))
).click()

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "POS_PURCHASE_ORDERS"))
).click()

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "SrchBtn"))
).click()

take_input = input("Fill details and press enter")

wb = Workbook()
ws = wb.active
ws.title = "POs with Billed < Total"
ws.append(["PO Number"])  # Header row
po_numbers_to_save = []

running = True

while running:
    # Find all row elements freshly after each page load
    try:
        table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
        )
        rows = table.find_elements(By.TAG_NAME, 'tr')
        num_rows = len(rows)

        for i in range(0, num_rows):
            try:
                # Check Document Type
                doc_type_elem = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, f"//span[contains(@id, 'PosDocumentType:{i}')]"))
                )
                doc_type_text = doc_type_elem.text.strip()
                if doc_type_text == "Global Blanket Agreement":
                    continue  # skip this row
                link = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//a[contains(@id, 'PosPoNumber:{i}')]"))
                )
                po_number = link.text
                driver.execute_script("window.open(arguments[0], '_blank');", link.get_attribute('href'))

                # Switch to new tab
                driver.switch_to.window(driver.window_handles[1])

                try:
                    total_amt = float(WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "TotalAmt"))
                    ).text)

                    billed_amt = float(WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "AmtBilled"))
                    ).text)

                    if billed_amt < total_amt:
                        po_numbers_to_save.append(po_number)

                except Exception as e:
                    print(f"Error reading PO details for {po_number}: {e}")

                finally:
                    # Close the tab and return to main window
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

                    # Optional: wait for table to ensure it’s still visible before proceeding
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
                    )
            except Exception as e:
                print(f"Error processing PO {i}: {e}")

        # Try to click the next button
        try:
            next_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Next')]"))
            )
            next_button.click()

            #Wait for the new page to load (table to refresh)
            WebDriverWait(driver, 10).until(
               EC.staleness_of(table)  # old table must go stale
            )
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
            )

        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException): # type: ignore
            print("No more pages left or Next button not clickable.")
            running = False

    except Exception as e:
        print(f"Error loading table: {e}")
        running = False
        
for po_num in po_numbers_to_save:
    ws.append([po_num])
wb.save("POs_with_billed_less_than_total.xlsx")

input("Press Enter to close the browser...")
