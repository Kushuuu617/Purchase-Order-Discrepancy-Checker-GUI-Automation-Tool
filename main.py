import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException, WebDriverException
)
from openpyxl import Workbook
import threading
import traceback


def start_automation(username, password, status_text):
    def update_status(msg):
        status_text.insert(tk.END, msg + "\n")
        status_text.see(tk.END)

    driver = None

    try:
        options = Options()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(options=options)
        update_status("Browser launched successfully.")

        driver.get("https://induserp.industowers.com/OA_HTML/AppsLocalLogin.jsp")
        update_status("Navigating to login page...")

        driver.find_element(By.ID, 'usernameField').send_keys(username)
        driver.find_element(By.ID, 'passwordField').send_keys(password)
        driver.find_element(By.CLASS_NAME, "OraButton").click()

        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, "//div[text()='India Local iSupplier']"))
        ).click()
        update_status("Logged in successfully.")

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

        messagebox.showinfo("Filter Step", "Apply the filters manually and press OK to continue.")

        wb = Workbook()
        ws = wb.active
        ws.title = "Discrepant POs"
        ws.append(["PO Number"])
        discrepant_po_numbers = []

        while True:
            try:
                table = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
                )
                rows = table.find_elements(By.TAG_NAME, 'tr')

                for i in range(len(rows)):
                    try:
                        doc_type = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, f"//span[contains(@id, 'PosDocumentType:{i}')]"))
                        ).text.strip()
                        if doc_type == "Global Blanket Agreement":
                            continue

                        po_link = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, f"//a[contains(@id, 'PosPoNumber:{i}')]"))
                        )
                        po_number = po_link.text

                        driver.execute_script("window.open(arguments[0], '_blank');", po_link.get_attribute('href'))
                        driver.switch_to.window(driver.window_handles[1])

                        try:
                            total_amount = float(WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "TotalAmt"))
                            ).text.replace(',', ''))

                            billed_amount = float(WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "AmtBilled"))
                            ).text.replace(',', ''))

                            if billed_amount < total_amount:
                                discrepant_po_numbers.append(po_number)
                                update_status(f"Discrepancy found: {po_number}")

                        except Exception as e:
                            update_status(f"Error reading PO details for {po_number}: {str(e)}")

                        finally:
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.ID, 'ResultRN.PosVpoPoList:Content'))
                            )

                    except Exception as e:
                        update_status(f"Error processing PO index {i}: {str(e)}")

                try:
                    next_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Next')]"))
                    )
                    next_button.click()
                    WebDriverWait(driver, 10).until(EC.staleness_of(table))

                except (TimeoutException, NoSuchElementException, ElementClickInterceptedException):
                    update_status("✅ All pages processed.")
                    break

            except Exception as e:
                update_status(f"Error loading table: {str(e)}")
                break

        for po in discrepant_po_numbers:
            ws.append([po])
        wb.save("POs_with_billed_less_than_total.xlsx")
        update_status("✔ Excel file saved successfully.")

    except WebDriverException as we:
        update_status(f"Browser/driver issue: {we}")
    except Exception as e:
        update_status("Fatal error occurred.\n" + traceback.format_exc())
    finally:
        if driver:
            try:
                driver.quit()
                update_status("Browser closed.")
            except:
                update_status("Failed to close browser properly.")


def run_gui():
    root = tk.Tk()
    root.title("PO Discrepancy Checker")
    root.geometry("550x450")

    tk.Label(root, text="Username:").pack(pady=5)
    username_entry = tk.Entry(root, width=50)
    username_entry.pack()

    tk.Label(root, text="Password:").pack(pady=5)
    password_entry = tk.Entry(root, width=50, show="*")
    password_entry.pack()

    tk.Label(root, text="Status:").pack(pady=5)
    status_text = tk.Text(root, height=15, width=65, wrap=tk.WORD)
    status_text.pack()

    def on_start():
        username = username_entry.get()
        password = password_entry.get()
        if not username or not password:
            messagebox.showerror("Error", "Username and Password are required.")
            return
        threading.Thread(target=start_automation, args=(username, password, status_text), daemon=True).start()

    tk.Button(root, text="Start Automation", command=on_start).pack(pady=10)
    root.mainloop()


if __name__ == "__main__":
    run_gui()

