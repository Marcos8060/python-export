
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd


def migrate_excel_to_db():
    excel_file: str = "export.xlsx"

    report_df = pd.read_excel(excel_file,)
    report_df.head()


def extract_totara_report():
    driver = webdriver.Chrome()
    try:
        driver.get("https://careskillslearning.co.uk/login/index.php")
        wait = WebDriverWait(driver, 10)

        username_email_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "loginbtn")))

        username_email_field.clear()
        username_email_field.send_keys("mismgr@venus.sg")
        time.sleep(5)

        password_field.clear()
        password_field.send_keys("kBNGBaY98P7WkfJ")
        time.sleep(5)

        login_btn.click()
        time.sleep(5)

        # Navigating to the specified route after successful login
        driver.get("https://careskillslearning.co.uk/totara/reportbuilder/report.php?id=391")
        time.sleep(5)

        # Clicking the export button after navigating to the route
        export_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "export-button")))
        export_button.click()
        time.sleep(5)

    except Exception as e:
        print(e)
    finally:
        driver.close()


if __name__ == '__main__':
    migrate_excel_to_db()


