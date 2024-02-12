import time

from mysql.connector.pooling import PooledMySQLConnection
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd

from mysql import connector
from mysql.connector.abstracts import MySQLConnectionAbstract


def create_table_query(table_name: str) -> str:
    created_table_cmd = (
        "CREATE TABLE {table_name}("
        "first_name VARCHAR(255),"
        "last_name VARCHAR(255),"
        "username VARCHAR(255),"
        "organisation TEXT,"
        "course_name TEXT,"
        "completed_date TEXT,"
        "due_expiry_date TEXT,"
        "csa_status TEXT,"
        "certification_id_number TEXT,"
        "user_id INTEGER,"
        "is_current INTEGER"
        ");"
    ).format(table_name=table_name)
    return created_table_cmd


def insert_into_query(table_name: str) -> str:
    insert_data_cmd = f'''
    INSERT INTO {table_name} (first_name, last_name, username, organisation, course_name, completed_date, due_expiry_date, csa_status, certification_id_number, user_id, is_current) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    '''
    return insert_data_cmd


def drop_table_query(table_name: str) -> str:
    drop_table_cmd = f'''
    DROP TABLE IF EXISTS {table_name};
    '''
    return drop_table_cmd


def connect_to_database() -> PooledMySQLConnection | MySQLConnectionAbstract | None:
    db_user = 'vision'
    db_password = 'fvL5yPNaeHhHWPHQ'
    db_host = 'localhost'
    db_port = 3366
    db_name = 'vision'

    try:
        connection = connector.connect(
            user=db_user,
            password=db_password,
            host=db_host,
            port=db_port,
            database=db_name,
            auth_plugin="mysql_native_password",
        )

        return connection
    except Exception as e:
        print(e)
        return None


def migrate_excel_to_db():
    # connect to db
    connection = connect_to_database()

    if connection is None:
        raise ConnectionError("Problem connecting with database")

    excel_file: str = "export.xlsx"
    # created a dataframe
    report_df = pd.read_excel(excel_file,)
#     replace NaN values in df with None
    report_df_copy = report_df.replace({pd.NA: None})

    cursor = connection.cursor()
    
    table_name = "skill_reports"
    # drop table
    drop_query = drop_table_query(table_name)
    cursor.execute(drop_query)

#     create table
    create_query = create_table_query(table_name)
    cursor.execute(create_query)

#     insert data
    insert_query = insert_into_query(table_name)

    for index, row in report_df_copy.iterrows():
        if index % 10 == 0:
            print(f"{index} rows added")
            time.sleep(2)

        row_dict = row.to_dict()
        row_values = []
        for col in report_df.columns:
            row_values.append(row_dict[col])

        cursor.execute(insert_query, tuple(row_values))
        print(index)


def extract_totara_report():
    driver = webdriver.Chrome()
    try:
        driver.get("htVARCHAR(255)tps://careskillslearning.co.uk/login/index.php")
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