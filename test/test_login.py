from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
import pytest
import os
from pages.login_page import LoginPage


def get_login_data():
    login_data = []
    # Read data from the Excel file
    df = pd.read_excel('data/login_data.xlsx')

    # Iterate through the rows of the DataFrame
    for index, row in df.iterrows():
        username = row['username']
        password = row['password']
        login_data.append((username, password))

    return login_data

    # update the data into excel file


def update_excel(username, password, result, tester_name):
    current_date = datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M:%S")
    file_path = 'data/login_data.xlsx'
    temp_file_path = 'data/login_data_temp.xlsx'

    try:
        # Loading the Excel file into a DataFrame
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return

    # Ensure the columns are of type string
    for col in ['Date', 'Time of Test', 'Name of Tester', 'Test Result']:
        if col not in df.columns:
            df[col] = ""  # Adding missing columns with empty strings

    df['Date'] = df['Date'].astype(str)
    df['Time of Test'] = df['Time of Test'].astype(str)
    df['Name of Tester'] = df['Name of Tester'].astype(str)
    df['Test Result'] = df['Test Result'].astype(str)

    # Updating the DataFrame with the test results obtained
    record_found = False
    for index, row in df.iterrows():
        if row['username'] == username and row['password'] == password:
            df.at[index, 'Date'] = current_date
            df.at[index, 'Time of Test'] = current_time
            df.at[index, 'Name of Tester'] = tester_name
            df.at[index, 'Test Result'] = result
            record_found = True
            break

    if not record_found:
        print(f"No matching record found for username: {username} and password: {password}")
        return

    try:
        # Saving the updated DataFrame to a temporary file
        with pd.ExcelWriter(temp_file_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Replacing the original file with the temporary file
        os.replace(temp_file_path, file_path)
    except Exception as e:
        print(f"Error writing to the Excel file: {e}")
        #  handling the temporary file cleanup
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)


@pytest.mark.usefixtures("setup")
class TestLogin:
    @pytest.mark.parametrize("username, password", get_login_data())
    def test_login(self, username, password):
        login_page = LoginPage(self.driver)
        login_page.login(username, password)

        tester_name = 'Revathy'

        if "dashboard" in self.driver.current_url:
            update_excel(username, password, "Passed", tester_name)
        else:
            update_excel(username, password, "Failed", tester_name)

        assert "dashboard" in self.driver.current_url
