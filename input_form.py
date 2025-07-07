"""Run input form in RPA Challenge"""
import os
import time

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

XPATH_DOWNLOAD = "//a[contains(text(), 'Download Excel')]"
XPATH_START_BUTON = "//button[contains(text(), 'Start')]"

"""Class InputForm"""


class InputForm:
    """
        This class handles all the input forms
    """

    def __init__(self, download_dir: str, url: str):
        self.data = {}
        self.url = url
        self.base_name = {
            "First Name": "A",
            "Last Name": "B",
            "Company Name": "C",
            "Role in Company": "D",
            "Address": "E",
            "Email": "F",
            "Phone Number": "G",
        }

        chrome_options = self.setting_option(download_dir=download_dir)
        self.driver = webdriver.Chrome(options=chrome_options)

    @staticmethod
    def setting_option(download_dir: str):
        """
        Setting options for chrome driver

        :param download_dir: Path to the download folder.
        :return: None
        """
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        return chrome_options

    def set_link_web(self):
        """
        Setting link for web browser

        :return: None
        """
        self.driver.get(self.url)

    def download_data(self):
        """
        Download data from web browser

        :return: None
        """
        download_button = self.driver.find_element(By.XPATH, XPATH_DOWNLOAD)
        download_button.click()

    def get_data_from_file(self, name_file: str):
        """
        Get data from file
        :param name_file: Name of the file to be downloaded.
        :return: None
        """
        wb = openpyxl.load_workbook(name_file)
        list_sheet_names = wb.sheetnames
        sheet = wb[list_sheet_names[0]]
        data = []

        for i in range(2, 12):
            data_item = {}
            for key, value in self.base_name.items():
                name = value + str(i)
                value = sheet[name].value
                data_item[key] = value
            data.append(data_item)

        self.data = data

    def start_the_process(self):
        """
        Start the process

        :return: None
        """
        start_button = self.driver.find_element(By.XPATH, XPATH_START_BUTON)
        start_button.click()

    def run_the_process(self):
        """
        Run the process

        :return: None
        """
        for user in self.data:
            form = self.driver.find_element(By.TAG_NAME, 'form')
            submit_button = form.find_element(By.CSS_SELECTOR, 'input[type="submit"]')

            input_forms = form.find_element(By.CLASS_NAME, 'row')

            list_items_input = input_forms.find_elements(By.XPATH, 'div')

            for item in list_items_input:
                label = item.find_element(By.TAG_NAME, 'label')
                label_text = label.text
                input_item = item.find_element(By.TAG_NAME, 'input')
                input_item.send_keys(user[label_text])
            submit_button.click()

    def run_input_form(self, download_dir: str, name_file: str):
        """
        Full follow input form
        :param download_dir: Path to the download folder.
        :param name_file: Name of the file to be downloaded.
        :return: None
        """
        self.set_link_web()
        self.download_data()

        if not self.wait_for_file_download(
                download_dir=download_dir,
                filename=name_file
            ):
            raise FileNotFoundError()

        self.get_data_from_file(name_file=name_file)
        self.start_the_process()
        self.run_the_process()

    @staticmethod
    def wait_for_file_download(download_dir: str, filename: str,
                               timeout: int = 30, poll_interval: float = 0.5):
        """
        Waits until the file appears in the download directory or until timeout.

        :param download_dir: Path to the download folder.
        :param filename: Expected name of the file to be downloaded.
        :param timeout: Maximum time (in seconds) to wait for the download.
        :param poll_interval: Time interval (in seconds) between checks.
        :return: Full path to the file if downloaded successfully, or None if timed out.
        """
        file_path = os.path.join(download_dir, filename)
        start_time = time.time()

        while time.time() - start_time < timeout:
            if os.path.exists(file_path) and not file_path.endswith(".crdownload"):
                return file_path
            time.sleep(poll_interval)

        return None


if __name__ == '__main__':
    NAME_FILE_DOWNLOAD = 'challenge.xlsx'

    this_path_dir = os.path.dirname(os.path.abspath(__file__))
    input_form = InputForm(download_dir=this_path_dir, url='https://www.rpachallenge.com/')
    input_form.run_input_form(download_dir=this_path_dir, name_file=NAME_FILE_DOWNLOAD)
