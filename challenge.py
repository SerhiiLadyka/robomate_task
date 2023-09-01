import time
from typing import List, Any
from abc import ABC, abstractmethod

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

import pandas as pd


class AutoFillFormService(ABC):
    def __init__(self, url: str, download_dir: str = None, detach_flag: bool = False):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", detach_flag)

        if download_dir:
            chrome_options.add_experimental_option("prefs", {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True
            })

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.get(url)

    @abstractmethod
    def _process_data_for_form(self) -> Any:
        pass

    @abstractmethod
    def _fill_form(self, fill_data: Any):
        pass

    def run(self):
        data = self._process_data_for_form()
        self._fill_form(data)
        self.driver.quit()


class RPAChallengeFillService(AutoFillFormService):
    driver = None
    SITE_URL = "https://rpachallenge.com/"
    FILE_DIR_PATH = "/Users/sergijladika/Desktop/PROJECTS/robomate_task"

    def __init__(self):
        super().__init__(url=self.SITE_URL, download_dir=self.FILE_DIR_PATH, detach_flag=True)

    def _process_data_for_form(self) -> List[dict]:
        download_button = self.driver.find_element(By.XPATH, '//a[contains(text(), "Download")]')
        download_button.click()

        start_button = self.driver.find_element(By.XPATH, '//button[contains(text(), "Start")]')
        start_button.click()
        time.sleep(5)

        file_path = self.FILE_DIR_PATH + "/challenge.xlsx"
        file_data_frame = pd.read_excel(file_path)
        fill_data = file_data_frame.to_dict('records')
        return fill_data

    def _fill_form(self, fill_data: List[dict]):
        for row in fill_data:
            for key, value in row.items():
                attribute_value = key.strip()

                elem = self.driver.find_element(
                    "xpath",
                    f'//*[@ng-reflect-dictionary-value="{attribute_value}"]/div/input'
                )

                elem.clear()
                elem.send_keys(value)

            submit_button = self.driver.find_element(By.XPATH, '//*[@value="Submit"]')
            submit_button.click()
            time.sleep(2)


if __name__ == "__main__":
    bot = RPAChallengeFillService()
    bot.run()

