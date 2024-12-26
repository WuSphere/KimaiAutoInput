from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import logging


class BaseUtils:
    @classmethod
    # 要素がクリック可能になるまで待機する関数
    def wait_and_click(cls, driver, by, selector, timeout=10):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, selector))
            )
            element.click()
            logging.info(f"要素がクリックされました: {selector}")
            return element
        except Exception as e:
            logging.error(f"クリック待機中にエラーが発生しました: {selector} - {e}")
            raise

    @classmethod
    # 要素が表示されるまで待機して値を入力する関数
    def wait_and_send_keys(cls, driver, by, selector, value, timeout=10):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((by, selector))
            )
            element.clear()
            element.send_keys(value)
            logging.info(f"入力が完了しました: {selector} - {value}")
        except Exception as e:
            logging.error(f"入力待機中にエラーが発生しました: {selector} - {e}")
            raise

    @classmethod
    # 要素が表示されるまで待機して値を入力する関数
    def wait_and_select_value(
        cls, driver, by, selector, value, escapeFlg=0, timeout=10
    ):
        try:
            ele_select = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, selector))
            )
            ele_select.click()

            ele_option = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//div[@data-selectable and text()='{value}']")
                )
            )
            ele_option.click()
            if escapeFlg == 1:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception as e:
            logging.error(f"オプション選択中にエラーが発生しました: {selector} - {e}")
            raise

    @classmethod
    # 要素が存在するまで待機する関数
    def wait_until_present(cls, driver, by, selector, timeout=10):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
            logging.info(f"要素が表示されました: {selector}")
            return element
        except Exception as e:
            logging.error(f"要素待機中にエラーが発生しました: {selector} - {e}")
            raise
