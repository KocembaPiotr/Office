from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


def browser_change_zoom_edge(browser, button_id: str) -> None:
    """
    Function to change zoom in MS Edge websites.
    :param button_id: id of button which on edge size is responsible for change zoom
    :param browser: webdriver from Selenium for Edge
    :return: None
    """
    browser.maximize_window()
    browser.get('edge://settings/appearance')
    bt = browser.find_elements(By.XPATH, f"//button[@ID='{button_id}']")
    bt[0].find_element(By.TAG_NAME, 'div')
    bt[0].click()
    bt[0].send_keys(Keys.ARROW_UP)
    bt[0].send_keys(Keys.ARROW_UP)
    bt[0].send_keys(Keys.ENTER)