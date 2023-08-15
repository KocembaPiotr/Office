from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


def browser_change_zoom_edge(browser) -> None:
    """
    Function to change zoom in MS Edge websites.
    :param browser: webdriver from Selenium for Edge
    :return: None
    """
    browser.maximize_window()
    browser.get('edge://settings/appearance')
    bt = browser.find_elements(By.XPATH, "//button[@ID='selecttrigger-47']")
    bt[0].find_element(By.TAG_NAME, 'div')
    bt[0].click()
    bt[0].send_keys(Keys.ARROW_UP)
    bt[0].send_keys(Keys.ARROW_UP)
    bt[0].send_keys(Keys.ENTER)