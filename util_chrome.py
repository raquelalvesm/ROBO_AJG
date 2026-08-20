from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


def criar_driver(user_data_dir=None, prefs=None, debugger_address=None):
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--remote-debugging-port=0')
    if debugger_address:
        options.add_argument(f'--remote-debugging-address={debugger_address}')
        options.add_experimental_option('debuggerAddress', debugger_address)
    elif user_data_dir:
        options.add_argument(f'--user-data-dir={user_data_dir}')
    if prefs:
        options.add_experimental_option('prefs', prefs)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.implicitly_wait(10)
    return driver
