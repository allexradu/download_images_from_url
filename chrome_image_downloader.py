from urllib.request import urlopen
import time
import requests
from selenium import webdriver
from distutils.spawn import find_executable
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import excel
import platform
import excel
import os

read_urls_excel_cell_letter = 'B'
read_product_name_cell_letter = 'A'
image_names_cell_letters = ['C']

image_file_names = []

excel.read_product_names(read_product_name_cell_letter)
excel.read_image_urls(read_urls_excel_cell_letter)

chrome_options = Options()
# chrome_options.add_argument("--no-sandbox")
# chrome_options.add_argument("--disable-setuid-sandbox")
# chrome_options.add_argument("--headless")
# chrome_options.add_argument("--disable-gpu")
# chrome_options.add_argument("window-size=1400,2100")
# chrome_options.add_argument('--disable-dev-shm-usage')

chromium = find_executable('chromium-browser')

# if chromium:
#     chrome_options.binary_location = chromium
#
# chrome_options.add_argument('--browser.download.folderList=2')
# chrome_options.add_argument(
#     '--browser.helperApps.neverAsk.saveToDisk=application/octet-stream')
# path = os.path.join(os.getcwd(), 'excel', 'photos')
# prefs = {'download.default_directory': f'{path}'}
# chrome_options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(ChromeDriverManager().install(), options = chrome_options)
# driver = webdriver.Chrome(options = chrome_options)

# url = 'https://www.assets.signify.com/is/content/PhilipsLighting/fp913703013809-pss-global'
# driver.get(url)
# time.sleep(2)

for i in range(1, len(excel.excel_product_image_url)):
    print(f' row {i} / {len(excel.excel_product_image_url)}')
    if excel.excel_product_image_url[i] is not None:
        driver.get(excel.excel_product_image_url[i])
        time.sleep(0.2)
        img = driver.find_element_by_css_selector('img')
        action = ActionChains(driver)
        action.context_click(on_element = img).send_keys(Keys.ARROW_DOWN).send_keys(
            Keys.RETURN).perform()

        time.sleep(100)
