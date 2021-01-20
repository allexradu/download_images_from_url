from urllib.request import urlopen
import time
import requests
from selenium import webdriver
from distutils.spawn import find_executable
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
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
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-setuid-sandbox")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("window-size=1400,2100")
chrome_options.add_argument('--disable-dev-shm-usage')

chromium = find_executable('chromium-browser')

if chromium:
    chrome_options.binary_location = chromium

chrome_options.add_argument('--browser.download.folderList=2')
chrome_options.add_argument(
    '--browser.helperApps.neverAsk.saveToDisk=application/octet-stream')
path = os.path.join(os.getcwd(), 'excel', 'photos')
prefs = {'download.default_directory': f'{path}'}
chrome_options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(ChromeDriverManager().install(), options = chrome_options)
# driver = webdriver.Chrome(options = chrome_options)

# url = 'https://mall.industry.siemens.com/mall/en/ww/Catalog/DatasheetDownload?downloadUrl=teddatasheet%2F%3Fformat%3DPDF%26caller%3DMall%26mlfbs%3D3NA3360%26language%3Den'

for i in range(1, len(excel.excel_product_image_url)):
    print(f' row {i} / {len(excel.excel_product_image_url)}')
    driver.get(excel.excel_product_image_url[i])
    time.sleep(2)
