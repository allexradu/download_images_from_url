import urllib.request
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
from urllib.error import HTTPError

read_urls_excel_cell_letter = 'B'
read_product_name_cell_letter = 'A'
image_names_cell_letters = ['C']

image_file_names = []

excel.read_product_names(read_product_name_cell_letter)
excel.read_image_urls(read_urls_excel_cell_letter)

for i in range(2645, len(excel.excel_product_image_url)):
    print(f'image {i}/{len(excel.excel_product_image_url)}')
    if excel.excel_product_image_url[i] is not None:
        try:
            urllib.request.urlretrieve(excel.excel_product_image_url[i],
                                       os.path.join(os.getcwd(), 'excel', 'photos',
                                                    excel.excel_product_image_url[i].split('/')[-1]))
        except HTTPError:
            pass
