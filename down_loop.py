import excel
import platform
import urllib
from urllib.error import HTTPError
from urllib.error import URLError
from urllib.request import urlretrieve
from socket import error as SocketError
from urllib.request import Request, urlopen

import requests
from requests.exceptions import ConnectionError
import shutil
import os

# read_urls_excel_cell_letters = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
# read_product_name_cell_letter = 'A'
# image_names_cell_letters = ['N', 'O', 'P', 'Q', 'R', 'S', 'T']

read_urls_excel_cell_letters = ['B']
read_product_name_cell_letter = 'A'
image_names_cell_letters = ['C']

image_file_names = []


excel.read_product_names(read_product_name_cell_letter)


def download_images(img_file_names, image_multiplier):
    for i in range(1, len(excel.excel_product_image_url)):
        if excel.excel_product_image_url[i] != '':
            if excel.excel_product_image_url[i] != 'none':
                if excel.excel_product_image_url[i] is not None:
                    try:
                        url = excel.excel_product_image_url[i]
                        split_url = url.split('/')
                        img_file_name_raw = split_url[-1]
                        split_file_name = img_file_name_raw.split('.')
                        img_suffix_raw = split_file_name[-1]
                        # img_suffix = img_suffix_raw[:3]
                        img_suffix = 'pdf'

                        img_file_name = excel.sanitise_product_names(
                            excel.excel_product_names[i]) + image_multiplier + '.' + img_suffix
                        if len(img_file_name) > 100:
                            img_file_name = excel.sanitise_product_names(
                                excel.excel_product_names[i][0:100]) + image_multiplier + '.' + img_suffix

                        # img_file_name = img_file_name_raw
                        system_prefix = 'excel\\photos\\' if platform.system() == 'Windows' else 'excel/photos/'

                        # response = requests.get(excel.excel_product_image_url[i], stream = True, verify = False)

                        # response = requests.get(excel.excel_product_image_url[i])
                        # print(response)

                        after_slash = excel.excel_product_image_url[i].split('/')[-1]
                        param1_key = after_slash[after_slash.find('?') + 1: after_slash.find('=')]
                        param1_value = after_slash[after_slash.find('=') + 1:]
                        # param1_value = after_slash[after_slash.find('=') + 1: after_slash.find('&')]
                        # param2_key = after_slash[
                        #              after_slash.find('&') + 1: after_slash.find('=', after_slash.find('&'))]
                        # param2_value = after_slash[after_slash.find('=', after_slash.find('&')) + 1:]
                        # payload = {param1_key: param1_value, param2_key: param2_value}

                        payload = {param1_key: param1_value}
                        headers = {
                            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0'}

                        response = requests.get(
                            'https://mall.industry.siemens.com/mall/en/ww/Catalog/DatasheetDownload', params = payload,
                            stream = True, verify = False, headers = headers, timeout = 2, allow_redirects = True)

                        cwd = os.path.abspath(os.curdir)

                        if img_file_name.find(' ') != -1:
                            file_name = system_prefix + img_file_name if platform.system() == 'Windows' \
                                else cwd + system_prefix + img_file_name.replace(' ', r'_')
                        else:
                            file_name = system_prefix + img_file_name

                        with open(system_prefix + excel.excel_product_names[i] + image_multiplier + '.' + img_suffix,
                                  'wb') as out_file:
                            shutil.copyfileobj(response.raw, out_file)

                        # open(system_prefix + excel.excel_product_names[i] + image_multiplier + img_suffix, 'wb').write(
                        #     response.content)
                        # print('downloading image: ', excel.excel_product_image_url[i] + ' index:' + str(i))
                        # img_file_names.append(img_file_name)

                        # with open(file_name, 'wb') as out_file:
                        #     shutil.copyfileobj(response.raw, out_file)
                        #     print('downloading image: ', excel.excel_product_image_url[i] + ' index:' + str(i))
                        # img_file_names.append(img_file_name)

                        # except FileNotFoundError:
                        #     print('error file not found')
                        #     img_file_names.append('n/a')
                    except (HTTPError, URLError, SocketError, ConnectionError) as e:
                        img_file_names.append('n/a')
                else:
                    print('error2')
                    img_file_names.append('n/a')


def down_loop():
    global image_file_names
    for i in range(len(read_urls_excel_cell_letters)):
        excel.read_image_urls(read_urls_excel_cell_letters[i])
        download_images(image_file_names, f'_{i}')
        excel.write_file_image_to_excel(image_file_names, image_names_cell_letters[i])
        image_file_names = []


down_loop()
