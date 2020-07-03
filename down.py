import excel
import platform
from urllib.error import HTTPError
from urllib.error import URLError
from urllib.request import urlretrieve
from socket import error as SocketError
from urllib.request import Request, urlopen
import requests
import shutil

read_urls_excel_cell_letter = 'T'
read_product_name_cell_letter = 'A'
image_names_cell_letter = "U"

image_file_names = []

excel.read_image_urls(read_urls_excel_cell_letter)
excel.read_product_names(read_product_name_cell_letter)


def download_images(img_file_names):
    for i in range(1, len(excel.excel_product_image_url)):
        if excel.excel_product_image_url[i] != '':
            if excel.excel_product_image_url[i] != 'none':
                if excel.excel_product_image_url[i] is not None:
                    try:
                        url = excel.excel_product_image_url[i]
                        split_url = url.split('/')
                        img_file_name_raw = split_url[-1]
                        split_file_name = img_file_name_raw.split('.')
                        img_suffix = split_file_name[-1]
                        img_file_name = excel.sanitise_product_names(excel.excel_product_names[i]) + '.' + img_suffix

                        img_file_names.append(img_file_name)

                        system_prefix = 'excel\\photos\\' if platform.system() == 'Windows' else 'excel/photos/'

                        response = requests.get(excel.excel_product_image_url[i], stream = True)
                        with open(system_prefix + img_file_name, 'wb') as out_file:
                            shutil.copyfileobj(response.raw, out_file)
                            print('downloading image: ', excel.excel_product_image_url[i] + ' index:' + str(i))
                    except FileNotFoundError or HTTPError or URLError or SocketError:
                        img_file_names.append('n/a')
                else:
                    img_file_names.append('n/a')


download_images(image_file_names)

excel.write_file_image_to_excel(image_file_names, image_names_cell_letter)
