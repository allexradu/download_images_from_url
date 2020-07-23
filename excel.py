import openpyxl
import platform
import random
import string

from openpyxl import load_workbook
import extra_functions

from openpyxl.utils.exceptions import IllegalCharacterError

excel_product_image_url = []
excel_product_names = []
work_sheet_index = 2
table_location = 'excel\\a2.xlsx' if platform.system() == 'Windows' else 'excel/a2.xlsx'


def sanitise_product_names(string_text):
    """ Replacing all the bad characters that inhibit search"""
    if string_text is not None:
        replace_commas = string_text.replace(',', '') if string_text.find(',') != -1 else string_text
        replace_stars = replace_commas.replace('*', ' ') if replace_commas.find('*') != -1 else replace_commas
        replace_slashes = replace_stars.replace(r'/', ' ') if replace_stars.find(r'/') != -1 else replace_stars
        replace_dollar_signs = replace_slashes.replace(' $', '') if replace_slashes.find(
            ' $') != -1 else replace_slashes
        replace_plus_sign = replace_dollar_signs.replace('+', ' ') if replace_dollar_signs.find(
            '+') != -1 else replace_dollar_signs
        replace_dots = replace_plus_sign.replace('.', ' ') if replace_plus_sign.find('.') != -1 else replace_plus_sign
        replace_dashes = replace_dots.replace('-', ' ') if replace_dots.find('-') != -1 else replace_dots
        replace_small = replace_dashes.replace('<', ' ') if replace_dashes.find('<') != -1 else replace_dashes
        replace_big = replace_small.replace('>', ' ') if replace_small.find('>') != -1 else replace_small
        replace_percentage = replace_big.replace(r"%", ' ') if replace_big.find(r"%") != -1 else replace_big
        replace_pipes = replace_percentage.replace('|', ' ') if replace_percentage.find(
            '|') != -1 else replace_percentage
        replace_check_marks = replace_pipes.replace('✅', ' ') if replace_pipes.find('✅') != -1 else replace_pipes
        replace_double_quotes = replace_check_marks.replace('"', '') if replace_check_marks.find(
            '"') != 1 else replace_check_marks
        replace_colon = replace_double_quotes.replace(':', '') if replace_double_quotes.find(
            ':') != -1 else replace_double_quotes
        replace_question_mark = replace_colon.replace('?', '') if replace_colon.find('?') != -1 else replace_colon
        replace_back_slash = replace_question_mark.replace('\\', '') if replace_question_mark.find(
            '\\') != 1 else replace_check_marks
        return replace_back_slash
    else:
        def get_random_string(length):
            letters = string.ascii_lowercase
            return ''.join(random.choice(letters) for i in range(length))

        return get_random_string(8)


def write_file_image_to_excel(image_file_names, image_names_cell_letter):
    global work_sheet_index

    work_sheet_index = 2

    wb = openpyxl.load_workbook(table_location)
    ws = wb.active

    for i in range(len(image_file_names)):
        print('wbind: ', work_sheet_index)
        product_brand_key = extra_functions.value_key(image_names_cell_letter, work_sheet_index)
        ws[product_brand_key] = image_file_names[i]

        work_sheet_index += 1

    wb.save(table_location)


def sanitise_string(string):
    if isinstance(string, str):
        sanitised_string = string.encode('unicode_escape').decode('utf-8')
        return sanitised_string
    else:
        return string


def read_image_urls(cell_letter):
    global excel_product_image_url
    get_work_sheet_index()
    excel_product_image_url = get_all_the_rows_from_column(cell_letter)


def read_product_names(cell_letter):
    global excel_product_names
    get_work_sheet_index()
    excel_product_names = get_all_the_rows_from_column(cell_letter)


def get_all_the_rows_from_column(cell_letter):
    wb = load_workbook(table_location)  # Work Book
    ws = wb.get_sheet_by_name('Sheet1')  # Work Sheet
    column = ws[cell_letter]  # Column
    column_list = [column[x].value for x in range(len(column))]
    return column_list


def get_work_sheet_index():
    global work_sheet_index
    wb = load_workbook(table_location)  # Work Book
    ws = wb.get_sheet_by_name('Sheet1')  # Work Sheet
    column = ws['A']  # Column
    work_sheet_index = len(column) + 1
    print('Worksheet index: ', work_sheet_index)
