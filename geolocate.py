import re
import math
import pickle
from time import sleep
from os.path import join
from urllib.parse import quote
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from xlsx_read import beep, PICKLE_PATH

SEARCH_PATH = 'https://yandex.ru/maps/213/moscow/?mode=search&text='
COORD_PATH = 'meta[property="og:image"]'

ONLY_RES_PATH = 'div[class="clipboard__content"]'
MULTI_RES_PATH = 'div[class="search-snippet-view__title"]'
NO_RES_PATH = 'div[class="search-list-view__warning"]'
COLS_XLSX = ['addressid', 'clientid', 'city', 'address', 'posname', 'detect', 'near_dspv']

MATCHED_XLSX = join('data', 'matched.xlsx')
N_FROM = 0
N_TO = 9999

ch_list, np_list = [], []


def pickle_to_lists():
    # load lists from pickled file
    global ch_list, np_list    
    with open(PICKLE_PATH, 'rb') as f:
        ch_list = pickle.load(f)
        np_list = pickle.load(f)


def lists_to_xlsx():
    # save lists as excel file
    wb = Workbook()
    ws = wb.active
    ws.append(COLS_XLSX)
    for np in np_list[N_FROM:N_TO]:
        ws.append([np[key] for key in COLS_XLSX])
    wb.save(filename = MATCHED_XLSX)
    wb.close()


def detect_coords():
    # get coordinates of new points
    driver = webdriver.Chrome()
    for i, np in enumerate(np_list[N_FROM:N_TO]):
        verbose_addr = ' '.join((np['city'].split())[:-1]) + ' ' + np['address']
        shallow_addr = ' '.join((np['city'].split())[:-1])
        np['detect'] = True
        for addr in (verbose_addr, shallow_addr):
            while True:
                try:
                    driver.get(SEARCH_PATH + quote(addr))
                except:
                    sleep(5)
                else:
                    break
            sleep(6)
            try:
                elem = driver.find_element_by_css_selector(ONLY_RES_PATH)
            except:
                try:
                    driver.find_element_by_css_selector(NO_RES_PATH)
                except:
                    try:
                        driver.find_element_by_css_selector(MULTI_RES_PATH).click()
                        elem = driver.find_element_by_css_selector(ONLY_RES_PATH)
                    except:
                        pass
                    else:
                        break
            else:
                break
        else:
            np['detect'] = False
        print('{:>8}) {} {} '.format(i, np['city'], np['address']), end='')
        if np['detect']:
            np['lon'] = float(elem.text.split(sep=', ')[0])
            np['lat'] = float(elem.text.split(sep=', ')[1])
            print('was detected')
        else:
            print('was not detected')
    driver.close()


def match_points():
    # match np with nearest dspv
    for np in np_list[N_FROM:N_TO]:
        if not np['detect']:
            np['near_dspv'] = ''
            continue
        min_dist = 9999
        for ch in ch_list:
            if ch['detect']:
                dist = math.sqrt((np['lat'] - ch['lat'])**2 + (np['lon'] - ch['lon'])**2)
                if  dist < min_dist:
                    min_dist = dist
                    np['near_dspv'] = ch['dspv2']


def main():
    pickle_to_lists()
    beep(1)
    detect_coords()
    beep(2)
    match_points()
    beep(3)
    lists_to_xlsx()
    beep(4)


if __name__ == '__main__':
    main()
