#!/usr/bin/python
import pandas
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import numpy as np
import sys
import os
import datetime
import re
from time import sleep
from calendar import monthrange
from math import isnan
import json

def raise_in_ternary(error):
    raise error

def get_qtd_days_in_month(month_index):
    date_of_search_month = datetime.datetime(datetime.datetime.now().year, month_index + 1, datetime.datetime.now().day)
    _, qtd_days_in_month = monthrange(date_of_search_month.year, date_of_search_month.month)
    return qtd_days_in_month

def get_months_br():
    months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    return months

def get_excel(arq_name):
    return pandas.ExcelFile(arq_name);

def get_header_VPN():
    data_to_value = lambda v: datetime.datetime.strptime(str(v), '%Y-%m-%d %H:%M:%S')
    str_to_value = lambda v: '' if str(v).upper() == 'NAN' else str(v)

    return {
        'Data': data_to_value,
        'Chegada': str_to_value,
        'Início Almoço': str_to_value,
        'Fim Almoço': str_to_value,
        'Saída': str_to_value,
        'Total do Dia': str_to_value,
        'Descontar': str_to_value, 
        'Extras': str_to_value,
        'Descrição': str_to_value
    }

def get_configs_VPN(month_index):
    headers = get_header_VPN()
    qtd_days_in_month = get_qtd_days_in_month(month_index)
    return { 'header': 7, 'names': headers.keys(), 'converters': headers, 'nrows': qtd_days_in_month }

def get_excel_sheet(excel, args_to_excel_header=None, month_index=None):
    month_name = get_months_br()[month_index if not month_index is None else datetime.datetime.now().month - 1]
    args_excel = args_to_excel_header if not args_to_excel_header is None else get_configs_VPN(month_index if not month_index is None else datetime.datetime.now().month - 1)
    excel_sheet_month = excel.parse(sheet_name=month_name, **args_excel)
    return excel_sheet_month

def get_total_hours_worked(total_hour_str):
    total_hour_dt = datetime.datetime.strptime(str(total_hour_str), '%H:%M:%S')
    return total_hour_dt.hour + total_hour_dt.minute / 60

def get_obj_of_day(date, total_hour):
    return {
        'date': date.strftime('%d/%m/%Y'),
        'total': 0 if total_hour == '' or (str(total_hour).isdigit() and isnan(total_hour)) else get_total_hours_worked(total_hour)
    }

def get_tuple_ziped_VPN(excel_sheet):
    return zip(excel_sheet["Data"], excel_sheet["Total do Dia"])

def get_obj_hours_per_months(excel_sheet, get_excel_zip=lambda all_rows: all_rows, convert_row=lambda v: v):
    list_of_days = []
    for excel_row_data, excel_row_total in get_excel_zip(excel_sheet):
        list_of_days.append(convert_row(excel_row_data, excel_row_total))
    return list_of_days

def get_selenium():
    web_driver = webdriver.Chrome(ChromeDriverManager().install())
    return web_driver

def browser_to_vpn_site(web_driver):
    tasks_VPN_link = 'https://venhapranuvem.sharepoint.com/sites/pwa/_layouts/15/pwa/Timesheet/MyTSSummary.aspx'
    web_driver.get(tasks_VPN_link)

def acess_list_of_days(web_driver_browsed):
    WebDriverWait(web_driver_browsed, 1000).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.XmlGridTable')))
    return web_driver_browsed.find_elements_by_css_selector('.XmlGridTable tbody tr:not(.XmlGridTitleRow) td a')

def login_vpn(web_driver_browsed, email, password):
    WebDriverWait(web_driver_browsed, 1000).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[value=Avançar]')))
    web_driver_browsed.find_element_by_css_selector('input[name=loginfmt]').send_keys(email)
    web_driver_browsed.find_element_by_css_selector('input[value=Avançar]').click()
    WebDriverWait(web_driver_browsed, 1000).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[value=Entrar]')))
    web_driver_browsed.find_element_by_css_selector('input[type=password][name=passwd]').send_keys(password)
    web_driver_browsed.find_element_by_css_selector('input[value=Entrar]').click()
    WebDriverWait(web_driver_browsed, 1000).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type=submit][value=Sim]')))
    web_driver_browsed.find_element_by_css_selector('input[type=submit][value=Sim]').click()

def get_element_indicating_if_filled(elements):
        get_next_element_text = False
        list_of_dates = []
        regex_search_date = re.compile(r'\((\d{2}\/\d{2}\/\d{4})\s*\-\s*(\d{2}\/\d{2}\/\d{4})\)')
        date_str_to_date = lambda date_str: datetime.datetime.strptime(date_str, '%d/%m/%Y')
        for element in elements:
            if get_next_element_text:
                start_date, end_date = regex_search_date.search(str(element.text)).groups()
                list_of_dates.append((date_str_to_date(start_date), date_str_to_date(end_date)))
                get_next_element_text = False

            if(str(element.text).upper().replace(' ', '') == 'CLIQUE PARA CRIAR'.replace(' ', '')):
                get_next_element_text = True
        return list(filter(lambda dates: datetime.datetime.now() > dates[1], list_of_dates))

def get_months_to_be_searched(list_of_days):
    copy_of_list_of_days = list(list_of_days)
    copy_of_list_of_days.sort(key = lambda dates: dates[0])
    (first_date, _), (_, last_date) = copy_of_list_of_days[0], copy_of_list_of_days[len(list_of_days) - 1]
    first_date = datetime.datetime(first_date.year, first_date.month, 1)
    last_date = datetime.datetime(last_date.year, last_date.month, get_qtd_days_in_month(last_date.month - 1))
    month_list = pandas.date_range(start=first_date, end=last_date, freq='MS').to_pydatetime().tolist()
    return month_list

def get_all_excels_sheets_by_months(excel, months):
    return [item for sublist in list(
                map(lambda month_date: 
                    get_obj_hours_per_months(
                        get_excel_sheet(
                            excel, month_index=month_date.month-1), 
                            get_tuple_ziped_VPN, 
                            get_obj_of_day), 
                months)) for item in sublist]

def get_number_of_child_of_element_search(web_driver_browsed):
    web_driver_browsed.find_element_by_css_selector

def scroll_to_end(web_driver_browsed):
    vertical_scroll_button = web_driver_browsed.find_element_by_css_selector('div[id$=GridControl] > div:nth-child(4) > div:nth-child(4) > div > div > div:nth-child(5) > div.vert-scroll-bar-arrow-cell.scroll-bar-arrow-cell > div')
    for _ in range(12):
        vertical_scroll_button.click() 

def scroll_to_right(web_driver_browsed):
    return
    try:
        horizontal_scroll = web_driver_browsed.find_elements_by_css_selector('div[id$=rightpane] > div:nth-child(2) > div > div > div:nth-child(3) > div.horiz-scroll-bar-outer-box.scroll-bar-outer-box')
        last_element_in_page_width_right = (web_driver_browsed.get_window_size())['width']
        webdriver.ActionChains(web_driver_browsed).click_and_hold(horizontal_scroll).pause(2).move_by_offset(last_element_in_page_width_right, 0).perform()
    except:
        pass

def get_hour_by_index_of_week(dates, excel_sheet, week_number):
    sunday, saturday = tuple(map(lambda date: date.strftime('%d/%m/%Y'), dates))
    dates_of_week_list = pandas.date_range(start=sunday, end=saturday, freq="D").to_pydatetime().tolist()
    date_searched = dates_of_week_list[week_number - 1].strftime('%d/%m/%Y')
    total_hours = [row['total'] for row in excel_sheet if row['date'] == date_searched][0]
    return str(round(total_hours, 1)).replace('.', ',')

def get_row_to_fulfill_from_cliqueAqui_button_VPN(web_driver_browsed, cliqueAqui_button, get_number_of_child_of_element_search):
    webdriver.ActionChains(web_driver_browsed).click(cliqueAqui_button).pause(3).perform()
    scroll_to_end(web_driver_browsed)
    number_of_child_of_element_search = str(get_number_of_child_of_element_search(web_driver_browsed))
    WebDriverWait(web_driver_browsed, 1000).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'table[id$=GridControl_rightpane_mainTable] > tbody > tr:nth-child({number_of_child_of_element_search})')))
    row_to_fulfill = web_driver_browsed.find_element_by_css_selector(f'table[id$=GridControl_rightpane_mainTable] > tbody > tr:nth-child({number_of_child_of_element_search})')

def fill_a_week_in_browser(web_driver_browsed, list_of_elements, excel_sheet, tuple_of_days, get_number_of_child_of_element_search=lambda wd: 22):
    #get first elemenet avaible with button wrote "CLIQUE AQUI" and date before today
    element_to_click = next(iter([e for i, e in enumerate(list_of_elements) if str(e.text).upper().replace(' ', '') == 'CLIQUE PARA CRIAR'.replace(' ', '') and i < len(list_of_elements) - 1 and tuple_of_days[0].strftime('%d/%m/%Y') in str(list_of_elements[i + 1].text)]), None)

    if not element_to_click is None:
        row_to_fulfill = get_row_to_fulfill_from_cliqueAqui_button_VPN(web_driver_browsed, element_to_click, get_number_of_child_of_element_search)
        for index, day_of_week_row in enumerate(list(row_to_fulfill.find_elements_by_css_selector('td'))):
            if index == 0:
                continue
            else:
                scroll_to_right(web_driver_browsed)

            hour_str = get_hour_by_index_of_week(tuple_of_days, excel_sheet, week_number=index)
            day_of_week_row.click()
            day_of_week_row.click()
            WebDriverWait(web_driver_browsed, 1000).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[id^=jsgrid_editbox]')))
            input_to_send = web_driver_browsed.find_element_by_css_selector('input[id^=jsgrid_editbox]')
            webdriver.ActionChains(web_driver_browsed).pause(1).send_keys_to_element(input_to_send, hour_str).pause(1).perform()

def send_data_to_VPN(web_driver_browsed):
    webdriver.ActionChains(web_driver_browsed).pause(2).key_down(Keys.CONTROL).key_down(Keys.LEFT_SHIFT).send_keys('S').perform()
    webdriver.ActionChains(web_driver_browsed).pause(3).send_keys(Keys.ENTER).pause(3).perform()

def get_config_json():
    arq = None
    try:
        arq = open('login_config.json', 'r')
        config_json = json.load(arq)
        return (config_json['email'], config_json['password'])
    except:
        print('Erro ao obter arquivo de login!')
    finally:
        if not arq is None:
            arq.close()

def main():
    excel_folha_de_pontos = None
    try:
        name_arq_excel = sys.argv[1] if len(sys.argv) > 1 else raise_in_ternary(ValueError("Necessita caminho do arquivo excel!"))
        excel_folha_de_pontos = get_excel(name_arq_excel)

        selenium_wd = None
        try:
            selenium_wd = get_selenium()
            selenium_wd.maximize_window()
            browser_to_vpn_site(selenium_wd)
            email, password = get_config_json()
            login_vpn(selenium_wd, email, password)

            list_of_elements = acess_list_of_days(selenium_wd)
            list_of_days = get_element_indicating_if_filled(list_of_elements)

            months = get_months_to_be_searched(list_of_days)
            excels_sheets = get_all_excels_sheets_by_months(excel_folha_de_pontos, months)
            excel_folha_de_pontos.close()

            for week in list_of_days:
                fill_a_week_in_browser(selenium_wd, list_of_elements, excels_sheets, week)
                send_data_to_VPN(selenium_wd)
                selenium_wd.back()
                list_of_elements = acess_list_of_days(selenium_wd)
            
            selenium_wd.close()
        except Exception as ex:
            print(ex)
            print("Erro ao usar selenium!")
        finally:
            if not selenium_wd is None:
                selenium_wd.close()
    except Exception as ex:
        print(ex)
        print("Erro ao obter arquivo excel!")
    finally:
        if not excel_folha_de_pontos is None:
            excel_folha_de_pontos.close()

if __name__ == "__main__":
    main()