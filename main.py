#!/usr/bin/python
# -*- coding:utf-8 -*-

import xlrd, arrow, urllib, re
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


def openexcel(file):
    """
    open excel file
    :param file: excel file
    :return: excelojb
    """
    try:
        book = xlrd.open_workbook(file)
        return book
    except Exception as e:
        print "open excel file failed" + str(e)


def readsheets(file):
    """
    read sheet
    :param file: excel obj
    :return: sheet obj
    """
    try:
        book = openexcel(file)
        sheet = book.sheets()
        return sheet
    except Exception as e:
        print "read sheet failed" + str(e)


def readdata(sheet, n=0):
    """
    data read
    :param sheet: excel sheet
    :param n: rows
    :return: data list
    """
    dataset = []
    for r in range(sheet.nrows):
        col = sheet.cell(r, n).value
        # 如果有表头
        if r != 0:
            dataset.append(col)
    return dataset


def browserdriver():
    """
    start driver
    :return: driver obj
    """
    dcap = dict(DesiredCapabilities.PHANTOMJS)
    dcap["phantomjs.page.settings.userAgent"] = (
        "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36"
    )
    driver = webdriver.PhantomJS(executable_path='lib/phantomjs', desired_capabilities=dcap)
    return driver


def tyc_data(driver, url, keyword):
    """
    get page source code
    :param driver: brower
    :param url: url
    :param keyword: keyword
    :return: tyc date
    """
    driver.get(url)
    try:
        element = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CLASS_NAME, "new-foot-v1"))
        )
    except Exception as e:
        print e
    finally:
        source = driver.page_source.encode("utf-8")
        tycsoup = BeautifulSoup(source, 'html.parser')
        name = tycsoup.select(
            "div.search_result_single > div.search_right_item > div > a.query_name > span > em")
        hname = tycsoup.select(
            "div.search_result_single > div.search_right_item > div.search_row_new > div > div.add > span.over-hide > em"
        )
        cmname = name[0].text if len(name) > 0 else None
        hiscmsname = hname[0].text if len(hname) > 0 else None
        if cmname == keyword or hiscmsname == keyword:
            company_url = tycsoup.select('div.search_result_single > div.search_right_item > div > a.query_name')[0].get('href')
            driver.get(company_url)
            try:
                element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "new-foot-v1"))
                )
            except Exception as e:
                print e
            finally:
                source = driver.page_source.encode("utf-8")
                tycdata = BeautifulSoup(source, 'html.parser')
                lpblock = tycdata.select("div.human-top > div > div > a")
                lpname = lpblock[0].text
                cpstatus = tycdata.find_all("div", class_=re.compile(r"\bstatusType\d"))[0].text
                print lpname, cpstatus




def main(logfile, excelfile):
    try:
        driver = browserdriver()
    except Exception as e:
        print e
    now = arrow.now()
    for sheet in readsheets('cxgs.xlsx'):
        for cmyname in readdata(sheet):
            keyword = urllib.quote(cmyname.encode("utf-8"))
            tycurl = "https://www.tianyancha.com/search?key=" + keyword + "&checkFrom=searchBox"
            tyc_data(driver, tycurl, cmyname)


if __name__ == '__main__':
    logfile = 'log.txt'
    excel = 'cxgs.xlsx'
    main(logfile, excel)
