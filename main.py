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
from openpyxl.workbook import Workbook
import random, time

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
    get Tianyancha Data
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
                reginfo = tycdata.select(
                    "div#_container_baseInfo > div > div > table.companyInfo-table > tbody > tr > td > div > div > div.baseinfo-module-content-value > text.tyc-num"
                )
                cpinfo = tycdata.select(
                    "div#_container_baseInfo > div > div.base0910 > table.companyInfo-table > tbody > tr > td"
                )
                lineob = tycdata.select(
                    "div#_container_baseInfo > div > div.base0910 > table.companyInfo-table > tbody > tr > td > span > span > span.hidden"
                )
                print keyword
                print lpname
                print cpstatus
                binfo = [
                    keyword, cpstatus, lpname
                ]
                for regdata in reginfo:
                    print regdecode(regdata.text)
                    binfo.append(regdecode(regdata.text))
                print regdecode(cpinfo[16].text)
                binfo.append(regdecode(cpinfo[16].text))
                for a in [1, 3, 6, 8, 10, 12, 14, 18, 22]:
                    print cpinfo[a].text
                    binfo.append(cpinfo[a].text)
                print lineob[0].text
                binfo.append(lineob[0].text)
                return binfo
        else:
            print '暂无信息'
            binfo = [keyword, "暂无信息"]
            return binfo


def regdecode(str):
    codedict = {
        '5': '.',
        '6': '0',
        '8': '1',
        '.': '2',
        '1': '3',
        '0': '4',
        '9': '5',
        '3': '6',
        '2': '7',
        '4': '8',
        '7': '9'
    }
    strlist = list(str)
    regdata = []
    for stra in strlist:
        if stra in codedict.keys():
            regdata.append(codedict[stra])
        else:
            print 'error'
            regdata.append(stra)
    return "".join(regdata)


def main(logfile, excelfile):
    try:
        driver = browserdriver()
    except Exception as e:
        print e
    now = arrow.now()
    newexcelfile =  "" + arrow.now().format("YYYY-MM-DD HH_mm_ss") + ".xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append([
        "公司名称", "公司状态", "法人名称", "注册资本", "注册时间", "核准时间", "工商注册号", "组织机构代码", "信用识别代码",
        "公司类型", "纳税人识别号", "行业", "营业期限", "登记机关", "注册地址", "经营范围"
    ])
    for sheet in readsheets('cxgs.xlsx'):
        for cmyname in readdata(sheet):
            keyword = urllib.quote(cmyname.encode("utf-8"))
            tycurl = "https://www.tianyancha.com/search?key=" + keyword + "&checkFrom=searchBox"
            binfo = tyc_data(driver, tycurl, cmyname)
            if binfo is None:
                pass
            else:
                ws.append(binfo)
            a = random.randint(10, 120)
            print "采集完毕，等待" + str(a) + "秒"
            time.sleep(a)
    wb.save(filename=newexcelfile)



if __name__ == '__main__':
    logfile = 'log.txt'
    excel = 'cxgs.xlsx'
    main(logfile, excel)
