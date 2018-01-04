#!/usr/bin/python
# -*- coding:utf-8 -*-

import xlrd, arrow, urllib, re, os, requests
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from openpyxl.workbook import Workbook
import random, time
from PIL import Image, ImageFont, ImageDraw
import pytesseract


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
        print("open excel file failed", str(e))


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
        print("read sheet failed", str(e))


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
    ua = "Mozilla/5.0 (Linux; Android 4.4.2; Nexus 4 Build/KOT49H) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.114 Mobile Safari/537.36"
    profile = webdriver.FirefoxProfile()
    profile.set_preference(
        "general.useragent.override", ua)
    driver = webdriver.Firefox(profile, executable_path='lib/geckodriver.exe')
    return driver


# def browserdriver():
    """
    start driver
    :return: driver obj
    """
    """
    dcap = DesiredCapabilities.PHANTOMJS.copy()
    dcap['phantomjs.page.customHeaders.Referer'] = 'https://www.baidu.com/'
    dcap["phantomjs.page.settings.userAgent"] = 'Mozilla/5.0 (Linux; Android 4.4.2; Nexus 4 Build/KOT49H) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.114 Mobile Safari/537.36'
    dcap["phantomjs.page.settings.loadImages"] = False
    dcap["phantomjs.page.customHeaders.User-Agent"] = 'Mozilla/5.0 (Linux; Android 4.4.2; Nexus 4 Build/KOT49H) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.114 Mobile Safari/537.36'
    dcap["phantomjs.page.settings.disk-cache"] = True
    driver = webdriver.PhantomJS(executable_path='lib/phantomjs', desired_capabilities=dcap)
    return driver
"""


def request(url):
    header = {
        "Referer": "https://www.baidu.com/",
        "User-Agent": "Mozilla/5.0 (Linux; Android 4.4.2; Nexus 4 Build/KOT49H) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.114 Mobile Safari/537.36"
    }
    r = requests.get(url, headers=header)
    return r.text


def tyc_data(driver, url, keyword, maping):
    """
    get Tianyancha Data
    :param driver: brower
    :param url: url
    :param keyword: keyword
    :return: tyc date
    """
    fails = 1
    while fails < 31:
        try:
            driver.set_page_load_timeout(10)
            driver.get("about:blank")
            driver.get(url)
            break
        except Exception as e:
            print(e)
            print('error wait 30 secs tyc_data1')
            time.sleep(30)
    else:
        raise
    try:
        WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CLASS_NAME, "footerV2"))
        )
    except Exception as e:
        print(e)
    finally:
        source = driver.page_source.encode("utf-8")
        tycsoup = BeautifulSoup(source, 'html.parser')
        name = tycsoup.select(
            "div > div > div > div > a.query_name > span > em")
        cmname = name[0].text if len(name) > 0 else None
        print(cmname)
        if cmname == keyword:
            company_url = "https://m.tianyancha.com" + tycsoup.select('div > div > div > div > a.query_name')[0].get('href')
            print(company_url)
            fails = 1
            while fails < 31:
                try:
                    driver.set_page_load_timeout(10)
                    driver.get("about:blank")
                    driver.get(company_url)
                    break
                except Exception as e:
                    print(e)
                    print('error wait 30 secs tyc_data2')
                    time.sleep(30)
            else:
                raise
            try:
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "footerV2"))
                )
            except Exception as e:
                print(e)
            finally:
                noinfo = ""
                source = driver.page_source.encode("utf-8")
                tycdata = BeautifulSoup(source, 'html.parser')
                binfo = []
                reginfo = tycdata.select("div.item-line > span")
                bsocplist = tycdata.select("div.item-line > span > span > span.hidden > div")
                bscop = bsocplist[0].get_text() if len(bsocplist) > 0 else None
                if (reginfo[11].get_text().encode('utf-8')) == "民办非企业单位":
                    binfo = [
                        cmname,
                        noinfo,
                        reginfo[1].get_text() if len(reginfo[1].get_text()) > 0 else None,
                        # regdecode(maping, reginfo[3].get_text()) if len(reginfo[3].get_text()) > 0 else None,
                        # regdecode(maping, reginfo[5].get_text()) if len(reginfo[5].get_text()) > 0 else None,
                        reginfo[3].get_text() if len(reginfo[3].get_text()) > 0 else None,
                        reginfo[5].get_text() if len(reginfo[5].get_text()) > 0 else None,
                        noinfo,
                        reginfo[7].get_text(),
                        noinfo,
                        reginfo[9].get_text(),
                        reginfo[11].get_text(),
                        noinfo,
                        noinfo,
                        noinfo,
                        reginfo[13].get_text(),
                        noinfo,
                        noinfo
                    ]
                else:
                    binfo = [
                        cmname,
                        reginfo[3].get_text() if len(reginfo[3].get_text()) > 0 else None,
                        reginfo[1].get_text() if len(reginfo[1].get_text()) > 0 else None,
                        regdecode(maping, reginfo[7].get_text()) if len(reginfo[7].get_text()) > 0 else None,
                        regdecode(maping, reginfo[5].get_text()) if len(reginfo[5].get_text()) > 0 else None,
                        regdecode(maping, reginfo[23].get_text()) if len(reginfo[23].get_text()) > 0 else None,
                        reginfo[13].get_text() if len(reginfo[13].get_text()) > 0 else None,
                        reginfo[15].get_text() if len(reginfo[15].get_text()) > 0 else None,
                        reginfo[17].get_text() if len(reginfo[17].get_text()) > 0 else None,
                        reginfo[11].get_text() if len(reginfo[11].get_text()) > 0 else None,
                        reginfo[19].get_text() if len(reginfo[19].get_text()) > 0 else None,
                        reginfo[9].get_text() if len(reginfo[9].get_text()) > 0 else None,
                        reginfo[21].get_text() if len(reginfo[21].get_text()) > 0 else None,
                        reginfo[25].get_text() if len(reginfo[25].get_text()) > 0 else None,
                        reginfo[27].get_text() if len(reginfo[27].get_text()) > 0 else None,
                        bscop
                    ]
                for x in binfo:
                    print(x)
                return binfo
        else:
            print('暂无信息')
            binfo = [keyword, "暂无信息"]
            return binfo



def gettycfont():
    fails = 1
    while fails < 31:
        try:
            source = request("https://m.tianyancha.com/")
            break
        except Exception as e:
            print(e)
            print('error wait 30 secs getfont1')
            time.sleep(30)
    else:
        raise
    pagesouup = BeautifulSoup(source, 'html.parser')
    csssheet = pagesouup.find_all("link", rel="stylesheet")
    for link in csssheet:
        csshref = link.get('href')
        if 'main' in csshref:
            print(csshref)
            fails = 1
            while fails < 31:
                try:
                    csscode = request(csshref)
                    break
                except Exception as e:
                    print(e)
                    print('error wait 30 secs getfont2')
                    time.sleep(30)
            else:
                raise
            recode = re.findall(r"\btyc-num-\w*.ttf", csscode)
            if recode[0] is not None:
                print('download font')
                fonturl = "https://static.tianyancha.com/m-require-js/public/fonts/"+recode[0]
                urllib.urlretrieve(fonturl, recode[0])
            print(recode[0])
            return getmaping(recode[0]), recode[0]
        else:
            pass


def fonttest(fontfile):
    fails = 1
    while fails < 31:
        try:
            source = request("https://m.tianyancha.com/")
            break
        except Exception as e:
            print(e)
            print('error wait 30 secs fonttest1')
            time.sleep(30)
    else:
        raise
    pagesouup = BeautifulSoup(source, 'html.parser')
    csssheet = pagesouup.find_all("link", rel="stylesheet")
    for link in csssheet:
        csshref = link.get('href')
        if 'main' in csshref:
            print(csshref)
            fails = 1
            while fails < 31:
                try:
                    csscode = request(csshref)
                    break
                except Exception as e:
                    print(e)
                    print('error wait 30 secs fonttest2')
                    time.sleep(30)
            else:
                raise
            recode = re.findall(r"\btyc-num-\w*.ttf", csscode)
            if recode[0] is not None:
                print recode[0]
                if recode[0] == fontfile:
                    return True
                else:
                    return False


def getmaping(fontfile):
    text = r" 0 1 2 3 4 5 6 7 8 9 . "
    im = Image.new("RGB", (1000, 100), (255, 255, 255))
    dr = ImageDraw.Draw(im)
    font = ImageFont.truetype(fontfile, 32)
    dr.text((10, 10), text, font=font, fill="#000000")
    im.show()
    im.save("t.png")
    fontimage = Image.open("t.png")
    numlist = tuple(pytesseract.image_to_string(fontimage))
    print(pytesseract.image_to_string(fontimage))
    mapfont = {
        '0': str(numlist[0]),
        '1': str(numlist[1]),
        '2': str(numlist[2]),
        '3': str(numlist[3]),
        '4': str(numlist[4]),
        '5': str(numlist[5]),
        '6': str(numlist[6]),
        '7': str(numlist[7]),
        '8': str(numlist[8]),
        '9': str(numlist[9]),
        '.': str(numlist[10])
    }
    return mapfont


def regdecode(mapfont, regstr):
    strlist = list(regstr)
    regdata = []
    for stra in strlist:
        if stra in mapfont.keys():
            regdata.append(mapfont[stra])
        else:
            regdata.append(stra)
    return "".join(regdata)
        

def main(logfile, excelfile):
    try:
        driver = browserdriver()
    except Exception as e:
        print(e)
    now = arrow.now()
    maping, fontname = gettycfont()
    newexcelfile = "" + arrow.now().format("YYYY-MM-DD HH_mm_ss") + ".xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append([
        "公司名称", "公司状态", "法人名称", "注册资本", "注册时间", "核准时间", "工商注册号", "组织机构代码", "信用识别代码",
        "公司类型", "纳税人识别号", "行业", "营业期限", "登记机关", "注册地址", "经营范围"
    ])
    """
    keyword = "宁波凌科教育培训学校"
    tycurl = "https://m.tianyancha.com/search?key=" + keyword + "&checkFrom=searchBox"
    tyc_data(driver, tycurl, keyword, maping)
    """
    dcount = 0
    for sheet in readsheets('cxgs.xlsx'):
        for cmyname in readdata(sheet):
            print(dcount)
            if fonttest(fontname):
                print("encrypt not change")
            else:
                maping, fontname = gettycfont()
            keyword = urllib.quote(cmyname.encode("utf-8"))
            tycurl = "https://m.tianyancha.com/search?key=" + keyword + "&checkFrom=searchBox"
            binfo = tyc_data(driver, tycurl, cmyname, maping)
            if binfo is None:
                print("Error", binfo, "is None")
                pass
            else:
                ws.append(binfo)
                wb.save(filename=newexcelfile)
            a = random.randint(0, 0)
            print('采集完毕，等待' + str(a) + '秒')
            time.sleep(a)
            dcount += 1
            if dcount == 150:
                try:
                    driver.quit()
                except Exception as e:
                    print(e)
                finally:
                    dcount = 0
                    driver = browserdriver()
            else:
                pass
    wb.save(filename=newexcelfile)
    driver.quit()


if __name__ == '__main__':
    logfile = 'log.txt'
    excel = 'cxgs.xlsx'
    main(logfile, excel)
