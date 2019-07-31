# -*-coding:utf-8 -*-

import requests
import lxml
from lxml import etree
from pyspider.libs.base_handler import *
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
# available since 2.26.0
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from boss import *
#import xlwt
#from xlwt import Workbook
import cPickle as pickle
import xlsxwriter
from xlsxwriter import Workbook
from urllib2 import urlopen
from io import BytesIO
import traceback
import copy
import sys

g_sxall = []
useXlsxWriter = False
"""
   {"Name": "圣光勇士", "ID": 2265},
  {"Name": "猿猴", "ID": 2263},

    {"Name": "玉火大师", "ID": 2266},

    {"Name": "风灵", "ID": 2271},

    {"Name": "神选着", "ID": 2268},


    {"Name": "大王", "ID": 2272},

    {"Name": "大工匠", "ID": 2276},

    {"Name": "风墙", "ID": 2280},

    {"Name": "吉安娜", "ID": 2281}
   

    
"""
BossList = [
    {"Name": "圣光勇士", "ID": 2265},
    {"Name": "猿猴", "ID": 2263},
    {"Name": "玉火大师", "ID": 2266},

    {"Name": "风灵", "ID": 2271},

    {"Name": "神选着", "ID": 2268},


    {"Name": "大王", "ID": 2272},

    {"Name": "大工匠", "ID": 2276},

    {"Name": "风墙", "ID": 2280},

    {"Name": "吉安娜", "ID": 2281}
 
]
pic_map = {}
attr_map = {}

def show_help():
    print("用法: <WclSpider> <职业> <专精> <0---DPS,1---HPS>")


def set_style(wb, name, height, bold=False, align=None):
    style = wb.add_format()

    if bold:
        style.set_bold()
    if align:
        style.set_align(align)
    return style


def main():
    global attr_map
    global pic_map
    # all_class=[]
    # all_spec=[]
    try:
        wow_class = sys.argv[1]
        wow_spec = sys.argv[2]
        is_hps = sys.argv[3]
        
    except Exception,e:
        print(e)
        traceback.print_exc()
    
    print("class= " + wow_class)
    print("spec = " + wow_spec)
    print("is_hps=" + str(is_hps))
    
    try:
        with open("attr_"+ wow_class +"_"+wow_spec+".txt" , "rb") as f_attr:
            attr_map = pickle.load(f_attr)
    except Exception, e:
        # traceback.print_exc()
        attr_map = {}

    try:
        f = Workbook(wow_class+'_'+wow_spec+'.xlsx')
        pagecnt = 0
        for b in YongHengWangGong_BossList:
            if pagecnt < 10:
                add_sheet(f, b["Name"], b["ID"], wow_class, wow_spec, bool(int(is_hps)))
            pagecnt += 1

    except Exception, e:
        # print(str(e))
        traceback.print_exc()

    finally:
        f.close()


def add_sheet(f, bossName, bossID, wclass, spec, is_heal=False):
    global attr_map
    global pic_map

    print("add a sheet \n")
    
    
    
    try:
        DpsColum = "DPS"
        if is_heal:
            DpsColum = "HPS"

        sheet1 = f.add_worksheet(bossName.decode('utf-8'))
        row0 = ["排名", "名字", "装等", DpsColum, "日期", "时长",
                "主属性", "暴击", "急速", "精通", "全能", "天赋", "特质", "饰品" , "精华"]
        # 0      1       2          3    4       5       6       7       8   9       10      11   12
        # 写第一行
        for i in range(0, len(row0)):
            sheet1.write(0, i, row0[i].decode(
                'utf-8'), set_style(f, 'Times New Roman', 220, True, 'center'))

        driver = webdriver.PhantomJS("C:/windows/system32/phantomjs.exe")
        #chrome_options = Options()
        # chrome_options.add_argument('--headless')
        #driver = webdriver.Chrome(chrome_options=chrome_options)

        URL = "https://cn.warcraftlogs.com/zone/rankings/23#boss=%s&class=%s&spec=%s" % (
            bossID, wclass, spec)

        if is_heal:
            URL += "&metric=hps"

        # load from remote
        driver.get(URL)

        playerName = driver.find_elements_by_css_selector(
            "a[class*='main-table-link main-table-player']")
        itemLevel = driver.find_elements_by_css_selector("td[class*='ilvl-cell']")
        #itemLevel2 = driver.find_elements_by_xpath("//td[@class='ilvl-cell']")
        DPS = driver.find_elements_by_css_selector(
            "td[class*='main-table-number primary players-table-dps']")
        Date = driver.find_elements_by_css_selector(
            "td[class*='main-table-number players-table-date']")
        Duration = driver.find_elements_by_css_selector(
            "td[class*='main-table-number players-table-duration']")

        #tf_icon = driver.find_elements_by_xpath("//img[@class='tiny-icon']")
        playerGear = driver.find_elements_by_xpath("//td[@class='unique-gear']")
        
        # 第一列 No,Name
        number = 1
        i = 0
        for tt in playerName:
            sheet1.write(i+1, 0, number, set_style(f,
                                                'Times New Roman', 220, False, 'center'))
            s = tt.get_attribute('innerHTML')
            sheet1.write(i+1, 1, s.strip(), set_style(f,
                                                    'Times New Roman', 220, False, 'center'))
            number += 1
            i += 1

        # 第2列 itemLevel
        i = 0
        for itl in itemLevel:
            itemText = itl.get_attribute("innerHTML").strip()
            sheet1.write(i+1, 2, int(itemText), set_style(f,
                                                        'Times New Roman', 220, False, 'center'))
            i += 1
            # print(itemText)

        i = 0
        for d in DPS:
            sheet1.write(i+1, 3, d.get_attribute("innerHTML").strip(),
                        set_style(f, 'Times New Roman', 220, False, 'center'))
            i += 1
            # print(d.get_attribute("innerHTML").strip())

        i = 0
        for dd in Date:
            textContent = dd.get_attribute("textContent")
            start_pos = textContent.find('$')
            s = textContent[start_pos+1:]
            sheet1.write(
                i+1, 4, s, set_style(f, 'Times New Roman', 220, False, 'center'))
            i += 1

        i = 0
        for dura in Duration:
            s = dura.get_attribute("innerHTML")
            sheet1.write(i+1, 5, str(s).strip(), set_style(f,
                                                        'Times New Roman', 220, False, 'center'))
            i += 1

        i = 0
        idx = 0
        rowId = 0
        
        sheet1.set_column(1, 1, 15)
        sheet1.set_column(11, 11, 23)
        sheet1.set_column(12, 12, 20)
        
        for pg in playerGear:
            icon_group = pg.find_elements_by_class_name("tiny_icon")

        
        rowId =0
        
        for pg in playerGear:
            icon_group = pg.find_elements_by_class_name("tiny-icon")
            i = 0
            while i <=17 and i < len(icon_group):
                pic_url = icon_group[i].get_attribute('src')
                try:
                    if not pic_map.get(pic_url):
                        image_data = BytesIO(urlopen(pic_url).read())
                        pic_map[pic_url] = image_data
                        print("get file " + pic_url)
                    else:
                        image_data = pic_map[pic_url]

                    if image_data.getvalue()[:4] != '\xff\xd8\xff\xe0' and image_data.getvalue()[:4]!='\xff\xd8\xff\xe1': 
                        i = i+1
                        continue

                    if i <= 6:
                        sheet1.insert_image(rowId+1, 11, pic_url, {'image_data': image_data, 'x_scale': 0.25,
                                                                'y_scale': 0.25, 'x_offset': 19*i})
                    elif i <= 8:
                        sheet1.insert_image(rowId+1, 13, pic_url, {'image_data': image_data, 'x_scale': 0.25,
                                                                'y_scale': 0.25, 'x_offset': 19*(i-7)})
                    elif i <= 14:
                        sheet1.insert_image(rowId+1, 12, pic_url, {'image_data': image_data, 'x_scale': 0.25,
                                                                'y_scale': 0.25, 'x_offset': 19*(i-9)})
                    else:
                        sheet1.insert_image(rowId+1, 14, pic_url, {'image_data': image_data, 'x_scale': 0.25,
                                                                'y_scale': 0.25, 'x_offset': 19*(i-15)})                       
                except Exception,e:
                    print(e)
                    traceback.print_exc()
                    
                i+=1
            rowId += 1
            idx += 15

        # get player detail URL
        rowId = 0
        #chrome_options = Options()
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--ignore-certificate-errors-spki-list')
        chrome_options.add_argument("--log-level=3")
        # disable picture
        prefs = {"profile.managed_default_content_settings.images": 2}
        chrome_options.add_experimental_option("prefs", prefs)
        driverChrome = webdriver.Chrome(chrome_options=chrome_options)
        for u in playerName:
            href = u.get_attribute("href")
            if href:
                fullURL = href

                if not attr_map.get(fullURL):
                    driverChrome.maximize_window()
                    driverChrome.get(fullURL)
                    # Open Details page
                    try:
                        #men_menu = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//a[@data-tracking-id='men']")))
                        driverChrome.execute_script("window.scrollBy(0,750)")
                        xpath_str = "//a[text()='%s\n']" % (u.text)

                        WebDriverWait(driverChrome, 20).until(
                            EC.presence_of_element_located((By.XPATH, xpath_str)))
                        # ActionChains(driver).move_to_element(men_menu).perform()

                        p1 = driverChrome.find_elements_by_xpath(xpath_str)
                        p1[0].click()
                        driverChrome.execute_script("window.scrollBy(0,-500)")
                        # driverChrome.save_screenshot('screenie.png')
                        p2 = driverChrome.find_element_by_xpath("//a[text()='摘要']")
                        p2.click()
                        xpath_str = "//span[@class='composition-entry']"
                        WebDriverWait(driverChrome, 20).until(
                            EC.presence_of_element_located((By.XPATH, xpath_str)))

                    except Exception, e:
                        print(str(e))
                        traceback.print_exc()

                    sxall = driverChrome.find_elements_by_xpath("//span[@class='composition-entry']")

                    fsxText = ""
                    i = 0
                    sxText = []
                    for s in sxall:
                        fsxText += s.text
                        fsxText += ' '
                        try:
                            n = s.find_elements_by_class_name("estimate")
                            ns = n[0].text.replace(',', '').strip()

                            if i != 1:
                                sxText.append(ns)
                                columId = 6
                                if i == 0:
                                    columId = i+6
                                else:
                                    columId = i+5

                                if columId <= 10:
                                    sheet1.write(
                                        rowId+1, columId, ns, set_style(f, 'Times New Roman', 220, False, 'center'))
                        except Exception, e:
                            print(str(e))
                            traceback.print_exc()
                        i += 1
                    fsxText = fsxText.replace(',', '')
                    print(fsxText)

                    attr_map[fullURL] = sxText

                else:
                    sxText = attr_map[fullURL]
                    columId = 0
                    for s in sxText:
                        sheet1.write(rowId+1, columId+6, s, set_style(f,'Times New Roman', 220, False, 'center'))
                        columId += 1



                """
                try:
                    tianfuTable = driverChrome.find_element_by_id('summary-talents-0')
                    tianfuList = tianfuTable.find_elements_by_css_selector("img[class*='table-icon mCS_img_loaded']")
                    if len(tianfuList) > 7:
                        id = 7
                        while id < len(tianfuList):
                            picURL = tianfuList[id].get_attribute("src")
                                    
                            if not pic_map.get(picURL):
                                image_data = BytesIO(urlopen(picURL).read())
                                pic_map[picURL] = image_data

                            else:
                                image_data = pic_map[picURL]

                            sheet1.insert_image(rowId+1, 14, picURL, {'image_data': image_data, 'x_scale': 0.25,'y_scale': 0.25, 'x_offset': 19*(id-7)})
                            id += 1
                except Exception,e:
                    print(str(e))
                    traceback.print_exc()   
                """ 
                #sheet1.write(rowId+1, 6, fsxText, set_style(f,'Times New Roman', 220, False, 'center'))
            rowId += 1

        # set colum  width
        sheet1.set_column(1, 1, 15)
        sheet1.set_column(11, 11, 23)
        sheet1.set_column(12, 12, 20)
    except Exception,e:
        print(e)
        traceback.print_exc()
    
    try:
        with open("attr_"+ wclass +"_" +spec+".txt", "wb") as f_attr:
            pickle.dump(attr_map, f_attr)
    except:
        pass


if __name__ == "__main__":
    main()
