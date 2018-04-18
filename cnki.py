import os
import re
import time
import string
import random
import pandas as pd
import subprocess as sp
from io import StringIO
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from PIL import Image
from PIL import ImageEnhance

random.seed(1024)

os.environ["TESSDATA_PREFIX"] = os.path.join(os.getcwd(), "tesseract-4.0.0-alpha")


def orc(img, tesseract=os.path.join(os.getcwd(), 'tesseract-4.0.0-alpha/tesseract.exe')):
    if isinstance(img, str):
        img = Image.open(img)
    gray = img.convert('L')
    contrast = ImageEnhance.Contrast(gray)
    ctgray = contrast.enhance(3.0)
    bw = ctgray.point(lambda x: 0 if x < 1 else 255)
    bw.save('captcha_threasholded.png')
    process = sp.Popen(
        [tesseract, 'captcha_threasholded.png', 'out', '--psm 7', '--tessdata-dir ' + os.path.dirname(tesseract)],
        shell=True
    )
    process.wait()
    with open('out.txt', 'r') as f:
        words = ''.join(list(f.readlines())).rstrip()
    words = ''.join(c for c in words if c in string.ascii_letters + string.digits).lower()
    return words



ff = webdriver.Firefox(executable_path=r'./geckodriver.exe')
driver = ff
year = 2011
collage_name = '北京大学'

driver.get('http://kns.cnki.net/kns/brief/result.aspx?dbprefix=CJFQ')
driver.maximize_window()

def get_search_result_number(driver):
    # 返回总搜索结果
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    pagenumberspan = driver.find_element_by_class_name('pagerTitleCell')
    pn = pagenumberspan.get_attribute('innerHTML').split('&nbsp')[2].replace(';', '').replace(',', '')
    driver.switch_to.default_content()
    return pn


def get_info(driver, header=''):
    # 下载页面显示的表格结果
    timere = re.compile('[0-9]{4}-[0-9]{1,2}-[0-9]{1,2} [0-9]{1,2}:[0-9]{1,2}')
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    trs = driver.find_element_by_class_name('GridTableContent').find_elements_by_tag_name('tr')
    tabletext = list()
    tabletext.append( ' '.join(['序号', trs[0].text]).replace(' ', '\t'))
    for x in trs[1:]:
        text = x.text
        times = timere.findall(text)  # 找到 2017-08-14 14:08 时间中的空格并替换为下划线
        if len(times) > 0:
            for tx in times:
                text = text.replace(tx, tx.replace(' ', '_'))
        text = text.replace('; ', ';')
        tabletext.append(text.replace(' ', '\t'))  # 把空格替换为 \t
    driver.switch_to.default_content()
    return tabletext


def nextpage(driver):
    # 搜索结果的下一页，如果有下一页则跳转到下一页，并返回 True，如果没有下一页则返回 False
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  ## 滚动屏幕到按钮所在位置
    try:
        next_page = driver.find_element_by_id('Page_next')
        next_page.click()
    except NoSuchElementException:
        next_page = None
    finally:
        driver.switch_to.default_content()
        driver.execute_script('document.body.scrollIntoView()')  ## 滚动屏幕到最上层
    if next_page is None:
        return False
    else:
        return True


def checkcode(driver):
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  ## 滚动屏幕到按钮所在位置
    try:
        driver.find_element_by_id('CheckCodeImg')
    except NoSuchElementException:
        driver.switch_to.default_content()
        driver.execute_script('document.body.scrollIntoView()')
        return False
    try:
        while driver.find_element_by_tag_name('label').text == '请输入验证码：':
            checkcode_img = driver.find_element_by_id('CheckCodeImg')
            checkcode_img.click()
            checkcode_img.screenshot('checkcode.png')
            captcha = orc('checkcode.png')
            time.sleep(random.randint(1,3))
            captcha_entry = driver.find_element_by_id('CheckCode')
            driver.execute_script('arguments[0].value = arguments[1]', captcha_entry, captcha)
            captch_submit = driver.find_elements_by_tag_name('input')[1]
            captch_submit.click()
    except NoSuchElementException:
        driver.switch_to.default_content()
        driver.execute_script('document.body.scrollIntoView()')  ## 滚动屏幕到最上层
        return True




# 检索按钮
searchbtn = driver.find_element_by_id('btnSearch')

# 选择学校
collage_input = driver.find_element_by_id('au_1_value2')
driver.execute_script('arguments[0].value = arguments[1]', collage_input, collage_name)

# 选择年份
yearfrom = driver.find_element_by_name('year_from')
yearfrom.send_keys('{0}'.format(year))
yearto = driver.find_element_by_name('year_to')
yearto.send_keys('{0}'.format(year))

# 选择来源
journalsource = driver.find_element_by_id('ddSubmit').find_elements_by_class_name('dd01')[3]
jsall = journalsource.find_element_by_id('AllmediaBox')
if jsall.is_selected():
    jsall.click()
jsbox = journalsource.find_elements_by_tag_name('input')[2:]
jsbox[3].click()
jsbox[4].click()

# 选择类别
clear_btn = driver.find_element_by_id('XuekeNavi_Div').find_elements_by_class_name('btn')[0]
clear_btn.click()
category = driver.find_element_by_id('XuekeNavi_Div').find_element_by_id('A').find_element_by_id('Afirst')
category.click()
subcategorydl = driver.find_element_by_id('XuekeNavi_Div').find_element_by_id('Achild')

for subid in ['A' + '{0:03d}'.format(x) for x in range(1, 14)]:
    subcategorydl.find_element_by_id(subid + 'first').click()  # open subsubcategory list
    subsubcategorys = subcategorydl.find_element_by_id(subid + 'child').find_elements_by_id('selectbox')
    for subsub in subsubcategorys:
        subsubid = subsub.get_property('value')
        subsub.click()
        searchbtn.click()  # 搜索
        if int(get_search_result_number(driver)) > 3000:
            raise ValueError('搜索结果超出限额：' + get_search_result_number(driver))
        while True:
            if
            tabletext = get_info(driver)
            tabletext[0] = '科目\t' + tabletext[0]
            for row in range(1, len(tabletext)):
                tabletext[row] = '\t'.join([subsubid, tabletext[row]])
            with open('_'.join([collage_name, year, subsubid]), 'ab') as f:
                f.write('\n'.join(tabletext[1:]).encode('utf-8'))
                f.write('\n'.encode('utf-8'))
            if not nextpage(driver):
                break

        subsub.click()
    subcategorydl.find_element_by_id(subid + 'second').click()  # close subsubcategory list