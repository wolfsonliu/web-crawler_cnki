import os
import re
import time
import string
import random
import pandas as pd
import subprocess as sp
from selenium import webdriver
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
drive = ff
year = 2011

drive.get('http://kns.cnki.net/kns/brief/result.aspx?dbprefix=CJFQ')

def get_search_result_number(drive):
    drive.switch_to_frame('iframeResult')
    pagenumberspan = drive.find_element_by_class_name('pagerTitleCell')
    pn = pagenumberspan.get_attribute('innerHTML').split('&nbsp')[2].replace(';', '').replace(',', '')
    drive.switch_to.default_content()
    return pn

def get_info(drive):
    timere = re.compile('[0-9]{4}-[0-9]{1,2}-[0-9]{1,2} [0-9]{1,2}:[0-9]{1,2}')
    drive.switch_to_frame('iframeResult')
    trs = drive.find_element_by_class_name('GridTableContent').text.split('\n').find_elements_by_tag_name('tr')
    tabletext = list()
    tabletext[0] = ' '.join(['序号', trs[0].text]).replace(' ', '\t')
    drive.switch_to.default_content()

# 检索按钮
searchbtn = drive.find_element_by_id('btnSearch')

# 选择学校
collage = drive.find_element_by_id('au_1_value2')
drive.execute_script('arguments[0].value = arguments[1]', collage, '北京大学')

# 选择年份
yearfrom = drive.find_element_by_name('year_from')
yearfrom.send_keys('{0}'.format(year))
yearto = drive.find_element_by_name('year_to')
yearto.send_keys('{0}'.format(year))

# 选择来源
journalsource = drive.find_element_by_id('ddSubmit').find_elements_by_class_name('dd01')[3]
jsall = journalsource.find_element_by_id('AllmediaBox')
if jsall.is_selected():
    jsall.click()
jsbox = journalsource.find_elements_by_tag_name('input')[2:]
jsbox[3].click()
jsbox[4].click()

# 选择类别
clear_btn = drive.find_element_by_id('XuekeNavi_Div').find_elements_by_class_name('btn')[0]
clear_btn.click()
category = drive.find_element_by_id('XuekeNavi_Div').find_element_by_id('A').find_element_by_id('Afirst')
category.click()
subcategorydl = drive.find_element_by_id('XuekeNavi_Div').find_element_by_id('Achild')

for subid in ['A' + '{0:03d}'.format(x) for x in range(1, 14)]:
    subcategorydl.find_element_by_id(subid + 'first').click()  # open subsubcategory list
    subsubcategorys = subcategorydl.find_element_by_id(subid + 'child').find_elements_by_id('selectbox')
    for subsub in subsubcategorys:
        subsubid = subsub.get_property('value')
        subsub.click()

        subsub.click()
    subcategorydl.find_element_by_id(subid + 'second').click()  # close subsubcategory list