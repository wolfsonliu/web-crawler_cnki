import os
import re
import time
import string
import random
import shutil
import subprocess as sp
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from PIL import Image
from PIL import ImageEnhance

random.seed(1024)

os.environ["TESSDATA_PREFIX"] = os.path.join(os.getcwd(), "tesseract-4.0.0-alpha")



def orc(img, tesseract=os.path.join(os.getcwd(), 'tesseract-4.0.0-alpha/tesseract.exe')):
    # 识别简单验证码
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


def get_search_result_number(driver):
    # 返回总搜索结果
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    pagenumberspan = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'pagerTitleCell'))
    )
    pn = pagenumberspan.get_attribute('innerHTML').split('&nbsp')[2].replace(';', '').replace(',', '')
    driver.switch_to.default_content()
    return pn


def get_info(driver, header=''):
    # 下载页面显示的表格结果
    timere = re.compile('[0-9]{4}-[0-9]{1,2}-[0-9]{1,2} [0-9]{1,2}:[0-9]{1,2}')
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    time.sleep(1)
    gtc = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'GridTableContent'))
    )
    trs = gtc.find_elements_by_tag_name('tr')
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


def next_page(driver):
    # 搜索结果的下一页，如果有下一页则跳转到下一页，并返回 True，如果没有下一页则返回 False
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到按钮所在位置
    try:
        next_page = driver.find_element_by_id('Page_next')
        next_page.click()
    except NoSuchElementException:
        next_page = None
    finally:
        driver.switch_to.default_content()
        driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到最上层
    if next_page is None:
        return False
    else:
        return True


def checkcode(driver):
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到按钮所在位置
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
        driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到最上层
        return True


def select_items(driver):
    # 选择当前的条目
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到按钮所在位置
    selectbox = driver.find_element_by_id('selectCheckbox')  # 获得选择框
    selectbox.click()
    driver.switch_to.default_content()
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到最上层


def clear_items(driver):
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到按钮所在位置
    clear = driver.find_element_by_class_name('pageBar_top').find_element_by_class_name('zero')
    clear.click()
    driver.switch_to.default_content()
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到最上层


def export_items(driver, waittime):
    # 下载选择的文件
    driver.switch_to.default_content()
    driver.switch_to_frame('iframeResult')
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到按钮所在位置
    driver.find_element_by_class_name('SavePoint').find_element_by_xpath('//a[text()=" 导出/参考文献"]').click()
    # 打开下载页面
    window_ids = driver.window_handles  # 获取所有页面卡 ID
    driver.switch_to.window(window_ids[1])  # 切换到下载页面
    savinglist = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, 'SaveTypeList'))
    )
    savinglist.find_elements_by_tag_name('li')[-1].click()  # 点击自定义选项
    checktable = driver.find_element_by_class_name('checkTable')  # 选项表
    checkall = checktable.find_element_by_class_name('btnTd').find_elements_by_tag_name('input')[0]  # 全选
    checkall.click()
    exportbtn = driver.find_element_by_id('exportExcel')  # 导出按钮
    exportbtn.click()
    time.sleep(waittime)
    driver.close()  # 关闭文件下载页面
    driver.switch_to.window(window_ids[0])  # 返回原页面
    driver.switch_to.default_content()
    driver.execute_script('document.body.scrollIntoView()')  # 滚动屏幕到最上层


def crawler(collage_name, year_start, year_end, subject_id_list):
    subject_name = {
        'A001': '自然科学理论与方法', 'A002': '数学', 'A003': '非线性科学与系统科学', 'A004': '力学',
        'A005': '物理学', 'A006': '生物学', 'A007': '天文学', 'A008': '自然地理学和测绘学', 'A009': '气象学',
        'A010': '海洋学', 'A011': '地质学', 'A012': '地球物理学', 'A013': '资源科学'
    }
    subsubject_name = {
        'A001_1': '自然科学研究的方针政策', 'A001_2': '自然科学理论与方法论', 'A001_3': '自然科学现状及发展',
        'A001_4': '自然科学研究及自然科学史',
        'A002_1': '数学概论', 'A002_2': '古典数学', 'A002_3': '初、高等数学',
        'A002_4': '数理逻辑、数学基础', 'A002_5': '代数、数论、组合理论', 'A002_6': '数学分析',
        'A002_7': '几何、拓扑', 'A002_8': '动力系统理论', 'A002_9': '概率论 、数理统计', 'A002_A': '运筹学',
        'A002_B': '控制论、信息论(数学理论)', 'A002_C': '计算数学', 'A002_D': '应用数学',
        'A003_1': '非线性科学', 'A003_2': '系统科学',
        'A004_1': '力学概论', 'A004_2': '理论力学(一般力学)', 'A004_3': '振动理论', 'A004_4': '连续介质力学(变形体力学)',
        'A004_5': '固体力学', 'A004_6': '流体力学', 'A004_7': '流变学', 'A004_8': '爆炸力学', 'A004_9': '应用力学',
        'A005_1': '物理学总论', 'A005_2': '物理学现状、概况', 'A005_3': '物理学实验方法与设备', 'A005_4': '物理测量',
        'A005_5': '理论物理学', 'A005_6': '声学', 'A005_7': '光学', 'A005_8': '电磁学、电动力学', 'A005_9': '无线电物理学',
        'A005_A': '真空电子学(电子物理学)', 'A005_B': '半导体物理学', 'A005_C': '固体物理学', 'A005_D': '低温物理学',
        'A005_E': '高压与高温物理学', 'A005_F': '等离子体物理学', 'A005_G': '热学与物质分子运动论',
        'A005_H': '分子物理学、原子物理学', 'A005_I': '原子核物理学、高能物理学', 'A005_J': '应用物理学',
        'A006_1': '生物科学总论', 'A006_2': '普通生物学', 'A006_3': '细胞学', 'A006_4': '遗传学', 'A006_5': '生理学',
        'A006_6': '生物化学', 'A006_7': '生物物理学', 'A006_8': '分子生物学及遗传工程', 'A006_9': '生物工程学',
        'A006_A': '环境生物学', 'A006_B': '古生物学', 'A006_C': '微生物学理论及其应用', 'A006_D': '植物学',
        'A006_E': '动物学', 'A006_F': '昆虫学', 'A006_G': '人类学',
        'A007_1': '一般性问题', 'A007_2': '天文观测设备与观测资料', 'A007_3': '天体测量学', 'A007_4': '天体力学(理论天文学)',
        'A007_5': '天体物理学', 'A007_6': '恒星天文学、星系天文学、宇宙学', 'A007_7': '射电天文学(无线电天文学)',
        'A007_8': '空间天文学', 'A007_9': '太阳系', 'A007_A': '时间、历法',
        'A008_1': '自然地理学', 'A008_2': '测绘学',
        'A009_1': '气象工作及气象经济', 'A009_2': '气象基础科学', 'A009_3': '大气探测(气象观测)', 'A009_4': '气象要素及大气现象',
        'A009_5': '动力气象学', 'A009_6': '天气学', 'A009_7': '天气预报', 'A009_8': '气候学', 'A009_9': '气象灾害学',
        'A009_A': '应用气象学', 'A009_B': '气象设备仪器制造、使用', 'A009_C': '人工影响天气',
        'A010_1': '海洋学', 'A010_2': '海洋调查技术方法与设备', 'A010_3': '区域海洋学', 'A010_4': '海洋气象学',
        'A010_5': '海洋水文学', 'A010_6': '海洋物理学', 'A010_7': '海洋化学', 'A010_8': '海洋生物学', 'A010_9': '海洋地球物理学',
        'A010_A': '海洋地质学、海洋地貌学', 'A010_B': '海洋资源与开发', 'A010_C': '海洋工程', 'A010_D': '海洋环境及海洋保护学',
        'A010_E': '海洋灾害',
        'A011_1': '地质工作及地质经济', 'A011_2': '地质灾害及灾害地质学', 'A011_3': '矿物学', 'A011_4': '岩石学',
        'A011_5': '古生物学', 'A011_6': '历史地质学、地层学', 'A011_7': '动力地质学', 'A011_8': '构造地质学',
        'A011_9': '地质力学', 'A011_A': '区域地质学', 'A011_B': '矿床学', 'A011_C': '地球化学', 'A011_D': '普查勘探地质学',
        'A011_E': '勘探工程学', 'A011_F': '宇宙地质学', 'A011_G': '海洋地质学', 'A011_H': '水文地质学与工程地质学',
        'A011_I': '环境地质学', 'A011_K': '地质勘查仪器及设备',
        'A012_1': '一般理论', 'A012_2': '地球起源及演化', 'A012_3': '地球形状学及地球结构', 'A012_4': '地球重力学',
        'A012_5': '大地构造物理学、岩组学(构造岩石学)', 'A012_6': '地热学、火山学', 'A012_7': '地磁学',
        'A012_8': '地电学', 'A012_9': '水文学', 'A012_A': '空间物理', 'A012_B': '地震地球物理学', 'A012_C': '勘查地球物理学',
        'A012_D': '海洋地球物理学', 'A012_E': '工程及水文地球物理勘探', 'A012_F': '地球物理勘测仪器',
        'A013_1': '资源管理与利用', 'A013_2': '各种资源的开发与利用'
    }
    pagedatadir = './pagedata'
    filedatadir = './filedata'
    os.makedirs(pagedatadir, exist_ok=True)
    os.makedirs(filedatadir, exist_ok=True)
    driver.switch_to.window(driver.window_handles[0])
    driver.switch_to.default_content()
    driver.get('http://kns.cnki.net/kns/brief/result.aspx?dbprefix=CJFQ')
    driver.maximize_window()
    # 检索按钮

    # 选择学校
    collage_input = driver.find_element_by_id('au_1_value2')
    driver.execute_script('arguments[0].value = arguments[1]', collage_input, collage_name)

    # 选择年份
    time.sleep(2)
    yearfrom = driver.find_element_by_name('year_from')
    yearfrom.send_keys('{0}'.format(year_start))
    time.sleep(2)
    yearto = driver.find_element_by_name('year_to')
    yearto.send_keys('{0}'.format(year_end))
    time.sleep(2)

    # 选择来源
    journalsource = driver.find_element_by_id('ddSubmit').find_elements_by_class_name('dd01')[3]
    jsall = journalsource.find_element_by_id('AllmediaBox')
    if jsall.is_selected():
        jsall.click()
    jsbox = journalsource.find_elements_by_tag_name('input')[2:]
    jsbox[3].click()
    jsbox[4].click()
    time.sleep(2)

    # 选择类别
    clear_btn = driver.find_element_by_id('XuekeNavi_Div').find_elements_by_class_name('btn')[0]
    clear_btn.click()
    category = driver.find_element_by_id('XuekeNavi_Div').find_element_by_id('A').find_element_by_id('Afirst')
    category.click()
    time.sleep(1)

    for subid in subject_id_list:
        # subcategorydlfirst = WebDriverWait(driver, 6).until(
        #     EC.presence_of_element_located((By.ID, subid + 'first'))
        # )
        subcategorydlfirst = driver.find_element_by_id(subid + 'first')
        subcategorydlfirst.click()  # open subsubcategory list
        time.sleep(1)
        subsubcategorys = driver.find_element_by_id(subid + 'child').find_elements_by_id('selectbox')
        for subsub in subsubcategorys:
            subsubid = subsub.get_property('value')
            pagefilename = '{0}_{1}_{2}_{3}_{4}.txt'.format(
                collage_name, subsubid,
                subsubject_name[subsubid],
                str(year_start), str(year_end)
            )
            subsub.click()
            time.sleep(2)
            searchbtn = driver.find_element_by_id('btnSearch')
            searchbtn.click()  # 搜索  ShowDiv0
            time.sleep(2)
            WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.ID, 'Show0'))
            )
            result_number = int(get_search_result_number(driver))
            if result_number > 6000:
                raise ValueError(
                    ' '.join([collage_name, subsubid, '搜索结果超出限额：', get_search_result_number(driver)])
                )
            elif result_number == 0:
                with open(
                        os.path.join(
                            pagedatadir,
                            pagefilename
                        ),
                        'ab'
                ) as f:
                    f.write('zero result\n'.encode('utf-8'))
            else:
                while True:
                    checkcode(driver)
                    tabletext = get_info(driver)
                    tabletext[0] = '科目\t' + tabletext[0]
                    for row in range(1, len(tabletext)):
                        tabletext[row] = '\t'.join([subsubid, tabletext[row]])
                    with open(
                            os.path.join(
                                pagedatadir,
                                pagefilename
                            ),
                            'ab'
                    ) as f:
                        f.write('\n'.join(tabletext[1:]).encode('utf-8'))
                        f.write('\n'.encode('utf-8'))
                    clear_items(driver)
                    select_items(driver)
                    export_items(driver, 2)
                    filename = max(
                        [os.path.join(filedatadir, f) for f in os.listdir(filedatadir)],
                        key=os.path.getctime
                    ) # 获取创建最晚的文件
                    for fi in range(60):
                        newfilename = '{0}_{1}_{2}_{3}_{4}_{5}.xls'.format(
                            collage_name, subsubid, subsubject_name[subsubid], year_start, year_end, fi
                        )
                        if not os.path.exists(os.path.join(filedatadir, newfilename)):
                            shutil.move(filename, os.path.join(filedatadir, newfilename))
                            break
                    clear_items(driver)
                    time.sleep(random.randint(2, 5))
                    if not next_page(driver):
                        break
            subsub.click()
        time.sleep(1)
        subcategorydlsecond = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, subid + 'second'))
        )
        subcategorydlsecond.click()  # close subsubcategory list





###########################################


subject_name = {
    'A001': '自然科学理论与方法', 'A002': '数学', 'A003': '非线性科学与系统科学', 'A004': '力学',
    'A005': '物理学', 'A006': '生物学', 'A007': '天文学', 'A008': '自然地理学和测绘学', 'A009': '气象学',
    'A010': '海洋学', 'A011': '地质学', 'A012': '地球物理学', 'A013': '资源科学'
}
subsubject_name = {
    'A001_1': '自然科学研究的方针政策', 'A001_2': '自然科学理论与方法论', 'A001_3': '自然科学现状及发展',
    'A001_4': '自然科学研究及自然科学史',
    'A002_1': '数学概论', 'A002_2': '古典数学', 'A002_3': '初、高等数学',
    'A002_4': '数理逻辑、数学基础', 'A002_5': '代数、数论、组合理论', 'A002_6': '数学分析',
    'A002_7': '几何、拓扑', 'A002_8': '动力系统理论', 'A002_9': '概率论 、数理统计', 'A002_A': '运筹学',
    'A002_B': '控制论、信息论(数学理论)', 'A002_C': '计算数学', 'A002_D': '应用数学',
    'A003_1': '非线性科学', 'A003_2': '系统科学',
    'A004_1': '力学概论', 'A004_2': '理论力学(一般力学)', 'A004_3': '振动理论', 'A004_4': '连续介质力学(变形体力学)',
    'A004_5': '固体力学', 'A004_6': '流体力学', 'A004_7': '流变学', 'A004_8': '爆炸力学', 'A004_9': '应用力学',
    'A005_1': '物理学总论', 'A005_2': '物理学现状、概况', 'A005_3': '物理学实验方法与设备', 'A005_4': '物理测量',
    'A005_5': '理论物理学', 'A005_6': '声学', 'A005_7': '光学', 'A005_8': '电磁学、电动力学', 'A005_9': '无线电物理学',
    'A005_A': '真空电子学(电子物理学)', 'A005_B': '半导体物理学', 'A005_C': '固体物理学', 'A005_D': '低温物理学',
    'A005_E': '高压与高温物理学', 'A005_F': '等离子体物理学', 'A005_G': '热学与物质分子运动论',
    'A005_H': '分子物理学、原子物理学', 'A005_I': '原子核物理学、高能物理学', 'A005_J': '应用物理学',
    'A006_1': '生物科学总论', 'A006_2': '普通生物学', 'A006_3': '细胞学', 'A006_4': '遗传学', 'A006_5': '生理学',
    'A006_6': '生物化学', 'A006_7': '生物物理学', 'A006_8': '分子生物学及遗传工程', 'A006_9': '生物工程学',
    'A006_A': '环境生物学', 'A006_B': '古生物学', 'A006_C': '微生物学理论及其应用', 'A006_D': '植物学',
    'A006_E': '动物学', 'A006_F': '昆虫学', 'A006_G': '人类学',
    'A007_1': '一般性问题', 'A007_2': '天文观测设备与观测资料', 'A007_3': '天体测量学', 'A007_4': '天体力学(理论天文学)',
    'A007_5': '天体物理学', 'A007_6': '恒星天文学、星系天文学、宇宙学', 'A007_7': '射电天文学(无线电天文学)',
    'A007_8': '空间天文学', 'A007_9': '太阳系', 'A007_A': '时间、历法',
    'A008_1': '自然地理学', 'A008_2': '测绘学',
    'A009_1': '气象工作及气象经济', 'A009_2': '气象基础科学', 'A009_3': '大气探测(气象观测)', 'A009_4': '气象要素及大气现象',
    'A009_5': '动力气象学', 'A009_6': '天气学', 'A009_7': '天气预报', 'A009_8': '气候学', 'A009_9': '气象灾害学',
    'A009_A': '应用气象学', 'A009_B': '气象设备仪器制造、使用', 'A009_C': '人工影响天气',
    'A010_1': '海洋学', 'A010_2': '海洋调查技术方法与设备', 'A010_3': '区域海洋学', 'A010_4': '海洋气象学',
    'A010_5': '海洋水文学', 'A010_6': '海洋物理学', 'A010_7': '海洋化学', 'A010_8': '海洋生物学', 'A010_9': '海洋地球物理学',
    'A010_A': '海洋地质学、海洋地貌学', 'A010_B': '海洋资源与开发', 'A010_C': '海洋工程', 'A010_D': '海洋环境及海洋保护学',
    'A010_E': '海洋灾害',
    'A011_1': '地质工作及地质经济', 'A011_2': '地质灾害及灾害地质学', 'A011_3': '矿物学', 'A011_4': '岩石学',
    'A011_5': '古生物学', 'A011_6': '历史地质学、地层学', 'A011_7': '动力地质学', 'A011_8': '构造地质学',
    'A011_9': '地质力学', 'A011_A': '区域地质学', 'A011_B': '矿床学', 'A011_C': '地球化学', 'A011_D': '普查勘探地质学',
    'A011_E': '勘探工程学', 'A011_F': '宇宙地质学', 'A011_G': '海洋地质学', 'A011_H': '水文地质学与工程地质学',
    'A011_I': '环境地质学', 'A011_K': '地质勘查仪器及设备',
    'A012_1': '一般理论', 'A012_2': '地球起源及演化', 'A012_3': '地球形状学及地球结构', 'A012_4': '地球重力学',
    'A012_5': '大地构造物理学、岩组学(构造岩石学)', 'A012_6': '地热学、火山学', 'A012_7': '地磁学',
    'A012_8': '地电学', 'A012_9': '水文学', 'A012_A': '空间物理', 'A012_B': '地震地球物理学', 'A012_C': '勘查地球物理学',
    'A012_D': '海洋地球物理学', 'A012_E': '工程及水文地球物理勘探', 'A012_F': '地球物理勘测仪器',
    'A013_1': '资源管理与利用', 'A013_2': '各种资源的开发与利用'
}



profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.dir', os.path.join(os.getcwd(), 'filedata'))
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/x--tagged; charset=utf-8')
driver = webdriver.Firefox(firefox_profile=profile, executable_path=r'./geckodriver.exe')


year = 2011
collage_name = '北京大学'


crawler('北京大学', 2011, 2017, ['A' + '{0:03d}'.format(x) for x in range(4, 14)])

