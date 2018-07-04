# -*- coding: utf-8 -*-

import xlrd
import time
from appium import webdriver

def driver_data(i = 6):
    '''
    获取被测设备信息
    :param i: 选择第几行设备：0到N
    :return: 设备列表
    '''
    driver_data = {'platformName':"","platformVersion":"","deviceName":"",
                   "appPackage":"","appActivity":""}
    file = openfile()
    sheet = file.sheet_by_name("phone_data")

    if i >= 2:
        driver_data['platformName'] = sheet.row_values(i)[0]
        driver_data['platformVersion'] = sheet.row_values(i)[1]
        driver_data['deviceName'] = sheet.row_values(i)[2]
        driver_data['appPackage'] = sheet.row_values(i)[3]
        driver_data['appActivity'] = sheet.row_values(i)[4]
    else:
        print("参数必须大于等于2")

    return driver_data

def setup(noReset = True,host = "127.0.0.1"):
    '''
    在本地启动应用
    :return:
    '''
    desired_caps = driver_data()
    desired_caps['noReset'] = noReset
    # desired_caps = {}
    # desired_caps['platformName'] = 'Android'
    # desired_caps['platformVersion'] = '5.1.1'  # '5.1.1'#'4.4.4'
    # desired_caps['deviceName'] = 'emulator-5554'  # ''80a8d0db' '9fb08572'
    # desired_caps['appPackage'] = 'com.codemao.box'  # APK包名
    # desired_caps['appActivity'] = "com.codemao.box.module.welcome.FirstActivity"  # 'com.qihoo.util.StartActivity'
    # desired_caps['noReset'] = noReset
    driver = webdriver.Remote( "http://" + host + ":4723/wd/hub", desired_caps)
    #driver = webdriver.Remote('http://172.16.7.121:4723/wd/hub', desired_caps)

    return driver

def teardowm(driver):
    '''
    关闭应用
    :param driver:
    :return:
    '''

    time.sleep(5)
    driver.close_app()

def find(driver,by,valuse):
    '''
    对象
    定位控件
    :param driver:
    :param by:
    :param value:
    :return:
    '''
    if by == 'x':
        driver.implicitly_wait(15)
        return driver.find_element_by_xpath(valuse)

    elif by == 'id': #resource-id
        driver.implicitly_wait(15)
        return driver.find_element_by_id(valuse)

    elif by == "zb":
        driver.implicitly_wait(15)
        return driver.tap(valuse)

def click(element):
    '''
    动作
    点击行为
    :param element:
    :return:
    '''
    return element.click()

def send(element,values):
    '''
    动作
    输入行为
    :param element:
    :param values:
    :return:
    '''
    element.clear()
    return element.send_keys(values)

def clear(element):
    '''
    动作
    清空行为
    :param element:
    :return:
    '''
    return element.clear()

def openfile():
    '''
    打开excel
    :return:
    '''
    file = xlrd.open_workbook("testcase.xlsx")

    return file

def get_elememt_id(file):
    '''
    获取文档的所有控件数据
    :param file: 调用openfile
    :return: 控件字典
    '''
    elements ={}
    sheet = file.sheet_by_name("element_data")
    rows = sheet.nrows

    for i in range(1, rows):
        page = sheet.row_values(i)[0]
        if page != "":
            current_page = page
        element_name = sheet.row_values(i)[1]
        if element_name != "":
            element_id = sheet.row_values(i)[2]
            if element_id != "":
                elements[ current_page +"-"+ element_name ] = element_id

    return elements,"id"

def get_elememt_xpath(file):
    '''
    获取文档的所有控件数据
    :param file: 调用openfile
    :return: 控件字典
    '''
    elements ={}
    sheet = file.sheet_by_name("element_data")
    rows = sheet.nrows

    for i in range(1, rows):
        page = sheet.row_values(i)[0]
        if page != "":
            current_page = page
        element_name = sheet.row_values(i)[1]
        if element_name != "":
            element_id = sheet.row_values(i)[2]
            element_xpath = sheet.row_values(i)[3]
            if element_id == "" and element_xpath != "":
                elements[ current_page +"-"+ element_name ] = element_xpath

    return elements,"x"

def get_elememt_zb(file):
    '''
    获取文档的所有控件数据
    :param file: 调用openfile
    :return: 控件字典
    '''
    elements ={}
    sheet = file.sheet_by_name("element_data")
    rows = sheet.nrows

    for i in range(1, rows):
        page = sheet.row_values(i)[0]
        if page != "":
            current_page = page
        element_name = sheet.row_values(i)[1]
        if element_name != "":
            element_id = sheet.row_values(i)[2]
            element_xpath = sheet.row_values(i)[3]
            element_zb = sheet.row_values(i)[4]
            if element_id == "" and element_xpath == "" and element_zb != "":
                elements[ current_page +"-"+ element_name ] = element_zb

    return elements,"zb"

def assert_element(driver,by,valuse):
    '''
    控件断言，异常抛出
    :param driver:
    :param by:
    :param valuse:
    :return:
    '''
    result = "pass"
    try:
        find(driver,by,valuse)
    except Exception as e:
        result = "fail"
        print("error: ",e,valuse)
    return result

def result_return(assert_result,case_name,i,check_page_element,case_ispass=True):
    '''
    返回结果，包括是用例否通过，有问题的行数
    :param assert_result: 断言结果
    :param case_name: 用例名称
    :param i: 执行到的行数
    :param check_page_element: 断言用到的控件名
    :return:
    '''
    print("----------------------------------------------------------------------")
    print("Case_name:", case_name)
    if case_ispass:
        try:
            assert assert_result == "pass"
            print("Result:PASS")
            print("----------------------------------------------------------------------")
        except Exception as e:
            print("Result:FAIL")
            print("Problem:", "第", i, "行用例，没有找到控件: ", check_page_element)
            print("----------------------------------------------------------------------")
    else:
        print("Result:FAIL")
        print("Problem:", "第", i, "行用例，没有找到控件: ", check_page_element)
        print("----------------------------------------------------------------------")

def case_return(assert_result,i,check_page_element):
    '''
    单条用例的执行结果，如果有问题就返回异常，没有问题，就不展示任何东西
    :param assert_result: 断言结果
    :param case_name: 用例名称
    :param i: 执行到的行数
    :param check_page_element: 断言用到的控件名
    :return:
    '''
    case_ispass = True
    try:
        assert assert_result == "pass"
    except Exception as e:
        case_ispass = False
        print("步骤预期值问题:", "第", i, "行用例，没有找到预期控件: ", check_page_element)
        print("................................................")
    return case_ispass

def run_action(driver,action,by_way,element,test_data,check_element,case_name,i,check_page_element):
    '''
    根据行为进行操作，关键方法
    :param action: 行为
    :param driver: 驱动
    :param by_way: 定位方法
    :param element: 对象的控件id
    :param test_data: 测试数据
    :param check_element: 预期对象的控件id
    :param case_name: 用例名称
    :param i: 执行到第i行
    :param check_page_element: 预期对象的控件名称
    :return:
    '''
    case_ispass = True
    if action == "click":
        try:
            the_element = find(driver, by_way, element)
            click(the_element)

        except Exception as e:
            print("第 ", i, " 行用例，没有找到控件: ", e)
        if check_element != "":
            assert_result = assert_element(driver, by_way, check_element)
            case_ispass = case_return(assert_result, i, check_page_element)

    if action == "send_key":
        try:
            the_element = find(driver, by_way, element)
            if test_data != "":
                if isinstance(test_data, str):
                    send(the_element, test_data)
                else:
                    send(the_element, str(int(test_data)))
        except Exception as e:
            print("第 ", i, " 行用例，没有找到控件: ", e)

    if action == "clear":
        try:
            the_element = find(driver, by_way, element)
            clear(the_element)
        except Exception as e:
            print("第 ", i, " 行用例")
            print("没有找到控件: ", e," 或者控件非输入框，请重新确认！！")

    if action == "over_continue":

        assert_result = assert_element(driver, by_way, check_element)
        result_return(assert_result, case_name, i, check_page_element,case_ispass)
        teardowm(driver)
        driver.launch_app()

    if action == "over_reset_continue":

        assert_result = assert_element(driver, by_way, check_element)
        result_return(assert_result, case_name, i, check_page_element,case_ispass)
        teardowm(driver)
        driver.reset()

    if action == "over":

        assert_result = assert_element(driver, by_way, check_element)
        result_return(assert_result, case_name, i, check_page_element,case_ispass)
        driver.reset()
        print("测试完毕")
        print("正在清除测试数据...")
        teardowm(driver)
        print("数据清除完毕...")