# -*- coding: utf-8 -*-

from CommonMethod import *

def runtest(choice = "case1"):
    '''
    跑case的主函数
    :param choice:
    :return:
    '''
    driver = setup()

    file = openfile()
    elements_dict_id = get_elememt_id(file)[0]
    elements_dict_xpath = get_elememt_xpath(file)[0]
    elements_dict_zb = get_elememt_zb(file)[0]
    #print(elements_dict)
    sheet = file.sheet_by_name(choice)
    rows = sheet.nrows

    for i in range(3,rows):
        isrun = sheet.row_values(i)[0]
        case = sheet.row_values(i)[1]
        action = sheet.row_values(i)[2]
        page = sheet.row_values(i)[3]
        element = sheet.row_values(i)[4]
        test_data = sheet.row_values(i)[5]
        check_page = sheet.row_values(i)[6]
        check_element = sheet.row_values(i)[7]
        page_element = page + "-" + element
        check_page_element = check_page + '-' + check_element
        by_way = "id"

        if page != "" and element != "":
            if page_element in elements_dict_id:
                element = elements_dict_id[page_element]
                by_way = "id"

            elif page_element in elements_dict_xpath:
                element = elements_dict_xpath[page_element]
                by_way = "x"

            elif page_element in elements_dict_zb:
                element = elements_dict_zb[page_element]
                by_way = "zb"

        if check_page != "" and check_element != "":
            if check_page_element in elements_dict_id:
                check_element = elements_dict_id[check_page_element]
                by_way = "id"

            elif check_page_element in elements_dict_xpath:
                check_element = elements_dict_xpath[page_element]
                by_way = "x"

            elif check_page_element in elements_dict_zb:
                check_element = elements_dict_zb[check_page_element]
                by_way = "zb"

        if case != "":
            case_name = case

        if isrun == "on":
            run_action(driver, action, by_way,element, test_data, check_element, case_name, i, check_page_element)
            if action == "over":
                break

if __name__ == "__main__":
    runtest()