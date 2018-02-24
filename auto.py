# -*- coding: UTF-8 -*-

from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from unpresent_wait import unpresence_of_element_located
from xpath_utils import *
from excel_utils import *
import traceback


class CRMAuto:
    RETRY_TIMES = 5

    def __init__(self, excelPath, iniPath):
        self.driver = None
        self.wait = None
        self.row_index = 1
        self.interval = 1
        self.xpath = XPathHelper(iniPath)
        self.excel = ExcelHelper(excelPath)

    def set_excel(self, excelPath):
        self.excel = ExcelHelper(excelPath)

    def set_xpath(self, iniPath):
        self.xpath = XPathHelper(iniPath)

    def set_row_index(self, row_index):
        self.row_index = row_index

    def set_interval(self, interval):
        self.interval = interval

    def set_excel_path(self, path):
        self.excel = ExcelHelper(path)

    def set_init_path(self, path):
        self.xpath = XPathHelper(path)

    def set_value(self, xpath, input_xpath, value):
        times = 0
        success = False
        error = None
        while times < CRMAuto.RETRY_TIMES:
            times += 1
            try:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                elem.click()
                time.sleep(self.interval)
                elem.click()
                time.sleep(self.interval)
                input = self.wait.until(EC.presence_of_element_located((By.XPATH, input_xpath)))
                input.clear()
                input.send_keys(value)
                success = True
                break
            except Exception as e:
                error = e
                time.sleep(self.interval / 2)
                print("set_value失败，再次尝试")
                traceback.print_exc(e)
        if not success:
            raise error

    def get_elem_text(self, xpath):
        element = self.driver.find_element_by_xpath(xpath)
        return element.text

    def set_input_value(self, xpath, value):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.send_keys(value)

    def single_click(self, xpath):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.click()
        time.sleep(self.interval)

    def double_click(self, xpath):
        elem = self.driver.find_element_by_xpath(xpath)
        elem.click()
        time.sleep(self.interval)
        elem.click()

    def remove_readonly(self):
        times = 0
        success = False
        error = None
        while times < CRMAuto.RETRY_TIMES:
            times += 1
            try:
                elem = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, self.xpath.get_wifi_xpath("wifi_user_type"))))
                elem.click()
                time.sleep(self.interval)
                elem.click()
                time.sleep(self.interval)
                js_code = "document.getElementById('editor-combobox$text').removeAttribute('readonly')"
                self.driver.execute_script(js_code)
                elem = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, self.xpath.get_wifi_xpath("wifi_location"))))
                elem.click()
                time.sleep(self.interval)
                elem.click()
                time.sleep(self.interval)
                success = True
                break
            except Exception as e:
                error = e
                time.sleep(self.interval / 2)
                print("remove_readonly失败，再次尝试")
                traceback.print_exc(e)
        if not success:
            raise error

    def close_alert(self):
        self.driver.switch_to.alert.accept()

    def switch_input_frame(self):
        self.driver.switch_to.default_content()

        frame = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, self.xpath.get_frame_xpath("input_frame_layer1")))
        )
        self.driver.switch_to.frame(frame)
        frame2 = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, self.xpath.get_frame_xpath("input_frame_layer2")))
        )
        self.driver.switch_to.frame(frame2)

    def switch_addr_frame(self):
        self.driver.switch_to.default_content()
        frame = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, self.xpath.get_frame_xpath("addrframe_layer1")))
        )
        self.driver.switch_to.frame(frame)
        addrframe = self.wait.until(
            EC.presence_of_element_located(
                (By.ID, "addrframe"))
        )
        self.driver.switch_to.frame(addrframe)

    def wait_until_elem_appear(self, xpath):
        elem = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, xpath))
        )
        return elem

    # def wait_until_elem_disappear(self, xpath, interval, message):
    #     while True:
    #         try:
    #             time.sleep(interval)
    #             self.driver.find_element_by_xpath(xpath)
    #             continue
    #         except NoSuchElementException as e:
    #             print(message)
    #             break

    def auto_input(self):
        time.sleep(5)
        self.switch_input_frame()
        # wifi板块填写
        self.remove_readonly()
        self.set_value(self.xpath.get_wifi_xpath("wifi_user_type"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "用户类型"))
        self.set_value(self.xpath.get_wifi_xpath("wifi_location"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "场所名称"))
        self.set_value(self.xpath.get_wifi_xpath("user_phone"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "业主联系电话"))
        self.set_value(self.xpath.get_wifi_xpath("install_location"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "安装地址"))
        self.set_value(self.xpath.get_wifi_xpath("user_id"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "业主身份证号"))
        self.single_click(self.xpath.get_wifi_xpath("licence_box"))
        time.sleep(self.interval / 3)
        self.set_value(self.xpath.get_wifi_xpath("licence"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "工商执照"))
        self.single_click(self.xpath.get_wifi_xpath("wifi_switch_btn"))
        self.set_value(self.xpath.get_wifi_xpath("save_check"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "安全审计合作商"))
        self.single_click(self.xpath.get_wifi_xpath("wifi_switch_btn"))
        self.set_value(self.xpath.get_wifi_xpath("internet_type"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "上网方式"))
        self.set_value(self.xpath.get_wifi_xpath("manager_phone"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "客户经理联系电话"))
        self.set_value(self.xpath.get_wifi_xpath("manager_name"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "客户经理名称"))
        self.set_value(self.xpath.get_wifi_xpath("legal_phone"), self.xpath.editor_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "法人联系方式"))
        self.set_value(self.xpath.get_wifi_xpath("cust_type"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.WIFI, self.row_index, "客户类型"))

        # adsl板块填写
        self.single_click(self.xpath.get_adsl_xpath("adsl_switch_btn"))
        adsl_speed = self.get_elem_text(self.xpath.get_adsl_xpath("adsl_speed"))
        adsl_account = self.get_elem_text(self.xpath.get_adsl_xpath("adsl_media_accout"))
        relative_account = self.get_elem_text(self.xpath.get_adsl_xpath("adsl_account"))
        self.set_value(self.xpath.get_adsl_xpath("wx_type"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.ADSL, self.row_index, "外线方式"))
        self.single_click(self.xpath.get_adsl_xpath("adsl_switch_btn"))
        time.sleep(self.interval)
        self.double_click(self.xpath.get_adsl_xpath("adsl_install_location"))
        self.wait_until_elem_appear(self.xpath.get_adsl_xpath("install_btn"))
        self.single_click(self.xpath.get_adsl_xpath("install_btn"))
        # 切到弹窗中
        self.switch_addr_frame()
        # 弹窗更下一层
        addrframe = self.wait.until(
            EC.presence_of_element_located(
                (By.ID, "addrframe"))
        )
        self.driver.switch_to.frame(addrframe)

        self.wait_until_elem_appear(self.xpath.get_adsl_xpath("addr_input"))
        self.set_input_value(self.xpath.get_adsl_xpath("addr_input"),
                             self.excel.get_cell_value(self.excel.ADSL, self.row_index, "地址ID"))
        self.single_click(self.xpath.get_adsl_xpath("addr_search"))
        elem = self.wait_until_elem_appear(self.xpath.get_adsl_xpath("addr_item"))
        elem.click()
        elem = self.wait_until_elem_appear(self.xpath.get_adsl_xpath("fttx_radio"))
        elem.click()
        time.sleep(self.interval)

        # 切回弹窗
        self.switch_addr_frame()
        self.single_click(self.xpath.get_adsl_xpath("addr_confirm"))
        self.wait.until(unpresence_of_element_located((By.XPATH, self.xpath.get_adsl_xpath("addr_confirm"))))
        # 切回主frame
        self.switch_input_frame()
        self.set_value(self.xpath.get_adsl_xpath("adsl_password"), self.xpath.button_input,
                       self.excel.get_cell_value(self.excel.ADSL, self.row_index, "密码"))
        if relative_account == " ":
            self.single_click(self.xpath.get_adsl_xpath("adsl_account_box"))
        self.set_value(self.xpath.get_adsl_xpath("adsl_account"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.ADSL, self.row_index, "账号关联账号"))
        self.set_value(self.xpath.get_adsl_xpath("adsl_user_type"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.ADSL, self.row_index, "用户类型"))
        self.set_value(self.xpath.get_adsl_xpath("adsl_rent"), self.xpath.combobox_input,
                       self.excel.get_cell_value(self.excel.ADSL, self.row_index, "月租类型"))
        self.single_click(self.xpath.get_adsl_xpath("adsl_user_type"))
        self.single_click(self.xpath.get_adsl_xpath("ty_box"))

        # itv板块的录入
        self.single_click(self.xpath.get_itv_xpath("itv_switch"))
        self.single_click(self.xpath.get_itv_xpath("itv_service"))
        self.single_click(self.xpath.get_itv_xpath("itv_service"))
        self.wait_until_elem_appear(self.xpath.get_itv_xpath("service_btn"))
        time.sleep(self.interval)
        self.single_click(self.xpath.get_itv_xpath("service_btn"))
        elem = self.wait_until_elem_appear(self.xpath.get_itv_xpath("service_item"))
        elem.click()
        time.sleep(self.interval * 2)
        self.set_value(self.xpath.get_itv_xpath("itv_password"), self.xpath.button_input,
                       self.excel.get_cell_value(self.excel.ADSL, self.row_index, "密码"))

        # 回去填账号和速率
        self.single_click(self.xpath.get_wifi_xpath("wifi_switch_btn"))
        self.single_click(self.xpath.get_wifi_xpath("wifi_account_box"))
        self.single_click(self.xpath.get_wifi_xpath("speed_box"))
        time.sleep(self.interval)
        self.set_value(self.xpath.get_wifi_xpath("wifi_account"), self.xpath.editor_input, adsl_account)
        self.set_value(self.xpath.get_wifi_xpath("wifi_speed"), self.xpath.editor_input, adsl_speed)

        self.single_click(self.xpath.get_wifi_xpath("wifi_service"))
        self.single_click(self.xpath.get_wifi_xpath("wifi_service"))
        self.wait_until_elem_appear(self.xpath.get_wifi_xpath("service_btn"))
        time.sleep(self.interval)

        self.single_click(self.xpath.get_wifi_xpath("service_btn"))
        elem = self.wait_until_elem_appear(self.xpath.get_wifi_xpath("service_item"))
        elem.click()
        time.sleep(self.interval * 2)

        self.driver.switch_to.default_content()

    def login(self, account, password):
        self.driver = webdriver.Firefox()
        self.wait = WebDriverWait(self.driver, 15)
        self.driver.get("http://132.121.101.140:8001/crmbs/mgr0/sysmgr/login/login.jsp")
        # try:
        #     elem = self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'close')))
        #     elem.click()
        # except Exception as e:
        #     traceback.print_exc()
        #     print("没有弹窗")

        elem = self.driver.find_element_by_id("userid")
        elem.clear()
        elem.send_keys(account)
        elem = self.driver.find_element_by_id("password")
        elem.clear()
        elem.send_keys(password)
        elem = self.wait.until(EC.element_to_be_clickable((By.ID, 'loginBnt')))
        elem.click()
