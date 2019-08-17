#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019-07-20 14:06
# @Author  : Jame
# @Site    : 
# @File    : invoice_check.py
# @Software: PyCharm

import json
import os
import pickle
import shutil
import time
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium .webdriver.support import expected_conditions as EC
from pathlib import Path

import requests
import json
import base64
import cv2
import keyboard
import sys
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt



class BaiduInfo:
    """
        百度云增值税发票 ocr 接口，每天免费 500 张
        appid: 16940491
        client_id: 
        secret_key: 
    """
    def __init__(self, client_id="", client_secret=""):
        self.api_url = "https://aip.baidubce.com/oauth/2.0/token"
        self.data = {
            "grant_type": "client_credentials",
            "client_id":client_id,
            "client_secret":client_secret
        }
        self.headers = {
            "Content-Type": "application/json",
            "charset": "UTF-8"
        }

    def get_access_token(self):
        r = requests.post(self.api_url, data=self.data, headers=self.headers)
        access_token = json.loads(r.content)["access_token"]
        return str(access_token)


class DatasUtils:
    def __init__(self):
        self.api_url_invoice = "https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice"
        self.access_token = BaiduInfo().get_access_token()
        self.api_url_invoice = self.api_url_invoice + "?access_token=" + self.access_token
        self.headers = {
            "Content-Type":"application/x-www-form-urlencoded"
        }

        self.data = {

        }

    def image2base64(self, path):
        with open(path, 'rb') as f:
            image_data_base64 = base64.b64encode(f.read())
            return image_data_base64

    def scan_image(self, path=""):
        """
        需要字段为：发票代码，发票号码，开票时间，发票不含税金额
            识别情况需要做异常处理，识别之后的情况有以下几种：
            1、成功
            2、失败
                2.1 全部无法识别
                2.2 部分识别
                    2.2.1 只识别出来部分字段(发票代码|发票号码|开票时间|发票不含税金额),其他字段缺失
                    2.2.2 识别出来部分字段错误

        :param path:
        :return:
        """
        if not path:
            raise ValueError("图片路径不能为空")
        self.data["image"] = self.image2base64(path)
        r =None
        try:
            r = requests.post(self.api_url_invoice, data=self.data, headers=self.headers)
        except Exception as e:
            print(e)
        r = r.json()
        # print(r)
        result = json.loads("{}")
        result["result"] = "成功"
        result["InvoiceNum"] = ""
        result["InvoiceDate"] = ""
        result["InvoiceCode"] = ""
        result["TotalAmount"] = ""
        try:
        #获取所需字段信息
            result["InvoiceNum"] = r["words_result"]["InvoiceNum"]
            result["InvoiceDate"] = r["words_result"]["InvoiceDate"]
            result["InvoiceCode"] = r["words_result"]["InvoiceCode"]
        #对日期格式进行处理
            result["InvoiceDate"] = result["InvoiceDate"].replace("年", "")
            result["InvoiceDate"] = result["InvoiceDate"].replace("月", "")
            result["InvoiceDate"] = result["InvoiceDate"].replace("日", "")
            result["TotalAmount"] = str(round(float(r["words_result"]["TotalAmount"]),2))
        except Exception as e:
            print(e)
        return result

    def show_image(self, path):
        img = cv2.imread(path)
        cv2.startWindowThread()
        win_name = '票据-'+path.split('.')[0]
        # cv2.namedWindow(win_name)
        cv2.namedWindow(win_name, cv2.WINDOW_NORMAL)
        cv2.resizeWindow(win_name, 1000, 1000)
        cv2.moveWindow(win_name, 200, 200)
        cv2.imshow(win_name, img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

    def get_images(self, path):
        img_types = [
            "jpg",
            "jpeg",
            "png"
        ]
        self.files = []
        for type in img_types:
            self.files.extend(Path(path).glob("*."+type))
        return self.files

    def move_image(self, source, dest):
        if not os.path.exists(dest):
            os.mkdir(dest)
        shutil.move(source, dest)

#TODO:
"""

    1. 票据图像处理
        1.1 票据定位，去除空白位置，给票据增加遮罩
        2.2 票据增强，有的票据比较模糊
        
    2. 结果报表生成
    
#TODO2


"""
class SiteAction:
    def __init__(self, dir_datas="datas"):
        self.url = "https://inv-veri.chinatax.gov.cn/"
        chrome_options = ChromeOptions()
        chrome_options.add_argument('--ignore-certificate-errors')
        self.data_utils = DatasUtils()
        # chrome_options.add_argument('--headless')
        # chrome_options.add_argument(
        #     'user-agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"')
        # self.browser = webdriver.Chrome(options=chrome_options)
        # self.browser.minimize_window()
        self.images_files =self.data_utils.get_images(dir_datas)
        self.browser = None
        self.current_files_idx = 0
        self.status = None # 流程状态, init_succsss, input_success, code_success, check_success,submit_sucess, code_error, finished
        self.status = "init_success"
        self.current_invoice_info = None
        self.status_pause = False
        self.write_to_file_idx = 0
        self.go_to_index()
        keyboard.add_hotkey("enter", self.next_action)
        keyboard.add_hotkey("esc", self.invoice_next)
        keyboard.add_hotkey("shift+enter", self.info_reinput)
        keyboard.add_hotkey("shift+ctrl+enter", self.info_input)
        keyboard.add_hotkey("shift+o", self.browser_reopen)
        keyboard.add_hotkey("shift+p", self.pause)
        keyboard.add_hotkey("shift+[", self.invoice_previous)
        keyboard.add_hotkey("shift+f", self.current_file_check)
        # keyboard.hook_key("enter", self.next_action())

    # def browser_elment_text_alter(self, id="ktsm_tip", value=""):
    def browser_elment_text_alter(self, value=""):
        id="ktsm_tip"
        # script_text = 'document.getElementById(\"'+id+'\").children[0].innerHTML = \"<span style=\"color:red;font-size:30px;margin-left:3%;\">'+value+'</span>\"'
        script_text = 'document.getElementById(\"'+id+'\").children[0].innerHTML = \"'+value+'</span>\"'
        text_size = 'document.getElementById(\"' + id + '\").children[0].style.fontSize = "20px"'
        # self.browser.find_element_by_id()
        try:
            self.browser.execute_script(script_text)
            self.browser.execute_script(text_size)
        except Exception as e:
            print(e)
    def browser_alter(self, value=""):
        script_text = "alert(\"" + value + "\")"
        try:
            self.browser.execute_script(script_text)
        except Exception as e:
            print(e)

    def images_directory(self, des_success_datas = "datas\\查验成功", des_fail_datas = "datas\\查验失败"):  # cyjg
        self.browser.switch_to.frame("dialog-body")
        # print(des_fail_datas)
        if self.element_exist_by_id('cyjg'):
            try:
                shutil.move(str(self.images_files[self.current_files_idx]), des_fail_datas)
            except:
                pass
        else:
            try:
                shutil.move(str(self.images_files[self.current_files_idx]), des_success_datas)
            except:
                pass


    def add_hotkey(self):
        keyboard.add_hotkey("enter", self.next_action)

    def element_wait(self, timeout=10, id="", tag_name="", class_name="", xpath=""):
        """
            selenium 继续操作等待函数
        :param timeout: 超时时间，即最大等待时间
        :param id: 查找元素的id，找到了则停止等待
        :param tag_name: 查找元素的标签
        :param class_name: 查找元素的id，找到了则停止等待
        :return:
            id不为空: Fasle, None;
            id为空: Fasle, webElement;
        """
        element = None
        if  id:
            element = WebDriverWait(self.browser, timeout).until(
                EC.presence_of_element_located((By.ID, id))
            )
        elif  tag_name:
            element = WebDriverWait(self.browser, timeout).until(EC.presence_of_element_located((By.TAG_NAME, tag_name)))
        elif  class_name:
            element = WebDriverWait(self.browser, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, class_name)))
        elif xpath:
            element = WebDriverWait(self.browser, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

        return element

    def pause(self):
        if not self.status_pause:
            # self.browser_alter("已暂停")
            self.browser_elment_text_alter("已暂停")
            self.status_before_pause = self.status
            self.status = "pause"
            self.status_pause = True
        else:
            # self.browser.switch_to.alert.accept()
            self.browser_elment_text_alter("请继续")
            self.status = self.status_before_pause
            self.status_pause = False


    def text_fill(self, id, key):
        """
            向指定id发送数据
        :param id:
        :param key:
        :return:
        """
        element = self.element_wait(id=id)
        element.clear()
        element.send_keys(key)

    def button_click_by_id_wait(self, id):
        button = WebDriverWait(self.browser, 10).until(EC.presence_of_element_located((By.ID, id)))
        button.click()

    def button_click_by_id(self, id):
        button = self.browser.find_element_by_id(id)
        button.click()

    def button_click_by_contains(self, type, value):
        button = None
        if "id" in type:
            button = self.browser.find_element_by_xpath(("//button[contains(@id,"+value+")]"))
        elif "class" in type:
            button = self.browser.find_element_by_xpath(("//button[contains(@class,"+value+")]"))
        button.click()

    def button_click_by_xpath(self, xpath):
        button = WebDriverWait(self.browser, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        button.click()

    def element_exist_by_id(self, element_id="", timeout=3):
        try:
            button = WebDriverWait(self.browser, timeout).until(EC.presence_of_element_located((By.ID, element_id)))
            self.browser.find_element_by_id(element_id)
            return True
        except:
            return False

    def info_reinput(self):
        if not self.current_invoice_info:
            self.current_invoice_info = self.data_utils.scan_image(self.images_files[self.current_files_idx])
        self.element_wait(id="uncheckfp")
        self.browser_elment_text_alter("正在获取票据信息："+str(self.images_files[self.current_files_idx]).split("\\")[-1]+", 请稍等，票据信息获取需3s左右.....")
        print("正在获取票据信息："+str(self.images_files[self.current_files_idx]).split("\\")[-1]+", 请稍等，票据信息获取需3s左右.....")
        # self.current_invoice_info = self.data_utils.scan_image(self.images_files[self.current_files_idx])
        if not len(self.current_invoice_info):
            self.browser_elment_text_alter("票据信息获取失败！")
            print("票据信息获取失败！")
            return 0
        else:
            # print(self.current_invoice_info)
            self.browser_elment_text_alter("票据信息获取成功，正在填写")
            print("票据信息获取成功，正在填写")
            self.text_fill('fpdm', self.current_invoice_info["InvoiceCode"])
            self.text_fill('fphm', self.current_invoice_info["InvoiceNum"])
            self.text_fill('kprq', self.current_invoice_info["InvoiceDate"])
            self.text_fill('kjje', self.current_invoice_info["TotalAmount"])
            self.text_fill("yzm", "")
            self.browser_elment_text_alter("填写完成")
            print("填写完成")
            # keyboard.wait("enter")
            self.status = "input_success"
        # self.add_hotkey()

    def write_to_file(self, workbook ="result.xls"):
        if not os.path.exists(workbook):
            print("结果文件未存在！")
            workbook_obj = xlwt.Workbook()
            sheet = workbook_obj.add_sheet("查询结果")
            sheet.write(0, 0, "序号")
            sheet.write(0, 1, "发票代码")
            sheet.write(0, 2, "发票号码")
            sheet.write(0, 3, "开票日期")
            sheet.write(0, 4, "开具金额（不含税）")
            sheet.write(0, 5, "查询状态")
            sheet.write(0, 6, "文件名称")
            workbook_obj.save('result.xls')
        r_xls = open_workbook(workbook)  # 读取excel文件
        row = r_xls.sheets()[0].nrows  # 获取已有的行数
        excel = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
        table = excel.get_sheet(0)  # 获取要操作的sheet
        # 对excel表追加一行内容
        self.write_to_file_idx += 1
        table.write(row, 0, row)  # 括号内分别为行数、列数、内容
        table.write(row, 1, self.current_invoice_info["InvoiceCode"])
        table.write(row, 2, self.current_invoice_info["InvoiceNum"])
        table.write(row, 3, self.current_invoice_info["InvoiceDate"])
        table.write(row, 4, self.current_invoice_info["TotalAmount"])
        table.write(row, 5, self.current_invoice_info["result"])
        table.write(row, 6, str(self.images_files[self.current_files_idx]).split("\\")[-1])
        self.data_utils.move_image(str(self.images_files[self.current_files_idx]), "datas/成功")
        self.images_files.pop(self.current_files_idx)
        self.current_files_idx -= 1
        excel.save(workbook)  # 保存并覆盖文件

    def info_input(self):
        if self.current_files_idx >= len(self.images_files):
            self.browser.execute_script("alert(\"查验结束\")")
            return 1
        self.browser_elment_text_alter("正在获取票据信息："+str(self.images_files[self.current_files_idx]).split("\\")[-1]+", 请稍等，票据信息获取需3s左右.....")
        print("正在获取票据信息："+str(self.images_files[self.current_files_idx]).split("\\")[-1]+", 请稍等，票据信息获取需3s左右.....")
        self.current_invoice_info = self.data_utils.scan_image(str(self.images_files[self.current_files_idx]))
        self.info_reinput()

    def browser_reopen(self):
        self.browser = webdriver.Chrome()
        self.browser.get(self.url)

    def code_refresh(self):
        try:
            #TODO: 111
            time.sleep(1)
            self.browser.find_element_by_id("yzm_img").click()
            self.text_fill("yzm", "")
        except:
            try:
                self.browser.find_element_by_id("yzm_unuse_img").click()
                self.text_fill("yzm", "")
            except:
                print("code_refresh fail!")
                self.popup_win_close()



    def submit(self):
        if self.info_input_check():
            try:
                self.button_click_by_id_wait("checkfp")
                # time.sleep(1)
                try:
                    self.browser.find_element_by_id("popup_ok")
                    self.status = "code_error"
                except:
                    pass
            except:
                self.status = "code_error"
                print("checkfp点击失败")
                return 0
            time.sleep(2)
            try:
                self.browser.find_element_by_tag_name("iframe")
                self.status = "submit_success"
            except:
                self.status = "input_success"
                self.text_fill("yzm","")
                self.code_refresh()
        else:
            self.browser_elment_text_alter("验证码未输入或输入错误！")
            print("验证码未输入或输入错误！")

    def popup_win_close(self):
        try:
            self.button_click_by_id("popup_ok")
        except:
            pass

    def invoice_next(self):
        self.browser_elment_text_alter("正在操作，请稍等！")
        try:
            if self.element_exist_by_id("cycs"):
                self.current_invoice_info["result"] = "成功"
            elif self.element_exist_by_id("cyjg"):
                self.current_invoice_info["result"] = "不一致"
            self.status = "init_success"
        except:
            self.status = "init_success"
            self.go_to_index()
        self.invoice_skip()

    def info_input_check(self):
        invoice_code = self.browser.find_element_by_id("fpdm").get_attribute("value")
        invoice_num = self.browser.find_element_by_id("fphm").get_attribute("value")
        invoice_date = self.browser.find_element_by_id("kprq").get_attribute("value")
        invoice_amount = self.browser.find_element_by_id("kjje").get_attribute("value")
        code_values = self.browser.find_element_by_id("yzm").get_attribute("value")
        if invoice_code == "":
            self.browser_elment_text_alter("请输入发票代码!")
            return False
        if invoice_num == "":
            self.browser_elment_text_alter("请输入发票号码!")
            return False
        if invoice_date == "YYYYMMDD":
            self.browser_elment_text_alter("请输入开票日期!")
            return False
        if invoice_amount == "":
            self.browser_elment_text_alter("请输入发票金额!")
            return False
        if  code_values == "请输入验证码" or code_values == "":
            # print(self.browser.find_element_by_id("yzm").get_attribute("value"))
            self.browser_elment_text_alter("请输入验证码！")
            print("请输入验证码！")
            return False
        self.status = "code_input_success"
        return True

    def invoice_skip(self):
        if self.current_files_idx < len(self.images_files)-1:
            self.current_files_idx += 1
            self.go_to_index()
        else:
            self.status = "finishes"
            self.browser.execute_script("alert(\"查验结束\")")
            print("票据查验结束..")

    def invoice_previous(self):
        self.status = "init_success"
        if self.current_files_idx == 0:
            self.browser_elment_text_alter("已经是第一张")
            self.go_to_index()
        else:
            self.current_files_idx -= 1
            self.go_to_index()

    def status_check(self):
        """
        浏览器状态校验，手动操作也能自动更新状态
        :return:
        """
        pass

    def next_action(self):
        print("next_action triggerd")
        if self.status == 'init_success':
            print("init_success")
            self.info_input()
        elif self.status == 'input_success':
            print("input_success")
            # print(self.info_input_check())
            if self.info_input_check():
                print("info_input_check true and submit")
                self.submit()
            else:
                print("code_refresh")
                self.code_refresh()
                # keyboard.wait("enter")
        elif self.status == 'submit_success':
            print("submit_success")
            print(str(self.images_files[self.current_files_idx])+" 数据写入")
            self.write_to_file()
            print(str(self.images_files[self.current_files_idx])+" 数据写入完成")
            self.invoice_next()
        elif self.status == "code_error":
            # keyboard.wait("enter")
            self.popup_win_close()
        # self.add_hotkey()
        # keyboard.wait("enter")

    def invoice_check(self):
        """
        ctrl+q退出程序
        :return:
        """
        keyboard.wait("ctrl+q")

    def go_to_index(self):
        if self.browser == None:
            self.browser = webdriver.Chrome()
            self.browser.maximize_window()
        self.browser.get(self.url)
        self.status = "init_success"
        print("status update success")
        self.element_wait()
        self.browser_elment_text_alter("当前为第"+str(self.current_files_idx+1)+"张，共"+str(len(self.images_files))+"张")

    def current_file_check(self):
        # print(str(self.images_files[self.current_files_idx]))
        # self.browser_alter(str(self.images_files[self.current_files_idx]).split("\\")[-1])
        self.browser_elment_text_alter("当前票据为：第"+str(self.current_files_idx+1)+"张，"+str(self.images_files[self.current_files_idx]).split("\\")[-1])
        self.data_utils.show_image(str(self.images_files[self.current_files_idx]))


if __name__ == "__main__":
    check = SiteAction()
    try:
        check.invoice_check()
    except Exception as e: # au = DatasUtils()
        print(e)
    finally:
        check.browser.quit()

