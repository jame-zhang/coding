#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019-07-04 11:59
# @Author  : Jame
# @Site    :
# @File    : ftb_python2.py
# @Software: PyCharm
import pandas as pd
from glob import glob
import json
from tqdm import tqdm
from pandas import ExcelWriter
import os


class FTB_PYTHON:
    def __init__(self, bank_statement, organization_info, result="result.xlsx"):
        self.bank_statement_path = bank_statement
        self.organization_info_path = organization_info
        self.organization_info_json_path = organization_info.split('.')[0] + '.json'
        self.result_path = result
        self.bank_statement = None
        self.organization_info_json = None
        self.result_sheet1 = pd.DataFrame(columns=["序号", "行号", "币种序号", "币种", "账号", "户名", "证件号码", "科目", "申报金额", "是否申报"])
        self.result_sheet2 = pd.DataFrame({}, columns= ["行号", "申报数量", "申报总金额"])
        self.currency_type2code = {
            "人民币":"01",
            "英镑":"12",
            "港币":"13",
            "美元":"14",
            "瑞士法郎":"15",
            "新加坡元":"18",
            "瑞典克朗":"21",
            "挪威克朗":"23",
            "日元":"27",
            "加拿大元":"28",
            "澳大利亚元":"29",
            "欧元":"38"
        }

    def get_organization_info(self):
        self.organization_info = pd.read_excel(self.organization_info_path)
        # 数据预处理，去空格
        self.organization_info[["主账号", "子账号"]] = self.organization_info[["主账号", "子账号"]].apply(lambda x: x.str.strip())
        self.organization_info["子账号"] = self.organization_info["子账号"].replace("1", pd.np.nan)

    def get_organization_info_json(self):
        if not glob(self.organization_info_json_path):
            print("开户信息表格JSON文件不存在，正在创建.")
            self.get_organization_info()
            self.save_organization_info_json()
        else:
            print("开户信息表格JSON文件已创建，正在载入.")
            with open(self.organization_info_json_path, 'r', encoding='utf-8') as f:
                self.organization_info_json = json.load(f)
            print("载入完成！")

    def get_bank_statement(self):
        self.bank_statement = pd.read_excel(self.bank_statement_path)
        # 数据预处理，丢弃无关信息，替换逗号等
        self.bank_statement.drop(self.bank_statement.columns.tolist()[1:], axis=1, inplace=True)
        self.bank_statement.columns = ["内容"]
        self.bank_statement = self.bank_statement.apply(lambda x: x.str.replace(',', ''))

    def save_organization_info_json(self):
        self.organization_info_json = json.loads('{}')
        for idx in tqdm(range(len(self.organization_info))):
            account_info = json.loads('{}')
            account_info["户名"] = self.organization_info.iloc[idx]["户名"]
            account_info["证件号码"] = str(self.organization_info.iloc[idx]["证件号码"])
            self.organization_info_json[self.organization_info.iloc[idx]["主账号"]] = account_info
            if not self.organization_info.iloc[idx]["子账号"]:
                print(self.organization_info.iloc[idx]["子账号"])
                self.organization_info_json[self.organization_info.iloc[idx]["子账号"]] = account_info
        with open(self.organization_info_json_path, 'w', encoding='utf-8') as f:
            json.dump(self.organization_info_json, f, indent=4, ensure_ascii=False)

    def get_result_sheet1(self):
        self.get_organization_info_json()
        self.get_bank_statement()
        order = 1
        currency_type = ""
        bank_num = ""
        print("正在处理数据，请稍等")
        for idx in tqdm(range(len(self.bank_statement))):
            result_row = {
                "序号": "",
                "行号": "",
                "币种序号": "",
                "币种": "",
                "账号": "",
                "户名": "",
                "证件号码": "",
                "科目": "",
                "申报金额": "",
                "是否申报": ""

            }
            content = str(self.bank_statement.iloc[idx].values[0]).strip()
            # 检测表格开头
            if content.startswith("行号"):
                bank_num = content.split()[1]
                currency_type = content.split()[4]
            if content.startswith("FTN"):
                content = content.split()
                result_row["行号"] = bank_num
                result_row["申报金额"] = content[-5]
                if float(result_row["申报金额"]) > 0.5:
                    result_row["序号"] = order
                    order += 1
                    result_row["是否申报"] = '是'
                else:
                    if order == 1:
                        result_row["序号"] = order
                    else:
                        result_row["序号"] = order - 1
                    result_row["是否申报"] = '否'
                result_row["账号"] = content[0]
                try:
                    result_row["户名"] = self.organization_info_json[result_row["账号"]]["户名"]
                    result_row["证件号码"] = self.organization_info_json[result_row["账号"]]["证件号码"]
                except:
                    result_row["户名"] = "信息表无此账号"
                result_row["币种"] = currency_type
                result_row["币种序号"] = self.currency_type2code[currency_type]
                result_row["科目"] = content[1]
                self.result_sheet1 = self.result_sheet1.append(result_row, ignore_index=True)
        print("数据处理完成！")

    def save_result_file(self):
        writer = ExcelWriter(self.result_path)
        if not len(self.result_sheet1):
            self.get_result_sheet1()
        if not len(self.result_sheet2):
            self.get_result_sheet2()
        self.result_sheet1.to_excel(writer, "sheet1", index=False)
        self.result_sheet2.to_excel(writer, "sheet2", index=False)
        writer.save()
        print("保存成功！")

    def get_result_sheet2(self):
        if not len(self.result_sheet1):
            self.get_result_sheet1()
        self.result_sheet2[["行号", "申报数量"]] = ftb.result_sheet1[["行号", "申报金额"]].groupby(["行号"]).count().reset_index()
        self.result_sheet1[["申报金额"]] = self.result_sheet1[["申报金额"]].astype(float)
        self.result_sheet2[["行号","申报总金额"]] = ftb.result_sheet1[["行号", "申报金额"]].groupby(["行号"]).sum().reset_index() 


if __name__ == '__main__':
    BANK_STATEMENT_FILE_NAME = '20190621.xls'
    ORGANIZATION_INFO_FILE_NAME = 'FT开户信息表.xls'
    bank_statment = os.path.join('datas', BANK_STATEMENT_FILE_NAME)
    organization_info = os.path.join('datas', ORGANIZATION_INFO_FILE_NAME)
    ftb = FTB_PYTHON(bank_statment, organization_info)
    ftb.save_result_file()
