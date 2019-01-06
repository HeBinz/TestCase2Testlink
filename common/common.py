#! /usr/bin/python
# coding:utf-8
"""
@author:Bingo.he
@file: common.py
@time: 2018/12/05
"""
import time

import xlrd
import os
import sys
import testlink
import json
from xlutils import copy
from common.logger import Log
from collections import Iterable

logger = Log(os.path.basename("TestLink"))


class TestLinkOperate:
    def __init__(self, server, api_key):
        self.tlc = testlink.TestlinkAPIClient(server, api_key)
        self.count_fail = 0
        self.count_success = 0

    def get_projects_info(self):
        """
        获取项目信息
        :return:
        """
        project_info = {}
        projects = self.tlc.getProjects()
        for project in projects:
            project_info[project['name']] = project['id']
        return json.dumps(project_info, indent=4, ensure_ascii=False)

    def get_projects_id(self):
        project_ids = []
        projects = self.tlc.getProjects()
        for project in projects:
            project_ids.append(project['id'])
        return project_ids

    def get_suites(self, suite_id):
        """
        获取用例集
        :return:
        """
        try:
            suites = self.tlc.getTestSuiteByID(suite_id)
            return suites
        except testlink.testlinkerrors.TLResponseError as e:
            # traceback.print_exc()
            logger.warning(str(e).split('\n')[1])
            logger.warning(str(e).split('\n')[0])
            sys.exit(1)

    def create_testcase(self, test_project_id, suits_id, data):
        """
        :param test_project_id:
        :param suits_id:
        :param data:
        :return:
        """
        # 设置优先级及摘要 默认值
        if data['importance'] not in [1, 2, 3]:
            data['importance'] = 3
        if data["summary"] == "":
            data["summary"] = "无"

        # 检测测试数据的有效性
        for k, v in data.items():
            if v == "" and k not in ['summary', 'importance']:
                logger.warning(
                    "TestCase '{title}' param '{k}' is null".format(title=data['title'], k=self.format_param(k)))

        # 初始化测试步骤及预期结果
        for i in range(0, len(data["step"])):
            self.tlc.appendStep(data["step"][i][0], data["step"][i][1], data["automation"])

        try:
            self.tlc.createTestCase(
                data["title"],
                suits_id,
                test_project_id,
                data["authorlogin"],
                data["summary"],
                preconditions=data["preconditions"],
                importance=data['importance'],
                executiontype=data["automation"])

            self.count_success += 1
            if (self.count_success % 10) == 0:
                logger.info("已成功导入 {} 条数据，失败 {} 条".format(self.count_success, self.count_fail))
        except Exception as e:
            logger.error("用例" + data["title"] + "导入错误：" + str(e))
            self.count_fail += 1
            sys.exit(1)

    @staticmethod
    def read_excel(file_path, sheet_num):
        """
        读取用例数据
        :return:
        """
        case_list = []
        try:
            book = xlrd.open_workbook(file_path)
        except Exception as error:
            logger.error('路径不在或者excel不正确 : ' + str(error))
            sys.exit(1)
        else:
            sheet = book.sheet_by_index(sheet_num)  # 取第几个sheet页
            rows = sheet.nrows  # 取这个sheet页的所有行数
            for i in range(rows):
                if i != 0:
                    case_list.append(sheet.row_values(i))  # 把每一条测试用例添加到case_list中
        return case_list

    @staticmethod
    def format_testcase(test_case):
        """
        格式化用例数据
        :param test_case:
        :return:
        """
        switcher = {
            "低": 1,
            "中": 2,
            "高": 3,
            "自动化": 2,
            "手工": 1
        }
        return {
            "title": test_case[0],
            "preconditions": test_case[1],
            "step": list(zip(str(test_case[2]).split('；'), str(test_case[3]).split('；'))),  # 以换行符作为测试步骤的分界
            "automation": switcher.get(test_case[4]),  # 1  手工, 2 自动
            "authorlogin": test_case[5],
            "importance": switcher.get(test_case[6]),
            "summary": test_case[7]
        }

    @staticmethod
    def format_param(source_data):
        """
        转换用例字段为中文
        :param source_data:
        :return:
        """
        switcher = {
            "title": "标题",
            "preconditions": "前置条件",
            "step": "步骤",
            "automation": "用例类型",
            "authorlogin": "拥有者",
            "importance": "优先级",
            "summary": "摘要"
        }
        return switcher.get(source_data, "Param not defind")

    def upload(self, test_project_id, test_father_id, test_file_name, sheet_num):

        # 对project_id father_id 用例格式 做有效性判断
        if test_project_id not in self.get_projects_id():
            logger.error('project_id is not auth')
            sys.exit(1)

        if not self.get_suites(test_father_id):
            logger.error('father_id is not auth')
            sys.exit(1)

        test_cases = self.read_excel(os.path.join('testCase', test_file_name), sheet_num)
        if not isinstance(test_cases, Iterable):
            logger.error('test_cases is not Iterable')
            sys.exit(1)

        logger.info("正在导入数据，请耐心等待...")
        for test_case in test_cases:
            testcase_data = self.format_testcase(test_case)
            self.create_testcase(test_project_id, test_father_id, testcase_data)

        logger.info("本次操作共提交 {} 条数据，成功导入 {} 条，失败 {} 条".format(self.count_success + self.count_fail,
                                                              self.count_success, self.count_fail))

    @staticmethod
    def format_execution_type(source_data):
        switcher = {
            '2': "自动化",
            '1': "手工"
        }
        return switcher.get(source_data, "Param not defind")

    @staticmethod
    def format_importance(source_data):
        switcher = {
            '1': "低",
            '2': "中",
            '3': "高"
        }
        return switcher.get(source_data, "Param not defind")

    @staticmethod
    def format_auth(source_data):
        switcher = {
            '131': "chenwei6",
            '130': "hebin",
            '129': "zhangshuhuan",
        }
        return switcher.get(source_data, "user-id is not handle")

    def save_suits(self, file_path, datas, father_id):
        book = xlrd.open_workbook(file_path, formatting_info=True)  # 读取Excel
        new_book = copy.copy(book)  # 复制读取的Excel
        sheet = new_book.get_sheet(0)  # 取第一个sheet页
        line_num = 1
        for i in range(0, len(datas)):
            name, preconditions, actions, expected_results, execution_type, author, importance, summary = datas[i]
            sheet.write(line_num, 0, u'%s' % name)
            sheet.write(line_num, 1, u'%s' % preconditions)
            sheet.write(line_num, 2, u'%s' % actions)
            sheet.write(line_num, 3, u'%s' % expected_results)
            sheet.write(line_num, 4, u'%s' % execution_type)
            sheet.write(line_num, 5, u'%s' % author)
            sheet.write(line_num, 6, u'%s' % importance)
            sheet.write(line_num, 7, u'%s' % summary)
            line_num += 1
        report_path = os.path.abspath(os.path.join('download'))
        if not os.path.exists(report_path):
            os.makedirs(report_path)
        suits_name = self.get_suites(father_id)["name"]
        new_book.save(os.path.abspath(os.path.join(report_path, '【{}】download@{}.xlsx'.format(suits_name, time.strftime(
            '%Y.%m.%d@%H-%M-%S')))))  # 保存修改过后复制的Excel
        logger.info("用例集:【{}】下载成功，文件保存于download目录".format(suits_name))

    def download(self, father_id):
        datas = []
        logger.info("正在下载数据，请耐心等待...")
        for data in self.tlc.getTestCasesForTestSuite(father_id, True, 'full'):
            actions = []
            expected_results = []
            name = data["name"]
            summary = data["summary"]
            preconditions = data["preconditions"]
            importance = data["importance"]
            execution_type = data["execution_type"]
            author = data["author_id"]
            # print(json.dumps(data, indent=4))
            for i in range(len(data["steps"])):
                actions.append(data["steps"][i]["actions"])
                expected_results.append(data["steps"][i]["expected_results"])
            datas.append((name, preconditions, '；'.join(actions), '；'.join(expected_results),
                          self.format_execution_type(execution_type), self.format_auth(author),
                          self.format_importance(importance),
                          summary))
        self.save_suits(os.path.join('testCase', 'download_template.xls'), datas, father_id)
