#!/usr/bin/python
# -*- coding: utf-8 -*-
##################################################
# __python__: 3.9.x
# __Author__: AJF
# __Purpose__: Manage projects through Jira data visualization
##################################################

import os, sys, time, subprocess, xlwt
import calendar
import requests
import logging
import argparse
import numpy as np
from jira import JIRA
from datetime import datetime
from time import sleep
from collections import defaultdict, Counter
from contextlib import contextmanager
from rich.progress import track
from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel
# import pysnooper
requests.packages.urllib3.disable_warnings()

parser = argparse.ArgumentParser(description='********** Jira Data Visualization by Project **********', prog=sys.argv[0])
parser.add_argument('-v', '--version', action='version', version='V1.0.1')
parser.add_argument('--project-id', type=str, nargs='+', action='store', help='(可选参数)通过project_id来搜索Jira数据, ex: X32A0-T972')
parser.add_argument('--status',  type=str, nargs='+', action='store', help='(可选参数)通过status来搜索Jira数据, ex: OPEN, Resolved' )
parser.add_argument('--reporter', type=str, nargs='+', help='(可选参数)通过reporter来搜索Jira数据, ex: san.zhang')
parser.add_argument('--component', type=str, nargs='+', help='(可选参数)通过component来搜索Jira数据, ex: HDMI, Dolby Vsion')
parser.add_argument('--resolution', type=str, nargs='+', help='(可选参数)通过resolution来搜索Jira数据, ex: Resolved, Won\'t fix')
parser.add_argument('--priority', type=str, nargs='+', help='(可选参数)通过priority来搜索Jira数据, ex: P0, P1')
parser.add_argument('--severity', type=str, nargs='+', help='(可选参数)通过severity来搜索Jira数据, ex: Normal, Major')
parser.add_argument('--label', type=str, nargs='+', help='(可选参数)通过label来搜索Jira数据, ex: pmlist-zql-20230103')
parser.add_argument('--month', type=str, nargs='+', help='(可选参数)通过date月份日期来搜索Jira数据, ex: 2022-11')
parser.add_argument('--duration', type=str, nargs='+', help='(可选参数)通过date月份日期范围来搜索Jira数据, ex: 2022-11 2023-02')
parser.add_argument('--date-range', type=str, nargs='+', help='(可选参数)通过date月份日期来筛选目标时间范围内的Jira数据内容, ex: 2022-11 2023-02')
parser.add_argument('--testcase-id', type=str, nargs='+', help='(可选参数)通过testcase_id来搜索Jira数据, ex: TV-F3081F0001')
parser.add_argument('--testcase-check', action='store_true', help='(可选参数)加上该参数会进行testcase_id检测, 默认: False')
parser.add_argument('--active-check', action='store_true', help='(可选参数)搜索统计Jira数据中所有人员comment活跃度占比, 默认: False')
parser.add_argument('--label-check', type=str, nargs='+', help='(可选参数)搜索统计Jira数据中添加labels人员的占比, ex: Common_From_Project, SH-Support-2023')
parser.add_argument('--verify-check', action='store_true', help='(可选参数)搜索统计Jira数据中verified人员的占比')
parser.add_argument('--epic-check', action='store_true', help='(可选参数)搜索统计Jira所属的epic信息')
parser.add_argument('--di-count', action='store_true', help='(可选参数)搜索统计Severity并计算整体DI值')
parser.add_argument('--raw-command', nargs=1, metavar='JQL', help='(可选参数)通过JQL语句来搜索Jira数据, ex: "Project ID" = AM30A2-T950D4 AND status in (OPEN, Reopened)"')
parser.add_argument('-e', '--expand', action='store_true', help='(可选参数)搜索范围加入changelog的历史操作数据, 默认: False')
parser.add_argument('-o', '--output', action='store_true', help='(可选参数)保存数据到本地excel表格, 表格默认命名: Output_Result_YYYYMMDD_HHMMSS.xlsx, 默认: False')
parser.add_argument('--verbose', action='store_true', help='(可选参数)加上该参数会打印更多调试信息, 默认: False')
args = parser.parse_args()

console = Console()
################################################################################################
PROJECT_ID = args.project_id             # [01]获取project_id       -> string (ex: X32A0-T972)
STATUS = args.status                     # [02]获取status           -> string (ex: OPEN, Resolved)
REPORTER = args.reporter                 # [03]获取reporter         -> string (ex: San.Zhang Si.Li)
COMPONENT = args.component               # [04]获取component        -> string (ex: HDMI, Dolby Vision)
RESOLUTION = args.resolution             # [05]获取resolution       -> string (ex: resolved, done)
PRIORITY = args.priority                 # [06]获取priority         -> string (ex: P0, P1)
SEVERITY = args.severity                 # [07]获取severity         -> string (ex: Blocker, Critical)
LABEL = args.label                       # [08]获取label            -> string (ex: pmlist-zql-20230103 must-fix-0113)
MONTH = args.month                       # [09]获取year-month       -> string (ex: 2022-11) 
DURATION = args.duration                 # [10]获取时间范围区间       -> string (ex: 2022-11, 2023-02)  供JQL语句中限定时间范围使用
DATERANGE = args.date_range              # [11]获取时间范围区间       -> string (ex: 2022-11, 2023-02)  供筛查目标时间范围内数据使用 
TESTCASE_ID = args.testcase_id           # [12]获取testcase ID      -> string (ex: TV-A3011-F0001)
RAW_COMMAND = args.raw_command           # [13]获取raw JQL          -> string (ex: "Project ID" = AM30A2-T950D4 AND status in (OPEN, Reopened)")
TESTCASE_CHECK = args.testcase_check     # [14]获取testcase_check   -> bool   (ex: True | False)
VERIFY_CHECK = args.verify_check         # [15]获取verified信息      -> bool   (ex: True | False)
EPIC_CHECK = args.epic_check             # [16]获取epic ID          -> bool   (ex: True | False)
DI_COUNT = args.di_count                 # [17]获取Severity计算di值  -> bool   (ex: True | False)
LABEL_CHECK = args.label_check           # [18]获取label信息         -> string (ex: Common_From_Project, SH-Support-2023)
ACTIVE_CHECK = args.active_check         # [19]获取active_check     -> bool   (ex: True | False)
VERBOSE_FLAG = args.verbose              # [20]获取verbose_flag     -> bool   (ex: True | False)
EXPAND_FLAG = args.expand                # [21]获取expand_flag      -> bool   (ex: True | False) 
OUTPUT_FLAG = args.output                # [22]获取output_flag      -> bool   (ex: True | False)
################################################################################################

# @pysnooper.snoop()
def args_init() -> dict:
    """parser args value and print from stdin"""

    global ARGS_DICT
    ARGS_DICT = {   'Project ID': PROJECT_ID,          #[1]
                    'Status': STATUS,                  #[2]
                    'Reporter': REPORTER,              #[3]
                    'Component': COMPONENT,            #[4]
                    'Resolution': RESOLUTION,          #[5]
                    'Priority': PRIORITY,              #[6]
                    'Severity': SEVERITY,              #[7]
                    'Label': LABEL,                    #[8]
                    'Month': MONTH,                    #[9]
                    'Duration': DURATION,              #[10]
                    'Date Range': DATERANGE,           #[11]  
                    'TestCase ID': TESTCASE_ID,        #[12]
                    'Raw Command': RAW_COMMAND,        #[13]
                    'TestCase Check': TESTCASE_CHECK,  #[14]
                    'Verify Check': VERIFY_CHECK,      #[15]
                    'Epic Check': EPIC_CHECK,          #[16]
                    'DI Count': DI_COUNT,              #[17]
                    'Label Check': LABEL_CHECK,        #[18]
                    'Active Check': ACTIVE_CHECK,      #[19]
                    'Verbose': VERBOSE_FLAG,           #[20]
                    'Expand': EXPAND_FLAG,             #[21]
                    'Output': OUTPUT_FLAG,             #[22]
                }

    with _wrapper(50):
        for index, (arg, value) in enumerate(ARGS_DICT.items(), 1):
            logging.debug('[{}] {} = {}'.format(str(index).zfill(2), arg.ljust(14, ' '), value))
    return ARGS_DICT

@contextmanager
def _wrapper(num):
    """warp horizontal line with NUM"""
    logging.debug('=' * num)
    yield
    logging.debug('=' * num)

def creat_local_file(filename):
    excel_file = '{}_{}.xlsx'.format(filename, curr_time)
    return excel_file

def logging_init() -> None:
    logger = logging.getLogger()
    if VERBOSE_FLAG:
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)
    global curr_time, log_file
    curr_time = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
    # log_folder = 'Script_logs'
    # if not os.path.exists(log_folder):
    #     os.mkdir(log_folder)
    # log_folder = os.path.join(os.getcwd(), log_folder)
    # log_folder_path = ''.join(log_folder)
    # log_folder_path = os.path.join(log_folder_path, curr_time)
    # log_name = (log_folder_path, '_Script.log')
    # log_file = ''.join(log_name)
    # file_handler = logging.FileHandler(log_file, mode='w')
    # # file_handler.setLevel(level=logging.DEBUG)
    # formatter1 = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(funcName)s - %(levelname)s: %(message)s")
    # file_handler.setFormatter(formatter1)
    # logger.addHandler(file_handler)
    logger.addHandler(RichHandler(rich_tracebacks=True, tracebacks_show_locals=True))

# @pysnooper.snoop()
class AmlJiraSystem(object):
    """Capture data from Amlogic Jira system"""
    def __init__(self, username:str, password:str):
        """Init Amlogic Jira server website"""
        super().__init__()
        self.username = username
        self.password = password
        self.jira_server = 'https://jira.amlogic.com'
        self.di_rules = { 'Blocker': 10, 'Critical': 3, 'Major':1, 'Normal': 0.1 }
        self.args_list = {}

    def login_jira(self) -> str:
        """Use username and password to login and return filter string"""
        try:
            self.myjira = JIRA(self.jira_server, basic_auth=(self.username, self.password), options={'Verify': False, 'delay-reload': 5, 'headers': {'Cache-Control': 'max-age','Content-Type': 'application/json'}}, timeout=300, max_retries=10)
            return self.myjira
        except Exception as err:
            logging.error(err)
            logging.error(f'Can not access {self.jira_server}, pls check!')
    
    def init_actual_args(self, **kwargs:'dict') -> 'dict':
        """Return ARGS_DICT form external args"""
        kwargs = ARGS_DICT
        live_args = {}
        # 过滤筛选实际的外部参数, 组装成新的Dict, 并返回
        if kwargs is not None:
            for kw_index, kw_value in kwargs.items():
                if not kw_value:
                    continue
                else:
                    live_args[kw_index] = kw_value
            logging.debug(f'Actual available args: {live_args}')
            return live_args

    def formatting_month(self, obj:'list') -> 'list':
        """formatting YYYY-MM or YYYY-MM YYYY-NN to List(YYYY, MM, DD) or List(YYYY, MM, YYYY, MM, DD)"""

        # logging.debug(f'=> Input month or duration: {obj}, type: {type(obj)}, len: {len(obj)}')        
        # ex: 2022-12
        if len(obj) in (6, 7):
            try:
                data = obj.split('-')    # ex: 2022-12
                # 返回(year,month)的起始日期与结束日期, ex:(1, 31)
                _, _end_day = calendar.monthrange(int(data[0]), int(data[1]))
                logging.debug(f'Actual month args: {obj}, year: {data[0]}, month: {data[1]}, day: {_end_day}')
                # 返回list, 长度=3
                return int(data[0]), int(data[1]), int(_end_day)
            except Exception as err:
                logging.error('Illegal data: {}, Error: {}'.format(obj, err))
                sys.exit(-1)

       #ex: 2022-12 2023-02  
        elif len(obj) == 2:
            try:
                data_p1, data_p2 = obj   #ex: 2022-12 2023-02 
                if '-' in data_p1:
                    _data_p1, _data_p2 = data_p1.strip(',').split('-')    # ex: _data_p1 = 2022, _data_p2 = 12
                    logging.debug(f'Actual part1 of duration: {data_p1}, year: {_data_p1}, month: {_data_p2}, day: {None}')   # 起始时间为当月的1号，所以day=None
                if '-' in data_p2:
                    _data_p3, _data_p4 = data_p2.strip(',').split('-')    # ex: _data_p3 = 2023, _data_p4 = 02 ,_end_day2 depend on month
                    # 返回(year,month)的起始日期与结束日期, ex:(1, 31)
                    _, _end_day2 = calendar.monthrange(int(_data_p3), int(_data_p4)) 
                    logging.debug(f'Actual part2 of duration: {data_p2}, year: {_data_p3}, month: {_data_p4}, day: {_end_day2}')
                # 返回list, 长度=5
                return int(_data_p1), int(_data_p2), int(_data_p3), int(_data_p4), int(_end_day2)  # [ 2022, 12, 2023, 02, 28 ]
            except Exception as err:
                logging.error('Illegal data: {}, Error: {}'.format(obj, err))
                sys.exit(-1)
        else:
            logging.error('Illegal data: {}, pls input data as following: 2022-12 2023-02'.format(obj))
            sys.exit(-1)
    
    def formatting_date_range(self, obj:'list') -> 'list':
        """formatting YYYY-MM YYYY-NN to List(YYYY, MM, YYYY, MM, DD)"""

        # logging.debug(f'=> Input date range: {obj}, type: {type(obj)}, len: {len(obj)}')        

        #ex: 2022-12 2023-02  
        if len(obj) == 2:
            try:
                data_p1, data_p2 = obj   #ex: 2022-12 2023-02 
                if '-' in data_p1:
                    _data_p1, _data_p2 = data_p1.strip(',').split('-')    # ex: _data_p1 = 2022, _data_p2 = 12
                    # logging.debug(f'Actual part1 of duration: {data_p1}, year: {_data_p1}, month: {_data_p2}, day: {None}')   # 起始时间为当月的1号，所以day=None
                if '-' in data_p2:
                    _data_p3, _data_p4 = data_p2.strip(',').split('-')    # ex: _data_p3 = 2023, _data_p4 = 02 ,_end_day2 depend on month
                    # 返回(year,month)的起始日期与结束日期, ex:(1, 31)
                    _, _end_day2 = calendar.monthrange(int(_data_p3), int(_data_p4)) 
                    # logging.debug(f'Actual part2 of duration: {data_p2}, year: {_data_p3}, month: {_data_p4}, day: {_end_day2}')
                # 返回list, 长度=5
                return int(_data_p1), int(_data_p2), int(_data_p3), int(_data_p4), int(_end_day2)  # [ 2022, 12, 2023, 02, 28 ]
            except Exception as err:
                logging.error('Illegal data: {}, Error: {}'.format(obj, err))
                sys.exit(-1)
        else:
            logging.error('Illegal data: {}, pls input data as following: 2022-12 2023-02'.format(obj))
            sys.exit(-1)

    def packaging_filter_from(self, kwargs):
        """Return a filter string from external args"""

        '''
        kwargs(dict): from input parameters
        return(str): JQL
        '''
        
        # 'ORDER BY created DESC'
        self.filter = ''
        sort_rule = 'ORDER BY created DESC'
        
        # [1]Project ID: 支持多参数
        if kwargs['Project ID']:       # Project ID为required参数，默认存在
            kw_project_id = ','.join(kwargs['Project ID'])
            kw_project_id = '"project id" in ({})'.format(kw_project_id)
        else:
            kw_project_id = None
        
        # [2]Status: 支持多参数
        if kwargs['Status']:
            kw_status = ','.join(kwargs['Status'])
            kw_status = 'status in ({})'.format(kw_status)
        else:
            kw_status = None

        # [3]Reporter: 支持多参数
        if kwargs['Reporter']:
            kw_reporter = ','.join(kwargs['Reporter'])
            kw_reporter = 'reporter in ({})'.format(kw_reporter)
        else:
            kw_reporter = None

        # [4]Component: 支持多参数
        if kwargs['Component']:
            kw_component = ','.join(kwargs['Component'])
            kw_component = 'component in ({})'.format(kw_component)
        else:
            kw_component = None

        # [5]Resolution: 支持多参数
        if kwargs['Resolution']:
            kw_resolution = ','.join(kwargs['Resolution'])
            kw_resolution = 'resolution in ({})'.format(kw_resolution)
        else:
            kw_resolution = None

        # [6]Priority: 支持多参数
        priority_map = {'P0': 'Highest', 'P1': 'High', 'P2': 'Medium', 'P3': 'Low', 'P4': 'Lowest'}
        priority_list = list()
        if kwargs['Priority']:
            for p in kwargs['Priority']:
                if p in priority_map.keys():
                    priority_list.append(priority_map[p])
                    logging.debug(priority_list)
            kw_priority = ','.join(priority_list)
            kw_priority = 'priority in ({})'.format(kw_priority)
        else:
            kw_priority = None

        # [7]Severity: 支持多参数
        if kwargs['Severity']:
            kw_severity = ','.join(kwargs['Severity'])
            kw_severity = 'severity in ({})'.format(kw_severity)
        else:
            kw_severity = None

        # [8]Label: 支持多参数
        if kwargs['Label']:
            kw_label = ','.join(kwargs['Label'])
            kw_label = 'labels in ({})'.format(kw_label)
        else:
            kw_label = None
        
        # [9]Month
        if kwargs['Month']:
            try:
                _year, _month, _day = self.formatting_month(kwargs['Month'][0])
                kw_month = 'created >= {}-{}-01 AND created <= {}-{}-{}'.format(_year, _month, _year, _month, _day)
            except ValueError:
                logging.error(f'Illegal data!')
        else:
            kw_month = None

        # [10]Duration 
        if kwargs['Duration']:
            self._year_s, self._month_s, self._year_e, self._month_e, self._day_e = self.formatting_month(kwargs['Duration'])    # [ 2022, 12, 2023, 02, 28 ]
            kw_duration = 'created >= {}-{}-01 AND created <= {}-{}-{}'.format(self._year_s, self._month_s, self._year_e, self._month_e, self._day_e)
        else:
            kw_duration = None

        for item in (kw_project_id, kw_status, kw_reporter,       # [1-3]
                     kw_component, kw_resolution, kw_priority,    # [4-6]
                     kw_severity, kw_label, kw_month,             # [7-9]
                     kw_duration):                                # [10]
            self.filter += f'{item} AND '
        self.filter += sort_rule
        self.filter = self.filter.replace('AND None ', '')         # 去掉无效的field=None的字符
        self.filter = self.filter.replace('None AND ', '')         # 去掉无效的field=None的字符
        self.filter = self.filter.replace('AND ORDER', 'ORDER')    # 去掉ORDER前的AND字符
        return self.filter

    @contextmanager
    def _wrapper(self, num):
        logging.info('-' * num)
        yield
        logging.info('-' * num)

    def _generator_warp(self, input:'list') -> str('generator'):
        for inp in input:
            yield inp
    
    def str2Time(self, string):
        '''将Jira日期时间字符串格式化为datetime类型: %Y-%m-%d %H:%M:%S'''
        if string and 'T' in string:
            timePart1 = string.split('T')
            timePart2 = timePart1[1].split('.')
            transform = timePart1[0] + ' ' + timePart2[0]
            return datetime.strptime(transform, '%Y-%m-%d %H:%M:%S')   # '2022-12-09'
        return string

    def nameUpper(self, name):
        '''将姓名首字母大写,及Software Version去掉'-'后面的部分'''
        if '-' in name:
            return name.split('-')[0]     # Android P-9.0 -> Android P
        if '.' in name:
            front_part = name.split('.')[0].capitalize()
            end_part = name.split('.')[-1].capitalize()
            return front_part + '.' + end_part      # san.zhang -> San.Zhang
        return name

    def format_daterange(self, object):
        '''格式化date range参数,返回起始日期区间
        :param: 
            object: ex: [ '2022-12', '2023-02' ]
        '''
        try:
            _year_s, _month_s, _year_e, _month_e, _day_e = self.formatting_date_range(object)    # [ 2022, 12, 2023, 02, 28 ]
            _duration_s = datetime.strptime('{}-{}-1'.format(_year_s, _month_s), '%Y-%m-%d')
            _duration_e = datetime.strptime('{}-{}-{}'.format(_year_e, _month_e, _day_e), '%Y-%m-%d')
        except Exception as err:
            logging.error('Date Range Error: {}'.format(err))    
        return _duration_s, _duration_e 

    def get_diff(self, a, b):
        '''两个labels list找出相差项'''
        return list(set(b).difference(set(a)))

    def calculate_severity(self, object):
        total_di = 0
        for k, v in object.items():
            if k == 'Blocker':
                total_di += v * 10
            if k == 'Critical':
                total_di += v * 3
            if k == 'Major':
                total_di += v * 1
            if k == 'Normal':
                total_di += v * 0.1
        return total_di

    def get_customize_fields(self):
        '''根据参数决定search_issues中fields的具体内容'''
        object = ['priority']
        custom_fields = object
        if any([ACTIVE_CHECK, VERIFY_CHECK]):
            custom_fields = object.append('comment') 
        if LABEL_CHECK:
            custom_fields = object.append('labels')
        if DI_COUNT:
            custom_fields = object.append('customfield_10109')
        if TESTCASE_CHECK:
            custom_fields = object.append('customfield_11604')
        if EPIC_CHECK:
            custom_fields = object.append('customfield_10102') 
        if OUTPUT_FLAG:
            custom_fields = [ 'customfield_10107', 'customfield_10407', 'issue_id', 'component', 'status', 'priority', 'assignee', 'customfield_10700', 'created', 'updated', 'finish_date', 'cost' ]
        return custom_fields

    # @pysnooper.snoop()
    def process_search(self, jql) -> list:
        """Return a generator object from an advance filter"""
        
        logging.info(f'=> JQL: {jql}')

        start_at = 0                            # search起始值
        max_results = 1000                      # search最大值
        
        self.segment = 0                        # JQL分段计数
        self.jql_total = 0                      # JQL总数
        self.fields = {}                        # JQL中fields总数
        self.commentor_all_count = []           # JQL中所有comments人员        ex: [ 'Zanbo.Huang', 'Maoguo.Xie' ]
        self.verified_all_count = Counter()     # JQL中所有verified人员        ex: { 'Zanbo.Huang': 3, 'Maoguo.Xie': 2 }
        self.verified_QA_count = Counter()      # JQL中所有QA verified人员     ex: { 'Zanbo.Huang': 3, 'Maoguo.Xie': 2 }
        self.support_label_count = Counter()    # JQL中所有打上特定Label的人员   ex: { 'Zanbo.Huang': 3, 'Maoguo.Xie': 2 }
        self.severity_count = Counter()
        self.testcase_count = Counter()
        self.addcase_count = Counter()
        self.othercase_count = Counter()
        self.nonecase_count = Counter()

        while True:
            jql_results = self.myjira.search_issues(jql_str=jql, 
                                                    startAt=start_at,
                                                    maxResults=max_results,           # -1等价于maxResults=5000 
                                                    json_result=True,                 # 返回的数据格式为json
                                                                                                       # 'expand': 'operations,versionedRepresentations,editmeta,changelog,renderedFields'
                                                    expand='changelog' if EXPAND_FLAG else None,       # 如参数带上"-e", 则返回的json数据中会包含['changelog']该部分的数据
                                                    fields=self.get_customize_fields(),
                                                    )
                                                    # fields=[
                                                            # 'summary',
                                                            # 'issuetype',
                                                            # 'components', 
                                                            # 'customfield_10407',      # project id
                                                            # 'customfield_10107',      # product, ex: TV reference
                                                            # 'status', 
                                                            # 'priority', 
                                                            # 'customfield_10109',      # severity
                                                            # 'customfield_10300',      # sw version
                                                            # 'customfield_10108',      # hw version
                                                            # 'customfield_11703',      # compare status
                                                            # 'customfield_11705',      # common issue, ex: confirmed yes or no
                                                            # 'resolution',
                                                            # 'fixVersions',
                                                            # 'assignee', 
                                                            # 'customfield_10700',      # rd manager
                                                            # 'reporter', 
                                                            # 'description', 
                                                            # 'attachments', 
                                                            # 'comment', 
                                                            # 'duedate',
                                                            # 'created',
                                                            # 'updated',
                                                            # 'labels',
                                                            # 'issuelinks',
                                                            # 'customfield_12200',      # report channel and role, ex: self-test and QA
                                                            # 'customfield_11604'       # test case
                                                            # ])
            if not jql_results:
                break
            
            # 所需的目标Json数据内容
            jql_issues = jql_results.get('issues')
            
            if jql_issues:

                _fields, _commentor_all_count, _verified_all_count, _verified_QA_count, _support_label_count, _severity_count, _testcase_count, _addcase_count, _othercase_count, _nonecase_count = self.get_fields_data(jql_issues)
                self.segment += 1
                logging.info('>>> [{}] - Total issues: {}'.format(self.segment, len(jql_issues)))
                
                # Comments活跃度
                if ACTIVE_CHECK:
                    logging.info('>>> [{}] - Total Comments histories: {}'.format(self.segment, len(_commentor_all_count)))
                
                # Verified记录
                if VERIFY_CHECK:
                    logging.info('>>> [{}] - Total Verified histories: {}'.format(self.segment, sum([x for x in dict(self.verified_all_count).values()])))
                    logging.info('>>> [{}] - Total QA Verified Count: {}'.format(self.segment, sum([x for x in _verified_QA_count.values()])))
                
                # Label添加记录
                if LABEL_CHECK:
                    logging.info('>>> [{}] - Total Label Count: {}'.format(self.segment, sum([x for x in _support_label_count.values()])))

                if DI_COUNT:
                    logging.info('>>> [{}] - Total DI Count: {}'.format(self.segment, sum([x for x in _severity_count.values()])))
                    logging.info('>>> [{}] - Total DI Count: {}'.format(self.segment, dict(_severity_count)))
                    logging.info('>>> [{}] - Total DI Count: {}'.format(self.segment, self.calculate_severity(dict(_severity_count))))
            
            else:
                # logging.warning('No issues were found')
                break
            
            # 搜索起始值按1000往上递增
            start_at += max_results
            
            # 整体数据合并(dict)
            self.fields.update(_fields)
            # comments分段相加(list)
            self.commentor_all_count += _commentor_all_count
            # verified人员分段相加(dict)
            self.verified_all_count += Counter(_verified_all_count)
            # QA verified人员分段相加(dict)
            self.verified_QA_count += Counter(_verified_QA_count)
            # SH-Support-2023 Label统计
            self.support_label_count += Counter(_support_label_count)
            # Severity分段相加(dict)
            self.severity_count += Counter(_severity_count)
            # 有效TestCase ID分段相加
            self.testcase_count += Counter(_testcase_count)
            # Add TestCase分段相加
            self.addcase_count += Counter(_addcase_count)
            # None TestCase分段相加
            self.nonecase_count += Counter(_nonecase_count)
            # Other TestCase分段相加
            self.othercase_count += Counter(_othercase_count)
            # JQL的总和计数(int)
            self.jql_total += len(jql_issues)

            with self._wrapper(50):
                logging.info('[01]------------Total Issues: {}'.format(self.jql_total))
                logging.info('[02]Total Comments Histories: {}'.format(len(self.commentor_all_count)))
                logging.info('[03]Total Verified Histories: {}'.format(sum([x for x in dict(self.verified_all_count).values()])))
                logging.info('[04]--------------Date Range: {}'.format(DATERANGE))
                logging.info('[05]-Total QA Verified Count: {}'.format(sum([x for x in dict(self.verified_QA_count).values()])))
                logging.info('[06]------------Target Label: {}'.format(LABEL_CHECK))
                logging.info('[07]-------Total Label Count: {}'.format(sum([x for x in dict(self.support_label_count).values()])))
                logging.info('[08]-All Labels Distribution: {}'.format(self.support_label_count))
                logging.info('[09]----------Total DI Value: {}'.format(self.calculate_severity(dict(self.severity_count))))
                logging.info('[10]----Valid TestCase Count: {}, Total: {}, Ratio: {:.1%}'.format(self.testcase_count, sum(self.testcase_count.values()), sum(self.testcase_count.values()) / self.jql_total))
                logging.info('[11]--AddCase TestCase Count: {}, Total: {}, Ratio: {:.1%}'.format(self.addcase_count, sum(self.addcase_count.values()), sum(self.addcase_count.values()) / self.jql_total))
                logging.info('[12]-NoneCase TestCase Count: {}, Total: {}, Ratio: {:.1%}'.format(self.nonecase_count, sum(self.nonecase_count.values()), sum(self.nonecase_count.values()) / self.jql_total))
                logging.info('[13]OtherCase TestCase Count: {}, Total: {}, Ratio: {:.1%}'.format(self.othercase_count, sum(self.othercase_count.values()), sum(self.othercase_count.values()) / self.jql_total))
                # logging.info('[10]----Valid TestCase Count: {}, Ratio: {:.1%}'.format(self.testcase_count, (self.testcase_count / self.jql_total)))
                # logging.info('[11]--AddCase TestCase Count: {}, Ratio: {:.1%}'.format(self.addcase_count, (self.addcase_count / self.jql_total)))

    # @pysnooper.snoop()
    def get_fields_data(self, jql_issues):
        """Return json object from jira.fields.TARGET"""
        field = {}                              # 每个issue的数据集合, 每个循环会清零
        issuelink = []                          # 每个issue中issuelink的数据集合, 每个循环会清零
        comments = defaultdict(list)            # 每个issue中comments的数据集合, 每个循环会清零
        histories = defaultdict(list)           # 每个issue中changelog中histories的数据集合, 每个循环会清零
        commentor_all_count = []                # 所有issue中的comments作者, 不会清零
        verified_all_count = defaultdict(int)   # 所有issue中的verified操作人员, 不会清零
        verified_QA_count = defaultdict(int)    # 所有issue中目标时间范围内TV FAE-QA进行verified的操作人员, 不会清零
        support_label_count = defaultdict(int)  # 所有issue中过滤目标label, ex: SH-Support-2023, Common
        severity_count = defaultdict(int)
        fields = defaultdict(list)              # 所有issue的数据集合, 不会清零

        testcase_count = defaultdict(int)
        addcase_count = defaultdict(int)
        othercase_count = defaultdict(int)
        nonecase_count = defaultdict(int)

        #* Sample: https://jira.amlogic.com/rest/api/2/issue/TV-64205?expand=changelog
        for i, d in track(enumerate(jql_issues, 1), description='[green]Processing[/green]', total=len(jql_issues)):
            field['issue_id'] = d['key']                                                                                                             # 01 issue id -> str
            field['priority'] = d['fields']['priority']['name']                                                                                      # 08 issue priority -> str
            
            if ACTIVE_CHECK:
                field['comments'] = d['fields']['comment']['comments']                                                                                  # 22 comments -> list
            if LABEL_CHECK:
                field['labels'] = d['fields']['labels']                                                                                                 # 25 labels -> list
            if EPIC_CHECK:
                field['epic'] = d['fields'].get('customfield_10102')
            if DI_COUNT:
                field['severity'] = d['fields']['customfield_10109']['value'] if d['fields'].get('customfield_10109') else None                       # 09 severity -> str
            if OUTPUT_FLAG:
                field['product'] = d['fields']['customfield_10107'][0]['value'] if d['fields'].get('customfield_10107') else None                     # 06 product -> str
                field['project_id'] = d['fields']['customfield_10407'][0]['value'] if d['fields'].get('customfield_10407') else None                  # 05 project id -> str
                field['component'] = d['fields']['components'][0]['name'] if d['fields'].get('components') else None                                  # 04 component -> str
                field['status'] = d['fields']['status']['name']                                                                                         # 07 issue status -> str
                field['assignee'] = self.nameUpper(d['fields']['assignee']['name']) if d['fields'].get('assignee') else None                          # 20 assignee -> str
                field['rd_manager'] = self.nameUpper(d['fields']['customfield_10700']['name']) if d['fields'].get('customfield_10700') else None      # 21 rd manager -> str
                field['created'] = self.str2Time(d['fields']['created'])                                                                              # 24 creat time -> str
                field['updated'] = self.str2Time(d['fields']['updated'])
            if TESTCASE_CHECK:
                field['testcase'] = d['fields']['customfield_11604'] if d['fields'].get('customfield_11604') else None                                # 28 test case -> str
                if field['testcase'] and 'TV-' in field['testcase']:
                    testcase_count['{}'.format(field['testcase'])] += 1
                    # logging.debug('->>> {}'.format(testcase_count))
                elif field['testcase'] and 'case' in field['testcase']:
                    addcase_count['{}'.format(field['testcase'])] += 1
                elif field['testcase'] and 'Case' in field['testcase']:
                    addcase_count['{}'.format(field['testcase'])] += 1
                    # logging.debug('-<<< {}'.format(addcase_count))
                elif not field['testcase']:
                    nonecase_count['{}'.format(field['testcase'])] += 1
                else:
                    othercase_count['{}'.format(field['testcase'])] += 1

            logging.debug('================== {} =================='.format(i))
            logging.debug('[01] ------------Issue ID: {}'.format(field['issue_id']))
            if EPIC_CHECK:
                logging.debug('[38] -------------Epic ID: {}'.format(field['epic']))
            if TESTCASE_CHECK:
                logging.debug('[39] -----------Test Case: {}'.format(field['testcase']))

            # field['summary'] = d['fields']['summary']                                                                                             # 02 summary -> str
            # field['issue_type'] = d['fields']['issuetype']['name']                                                                                # 03 issue type -> str
            # field['severity'] = d['fields']['customfield_10109']['value'] if d['fields'].get('customfield_10109') else None                       # 09 severity -> str
            # field['sw_version'] = self.nameUpper(d['fields']['customfield_10300'][0]['value']) if d['fields'].get('customfield_10300') else None  # 10 software version -> str
            # field['hw_version'] = d['fields']['customfield_10108'][0]['value'] if d['fields'].get('customfield_10108') else None                  # 11 hardware version -> str
            # field['description'] = [ x.strip() for x in d['fields']['description'].replace('\r','').split('\n') if x ]                            # 12 description -> list
            # field['compare_status'] = d['fields']['customfield_11703'][0]['value'] if d['fields'].get('customfield_11703') else None              # 13 compare status -> str
            # field['common_issue'] = d['fields']['customfield_11705']['value'] if d['fields']['customfield_11705'] else None                       # 14 common issue -> str
            # field['resolution'] = d['fields']['resolution']['name'] if d['fields'].get('resolution') else None                                    # 15 resoltion -> str
            # field['fix_version'] = d['fields']['fixVersions'][0]['name'] if d['fields'].get('fixVersions') else None                              # 16 fix version -> str
            # field['reporter'] = self.nameUpper(d['fields']['reporter']['name'])                                                                   # 17 reporter -> str
            # field['report_channel'] = d['fields']['customfield_12200']['value'] if d['fields'].get('customfield_12200') else None                 # 18 report channel -> str 
            # field['report_role'] = d['fields']['customfield_12200']['child']['value'] if d['fields'].get('customfield_12200') and d['fields']['customfield_12200'].get('child') else None  # 19 report role -> str
            # field['comments_count'] = len(field['comments']) 
            # field['duedate'] = self.str2Time(d['fields']['duedate']) if self.str2Time(d['fields'].get('duedate')) else None                       # 23 due data -> str
            # field['description_line'] = len(field['description'])                                                                                 # 26 desciption_flag -> int
            # field['attachment_count'] = len(d['fields']['attachment']) if d['fields']['attachment'] else '0'                                      # 27 attachment_flag -> int
            # field['testcase'] = d['fields']['customfield_11604'] if d['fields'].get('customfield_11604') else None                                # 28 test case -> str

            # field['issuelinks'] = len(d['fields']['issuelinks']) if d['fields'].get('issuelinks') else None                                       # 29 issuelinks -> int
            # if d['fields'].get('issuelinks'):
            #     issuelinks = d['fields']['issuelinks']
            #     for x in issuelinks:
            #         if x.get('type') and x.get('outwardIssue'):
            #             type = x['type']['outward']
            #             outward = x['outwardIssue']['key']
            #             issuelink.append('{}: {}'.format(type, outward))   # ['clones: TV-33435', 'clones: TV-33461'] 
            #         if x.get('type') and x.get('inwardIssue'):
            #             type = x['type']['outward']
            #             inward = x['inwardIssue']['key']
            #             issuelink.append('{}: {}'.format(type, inward))    # ['clones: TV-33435', 'clones: TV-33461'] 
            # field['issuelinks_group'] = issuelink if issuelink else None

            
            #* 对comments内容进行进一步处理
            if ACTIVE_CHECK:
                if field['comments']:
                    for c, x in enumerate(field['comments'], 1):
                        comments[c].append(str(self.str2Time(x['created'])))      # created时间
                        comments[c].append(self.nameUpper(x['author']['name']))   # author对象
                        comments[c].append(x['body'])                             # comment内容
                        commentor_all_count.append(self.nameUpper(x['author']['name']))
                field['_comments'] = comments if comments else None
                # field['last_comment'] = list(field['_comments'].values())[-1] if field.get('comments') else ['None','None','None']  # 最后一条Comment的 [ 时间, 作者, 内容 ]
            
            if DI_COUNT:
                if field['severity'] in self.di_rules.keys():
                    logging.debug('Severity: {}'.format(field['severity']))
                    severity_count['{}'.format(field['severity'])] += 1
                    logging.debug(dict(severity_count))

            #* 通过"-e, --expand"参数来控制changelog的内容是否加载
            if EXPAND_FLAG:
                # field['histories_count'] = d['changelog']['total']            # ex: 'total': 16
                field['histories'] = d['changelog']['histories']                # ex: list数据类型
                for h, x in enumerate(field['histories'], 1):                   # 遍历所有histories数据
                    histories[h].append(str(self.str2Time(x['created'])))       # created时间    'created': '2023-02-13T11:50:52.889+0800'
                    histories[h].append(self.nameUpper(x['author']['name']))    # author对象     'name': 'linguo.bu'
                    histories[h].append(x['items'][0]['field'])                 # 类型           'field': 'Link'
                    histories[h].append(x['items'][0]['fromString'])            # 原初始内容      'fromString': None
                    histories[h].append(x['items'][0]['toString'])              # 变更后的内容     'toString': 'TV-73461'

                    #* 判断verified操作, 并记录所有操作人员
                    if VERIFY_CHECK:
                        if x['items'][0]['field'] == 'status' and x['items'][0]['toString'] == 'Verified':
                            logging.debug('date: {}, verified author: {}'.format(self.str2Time(x['created']), self.nameUpper(x['author']['name'])))
                            verified_all_count['{}'.format(self.nameUpper(x['author']['name']))] += 1

                    if VERIFY_CHECK and DATERANGE:
                        #* 判断author对象是否为FAE QA, 并进行了Verified操作
                        if x['items'][0]['field'] == 'status' and x['items'][0]['toString'] == 'Verified':
                            if self.nameUpper(x['author']['name']) in tv_product_team:
                                _duration_s, _duration_e = self.format_daterange(DATERANGE)     # ex: 2022-12-01, 2023-02-28
                                if _duration_s <= self.str2Time(x['created']) <= _duration_e:
                                    logging.debug('date: {}, verified QA author: {}'.format(self.str2Time(x['created']), self.nameUpper(x['author']['name'])))
                                    verified_QA_count['{}'.format(self.nameUpper(x['author']['name']))] += 1

                    #* 判断labels, 如: Common_From_Project, SH-Support-2023 ...
                    if LABEL_CHECK and DATERANGE:
                        _duration_s, _duration_e = self.format_daterange(DATERANGE)     # ex: 2022-12-01, 2023-02-28
                        try:
                            label_from = x['items'][0]['fromString'].split(' ')
                            label_to = x['items'][0]['toString'].split(' ')
                        except AttributeError:
                            label_from, label_to = None, None
                        if x['items'][0]['field'] == 'labels' and LABEL_CHECK[0] in self.get_diff(label_from, label_to):
                            if self.nameUpper(x['author']['name']) in tv_product_team:
                                if _duration_s <= self.str2Time(x['created']) <= _duration_e:
                                    logging.debug('date: {}, add label author: {}'.format(self.str2Time(x['created']), self.nameUpper(x['author']['name'])))
                                    support_label_count['{}'.format(self.nameUpper(x['author']['name']))] += 1
                    
                    #* 判断finish date是否已设置, 已设置则获取相应的日期时间
                    if OUTPUT_FLAG:
                        if x['items'][0]['field'] == 'Finish date (WBSGantt)':
                            field['_finish_date'] = x['items'][0]['to']        # 2023-02-03
                            try:
                                field['finish_date'] = datetime.strptime(field['_finish_date'], '%Y-%m-%d')    # 2023-02-03 00:00:00
                                logging.debug('finish data: {}'.format(str(field['finish_date']).split(' ')[0]))
                            except Exception:
                                continue
                
                #! 保留所有changelog信息到单独issue的field中
                field['changelog'] = histories                                  # {1: ['2022-09-02 18:59:32', 'Jianfan.Ai', 'Link', None, 'This issue clones TV-58996'], }

            #* 如果该issue中已存在"Finish date (WBSGantt)"该参数时, 则开始计算相差时间
            field['now_date'] = datetime.strptime('{}-{}-{}'.format(datetime.now().year, datetime.now().month, datetime.now().day), '%Y-%m-%d')  # ex: now_date = 2023-02-05
            if field.get('finish_date'):
                field['cost'] = (field['now_date'] - field['finish_date']).days      # 计算(当前时间 - PMLIST计划解决时间)之间相差的天数
            else:
                field['finish_date'] = None
                field['cost'] = None

            # logging.debug('================== {} =================='.format(i))
            # logging.debug('[01] ------------Issue ID: {}'.format(field['issue_id']))
            # logging.debug('[02] -------------Summary: {}'.format(field['summary']))
            # logging.debug('[03] ---------Issue Type : {}'.format(field['issue_type']))
            # logging.debug('[04] -----------Component: {}'.format(field['component']))
            # logging.debug('[05] ----------Project ID: {}'.format(field['project_id']))
            # logging.debug('[06] -------------Product: {}'.format(field['product']))
            # logging.debug('[07] --------------Status: {}'.format(field['status']))
            # logging.debug('[08] ------------Priority: {}'.format(field['priority']))
            # logging.debug('[09] ------------Severity: {}'.format(field['severity']))
            # logging.debug('[10] ----------HW Version: {}'.format(field['hw_version']))
            # logging.debug('[11] ----------SW Release: {}'.format(field['sw_version']))
            # logging.debug('[12] ---------Description: {}'.format(field['description_line']))
            # logging.debug('[13] ------Compare Status: {}'.format(field['compare_status']))
            # logging.debug('[14] --------Common Issue: {}'.format(field['common_issue']))
            # logging.debug('[15] ----------Resolution: {}'.format(field['resolution']))
            # logging.debug('[16] ---------Fix Version: {}'.format(field['fix_version']))
            # logging.debug('[17] ------------Reporter: {}'.format(field['reporter']))
            # logging.debug('[18] ------Report Channel: {}'.format(field['report_channel']))
            # logging.debug('[19] ---------Report Role: {}'.format(field['report_role']))
            # logging.debug('[20] ------------Assignee: {}'.format(field['assignee']))
            # logging.debug('[21] ----------RD Manager: {}'.format(field['rd_manager']))
            # logging.debug('[22] ----Total Attachment: {}'.format(field['attachment_count']))
            # logging.debug('[23] --------Created Time: {}'.format(field['created']))
            # logging.debug('[24] --------Updated Time: {}'.format(field['updated']))
            # logging.debug('[25] -----Total Histories: {}'.format(field.get('histories_count')))
            # logging.debug('[26] -Changelog Histories: {}'.format('[ "Created", "Author", "Field", "fromString", "toString" ]' if field.get('histories_count') else None))
            # if field.get('changelog'):
            #     for k, v in field['changelog'].items():   # v[2]: 'Finish date (WBSGantt)', 'Start date (WBSGantt)'
            #         logging.debug('[27]-[{}] {}, {}, {}, {}, {}'.format(str(k).zfill(2), v[0], v[1], v[2], v[3], v[4]))
            # logging.debug('[28] ------Total Comments: {}'.format(field['comments_count']))
            # logging.debug('[29] ---Last Comment Date: {}'.format(field['last_comment'][0]))
            # logging.debug('[30] -Last Comment Author: {}'.format(field['last_comment'][-2]))
            # logging.debug('[31] Last Comment Content: {}'.format(field['last_comment'][-1].replace('\n','')))
            # logging.debug('[32] ---------------Label: {}'.format(field['labels']))
            # logging.debug('[33] ---------Issue Links: {}'.format(field['issuelinks']))
            # logging.debug('[34] -------Issuelinks ID: {}'.format(field['issuelinks_group']))
            # logging.debug('[35] ---------Verified QA: {}'.format(verified_QA_count))
            # logging.debug('[36] ------------Due Date: {}'.format(field['duedate']))
            # logging.debug('[37] >>>>>>>>>Finish date: {}'.format(field['finish_date']))
            # logging.debug('[38] >>>>>>>>>>>Cost days: {}'.format(field['cost']))
            # logging.debug('[39] -----------Test Case: {}'.format(field['testcase']))
            
            if OUTPUT_FLAG:
                if 'SWPL-' not in field['issue_id'] and field['priority'] not in ('P2', 'P3', 'P4') and field['cost']:
                    if field['cost'] > 0:
                        fields[i].append(field)
            else:
                fields[i].append(field)

            #! 清零操作(非常关键, 可以避免数据重复导致的错误)
            field = dict()
            issuelink = list()
            comments = defaultdict(list)
            histories = defaultdict(list) 

        # 问题总数
        # logging.info('>>> Total issues: {}'.format(len(jql_issues)))

        # Comments活跃度检测参数
        # if ACTIVE_CHECK:
        #     logging.info('>>> Total Comments histories: {}'.format(len(commentor_all_count)))
        #     logging.info('>>> Total Verified histories: {}'.format(len(verified_all_count)))
        
        # Changelog参数与目标日期范围参数
        # if EXPAND_FLAG and DATERANGE:
        #     logging.info('>>> Date range: {}, Total Verified QAs: {}'.format(DATERANGE, dict(verified_QA_count)))
        
        return fields, commentor_all_count, verified_all_count, verified_QA_count, support_label_count, severity_count, testcase_count, addcase_count, othercase_count, nonecase_count


    def show_chart(self, object, total, category, operate):
        '''show termgraph chart in stdout from list/dict(object)'''
        if object:
            if isinstance(object, list):
                _authors = sorted(Counter(object).items(),key=lambda x:x[1],reverse=True)   # 按Value数值大小重新排序,返回字典
            if isinstance(object, dict):
                _authors = sorted(object.items(),key=lambda x:x[1],reverse=True)            # 按Value数值大小重新排序,返回字典
            authors_array = np.array(_authors)
            _authors_array = [ int(x) for x in authors_array[:,1] ]                         # 仅取array数组的第二个元素组成新的list
            authors_total = sum(_authors_array)
            author_ratio =  [ '{:.1%}'.format(x / authors_total) for x in _authors_array ]  # 计算出每个author所占总数的百分比
            authors = np.insert(authors_array, 2, author_ratio, axis=1)                     # 往author_array数组中插入一列ratio数据
            authors = authors[:36]    # 仅截止Top20的数据用于显示, 其余数据Skipped

            #* 数据格式化
            pattern_author = ''
            for x in authors:      # authors为numpy.ndarray数据类型
                if str(x[0]) in tv_product_team:
                    pattern_author += (str('*{}({})'.format(x[0], x[2])) + ' ' + str(x[1]) + '\n')      # 格式化termgraph所识别的数据格式(TV FAE-QA人员名字前加上*号)
                else:
                    pattern_author += (str('{}({})'.format(x[0], x[2])) + ' ' + str(x[1]) + '\n')       # 格式化termgraph所识别的数据格式

            #* 数据可视化输出
            for string in (pattern_author, ):
                # termgraph_cmd = "echo '{}' | termgraph --format \{{\:.0f}} --title 'By Comments: {} people already commented in {}'".format(string.rstrip(), len(_authors), PROJECT_ID)    # 调用termgraph图形化数据 
                termgraph_cmd = "echo '{}' | termgraph --format \{{\:.0f}} --title 'By {}({}): {} people already {} in {}'".format(string.rstrip(), category, total, len(_authors), operate, PROJECT_ID)    # 调用termgraph图形化数据 
                try:
                    result = subprocess.Popen(termgraph_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True).communicate()[0].decode()
                    console.print(Panel.fit(result, width=1000))
                except Exception:
                    console.print('[red]Termgraph error![/red]\n')
        else:
            logging.warning('There\'s no data to work with Termgraph!')

class TableObject():
    def __init__(self):
        self.workbook = xlwt.Workbook(encoding='utf-8')  # 初始化实例
        self.worksheet = self.workbook.add_sheet('{}'.format(curr_time), cell_overwrite_ok=True)  # 新建sheet名为"curr_time"的页面，且覆盖写入=True
        
        #* style=正文内容格式
        self.style = xlwt.XFStyle()               # 初始化样式
        self.borders = xlwt.Borders()             # 初始化边框
        self.alignment = xlwt.Alignment()         # 初始化对齐样式
        self.font = xlwt.Font()                   # 初始化字体
        self.font.name = 'Arial'                  # 设定字体为Arial
        self.font.bold = True                     # 设定字体为粗体
        self.style.font = self.font               # 设定字体样式(Arial, 粗体)
        self.alignment.wrap = 1                   # 打开自动换行
        self.alignment.horz = 0x02                # 0x02=水平居中
        self.alignment.vert = 0x01                # 0x01=垂直居中
        self.style.alignment = self.alignment     # 设置对齐格式(打开自动换行, 水平居中, 垂直居中)
        self.borders.left = 6                     # 设置边框为6=双线
        self.borders.right = 6
        self.borders.top = 6
        self.borders.bottom = 6
        self.style.borders = self.borders

        #* style1=定义首行标题栏格式
        self.style1 = xlwt.XFStyle()
        self.pattern = xlwt.Pattern()
        self.font1 = xlwt.Font()
        self.pattern.pattern = xlwt.Pattern.SOLID_PATTERN   # 设置背景颜色模式
        self.pattern.pattern_fore_colour = 30               # 淡蓝色(Amlogic)
        self.font1.name = 'Arial'                           # 字体Arial
        self.font1.bold = True                              # 加粗
        self.font1.underline = True                         # 下划线
        self.font1.italic = True                            # 斜体
        self.font1.colour_index = 1                         # 红色字体
        self.style1.font = self.font1                       # 设定字体样式
        self.style1.alignment.wrap = 1                      # 打开自动换行
        self.style1.alignment.horz = 0x02                   # 0x02=水平居中
        self.style1.alignment.vert = 0x01                   # 0x01=垂直居中
        self.style1.borders = self.borders
        self.style1.pattern = self.pattern

        #* style2=定义第四列Summary靠左对齐
        self.style2 = xlwt.XFStyle()
        self.alignment2 = xlwt.Alignment()
        self.style2.font = self.font             # 设定字体样式
        self.alignment2.wrap = 1                 # 打开自动换行
        self.alignment2.horz = 0x01              # 0x01=靠左对齐
        self.alignment2.vert = 0x01              # 0x01=垂直居中
        self.style2.alignment = self.alignment2  # 设置对齐格式
        self.style2.borders = self.borders
        
        #* style3=定义正文红色字体
        self.style3 = xlwt.XFStyle()
        self.font3 = xlwt.Font()                 # 初始化字体
        self.font3.name = 'Arial'                # 设定字体为Arial
        self.font3.bold = True                   # 设定字体为粗体
        self.font3.colour_index = 0x0A           # 设定字体颜色为red
        self.style3.font = self.font3            # 设定字体样式
        self.style3.alignment = self.alignment   # 设置对齐格式
        self.style3.borders = self.borders

        #* 调整列宽
        self.worksheet.col(0).width = 6 * 256    # 第01列：Index，设置宽度
        self.worksheet.col(1).width = 20 * 256   # 第02列：Produect Line，设置宽度
        self.worksheet.col(2).width = 20 * 256   # 第03列：Project ID，设置宽度
        self.worksheet.col(3).width = 16 * 256   # 第04列：Issue Key，设置宽度
        self.worksheet.col(4).width = 24 * 256   # 第05列：Component，设置宽度
        self.worksheet.col(5).width = 16 * 256   # 第06列：Status，设置宽度
        self.worksheet.col(6).width = 12 * 256   # 第07列：Priority，设置宽度
        self.worksheet.col(7).width = 18 * 256   # 第08列：Assignee，设置宽度
        self.worksheet.col(8).width = 18 * 256   # 第09列：RD Manager，设置宽度
        self.worksheet.col(9).width = 16 * 256   # 第10列：Created，设置宽度
        self.worksheet.col(10).width = 16 * 256  # 第11列：Updated，设置宽度
        self.worksheet.col(11).width = 16 * 256  # 第12列：Finish Date，设置宽度
        self.worksheet.col(12).width = 12 * 256  # 第13列：Cost，设置宽度

        #* 调整首行格式
        self.first_style = xlwt.easyxf('font:height 720;')  # 设置行高
        self.tall_style = xlwt.easyxf('font:height 540;')   # 设置行高
        first_row = self.worksheet.row(0)                   # 第一行标题
        first_row.set_style(self.first_style)               # 第一行设置格式

    def write2excel(self, row=None, col=None, input=None, style=None) -> object:
        self.worksheet.write(row, col, input, style)    # 写入数据格式: (行, 列, 具体内容, 字体格式样式)

def write2file(iteration=None, limit=None):
    '''写入Fields数据到excel表格中'''

    #* 新建excel表格
    excel_file =  creat_local_file(filename="Output_Result")

    #* xlwt excel样式定义初始化
    LoExcel = TableObject()

    #* Excel Style定义
    title_style = LoExcel.style1      # 表格标题样式(蓝底,白字,斜体,加粗,居中对齐)
    content_style = LoExcel.style     # 正文内容样式(Arial字体,加粗,居中对齐)
    left_style = LoExcel.style2       # 正文内容样式(Arial字体,加粗,靠左对齐)
    red_style = LoExcel.style3        # 正文内容样式(Arial字体,加粗,红色,居中对齐)
    
    #* Excel表格标题
    title = [   'Index',              # 01
                'Product Line',       # 02
                'Project ID',         # 03
                'Issue Key',          # 04
                'Component',          # 05
                'Status',             # 06
                'Priority',           # 07 
                'Assignee',           # 08
                'Manager',            # 09
                'Created',            # 10
                'Updated',            # 11
                'Finish date',        # 12
                'Cost Time'           # 13
            ] 

    #* Excel表格中写入首行标题栏
    for i in range(len(title)):
        LoExcel.worksheet.write(0, i, title[i], title_style)  # style1=标题栏格式（蓝底白字)
        LoExcel.workbook.save(excel_file) 

    for i,(k, v) in track(enumerate(iteration.items(),1),  description="[green]Processing[/green]", total=(len(fields))):
        LoExcel.write2excel(row=i, col=0, input=i, style=content_style)
        LoExcel.write2excel(row=i, col=1, input=v[0]['product'], style=left_style)
        LoExcel.write2excel(row=i, col=2, input=v[0]['project_id'], style=left_style)
        LoExcel.write2excel(row=i, col=3, input=v[0]['issue_id'], style=content_style)
        LoExcel.write2excel(row=i, col=4, input=v[0]['component'], style=left_style)
        LoExcel.write2excel(row=i, col=5, input=v[0]['status'], style=left_style)
        LoExcel.write2excel(row=i, col=6, input=v[0]['priority'], style=content_style)
        LoExcel.write2excel(row=i, col=7, input=v[0]['assignee'], style=left_style)
        LoExcel.write2excel(row=i, col=8, input=v[0]['rd_manager'], style=left_style)
        LoExcel.write2excel(row=i, col=9, input=str(v[0]['created']).split(' ')[0], style=content_style)
        LoExcel.write2excel(row=i, col=10, input=str(v[0]['updated']).split(' ')[0], style=content_style)
        LoExcel.write2excel(row=i, col=11, input=str(v[0].get('finish_date')).split(' ')[0], style=content_style)
        if v[0]['cost'] >= limit:
            LoExcel.write2excel(row=i, col=12, input=v[0]['cost'], style=red_style)
        else:
            LoExcel.write2excel(row=i, col=12, input=v[0]['cost'], style=content_style)
        LoExcel.workbook.save(excel_file) 
    LoExcel.workbook.save(excel_file) 
    logging.info('Saved to: {}'.format(os.path.join(os.getcwd(), excel_file))) 

def showLogo():
    console.print("""
  .d8888888b.                888                   d8b                 d8888          888            
 d88P"   "Y88b               888                   Y8P                d88888          888            
 888  d8b  888               888                                     d88P888          888            
 888  888  888 88888b.d88b.  888  .d88b.   .d88b.  888  .d8888b     d88P 888 888  888 888888 .d88b.  
 888  888bd88P 888 "888 "88b 888 d88""88b d88P"88b 888 d88P"       d88P  888 888  888 888   d88""88b 
 888  Y8888P"  888  888  888 888 888  888 888  888 888 888        d88P   888 888  888 888   888  888 
 Y88b.     .d8 888  888  888 888 Y88..88P Y88b 888 888 Y88b.     d8888888888 Y88b 888 Y88b. Y88..88P 
  "Y88888888P" 888  888  888 888  "Y88P"   "Y88888 888  "Y8888P d88P     888  "Y88888  "Y888 "Y88P"  
                                               888                                                   
                                          Y8b d88P                                                   
                                           "Y88P"                                                                                                   
""", style='blue')
    sleep(1)


if __name__ == '__main__':

    start_time = time.time()   # 计时起点
    logging_init()             # logging初始化
    showLogo()                 # show logo
    tv_product_team = [ 'Jianfan.Ai','Bo.Ren','Zhewu.Tao','Zanbo.Huang','Linguo.Bu','Maoguo.Xie','Cong.Zhang','Jianhui.Peng','Shuangxiao.Hu','Jie.Xiong','Xinying.Yang','Haolin.Li','Ying.Li','Will.Chen','Huinan.Liang','Jianhua.Huang','Binbin.Gao','Zhendong.Zhou',  # SZ
                        'Meiling.Zhu','Xiaofeng.Li','Xuejiao.Li','Xiaoshuang.Ni','Qin.Zhang','Xinyue.Yu','Mingdong.Wang','Yunzhu.Zhang','Hongyu.Wang','Zonghao.Ma','Zihan.Wang',                                            # BJ
                        'Tracy.Chen','Haiying.Liu','Yueming.Xu','Qianyi.Liu',                                                                                                                                               # SH
                        'Changwen.Dai', 'Jinbo.Du', 'Chunyan.Liu', 'Menghui.Liu', 'Jiajia.Mu'  ]                                                                                                                                             

    #* Step1: 获取外部参数并格式化
    external_args_dict = args_init()

    #* Step2: AmlJiraSystem实例
    Viz = AmlJiraSystem('jianfan.ai', 'Amlogic1234!')
    
    #* Step3: 返回myjira对象
    Viz.login_jira() 

    #* Step4: 通过外部参数组合生成JQL搜索语句
    #* external_args_dict数据类型为dict, 由外部参数组合而成的dict, 返回的jql数据类型为str 
    jql = Viz.packaging_filter_from(external_args_dict)

    #* Step5: 传入JQL字段,该数据来自外部参数传入,要么是自行定义的参数 or Raw command (注意: 这里有实际访问Jira)
    #* RAW_COMMAND数据类型为str, 即JQL搜索语句. ex: "project id"=AB30A8-T962X3Z AND status in (OPEN)
    if RAW_COMMAND:
        Viz.process_search(RAW_COMMAND[0])
    else:
        Viz.process_search(jql)
 
    #* Step6: 将filter_list传入Jira处理函数, 返回: fields为所有Jira数据, commentor_all_count为所有comment作者的记录
    fields = Viz.fields
    commentor_all_count = Viz.commentor_all_count
    verified_all_count = Viz.verified_all_count
    verified_QA_count = Viz.verified_QA_count
    support_label_count = Viz.support_label_count
    severity_count = Viz.severity_count

    #* Step7: 如果ACTIVE_CHECK=True, 则将commentor_all_count转化为termgraph图形打印到stdout
    if ACTIVE_CHECK and commentor_all_count:
        Viz.show_chart(commentor_all_count, len(commentor_all_count), "Comments", "added comments")
    
    if EXPAND_FLAG and VERIFY_CHECK and DATERANGE and verified_all_count:
        Viz.show_chart(dict(verified_all_count), sum([x for x in dict(verified_all_count).values()]), "Verified", "changed status to verified")

    if EXPAND_FLAG and VERIFY_CHECK and DATERANGE and verified_QA_count:
        Viz.show_chart(dict(verified_QA_count), sum([x for x in dict(verified_QA_count).values()]), "QA Verified", "changed status to verified")      # verified_QA_count为Counter()数据类型, 需要转化成dict()
    
    if EXPAND_FLAG and LABEL_CHECK and DATERANGE:
        Viz.show_chart(dict(support_label_count), sum([x for x in dict(support_label_count).values()]), "Label", "added label < {} >".format(LABEL_CHECK[0]))
    
    if DI_COUNT:
        Viz.show_chart(dict(severity_count), sum([x for x in dict(severity_count).values()]), "Severity", "changed severity")

    if OUTPUT_FLAG:
        write2file(iteration=fields, limit=30)       # 迭代对象=fields, 红色Highlight字体时间限制>=30天
        
    #* 计时终点
    end_time = time.time()
    cost = end_time - start_time
    logging.info('>>> Cost: {:.1f}s'.format(cost)) 
