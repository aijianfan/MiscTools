# !/usr/bin/python
# -*- coding: UTF-8 -*-
##################################################
# __Date__: 2024/4/3
# __Python__: 3.9.6
# __Phidget22__: v1.17.20231004(http://www.phidgets.com)
# __Opencv__: v4.6.0
# __Author__: Jianfan.Ai
##################################################
# Usage:
# 1. 如何进入学习模式. ex: python3 iRemote.py -l -c Hisense                 // -l:进入学习模式, -c Hisense: 指定为Hisense厂商的Remote 
# 2. 如何进行发送键值. ex: python3 iRemote.py -s -c Hisense -k Home         // -s:进入发送模式, -c Hisense: 指定为Hisense厂商的Remote, -k Home: 发送Hisense Remote的Home键 
# 3. 如何加载所选定的测试脚本文件配置文件来进行测试 . ex: python3 iRemote.py -f test_atv_playback.yaml   // 加载该配置文件中的所有信息并开始测试,包括: 厂商Remote、循环次数、具体按键、延迟时间等
##################################################

import os
import sys
import time
import yaml
import logging
import argparse
import platform
import signal
import random
from rich.progress import track
from rich.logging import RichHandler
from rich.console import Console
from Phidget22.Devices.IR import IR, CodeInfo
from Phidget22.PhidgetException import PhidgetException

class LoadConfig:
    """ 初始化脚本参数、logging等配置 """
    def __init__(self):
        self.console = Console()         # rich console对象
        self.parse_args()                # 解析脚本外部参数
        self.show_logo()                 # 加载Amlogic Auto logo
        self.init_logging()              # 初始化logging配置
        self.print_args()                # 脚本外部参数的实际配置情况
        self.display_system_info()       # 显示OS、CPU、Python版本等系统信息
        if self.all:                     # 当设置 "-a" 参数时, 列出所有厂商键值表信息
            self.show_all_counts()

    def parse_args(self):
        """ 脚本参数配置 """
        parser = argparse.ArgumentParser(description='********** PhidgetIR Automation Tool **********', prog=sys.argv[0])
        parser.add_argument('-v', '--version', action='version', version='V1.0.0')
        parser.add_argument('-l', '--learn', dest='learn', action='store_true', help='Determine whether PhidgetIR needs to enter learning mode 1')
        parser.add_argument('-m', '--mode', dest='mode', action='store_true', help='Determine whether PhidgetIR needs to enter learning mode 2')
        parser.add_argument('-s', '--send', action='store_true', help='Determine to parse and send a specific remote key')
        parser.add_argument('-c', '--customer', type=str, help='Type the customer name you want, ex: XiaoMi, Hisense')
        parser.add_argument('-k', '--key', type=str, help='Type the remote key you want, ex: Power, Enter')
        parser.add_argument('-f', '--file', type=str, help='Type the YAML case you want to execute, ex: test_atv_playback.yaml')
        parser.add_argument('-a', '--all', action='store_true', help='Show all Manufacture and Keycode data from YAML file')
        parser.add_argument('-r', '--random', action='store_true', help='Send random keycode from the specific customer')
        parser.add_argument('--verbose', action='store_true', help='Show more information while running, default: False')
        
        args = parser.parse_args()

        # 将所有命令行参数存储到类属性中
        self.learn = args.learn             # PhidgetIR学习模式1
        self.mode = args.mode               # PhidgetIR学习模式2
        self.send = args.send               # PhidgetIR发射模式
        self.customer = args.customer       # 平台厂商信息
        self.key = args.key                 # 平台厂商遥控器键值信息
        self.file = args.file               # ./TestCases/目录下的YAML测试用例文件
        self.all = args.all                 # 列出所有的厂商与键值信息
        self.random = args.random           # 是否随机
        self.verbose = args.verbose         # 调试打印开关

        # 脚本不带参数执行时，返回usage的提示
        if not any(vars(args).values()):
            parser.exit(message=parser.format_usage())

    def init_logging(self):
        """ 初始化logging模块(包含动态参数的配置Verbose调试开关) """
        logging.basicConfig(level=logging.DEBUG if self.verbose else logging.INFO,
                            # format='%(message)s',
                            format='%(message)s',
                            handlers=[
                                RichHandler(rich_tracebacks=True,                                # rich_tracebacks开关
                                            tracebacks_show_locals=True,                         # Exception时展示代码段
                                            log_time_format="[%Y/%m/%d %H:%M:%S]",               # datetime时间格式
                                            omit_repeated_times=False,                           # False=逐行打印logging时间戳 
                                            keywords=['USB Camera', 'PhidgetIR', 'Hold']),       # 高亮Console中关键字
                                logging.FileHandler('Script.log', mode='a', encoding='utf-8')
                                                  ])

    def show_logo(self):
        """ 打印Amlogic Auto logo """
        self.console.print("""
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
        time.sleep(1)

    def print_args(self):
        """ 打印脚本参数设置情况 """
        args_dict = {
            'Learning Mode 1': self.learn,
            'Learning Mode 2': self.mode,
            'Sending Signal': self.send,
            'Customer': self.customer,
            'Remote Key': self.key,
            'YAML Case': self.file,
            'Show info': self.all,
            'is Random': self.random,
            'Verbose': self.verbose
        }

        logging.info('Get parameter configuration:')
        logging.info(f"{'*'*50}")
        for index, (arg, value) in enumerate(args_dict.items(), 1):
            logging.info(f'[{str(index).zfill(2)}] {arg.ljust(16, " ")} = {value}')
        logging.info(f"{'*'*50}\n")

    def show_all_counts(self):
        """ 统计Manufacturer、Keycode信息 """
        if self.all:
            # Load YAML file
            with open(YAML_KEYCODE_FILE, 'r', encoding='utf-8') as yaml_file:
                data = yaml.safe_load(yaml_file)
            
            # Count Manufacturer and Keycode
            manufacturer_count = len(data)
            keycode_count = sum(len(item) for item in data.values())

            # Print counts
            logging.info(f"{'*'*50}")
            logging.info(f"Manufacturer Count: {manufacturer_count}")
            logging.info(f"Keycode Count: {keycode_count}")
            logging.info(f"{'*'*50}")
            
            # Print manufacturer and keycode details
            logging.info("Manufacturer and Keycode details:")
            for manufacturer, details in data.items():
                logging.info(f"Manufacturer: < {manufacturer} >, Count: {len(details)} ")
                for keycode, _ in details.items():
                    logging.info(f"  - Keycode: {keycode}")

    def display_system_info(self):
        """ 获取操作系统相关信息 """
        os_info = platform.platform()
        processor_info = platform.processor()
        architecture = platform.architecture()
        python_version = platform.python_version()

        logging.debug("System Information:")
        logging.debug(f"{'*'*50}")
        logging.debug(f"{'# Operating System:' :<20} {os_info}")
        logging.debug(f"{'# Processor:' :<20} {processor_info}")
        logging.debug(f"{'# Architecture:' :<20} {architecture[0]} {architecture[1]}")
        logging.debug(f"{'# Python Version:' :<20} {python_version}")
        logging.debug(f"{'*'*50}\n")

        # 提供建议
        if "Windows" in os_info:
            logging.debug("Consider using Windows Subsystem for Linux (WSL) for development purposes.")
        elif "Linux" in os_info:
            logging.debug("Ensure your system is up-to-date with the latest security patches.")
        elif "Darwin" in os_info:
            logging.debug("Use Homebrew for managing software packages on macOS.")

        # 检查Python版本建议
        if sys.version_info < (3, 8):
            logging.warning("Consider upgrading to Python 3.8 or later for improved features and performance.")

def format_duration(seconds):
    """ 将秒转换为 x day(s) x hour(s) x min(s) x second(s) 的格式 """
    days, seconds = divmod(seconds, 86400)
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    return f"{days} day(s) {hours} hour(s) {minutes} min(s) {seconds:.1f} second(s)"

def monitorThread(signum, frame):
    """ 当按下Ctrl+C中断脚本时, 会捕获到KeyboardInterrupt异常 """
    logging.warning('You choose to stop running the script')
    sys.exit(0)

def summary_result():
    """ 统计脚本执行耗时 """
    end_time = time.time()
    elapsed_time = end_time - start_time
    logging.info(f"{'='*65}")
    logging.info(f"Running Time: {format_duration(elapsed_time)}")
    logging.info(f"{'='*65}")

def random_transmit(manufacturer=None, count=0):
    """ 随机发送指定厂商遥控器键值 """
    with open(file=YAML_KEYCODE_FILE, mode='r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    while True:
        count += 1
        logging.info(f"{'='*20} Loop: {count} {'='*20}")
        key = random.choice(list(data[manufacturer.capitalize()].keys()))
        if key == "Power":
            for _ in range(2):
                command = ir.code_transition(YAML_KEYCODE_FILE, manufacturer.capitalize(), key.capitalize())
                ir.transmit_code(*command)
                time.sleep(10)
        else:
            command = ir.code_transition(YAML_KEYCODE_FILE, manufacturer.capitalize(), key.capitalize())
            ir.transmit_code(*command)
        time.sleep(random.uniform(1, 2))

# @pysnooper.snoop()
def process_action(action_name, action_details, ir):
    """ 主程序 """
    logging.info(f"{'='*65}")
    logging.info(f"Processing: {action_name}")
    manufacturer = action_details['Manufacturer']
    cycle = action_details['Cycle']
    steps = action_details['Steps']

    # 从YAML测试文件中读取 Manufacturer、Cycle、Steps、Duration等数据, 然后进行相应的操作
    try:
        if isinstance(cycle, int):
            for i in range(cycle):
                logging.info(f"{'='*20} Start of loop: {i+1} of {cycle} {'='*20}")
                execute_steps(steps, ir, manufacturer, i+1)
        else:
            loop_count = 1
            while True:
                logging.info(f"{'='*20} Start of infinite loop: {loop_count} {'='*20}")
                execute_steps(steps, ir, manufacturer, loop_count)
                loop_count += 1
    except Exception as e:
        logging.error(e)

def execute_steps(steps, ir, manufacturer, count):
    """ 执行每一个红外键值发射的步骤 """
    for step in steps:
        action = step['Step']
        duration = step.get('Duration')
        command = ir.code_transition(YAML_KEYCODE_FILE, manufacturer.capitalize(), action.capitalize())
        ir.transmit_code(*command)
        time.sleep(duration)

# @pysnooper.snoop()
class PhidgetIR:
    """ Phidget IR测试设备 """
    def __init__(self):
        self.ir = IR()
        self.codeinfo = dict()
        self.rawlist = list()
        try:
            self.ir.openWaitForAttachment(5000)
            self.ir.setOnAttachHandler(self.onIR_attach())
            # self.ir.setOnDetachHandler(self.onIR_detach())
            # self.ir.setOnErrorHandler(self.onIR_Error)
            self.show_info()
        except Exception as e:
            logging.error(f"Failed to attach Phidget IR: {e}")
            config.console.print_exception()
            sys.exit(1)
    
    def onIR_attach(self):
        """ 设备正常挂载时返回信息 """
        logging.info("Phidget device attached successfully!")

    def show_info(self):
        """ PhidgetIR设备信息 """
        logging.debug(f"{'*'*50}")
        logging.debug(f"{'# Device Name:' :<18} {self.ir.getDeviceName()}")
        logging.debug(f"{'# Serial Number:' :<18} {self.ir.getDeviceSerialNumber()}")
        logging.debug(f"{'# Device Version:' :<18} {self.ir.getDeviceVersion()}")
        logging.debug(f"{'# Channel:' :<18} {self.ir.getChannel()}")
        logging.debug(f"{'# Is Attached:' :<18} {self.ir.getAttached()}")
        logging.debug(f"{'*'*50}\n")

    def onIR_detach(self):
        """ 设备卸载时返回信息 """
        logging.info("Phidget IR Detach!!")

    def on_learn(self, ir_device: str, code: str ,codeInfo: dict ) -> None:
        """ 设备进入学习模式 
            ir_device: Phidget IR's S/N
            code: remote code value
            codeInfo: data structure of remote code
        """
        logging.info('Detected new remote event!')
        self.button_name = input('Please type the button name(eg: power): ').capitalize()
        self.code = code
        self.codeinfo['bitCount'] = codeInfo.bitCount
        self.codeinfo['encoding'] = codeInfo.encoding
        self.codeinfo['zero'] = list(codeInfo.zero)
        self.codeinfo['one'] = list(codeInfo.one)
        self.codeinfo['header'] = list(codeInfo.header)
        self.codeinfo['trail'] = codeInfo.trail
        self.codeinfo['gap'] = codeInfo.gap
        # self.codeinfo['repeat'] = list(codeInfo.repeat)
        self.codeinfo['repeat'] = 0        # default set to 0
        self.codeinfo['minRepeat'] = codeInfo.minRepeat
        self.codeinfo['dutyCycle'] = codeInfo.dutyCycle
        self.codeinfo['toggleMask'] = codeInfo.toggleMask
        self.codeinfo['carrierFrequency'] = int(codeInfo.carrierFrequency)

        logging.info(f'Code: [ 0x{self.code} ]')
        logging.info(f'CodeInfo: [ {self.codeinfo} ]')

        # 保存厂商、按键名、键值数据、键值结构到指定的YAML文件
        self.to_file(manufacturer=config.customer.capitalize(), button_name=self.button_name.capitalize(), code=self.code, codeInfo=self.codeinfo, rawdata=None, file=YAML_KEYCODE_FILE)

    def parse_from_yaml(self, file: str, manufacturer: str, button_name: str):
        """ 从YAML文件中读取红外键值数据: YAML文件、厂商、按键名, 返回: 键值、键值数据 """
        try:
            with open(file, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                code_info = data[manufacturer.capitalize()][button_name.capitalize()]
                code = code_info.get('code', None)
                codeInfo = code_info.get('codeInfo', None)
                rawdata = code_info.get('rawdata', None)
                return code, codeInfo, rawdata
        except FileNotFoundError:
            logging.error(f'File not found: [{file}].')
            config.console.print_exception()
        except KeyError as e:
            logging.error(f'Missing key: {e} not in file [{file}]')
            config.console.print_exception()

    def onRawData(self, ir_device, rawdata) -> None:
        """ 保存int类型raw数据到List """
        logging.info('Detected new remote event!')
        self.rawlist.extend(filter(lambda x: x >= 0, rawdata))
        logging.info(f"rawdata: {rawdata}")

    def code_transition(self, local_file: str, customer: str, remote_button: str):
        """ 从YAML文件中提取客户、按键信息, 转化为所需的键值数据、键值结构 """
        try:
            logging.info(f'Attempt to send remote key from [ {customer} ]: < {remote_button} >')
            code, codeInfo_dict, rawdata =  self.parse_from_yaml(file=local_file, manufacturer=customer, button_name=remote_button)
        except FileNotFoundError:
            logging.error('YAML file not found.')
            config.console.print_exception()
        except KeyError as e:
            logging.error(f'Missing key in data: {e}')
            config.console.print_exception()
        
        # Phidget22特有的CodeInfo类型数据实例化
        codeInfo = CodeInfo()

        # CodeInfo数据反序列化并定义了默认初始值
        if code and codeInfo_dict:
            for attr, default_value in [
                ('bitCount', 32), ('encoding', 2), ('zero', [525, 594]),
                ('one', [525, 1708]), ('header', [4442, 4516]), ('trail', 525),
                ('gap', 107669), ('minRepeat', 1), ('dutyCycle', 0.5),
                ('toggleMask', ''), ('carrierFrequency', 38000)]:
                setattr(codeInfo, attr, codeInfo_dict.get(attr, default_value))

        logging.debug(f'Key code read from file: 0x{code}')
        logging.debug(f'Key codeInfo converted from file : {codeInfo}')
        return code, codeInfo, rawdata

    # def onIR_Error(self, code: str, description: str):
    #     """ 设备工作异常时返回信息 """
    #     logging.error("Code: " + ErrorEventCode.getName(code))
    #     logging.error("Description: " + str(description))
    #     logging.error("----------")
        
    def learn_code(self):
        """ 学习模式1: 注册监听Keycode、CodeInfo """
        logging.info('Attempt to enter learning mode 1 ...')
        logging.info('Please Hold your remote button for 4s!')
        self.ir.setOnLearnHandler(self.on_learn)

        while True:
            time.sleep(1)

    def learn_code2(self):
        """ 学习模式2: 注册监听键值获取Raw Data """
        logging.info('Attempt to enter learning mode 2 ...')
        logging.info('You must press the remote button in 10s, Please Hold your remote button at least for 1s!')
        self.button_name = input('Please type the button name(eg: power): ').capitalize()
        self.ir.setOnRawDataHandler(self.onRawData)

        for i in track(sequence=range(10) ,description="Count down...", refresh_per_second=0.5):
            time.sleep(1)
            logging.info(f"Rawdata: {self.rawlist}, Length: {len(self.rawlist)}")
            if i == 9:
                self.transmit_code(code=None, codeInfo=None, rawdata=self.rawlist)
                self.to_file(manufacturer=config.customer.capitalize(), button_name=self.button_name.capitalize(), code=None, codeInfo=None, rawdata=self.rawlist, file=YAML_KEYCODE_FILE)

    def transmit_code(self, code: str, codeInfo: dict, rawdata: list):
        """ 发射红外键值 """
        try:
            if code is not None:
                logging.debug(f"Transmitted code: {code}")
                self.ir.transmit(code, codeInfo)
            if rawdata is not None:
                logging.debug(f"Transmit Rawdata...")
                self.ir.transmitRaw(data=rawdata, carrierFrequency=0, dutyCycle=0, gap=100000)
        except PhidgetException as e:
            logging.error(f"Error during code transfering: {e}")
            raise
    
    def transmit_rawdata(self, rawdata: list, freq: int, period: int, gap: int) -> None:
        """ 发射红外Rawdata数据 """
        try:
            logging.info(f"Transmit Rawdata...")
            self.ir.transmitRaw(data=rawdata, carrierFrequency=freq, dutyCycle=period, gap=gap)
            logging.debug(f"Transmit rawdata: {rawdata}, config: carrierFrequency: {freq}, dutyCyle: {period}, gap: {gap}")
        except PhidgetException as e:
            logging.error(f"Error during code transfering: {e}")
            raise

    # def to_file2(self, manufacturer: str, button_name: str, rawdata: list, file: str) -> None:
    #     """ 向YAML文件中添加遥控器按键代码. """
    #    try:
    

    def to_file(self, manufacturer: str, button_name: str, code: str, codeInfo: dict, rawdata: list, file: str) -> None:
        """ 向YAML文件中添加遥控器按键代码. """
        # 尝试加载已有的YAML数据，如果文件不存在或为空，则初始化为空字典
        try:
            with open(file, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f) or {}
        except FileNotFoundError:
            data = {}

        previous_button_count = 0  # 初始化先前的按钮数量
        
        # 更新数据结构
        if manufacturer not in data:
            data[manufacturer] = {}
            data[manufacturer][button_name] = {
            'code': code,
            'codeInfo': codeInfo,
            'rawdata': rawdata
            }
        else:
            # 如果制造商已存在，记录先前的按钮数量
            previous_button_count = len(data[manufacturer])
            data[manufacturer][button_name] = {
            'code': code,
            'codeInfo': codeInfo,
            'rawdata': rawdata
            }

        # 计算新增后的按钮数量
        new_button_count = len(data[manufacturer])

        # 写回YAML文件, 如Keycode已存在则会覆盖
        with open(file, 'w', encoding='utf-8') as f:
            logging.info(f'Already saved to yaml file: [ {file} ], '
                         f'Manufacturer: [ {manufacturer} ],'
                         f'Button: [ {button_name} ], '
                         f'Previous count: {previous_button_count}, '
                         f'New count: {new_button_count}')
            yaml.dump(data, f, default_flow_style=False)

    def __del__(self):
        self.ir.close()

if __name__ == "__main__":

    start_time = time.time()        # 记录脚本执行的起始时间
    
    # 按键监测
    signal.signal(signal.SIGINT, monitorThread)
    signal.signal(signal.SIGTERM, monitorThread) 
    
    YAML_KEYCODE_FILE = 'remote_control_codes.yaml'   # 保存所有遥控键值的配置文件
    config = LoadConfig()             # 配置各类初始化参数
    
    try:
        # 拼接测试脚本文件的具体路径. ex: ./TestCases/test_atv_playback.yaml
        YAML_TEST_CASE = os.path.join('TestCases', config.file)  
    except TypeError:
        YAML_TEST_CASE = None

    # 当存在外部参数[-l, -m, -s, -f, -r]时, 才对PhidgetIR设备进行初始化
    if any([config.learn, config.mode, config.send, config.file, config.random]):
        ir = PhidgetIR()                

    if config.learn and not config.customer:
        logging.warning(f'Enter Learning Mode need execute script with add "-c CUSTOMER" param!')

    # 进入学习模式1(阻塞模式, 依赖参数:[-l, -c CUSTOMER]): 接收红外键值. ex: python3 iRemote.py -l -c Hisense
    if config.learn and config.customer:
        ir.learn_code()
    
    # 进入学习模式2(阻塞模式, 依赖参数:[-m, -c CUSTOMER]): 接收红外键值. ex: python3 iRemote.py -m -c Hisense
    if config.mode and config.customer:
        ir.learn_code2()
    
    # 进入单个键值发射模式(单次执行, 依赖参数:[-s, -c CUSTOMER, -k KEY]): 发送红外键值. ex: python3 iRemote.py -s -c Hisense -k Home
    if config.send and config.customer and config.key:
        ir.transmit_code(*ir.code_transition(YAML_KEYCODE_FILE, config.customer, config.key))

    # 进入加载YAML测试文件的测试模式. ex: python3 iRemote.py -f test_atv_playback.yaml
    if config.file:
        try:
            # 解析YAML Case文件, 并执行相应测试动作 
            with open(YAML_TEST_CASE, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)

                for key, details in data.items():
                    process_action(key, details, ir)   # 主程序
        
        # except Exception as e:
        #     logging.error(f'Unexpected exception: {e}')
        #     config.console.print_exception()
        finally:
            summary_result()  # 汇总脚本执行耗时等信息
    
    # 通过[-r, -c]组合来控制随机发送指定厂商遥控键值
    if config.random and config.customer:
        random_transmit(manufacturer=config.customer)