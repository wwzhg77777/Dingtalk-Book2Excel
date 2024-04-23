#!/usr/bin/python3
# -*- coding: utf-8 -*-
'''
@Author      :  ww1372247148@163.com
@AuthorDNS   :  wendirong.top
@CreateTime  :  2023-03-27 11:42:56
@FilePath    :  utils_logging.py
@FileVersion :  1.0
@LastEditTime:  2023-03-27 11:42:56
@FileDesc    :  提供logging日志记录的工具函数
'''

from . import *


class CustomLogger(logging.Logger):

    _fh: logging.StreamHandler
    _ch: logging.StreamHandler
    _logger: logging.Logger

    def __init__(self, log_name_: str, log_desc_: str = '', is_console_: bool = True, start_in_log_: bool = True):
        '''
        初始化logging日志记录的操作类
    log_name_           : 日志文件名
        log_desc_       : 日志的初始化描述
    is_console_         : 是否在控制台输出
        start_in_log_   : 是否记录初始化信息
        '''
        self._logger = logging.getLogger(log_name_ + str(int(time.time())))
        log_path = os.path.join(ROOTPATH, 'logs')
        if not os.path.exists(log_path):
            os.makedirs(log_path)

        self._logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s[line:%(lineno)d]: %(message)s')
        self._fh = logging.FileHandler(os.path.join(log_path, f"{log_name_}.log"), mode='a', encoding='utf-8')
        self._fh.setLevel(logging.INFO)
        self._fh.setFormatter(formatter)
        self._logger.addHandler(self._fh)

        if is_console_:
            self._ch = logging.StreamHandler()
            self._ch.setLevel(logging.INFO)
            self._ch.setFormatter(formatter)
            self._logger.addHandler(self._ch)
        self._logger.info('{} Function Start ......'.format(log_desc_)) if start_in_log_ else ''

    def get_logs(self):
        return self._logger

    def remove_logs(self):
        self._logger.removeHandler(self._fh)
        self._logger.removeHandler(self._ch)