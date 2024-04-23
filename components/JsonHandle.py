#!/usr/bin/python3
# -*- coding: utf-8 -*-
'''
@Author      :  ww1372247148@163.com
@AuthorDNS   :  wendirong.top
@CreateTime  :  2023-04-19 21:29:39
@FilePath    :  JsonHandle.py
@FileVersion :  1.0
@LastEditTime:  2023-04-19 21:29:39
@FileDesc    :  用于读写Json数据的Json操作类
'''

from . import *

class JsonUtils:
    @staticmethod
    def getJsonLogger(prefix_: str = ''):
        logger_hander_ = CustomLogger(prefix_, start_in_log_=False)
        logger_ = logger_hander_.get_logs()
        return logger_, logger_hander_

    @staticmethod
    def writeJson(json_fn_: str, json_data_: object, is_log: bool = True):
        import json

        logger, logger_hander = JsonUtils.getJsonLogger('json')
        with open(json_fn_, 'w+', encoding='utf-8') as f:
            f.write(json.dumps(json_data_, indent=4, ensure_ascii=False))
        logger.info('Success write {} JsonFile. path: {}'.format(os.path.basename(json_fn_), json_fn_)) if is_log else ''
        logger_hander.remove_logs()
