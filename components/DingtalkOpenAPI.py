#!/usr/bin/python3
# -*- coding: utf-8 -*-
'''
@Author      :  ww1372247148@163.com
@AuthorDNS   :  wendirong.top
@CreateTime  :  2023-07-21
@FilePath    :  DingtalkOpenAPI.py
@FileVersion :  1.4
@LastEditTime:  2023-12-07
@FileDesc    :  提供访问钉钉API接口的工具函数
'''

import json

import requests
import urllib3

from . import *

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class DingtalkOpenAPI:
    '''
    访问钉钉的APIs接口数据, 本地服务报错退出, 钉钉服务报错不退出
    APIs:
    get_accessToken             : 获取钉钉企业内部应用的accessToken

    get_userInfo                : 查询用户详情
    get_deptInfo                : 获取部门详情
    get_listParentByUser        : 获取指定用户的所有父部门列表
    get_listParentByDept        : 获取指定部门的所有父部门列表
    get_listsub_dept            : 获取部门列表
    get_listsub_user            : 查询部门用户完整信息

    '''

    # Public var
    #
    jsonHeaders = {'Content-Type': 'application/json', 'Connection': 'close'}

    # Private var
    #
    __g_logger_info: logging.Logger
    # __g_logger_debug: logging.Logger
    __g_api_uri_host: str
    '''钉钉API接口的host名称'''
    __app_key: str
    '''钉钉应用的appKey'''
    __app_secret: str
    '''钉钉应用的appSecret'''
    __access_token: str
    '''钉钉企业内部应用的accessToken'''
    __token_expireIn_unix10: int
    '''accessToken的过期时间, 单位秒. 记录到Unix10位时间戳'''
    __is_log: bool
    '''是否记录info日志'''

    # 初始api接口并自动存档accessToken
    def __init__(self, app_key_: str, app_secret_: str, g_api_uri_host_: str, is_log_: bool = True):
        '''
        构造函数:
        Load    : appKey, appSecret, Host
        '''
        self.__app_key = app_key_
        self.__app_secret = app_secret_
        self.__g_api_uri_host = g_api_uri_host_
        self.__is_log = is_log_

        logger_hander_info = CustomLogger(log_name_='DingOpenApi_info', is_console_=False, start_in_log_=False)
        self.__g_logger_info = logger_hander_info.get_logs()
        # logger_hander_debug = CustomLogger(log_name_='DingOpenApi_debug', is_console_=False, start_in_log_=False)
        # self.__g_logger_debug = logger_hander_debug.get_logs()
        self.get_accessToken()

    def set_headers(self, isHeaders: bool = True):
        '''
        补齐headers参数的x-acs-dingtalk-access-token字段
        返回 json + tokens 的headers请求头
        '''
        sh_result = DingtalkOpenAPI.jsonHeaders
        if int(time.time()) > self.__token_expireIn_unix10:
            self.get_accessToken(self.__g_api_uri_host)
        if isHeaders:
            sh_result['x-acs-dingtalk-access-token'] = self.__access_token
        return sh_result

    def get_accessToken(self, api_uri_host_: str = 'api.dingtalk.com', api_uri_path_: str = '/v1.0/oauth2/accessToken', is_https: bool = True, *args, **kwargs):
        '''
        api     : 获取钉钉企业内部应用的accessToken
        url     : POST /v1.0/oauth2/accessToken HTTP/1.1
        host    : api.dingtalk.com
        Load    : self.__app_key, self.__app_secret
        return  : (status_code:int, json:dict, onetime:time)
        '''
        get_token_result = ''
        _body = {'appKey': self.__app_key, 'appSecret': self.__app_secret}
        _headers = DingtalkOpenAPI.jsonHeaders
        try:
            get_token_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", json=_body, headers=_headers, verify=False)
            ret_json = get_token_result.json()
            if self.ret_error('get_accessToken', get_token_result.status_code, ret_json):
                self.__g_logger_info.info('request Api[get_accessToken] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                self.__access_token = ret_json['accessToken']
                self.__token_expireIn_unix10 = int(time.time()) + int(ret_json['expireIn']) - 200
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_accessToken]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))
            sys.exit()

    def get_userInfo(self, api_uri_host_: str = 'oapi.dingtalk.com', api_uri_path_: str = '/topapi/v2/user/get', is_https: bool = True, *args, **kwargs):
        '''
        api     : 查询用户详情
        url     : GET /topapi/v2/user/get?access_token=String HTTP/1.1
        host    : oapi.dingtalk.com
        Load    : self.__access_token, kwargs['userid']
        return  : (status_code:int, json:dict, onetime:time)
        '''
        get_userInfo_result = ''
        _body = {'language': 'zh_CN', 'userid': kwargs['userid']}
        try:
            _sT = time.time()
            _headers = self.set_headers(False)
            _params = {'access_token': self.__access_token}
            get_userInfo_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
            _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
            ret_json = get_userInfo_result.json()
            if self.ret_error('get_userInfo', get_userInfo_result.status_code, ret_json):
                self.__g_logger_info.info('request Api[get_userInfo] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                return (ret_json, _oT)
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_userInfo]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))
            sys.exit()

    def get_deptInfo(self, api_uri_host_: str = 'oapi.dingtalk.com', api_uri_path_: str = '/topapi/v2/department/get', is_https: bool = True, *args, **kwargs):
        '''
        api     : 获取部门详情
        url     : GET /topapi/v2/department/get?access_token=String HTTP/1.1
        host    : oapi.dingtalk.com
        Load    : self.__access_token, kwargs['dept_id']
        return  : (status_code:int, json:dict, onetime:time)
        '''
        get_deptInfo_result = ''
        _body = {'language': 'zh_CN', 'dept_id': kwargs['dept_id']}
        try:
            _sT = time.time()
            _headers = self.set_headers(False)
            _params = {'access_token': self.__access_token}
            get_deptInfo_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
            _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
            ret_json = get_deptInfo_result.json()
            if self.ret_error('get_deptInfo', get_deptInfo_result.status_code, ret_json):
                self.__g_logger_info.info('request Api[get_deptInfo] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                return (ret_json, _oT)
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_deptInfo]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))
            sys.exit()

    def get_listParentByUser(self, api_uri_host_: str = 'oapi.dingtalk.com', api_uri_path_: str = '/topapi/v2/department/listparentbyuser', is_https: bool = True, *args, **kwargs):
        '''
        api     : 获取指定用户的所有父部门列表
        url     : GET /topapi/v2/department/listparentbyuser?access_token=String HTTP/1.1
        host    : oapi.dingtalk.com
        Load    : self.__access_token, kwargs['userid']
        return  : (status_code:int, json:dict, onetime:time)
        '''
        get_listParentByUser_result = ''
        _body = {'language': 'zh_CN', 'userid': kwargs['userid']}
        try:
            _sT = time.time()
            _headers = self.set_headers(False)
            _params = {'access_token': self.__access_token}
            get_listParentByUser_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
            _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
            ret_json = get_listParentByUser_result.json()
            if self.ret_error('get_listParentByUser', get_listParentByUser_result.status_code, ret_json):
                self.__g_logger_info.info('request Api[get_listParentByUser] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                return (ret_json, _oT)
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_listParentByUser]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))
            sys.exit()

    def get_listParentByDept(self, api_uri_host_: str = 'oapi.dingtalk.com', api_uri_path_: str = '/topapi/v2/department/listparentbydept', is_https: bool = True, *args, **kwargs):
        '''
        api     : 获取指定部门的所有父部门列表
        url     : GET /topapi/v2/department/listparentbydept?access_token=String HTTP/1.1
        host    : oapi.dingtalk.com
        Load    : self.__access_token, kwargs['dept_id']
        return  : (status_code:int, json:dict, onetime:time)
        '''
        get_listParentByDept_result = ''
        _body = {'language': 'zh_CN', 'dept_id': kwargs['dept_id']}
        try:
            _sT = time.time()
            _headers = self.set_headers(False)
            _params = {'access_token': self.__access_token}
            get_listParentByDept_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
            _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
            ret_json = get_listParentByDept_result.json()
            if self.ret_error('get_listParentByDept', get_listParentByDept_result.status_code, ret_json):
                self.__g_logger_info.info('request Api[get_listParentByDept] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                return (ret_json, _oT)
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_listParentByDept]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))
            sys.exit()

    def get_listsub_dept(self, api_uri_host_: str = 'oapi.dingtalk.com', api_uri_path_: str = '/topapi/v2/department/listsub', is_https: bool = True, *args, **kwargs):
        '''
        api     : 获取部门列表
        url     : GET /topapi/v2/department/listsub?access_token=String HTTP/1.1
        host    : oapi.dingtalk.com
        Load    : self.__access_token, kwargs['dept_id']
        return  : (status_code:int, json:dict, onetime:time)
        '''
        get_listsub_dept_result = ''
        _body = {'language': 'zh_CN', 'dept_id': kwargs['dept_id']}
        try:
            _sT = time.time()
            _headers = self.set_headers(False)
            _params = {'access_token': self.__access_token}
            get_listsub_dept_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
            _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
            ret_json = get_listsub_dept_result.json()
            if self.ret_error('get_listsub_dept', get_listsub_dept_result.status_code, ret_json):
                self.__g_logger_info.info('request Api[get_listsub_dept] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                return (ret_json, _oT)
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_listsub_dept]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))
            sys.exit()

    def get_listsub_user(self, api_uri_host_: str = 'oapi.dingtalk.com', api_uri_path_: str = '/topapi/v2/user/list', is_https: bool = True, take_all: bool = False, *args, **kwargs):
        '''
        api     : 查询部门用户完整信息
        url     : GET /topapi/v2/user/list?access_token=String HTTP/1.1
        host    : oapi.dingtalk.com
        Load    : self.__access_token, kwargs['dept_id', 'cursor', 'size']
        return  : (status_code:int, json:dict, onetime:time)
        '''
        if take_all and not kwargs.__contains__('size'):
            self.__g_logger_info.error('request Api[get_listsub_user] is error. more info: param object is error.')
            return
        get_listsub_user_result = ''
        _body = {'language': 'zh_CN', 'dept_id': kwargs['dept_id'], 'cursor': kwargs['cursor'], 'size': kwargs['size']}
        try:
            _sT = time.time()
            _headers = self.set_headers(False)
            _params = {'access_token': self.__access_token}
            if take_all:
                ret_datas = []
                is_loop_get = True
                while is_loop_get:
                    get_listsub_user_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
                    ret_json = get_listsub_user_result.json()
                    if self.ret_error('get_listsub_user', get_listsub_user_result.status_code, ret_json):
                        [ret_datas.append(v) for v in ret_json['result']['list']]
                        if ret_json['result']['has_more']:
                            _body['cursor'] = ret_json['result']['next_cursor']
                        else:
                            is_loop_get = False
                _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
                self.__g_logger_info.info('request Api[get_listsub_user] success. response content: {}'.format(ret_datas)) if self.__is_log else ''
                return (ret_datas, _oT)
            else:
                get_listsub_user_result = requests.post(url=f"{'https' if is_https else 'http'}://{api_uri_host_}{api_uri_path_}", params=_params, json=_body, headers=_headers, verify=False)
                _oT = APITools.custom_random(str(time.time() - _sT), 3, 1)
                ret_json = get_listsub_user_result.json()
                if self.ret_error('get_listsub_user', get_listsub_user_result.status_code, ret_json):
                    self.__g_logger_info.info('request Api[get_listsub_user] success. response content: {}'.format(ret_json)) if self.__is_log else ''
                    return (ret_json, _oT)
        except Exception as e:
            self.__g_logger_info.error('localServerError: 接口请求参数异常, 检查[get_listsub_user]的入参. errmsg: {}.\n{}'.format(str(e), traceback.format_exc()))

    def ret_error(self, req_api_name_: str, req_status_code_: int, req_result_: object):
        '''
        参考钉钉官方文档API的错误码, 记录错误日志到服务器, 不退出本地服务
        '''
        if req_status_code_ == 500:
            self.__g_logger_info.error('dingtalkApiError: 钉钉Api接口服务异常. apiname: {a}. errcode: {c}. errmsg: {e}'.format(a=req_api_name_, c=req_status_code_, e=str(req_result_)))
        elif req_status_code_ != 200:
            self.__g_logger_info.error('dingtalkApiError. apiname: {a}. errcode: {c}. errmsg: {e}'.format(a=req_api_name_, c=req_status_code_, e=str(req_result_)))
        else:
            return True


class APITools:
    '''
    辅助工具类
    '''

    @staticmethod
    def custom_random(input_, n_, flag_):
        """
        自定义取小数点后n位, flag的 0: 四舍五入, 1: 向下取整, 2: 向上取整
        """

        # custom_random = lambda r,n: str(r)[:str(r).rfind('.') + n + 1]
        num = str(input_)
        mash = num[num.rfind('.') + 1 :]
        if flag_ == 0:
            for nx in range(len(mash)):
                if n_ == 0:
                    if mash[0] <= 4:
                        return num[: num.find('.')]
                    else:
                        return num[: num.find('.')] + 1
                else:
                    if nx <= n_ + 1:
                        if mash[nx] < 4:
                            return num[: num.find('.') + n_]
                        elif mash[nx] == 4:
                            continue
                        else:
                            # n=0 : +1
                            # n=1 : +0.1
                            # n=2 : +0.01
                            # ...
                            return num[: num.find('.') + n_ + 1] + (1 / 10**n_)
                    else:
                        print('#')
        elif flag_ == 1:
            return num[: num.find('.') + n_ + 1]
        elif flag_ == 2:
            return num[: num.find('.') + n_ + 1] + (1 / 10**n_)
        else:
            return "flag为错误值"

    @staticmethod
    def WriteJson(json_: object, filedir_: str, filename_: str, prefix_: str = 'OpenAPI_', flag_: int = -1):
        """
        输出json数据到文件
        json_       : 写入的json
        filedir_    : 写入的文件夹路径
        filename_   : 写入的文件名称
        flag_       : print输出执行结果
        """
        writeJson = {'writeTime': time.strftime("%Y-%m-%d %H:%M", time.localtime(time.time())), 'result': json_}
        with open(os.path.join(filedir_, f"{prefix_}{filename_}.json"), 'w', encoding='utf-8') as f:
            f.write(json.dumps(writeJson, indent=2, ensure_ascii=False))
        print(f"执行写入:{filedir_}\\{prefix_}{filename_}.json") if flag_ == -1 else ''

    @staticmethod
    def loop_dir(fullname_: str, num: int = 1):
        '''
        新建指定名称的文件夹
        xxxx-xx-xx xx.NODE.1
        xxxx-xx-xx xx.NODE.2
        xxxx-xx-xx xx.NODE....
        '''
        programFullPath = fullname_[: fullname_.rfind('.') + 1] + str(num)

        if os.path.exists(programFullPath):
            num += 1
            return APITools.loop_dir(fullname_=programFullPath, num=num)
        else:
            return programFullPath
