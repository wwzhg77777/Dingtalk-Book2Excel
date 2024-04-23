#!/usr/bin/python3
# -*- coding: utf-8 -*-
'''
@Author      :  ww1372247148@163.com
@AuthorDNS   :  wendirong.top
@CreateTime  :  2023-12-06
@FilePath    :  dingtalk_book2excel.py
@FileVersion :  1.0
@LastEditTime:  2023-12-06
@FileDesc    :  读取钉钉通讯录的用户、组织架构，合并导出Excel表
'''

from utils.utils_const import DD_CONST
from components.JsonHandle import JsonUtils
from components.DingtalkOpenAPI import DingtalkOpenAPI

def loop_read_dinginfo(data_: dict, ret_users_: list, ret_departments_: list, ownGroup_: str):
    '''
        递归读取Json数据到Dict对象
        data_               : 读取的json数据
        ret_users_          : 当前遍历的用户写入到list
        ret_departments_    : 当前遍历的部门写入到list
        ownGroup_           : 当前遍历的部门完整名称

    SUCCESS
    return int              : 成功返回 int对象: num_rows_

    ERROR
    return None             : 错误返回 None
    '''
    from copy import deepcopy

    if data_['type'] == 'department':
        if data_.__contains__('children'):
            for item in data_['children']:
                loop_read_dinginfo(item, ret_users_=ret_users_, ret_departments_=ret_departments_, ownGroup_=ownGroup_)
            tmp_data = deepcopy(data_)
            for v in tmp_data['children']:
                if v['type'] == 'department' and v.__contains__('children'):
                    del v['children']
            ret_departments_.append(tmp_data)
        else:
            ret_departments_.append(data_)
    elif data_['type'] == 'user':
        data_['topGroup'] = data_['ownGroup'].split('/')[0]
        data_['department'] = '-'.join(data_['ownGroup'].split('/')) if len(data_['ownGroup'].split('/')) > 1 else data_['ownGroup']
        ret_users_.append(data_)


def loop_get_dinginfo(each_dept: dict, ownDeptIds: str, ownGroup: str, _dingApi: DingtalkOpenAPI = None):
    '''
        递归获取钉钉通讯录
        each_dept       : 当前遍历的钉钉部门信息
        w_ding_dept     : 写入当前遍历的钉钉部门

    SUCCESS
    return dict         : 成功返回 dict对象, Json格式

    ERROR
    return None         : 错误返回 None
    '''
    deptlist = _dingApi.get_listsub_dept(dept_id=each_dept['dept_id'])[0]['result']
    if len(deptlist) > 0:
        childrens = []
        for dept in deptlist:
            ret_dinginfo = loop_get_dinginfo(dept, f"{ownDeptIds},{dept['dept_id']}", f"{ownGroup}/{dept['name']}", _dingApi)
            childrens.append(ret_dinginfo) if ret_dinginfo is not None else ''
        userlist =_dingApi.get_listsub_user(dept_id=each_dept['dept_id'], cursor=0, size=100, take_all=True)[0]
        [
            childrens.append(
                {
                    'id': str(user['userid']),
                    'name': user['name'],
                    'jobNumber': user['job_number'] if user.__contains__( 'job_number') else '',
                    'email': user[ 'org_email'] if user.__contains__('org_email') else '',
                    'position': user[ 'title'] if user.__contains__('title') else '',
                    'deptId': str( each_dept['dept_id']),
                    'avatar': user['avatar'] if user.__contains__('avatar') else '',
                    'ownGroup': ownGroup,
                    'type': 'user',
                }
            )
            for user in userlist
        ]
        if len(userlist) > 0 or len(childrens) > 0:
            return {'id': str( each_dept['dept_id']), 'name': each_dept['name'], 'ownDeptIds': ownDeptIds, 'type': 'department', 'ownGroup': ownGroup, 'children': childrens}
    else:
        userlist =  _dingApi.get_listsub_user(dept_id=each_dept['dept_id'], cursor=0, size=100, take_all=True)[0]
        if len(userlist) > 0:
            childrens = []
            [
                childrens.append(
                    {
                        'id': str(user['userid']),
                        'name': user['name'],
                        'jobNumber': user['job_number'] if user.__contains__('job_number') else '',
                        'email': user['org_email'] if user.__contains__('org_email') else '',
                        'position': user[ 'title'] if user.__contains__( 'title') else '',
                        'deptId': str(each_dept['dept_id']),
                        'avatar': user['avatar'] if user.__contains__('avatar') else '',
                        'ownGroup': ownGroup,
                        'type': 'user',
                    }
                )
                for user in userlist
            ]
            return {'id': str( each_dept['dept_id']), 'name': each_dept['name'], 'ownDeptIds': ownDeptIds, 'type': 'department', 'ownGroup': ownGroup, 'children': childrens}


def write_dinginfo(writepath_: str, loadJson_: bool,  writeExcel_: bool = False, writeJson_: bool = True, date_: str = None, *args, **kwargs):
    '''
        读取钉钉通讯录的钉钉用户、钉钉部门，合并导出Excel表
        writepath_      : 当前写入的目录
        loadJson_       : True: 从Json文件导入钉钉通讯录数据, False: 调用钉钉Api获取钉钉通讯录数据
        writeExcel_     : True: 导出Excel, False: 不导出Excel, 只导出Json
        *args       : 不定参数集. list列表允许输入多个参数
        **kwargs    : 不定参数集. dict集合允许输入多个键值对


    SUCCESS
    return dict      : 成功返回 dict对象, Json格式

    ERROR
    return None      : 错误返回 None
    '''
    import json
    from copy import deepcopy

    def get_deptlist_by_dingtalk(dingApi: DingtalkOpenAPI):
        deptlist = dingApi.get_listsub_dept(dept_id=1)[0]['result']
        for dept in deptlist:
            ret_dinginfo = loop_get_dinginfo(dept, f"{dept['parent_id']},{dept['dept_id']}", dept['name'], dingApi)
            if ret_dinginfo is not None:
                dingtalk_source[dept['dept_id']] = ret_dinginfo
        JsonUtils.writeJson(os.path.join(json_path, 'dingtalk_source.json'), dingtalk_source, True) if writeJson_ else ''

    json_path = os.path.join(writepath_, 'json')
    if not os.path.exists(json_path):
        os.makedirs(json_path)
    dingtalk_source = {}

    if loadJson_:
        dingtalk_source = json.load(open(os.path.join(json_path, 'dingtalk_source.json'), 'r+', encoding='utf-8'))
    else:
        dingApi = DingtalkOpenAPI(app_key_=DD_CONST.APP_KEY, app_secret_=DD_CONST.APP_SECRET, g_api_uri_host_=DD_CONST.API_URI_HOST)
        get_deptlist_by_dingtalk(dingApi)

    ret_users = []
    ret_departments = []
    for item in dingtalk_source.values():
        loop_read_dinginfo(item, ret_users_=ret_users, ret_departments_=ret_departments, ownGroup_=item['name'])
    ret_departments = sorted(ret_departments, key=lambda x: (x['ownGroup'], x['name']), reverse=False)
    ret_departments = {item['id']: item for item in ret_departments}

    # dingtalk列表去重 -- 处理用户兼任多部门问题
    dingtalk_source_lists = deepcopy(ret_users)
    dd_jobNumber_fullList = [v['jobNumber'] for v in dingtalk_source_lists if v['jobNumber'] != '']
    dd_name_fullList = [v['name'] for v in dingtalk_source_lists if v['jobNumber'] != '']
    dingtalk_uniqueList = sorted(
        [v for v in dingtalk_source_lists if dd_jobNumber_fullList.count(v['jobNumber']) == 1] + [v for v in dingtalk_source_lists if dd_jobNumber_fullList.count(v['jobNumber']) > 1 and dd_name_fullList.count(v['name']) == 1],
        key=lambda x: (x['ownGroup'], x['name'], x['jobNumber']),
        reverse=False,
    )
    dingtalk_repeatList = sorted([v for v in dingtalk_source_lists if dd_jobNumber_fullList.count(v['jobNumber']) > 1 and dd_name_fullList.count(v['name']) > 1], key=lambda x: (x['ownGroup'], x['name'], x['jobNumber']), reverse=False)
    dingtalk_unique_in_repeatList = sorted(list(set([v['jobNumber'] for v in dingtalk_repeatList])))
    for item in dingtalk_unique_in_repeatList:
        uniqueItem = {}
        for repeatItem in [v for v in dingtalk_repeatList if v['jobNumber'] == item]:
            if uniqueItem:
                uniqueItem = {k: (('%s\n%s' % (uniqueItem[k], v)) if uniqueItem[k] != v else uniqueItem[k]) for k, v in repeatItem.items()}
            else:
                uniqueItem = {k: v for k, v in repeatItem.items()}
        uniqueItem['unique'] = True
        uniqueItem['topGroup'] = '\n'.join(sorted(list(set(uniqueItem['topGroup'].split('\n')))))
        dingtalk_repeatList[[v['jobNumber'] for v in dingtalk_repeatList].index(item)] = uniqueItem
    dingtalk_repeatList = [v for v in dingtalk_repeatList if 'unique' in v.keys()]
    dingtalk_unique_list = sorted(dingtalk_repeatList + dingtalk_uniqueList, key=lambda x: (x['ownGroup'], x['name'], x['jobNumber']), reverse=False)
    for i in range(len(dingtalk_unique_list)):
        dingtalk_unique_list[i]['uid'] = i + 1
    dingtalk_unique_dict = {v['id']: v for v in dingtalk_unique_list}

    JsonUtils.writeJson(os.path.join(json_path, 'dingtalk_user_normal.json'), ret_users, True) if writeJson_ else ''
    JsonUtils.writeJson(os.path.join(json_path, 'dingtalk_user_unique.json'), dingtalk_unique_list, True) if writeJson_ else ''
    JsonUtils.writeJson(os.path.join(json_path, 'dingtalk_department.json'), ret_departments, True) if writeJson_ else ''

    if writeExcel_:
        from openpyxl import Workbook
        from components.ExcelHandle import ExcelStaticMethods
        import datetime

        xlsx_path = os.path.join(writepath_, 'xlsx')
        if not os.path.exists(xlsx_path):
            os.makedirs(xlsx_path)
        write_dingdata_excel_fn = os.path.join(xlsx_path, '钉钉通讯录_unique_%s.xlsx' % int(datetime.datetime.now().strftime('%Y%m%d%H%M%S')))
        wb = Workbook()
        ExcelStaticMethods.writeExcel(dingtalk_unique_list, write_dingdata_excel_fn, wb, 1)
        wb.save(write_dingdata_excel_fn)

    return ret_departments, dingtalk_unique_dict


if __name__ == '__main__':
    import os
    import time

    _sT = time.time()
    ROOTPATH = r'/www/wwwroot/dingtalk-book2excel/'
    write_path = os.path.join(ROOTPATH, 'assets')
    if not os.path.exists(write_path):
        os.makedirs(write_path)

    write_dinginfo(writepath_=write_path, loadJson_=False, writeExcel_=True)

    print(f"程序运行时间: {(time.time() - _sT)} s")
else:
    from . import *
