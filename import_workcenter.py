# -*- encoding: utf-8 -*-

import xmlrpclib
import psycopg2.extras
import csv
import os
import xlrd
import time
import datetime

import sys
import xlrd
default_encoding = 'utf-8'
if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

dbname = 'indowood'
pwd = 'INIT@!@#'
oe_ip = 'localhost:9999'

sock_common = xmlrpclib.ServerProxy('http://' + oe_ip + '/xmlrpc/common')
uid = sock_common.login(dbname, 'admin', pwd)
sock = xmlrpclib.ServerProxy('http://' + oe_ip + '/xmlrpc/object', allow_none=True)

if __name__ == '__main__':
    xlsfile = 'wc.xlsx'
    book = xlrd.open_workbook(xlsfile)
    sh = book.sheet_by_index(0)
    _cell_values = sh._cell_values
    number = 0
    for rx in range(1, sh.nrows):
        print rx
        vals = {
            'name': str(_cell_values[rx][0]),
            # 'resource_type': str(_cell_values[rx][2]),
            'resource_type': 'material',
            'calendar_id': 7,
            'department_id': 7,
            'time_cycle': int(_cell_values[rx][7]),
            'capacity_per_cycle': int(_cell_values[rx][8]),
        }
        section = sock.execute(dbname, uid, pwd, 'init.mrp.section', 'search', [('name', 'like', str(_cell_values[rx][5]))])
        job = sock.execute(dbname, uid, pwd, 'init.mrp.job', 'search', [('name', 'like', str(_cell_values[rx][6]))])
        print section, job
        # employee_id = sock.execute(dbname, uid, pwd, 'hr.employee', 'search', [('internal_code', '=', code)])
        # if len(employee_id):
        #     date = datetime.datetime.strptime(str(_cell_values[rx][3]), '%d/%m/%Y')
        #     vals = {
        #         'id_issue_date': date.strftime('%Y-%m-%d'),
        #         'id_issue_place': str(_cell_values[rx][4]),
        #     }
        #     print code
        #     sock.execute(dbname, uid, pwd, 'hr.employee', 'write', employee_id, vals)
        vals.update({
            'section_id': section[0],
            'job_id': job[0],
        })
        sock.execute(dbname, uid, pwd, 'mrp.workcenter', 'create', vals)

#     
