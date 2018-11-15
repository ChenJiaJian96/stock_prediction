from openpyxl import *
import baostock as bs
import pandas as pd
from datetime import *
from time import mktime, strptime

# 区间列表
section_list = [[0, 10], [5, 6], [5, 7], [5, 8], [5, 9], [5, 10]]
# 股票代码列表
code_list = ["600108", "600896", "600897", "600642", "600719", "600769", "600863", "600649", "600749", "600692",
             "600098", "600644", "600726", "600780", "600864", "600054", "600798", "600011", "600116", "600674",
             "600744", "600795", "600886", "600138", "600068", "600758", "000002", "000031", "600684", "600743",
             "600766", "600807", "600895", "600419", "600489", "600610", "600750", "600810", "600819", "600830",
             "600889", "600321", "600051", "600531"]
date_list = ['2009-06-29', '2011-07-23', '2014-08-02', '2015-08-12']


# data = load_workbook(r'./competition topic/事故灾害统计表.xlsx')
# sheet = data.worksheets[0]
# ncol_data = sheet.max_column
# nrow_data = sheet.max_row
# num_affair = nrow_data - 1  # 表明要计算的事件数
# first_row_list = []
# for i in range(1, ncol_data + 1):
# first_row_list.append(sheet.cell(1, i).value)
# ncol_year = first_row_list.index("iyear") + 1
# ncol_month = first_row_list.index("imonth") + 1
# ncol_day = first_row_list.index("iday") + 1


# 初始化股票数据
def get_stock_data():
    # 登陆系统
    lg = bs.login()
    # 显示登陆返回信息
    print('login respond error_code:' + lg.error_code)
    print('login respond  error_msg:' + lg.error_msg)

    # 获取沪深A股历史K线数据
    # 详细指标参数，参见“历史行情指标参数”章节

    # date	交易所行情日期
    # code	证券代码
    # open	开盘价
    # high	最高价
    # low	最低价
    # close	收盘价
    # preclose	昨日收盘价
    for code in code_list:
        code = "sh." + code
        rs = bs.query_history_k_data(code, "date,code,open,high,low,close,preclose", start_date='2009-06-19',
                                     end_date='2009-07-10')
        print('query_history_k_data respond error_code:' + rs.error_code)
        print('query_history_k_data respond  error_msg:' + rs.error_msg)
        data_list = []
        while (rs.error_code == '0') & rs.next():
            data_list.append(rs.get_row_data())
        rs = bs.query_history_k_data(code, "date,code,open,high,low,close,preclose", start_date='2011-07-13',
                                     end_date='2011-08-02')
        print('query_history_k_data respond error_code:' + rs.error_code)
        print('query_history_k_data respond  error_msg:' + rs.error_msg)
        while (rs.error_code == '0') & rs.next():
            # 获取一条记录，将记录合并在一起
            data_list.append(rs.get_row_data())
        rs = bs.query_history_k_data(code, "date,code,open,high,low,close,preclose", start_date='2014-07-23',
                                     end_date='2014-08-12')
        print('query_history_k_data respond error_code:' + rs.error_code)
        print('query_history_k_data respond  error_msg:' + rs.error_msg)
        while (rs.error_code == '0') & rs.next():
            # 获取一条记录，将记录合并在一起
            data_list.append(rs.get_row_data())
        rs = bs.query_history_k_data(code, "date,code,open,high,low,close,preclose", start_date='2015-08-02',
                                     end_date='2015-08-22')
        print('query_history_k_data respond error_code:' + rs.error_code)
        print('query_history_k_data respond  error_msg:' + rs.error_msg)
        while (rs.error_code == '0') & rs.next():
            # 获取一条记录，将记录合并在一起
            data_list.append(rs.get_row_data())

        result = pd.DataFrame(data_list, columns=rs.fields)
        result.to_excel(r'./stock_data/' + str(code) + '.xlsx', index=False)

    # 登出系统
    bs.logout()


# 计算正常收益模型的平均日收益率
def cal_R(code, cur_date):
    cur_path = stock_path(code)
    try:
        cur_book = load_workbook(cur_path)
    except Exception:
        return 0
    else:
        cur_sheet = cur_book.worksheets[0]
        cur_nrow = cur_sheet.max_row
        cur_ncol = cur_sheet.max_column
        cur_first_row_list = []
        for i in range(1, cur_ncol + 1):
            cur_first_row_list.append(cur_sheet.cell(1, i).value)
        ncol_date = cur_first_row_list.index("date") + 1
        ncol_open = cur_first_row_list.index("open") + 1
        ncol_close = cur_first_row_list.index("close") + 1
        ncol_preclose = cur_first_row_list.index("preclose") + 1
        num = 0
        total = 0
        cur_date = strptime(str(cur_date), '%Y-%m-%d')
        for i in range(2, cur_nrow + 1):
            stock_date = strptime(str(cur_sheet.cell(i, ncol_date).value), '%Y-%m-%d')
            # 事情发生的前五天内
            if 0 < int(mktime(cur_date) - mktime(stock_date)) < 432001:
                num += 1
                total += (float(cur_sheet.cell(i, ncol_close).value) - float(
                    cur_sheet.cell(i, ncol_open).value)) / float(
                    cur_sheet.cell(i, ncol_preclose).value)
        return total / num


# 计算每个股票每个日期的正常平均日收益率
def get_matrix_R():
    result_matrix = {}
    for code in code_list:
        for cur_date in date_list:
            result = cal_R(code, cur_date)
            if result != 0:
                result_matrix[code, cur_date] = result
    return result_matrix


def stock_path(str):
    return r'./stock_data/sh.' + str + '.xlsx'


# get_stock_data()
print(get_matrix_R())
