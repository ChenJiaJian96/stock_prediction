import baostock as bs
import pandas as pd

#### 登陆系统 ####
lg = bs.login()
# 显示登陆返回信息
print('login respond error_code:' + lg.error_code)
print('login respond  error_msg:' + lg.error_msg)

#### 获取沪深A股历史K线数据 ####
# 详细指标参数，参见“历史行情指标参数”章节
code_list = ["000557", "600108", "900919", "600896", "600897", "600642", "600719", "600769", "600863", "600649",
             "000610", "000888", "600749", "600023", "000663", "600692", "900938", "600692", "900938", "600098",
             "600644", "600726", "600780", "600864", "000428", "000613", "600054", "600929", "000415", "000735",
             "600087", "600798", "600011", "600116", "600674", "600744", "600795", "600886", "000430", "000802",
             "600138", "600942", "600068", "600758", "000002", "000031", "000548", "000609", "000667", "000838",
             "600684", "600743", "600766", "600807", "600895"]
for code in code_list:
    rs = bs.query_history_k_data(code,
                                 "date,code,open,high,low,close,preclose,volume,amount,adjustflag,turn,tradestatus,pctChg,isST",
                                 start_date='2009-06-24', end_date='2009-07-05',
                                 frequency="d", adjustflag="3")
    print('query_history_k_data respond error_code:' + rs.error_code)
    print('query_history_k_data respond  error_msg:' + rs.error_msg)
    data_list = []
    while (rs.error_code == '0') & rs.next():
        # 获取一条记录，将记录合并在一起
        data_list.append(rs.get_row_data())

    result = pd.DataFrame(data_list, columns=rs.fields)

#### 结果集输出到csv文件 ####
result.to_excel("C:\\Users\\Administrator.DESKTOP-GAKELJI\\Desktop\\data.xlsx", index=False)
print(result)

#### 登出系统 ####
bs.logout()
