from Data import *


def init_data():
    print("System started.")
    file_path = 'F:/PycharmProjects/stock_prediction/competition topic/测试数据.xlsx'
    local_data = Data(file_path)
    return local_data


data = init_data()
data.initial_data()