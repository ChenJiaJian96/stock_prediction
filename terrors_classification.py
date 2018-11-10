from Data import *


def init_data():
    print("System started.")
    file_path = 'F:/PycharmProjects/stock_prediction/competition topic/测试数据.xlsx'
    local_data = Data(file_path)
    return local_data


data = init_data()
casualty_array = data.get_casualty_by_attack_and_weapon()
print(casualty_array[0].size())
print(casualty_array[1].size())
