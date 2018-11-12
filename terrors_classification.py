from Data import *


def init_data():
    print("System started.")
    file_path = r'./competition topic/测试数据.xlsx'
    local_data = Data(file_path)
    return local_data


data = init_data()
data.initial_data()
data.get_casualty_by_attack_and_weapon()
