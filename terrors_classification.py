from Data import *


def init_data():
    print("System started.")
    file_path = r'./competition topic/测试数据3.xlsx'
    local_data = Data(file_path)
    return local_data


data = init_data()  # 创建实例
data.initial_data()  # 填充数据
