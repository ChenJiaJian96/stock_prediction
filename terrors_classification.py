from Data import *


def init_data():
    print("System started.")
    file_path = r'./competition topic/处理后数据.xlsx'
    local_data = Data(file_path)
    return local_data


data = init_data()  # 创建实例
# data.initial_data()  # 填充数据
# data.get_matrix_attacktype_score()
# data.get_matrix_weapontype_score()
# data.insert_score()
# data.cal_total_F()
print("finish!!!")
