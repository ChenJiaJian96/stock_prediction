from xlrd import *


# 用于打开文件，保存文件实例
class Data:
    def __init__(self, file_path):
        self.data = None
        self.table = None
        try:
            self.data = open_workbook(file_path)
        except XLRDError:
            print("打开文件格式错误")
        else:
            # 打开工作表'Data'
            self.table = self.data.sheet_by_index(0)
        self.first_col_list = self.table.row_values(0)

    # 获取指定表名的数据列（优先选取第一列）
    def get_data_list(self, col_name):
        m = self.first_col_list.index(col_name)
        return list(self.table.col_value(m, start_rowx=1, end_rowx=None))

    # 统计各攻击信息和武器信息综合的人员伤亡情况
    def get_casualty_by_attack_and_weapon(self):
        casualty_dict = {}

