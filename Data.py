from xlrd import *
import numpy as np

global NUM_TYPE_WEAPON
global NUM_TYPE_ATTACK
NUM_TYPE_WEAPON = 13  # 武器信息种类数量
NUM_TYPE_ATTACK = 9  # 攻击信息种类数量


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
            print("打开文件: " + file_path)
            self.table = self.data.sheet_by_index(0)
        self.first_col_list = self.table.row_values(0)
        # TODO: 数据清洗、填充

    # 获取指定表名的数据列（优先选取第一列）
    def get_data_list(self, col_name):
        m = self.first_col_list.index(col_name)
        return list(self.table.col_values(m, start_rowx=1, end_rowx=None))

    # 统计各攻击信息和武器信息综合的人员伤亡情况
    # 返回伤亡情况数组： 2 * 武器信息种类 * 攻击信息种类
    def get_casualty_by_attack_and_weapon(self):
        death_list = self.get_data_list('nkill')
        print("death_list length: " + str(len(death_list)))
        wound_list = self.get_data_list('nwound')
        print("wound_list length: " + str(len(wound_list)))
        casualty_array = np.zeros((2, NUM_TYPE_WEAPON + 1, NUM_TYPE_ATTACK + 1))
        for j in range(3):
            attack_list = self.get_data_list('attacktype' + str(j + 1))
            print("attack_list length: " + str(len(attack_list)))
            weapon_list = self.get_data_list('weaptype' + str(j + 1))
            print("weapon_list length: " + str(len(weapon_list)))
            # TODO: 后期补充后此处两个list不应该存在空元素，可以删掉判断
            for i in range(len(death_list)):
                if attack_list[i] != '' and weapon_list[i] != '':
                    if death_list[i] != '':
                        casualty_array[0, int(weapon_list[i]), int(attack_list[i])] = \
                            casualty_array[0][int(weapon_list[i])][int(attack_list[i])] + int(death_list[i])
                    if wound_list[i] != '':
                        casualty_array[1, int(weapon_list[i]), int(attack_list[i])] = \
                            casualty_array[1][int(weapon_list[i])][int(attack_list[i])] + int(wound_list[i])
            attack_list.clear()
            weapon_list.clear()
        return casualty_array
