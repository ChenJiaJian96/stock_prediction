from xlrd import *
import numpy as np


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
    def get_casualty_by_attack_and_weapon(self):
        death_list = self.get_data_list('nkill')
        print("death_list length: " + str(len(death_list)))
        wound_list = self.get_data_list('nwound')
        print("wound_list length: " + str(len(wound_list)))
        attack_list = self.get_data_list('attacktype1')
        print("attack1_list length: " + str(len(attack_list)))
        weapon_list = self.get_data_list('weaptype1')
        print("weapon1_list length: " + str(len(weapon_list)))
        # 伤亡情况数组为 2 * 武器信息种类 * 攻击信息种类
        casualty_array = np.zeros((2, len(weapon_list), len(attack_list)))
        # TODO: 后期补充后此处两个list不应该存在空元素，可以删掉判断
        for i in range(len(death_list)):
            if death_list[i] != '':
                casualty_array[0, int(attack_list[i]), int(weapon_list[i])] += int(death_list[i])
            if wound_list[i] != '':
                casualty_array[1, int(attack_list[i]), int(weapon_list[i])] += int(wound_list[i])
        attack_list.clear()
        weapon_list.clear()
        attack_list = self.get_data_list('attacktype2')
        print("attack2_list length: " + str(len(attack_list)))
        weapon_list = self.get_data_list('weaptype2')
        print("weapon2_list length: " + str(len(weapon_list)))
        for i in range(len(death_list)):
            if death_list[i] != '':
                if attack_list[i] != '':
                    if weapon_list[i] != '':
                        casualty_array[0, int(attack_list[i]), int(weapon_list[i])] += int(death_list[i])
            if wound_list[i] != '':
                casualty_array[1, int(attack_list[i]), int(weapon_list[i])] += int(wound_list[i])
        attack_list.clear()
        weapon_list.clear()
        attack_list = self.get_data_list('attacktype3')
        print("attack3_list length: " + str(len(attack_list)))
        weapon_list = self.get_data_list('weaptype3')
        print("weapon3_list length: " + str(len(weapon_list)))
        for i in range(len(death_list)):
            if death_list[i] != '':
                casualty_array[0, int(attack_list[i]), int(weapon_list[i])] += int(death_list[i])
            if wound_list[i] != '':
                casualty_array[1, int(attack_list[i]), int(weapon_list[i])] += int(wound_list[i])
        return casualty_array
