from xlrd import *
from xlwt import *
import numpy as np
from xlutils.copy import copy

global NUM_TYPE_WEAPON
global NUM_TYPE_ATTACK
NUM_TYPE_WEAPON = 13  # 武器信息种类数量
NUM_TYPE_ATTACK = 9  # 攻击信息种类数量


# 用于打开文件，保存文件实例
class Data:

    def __init__(self, file_path):
        self.data = None
        self.sheet = None
        try:
            self.data = open_workbook(file_path)
        except Exception:
            print("打开文件出错")
            quit()
        else:
            print("打开文件: " + file_path)
            self.sheet = self.data.sheet_by_index(0)
        self.first_col_list = self.sheet.row_values(0)
        self.nrow_data = self.sheet.nrows


    # 数据填充
    def initial_data(self):
        # 填充经济损失为0情况下的经济损失等级
        ncol_property = self.first_col_list.index('property')
        ncol_propextent = self.first_col_list.index('propextent')
        num = 0
        copy_book = copy(self.data)
        copy_first_sheet = copy_book.get_sheet(0)
        print("正在处理'propextent'行")
        for i in range(self.nrow_data):
            if self.sheet.cell_value(i, ncol_property) == 0:
                copy_first_sheet.write(i, ncol_propextent, 0)
                num = num + 1
        print(str(num) + "行设置为0")
        # 根据欧几里得距离计算相似度，匹配填充nkill,nwound,propextent



        copy_book.save(r'./competition topic/测试数据new.xlsx')

    # 获取指定表名的数据列（优先选取第一列）
    def get_data_list(self, col_name):
        m = self.first_col_list.index(col_name)
        return list(self.sheet.col_values(m, start_rowx=1, end_rowx=None))

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
