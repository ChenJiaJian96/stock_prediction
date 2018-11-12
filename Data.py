from xlrd import *
from xlwt import *
import numpy as np
from xlutils.copy import copy
from math import pow, sqrt

global NUM_TYPE_WEAPON
global NUM_TYPE_ATTACK
NUM_TYPE_WEAPON = 13  # 武器信息种类数量
NUM_TYPE_ATTACK = 9  # 攻击信息种类数量


# 用于打开文件，保存文件实例
class Data:

    def __init__(self, file_path):
        self.data = None
        self.sheet = None
        print("正在打开" + file_path)
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
        # 对用列序号
        self.ncol_iyear = self.first_col_list.index('iyear')
        self.ncol_region = self.first_col_list.index('region')
        self.ncol_attacktype1 = self.first_col_list.index('attacktype1')
        self.ncol_attacktype2 = self.first_col_list.index('attacktype2')
        self.ncol_attacktype3 = self.first_col_list.index('attacktype3')
        self.ncol_weaptype1 = self.first_col_list.index('weaptype1')
        self.ncol_weaptype2 = self.first_col_list.index('weaptype2')
        self.ncol_weaptype3 = self.first_col_list.index('weaptype3')
        self.ncol_targtype1 = self.first_col_list.index('targtype1')
        self.ncol_targtype2 = self.first_col_list.index('targtype2')
        self.ncol_targtype3 = self.first_col_list.index('targtype3')
        self.ncol_nkill = self.first_col_list.index('nkill')
        self.ncol_nwound = self.first_col_list.index('nwound')
        self.ncol_property = self.first_col_list.index('property')
        self.ncol_propextent = self.first_col_list.index('propextent')

    # 数据填充
    def initial_data(self):
        # 填充经济损失为0情况下的经济损失等级
        copy_book = copy(self.data)
        first_sheet = copy_book.get_sheet(0)
        for i in range(self.nrow_data):
            if self.sheet.cell_value(i, self.ncol_property) == 0:
                first_sheet.write(i, self.ncol_propextent, 0)
            if self.sheet.cell_value(i, self.ncol_attacktype2) == '':
                first_sheet.write(i, self.ncol_attacktype2, 0)
            if self.sheet.cell_value(i, self.ncol_attacktype3) == '':
                first_sheet.write(i, self.ncol_attacktype3, 0)
            if self.sheet.cell_value(i, self.ncol_weaptype2) == '':
                first_sheet.write(i, self.ncol_weaptype2, 0)
            if self.sheet.cell_value(i, self.ncol_weaptype3) == '':
                first_sheet.write(i, self.ncol_weaptype3, 0)
            if self.sheet.cell_value(i, self.ncol_targtype2) == '':
                first_sheet.write(i, self.ncol_targtype2, 0)
            if self.sheet.cell_value(i, self.ncol_targtype3) == '':
                first_sheet.write(i, self.ncol_targtype3, 0)
        copy_book.save(r'./competition topic/proceed.xlsx')
        self.data = open_workbook(r'./competition topic/proceed.xlsx')
        self.sheet = self.data.sheet_by_index(0)
        # 根据欧几里得距离计算相似度，匹配填充nkill,nwound,propextent
        # 数据二维列表，n行12列，计算保存相似度
        # [nrow,nyear,region,attacktype1,2,3,weapontype1,2,3,targtype1,2,3]
        data_list = []
        # 需要计算的行号
        num_need_cal_list = []
        # 从first_sheet中提取符合条件的数据保存至data_matrix中（三列均非空）
        for row in range(1, self.nrow_data):
            if self.sheet.cell_value(row, self.ncol_nkill) != '' \
                    and self.sheet.cell_value(row, self.ncol_nwound) != '' \
                    and self.sheet.cell_value(row, self.ncol_property) != -9:
                # 插入到data_matrix中
                insert_list = self.ret_datalist_cal_simi(row)
                data_list.append(insert_list)
            else:
                # 将序号插入到num_need_cal_list中
                num_need_cal_list.append(row)
        copy_book = copy(self.data)
        first_sheet = copy_book.get_sheet(0)
        # 开始计算
        for i in num_need_cal_list:
            print("正在计算第" + str(i) + "行数据")
            # 临时列表，用于保存当前处理的数据 nrow:原来xls中是第nrow行数据
            # [nrow,nyear,region,attacktype1,2,3,weapontype1,2,3,targtype1,2,3]
            temp_list = self.ret_datalist_cal_simi(i)
            # 相关度列表
            similarity_list = []
            for j in range(len(data_list)):
                similarity_list.append(self.ret_similarity(temp_list, data_list[j]))
            min_similarity = min(similarity_list)
            min_pos = data_list[similarity_list.index(min_similarity)][0]
            # 若i行数据为空，则填入min_pos行数据
            if self.sheet.cell_value(i, self.ncol_nkill) == '':
                kill_data = self.sheet.cell_value(min_pos, self.ncol_nkill)
                first_sheet.write(i, self.ncol_nkill, kill_data)
            if self.sheet.cell_value(i, self.ncol_nwound) == '':
                wound_data = self.sheet.cell_value(min_pos, self.ncol_nwound)
                first_sheet.write(i, self.ncol_nwound, wound_data)
            if self.sheet.cell_value(i, self.ncol_propextent) == '':
                prop_data = self.sheet.cell_value(min_pos, self.ncol_propextent)
                first_sheet.write(i, self.ncol_propextent, prop_data)
        copy_book.save(r'./competition topic/proceed.xlsx')

    # 返回填充数据时需要计算相似度的数据
    def ret_datalist_cal_simi(self, row):
        return [row, self.sheet.cell_value(row, self.ncol_iyear), self.sheet.cell_value(row, self.ncol_region),
                self.sheet.cell_value(row, self.ncol_attacktype1),
                self.sheet.cell_value(row, self.ncol_attacktype2),
                self.sheet.cell_value(row, self.ncol_attacktype3), self.sheet.cell_value(row, self.ncol_weaptype1),
                self.sheet.cell_value(row, self.ncol_weaptype2), self.sheet.cell_value(row, self.ncol_weaptype3),
                self.sheet.cell_value(row, self.ncol_targtype1), self.sheet.cell_value(row, self.ncol_targtype2),
                self.sheet.cell_value(row, self.ncol_targtype3)]

    # 返回两个列表之间的相似度
    def ret_similarity(self, list1, list2):
        if len(list1) != len(list2):
            print("Error!!!!!!")
            return 100
        else:
            res = pow((list1[1] - list2[1]), 2) + pow((list1[2] - list2[2]), 2) + 0.7 * pow((list1[3] - list2[3]), 2) + 0.2 * pow(
                (list1[4] - list2[4]), 2) + 0.1 * pow((list1[5] - list2[5]), 2) + 0.7 * pow((list1[6] - list2[6]), 2) + 0.2 * pow(
                (list1[7] - list2[7]), 2) + 0.1 * pow((list1[8] - list2[8]), 2) + 0.7 * pow((list1[9] - list2[9]), 2) + 0.2 * pow(
                (list1[10] - list2[10]), 2) + 0.1 * pow((list1[11] - list2[11]), 2)
            return sqrt(res)

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
