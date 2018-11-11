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
        ncol_iyear = self.first_col_list.index('iyear')
        ncol_region = self.first_col_list.index('region')
        ncol_attacktype1 = self.first_col_list.index('attacktype1')
        ncol_attacktype2 = self.first_col_list.index('attacktype2')
        ncol_attacktype3 = self.first_col_list.index('attacktype3')
        ncol_weaptype1 = self.first_col_list.index('weaptype1')
        ncol_weaptype2 = self.first_col_list.index('weaptype2')
        ncol_weaptype3 = self.first_col_list.index('weaptype3')
        ncol_targtype1 = self.first_col_list.index('targtype1')
        ncol_targtype2 = self.first_col_list.index('targtype2')
        ncol_targtype3 = self.first_col_list.index('targtype3')
        ncol_nkill = self.first_col_list.index('nkill')
        ncol_nwound = self.first_col_list.index('nwound')
        ncol_property = self.first_col_list.index('property')
        ncol_propextent = self.first_col_list.index('propextent')
        num = 0
        copy_book = copy(self.data)
        first_sheet = copy_book.get_sheet(0)
        print("正在处理'propextent'行")
        for i in range(self.nrow_data):
            if self.sheet.cell_value(i, ncol_property) == 0:
                first_sheet.write(i, ncol_propextent, 0)
                num = num + 1
        print("propextent" + str(num) + "行设置为0")
        # 根据欧几里得距离计算相似度，匹配填充nkill,nwound,propextent
        # 数据二维列表，n行13列，计算保存相似度
        # [nyear,region,attacktype1,2,3,weapontype1,2,3,targtype1,2,3,nkill,nwound,propextent,similarity]
        data_matrix = np.empty(shape=[12], dtype=int, order='C')
        # 临时二维列表，用于保存当前处理的数据 nrow:原来xls中是第nrow行数据
        # [nrow,nyear,region,attacktype1,2,3,weapontype1,2,3,targtype1,2,3,nkill,nwound,propextent]
        temp_matrix = np.zeros([1, 15], dtype=int, order='C')
        # 需要计算的行号
        num_need_cal_list = []
        # 从first_sheet中提取符合条件的数据保存至data_matrix中（三列均非空）
        for row in range(self.nrow_data):
            if self.sheet.cell_value(row, ncol_nkill) != '' \
                    and self.sheet.cell_value(row, ncol_nwound) != '' \
                    and self.sheet.cell_value(row, ncol_property) != -9:
                # 插入到data_matrix中
                insert_list = [row, self.sheet.cell_value(row, ncol_iyear), self.sheet.cell_value(row, ncol_region),
                               self.sheet.cell_value(row, ncol_attacktype1),
                               self.sheet.cell_value(row, ncol_attacktype2),
                               self.sheet.cell_value(row, ncol_attacktype3), self.sheet.cell_value(row, ncol_weaptype1),
                               self.sheet.cell_value(row, ncol_weaptype2), self.sheet.cell_value(row, ncol_weaptype3),
                               self.sheet.cell_value(row, ncol_targtype1), self.sheet.cell_value(row, ncol_targtype2),
                               self.sheet.cell_value(row, ncol_targtype3), 0]

                pass
            else:
                # 将序号插入到num_need_cal_list中
                num_need_cal_list.append(row)
        print(num_need_cal_list)

        # copy_book.save(r'./competition topic/测试数据new.xlsx')

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
