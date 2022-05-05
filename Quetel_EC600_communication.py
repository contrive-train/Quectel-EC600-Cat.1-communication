# -*- coding: utf-8 -*-
# @Time    : 2022/3/109:52
# @Author  : 周智勇
# @FileName:Quetel_EC600_communication.py
# @Software:PyCharm

from matplotlib import pyplot as plt
import pylab as mpl
import numpy as np
import xlrd

plt.rcParams['font.family'] = ['sans-serif']
mpl.rcParams['axes.unicode_minus'] = False

excel = xlrd.open_workbook(r'C:\Users\user\Desktop\Quectel_ec600.xls')
# x = [i for i in range(1, 42)]
x = np.arange(1, 42)  # 随着日期的改变，只需要更改参数42的值

plt.figure(figsize=(20, 20), dpi=80)

theoratical = [48 for i in range(1, 42)]  # 随着日期的改变，只需要更改参数42的值
print(theoratical)  # 三个表的理论通讯次数均为48，采集时间范围 08:00~07:30 (+1天),上传周期是30min


# 打开表格文件中的第一张表格，索引从0开始
sheet1 = excel.sheets()[0]
ncols = sheet1.ncols  # 获取第一张表格的列数赋值给ncols
for i in range(ncols):  # 用一个for循环遍历所有的lie数
    print(sheet1.col_values(i))  # 打印所有遍历到的列数的内容
print(sheet1.col_values(i)[::-1])

sheet2 = excel.sheets()[1]  # 打开表格文件中的第二张表格，索引为1
ncols = sheet2.ncols  # 获取第一张表格的列数赋值给ncols
for i in range(ncols):  # 用一个for循环遍历所有的列数
    print(sheet2.col_values(i))  # 打印所有遍历到的行数的内容
print(sheet2.col_values(i)[::-1])
# sheet2_reverse=sheet2.col_values(i).reverse()

sheet3 = excel.sheets()[2]  # 打开表格文件中的第三张表格，索引为2
ncols = sheet3.ncols  # 获取第一张表格的列数赋值给ncols
for i in range(ncols):  # 用一个for循环遍历所有的列数
    print(sheet3.col_values(i))  # 打印所有遍历到的行数的内容
print(sheet3.col_values(i)[::-1])
# # sheet3_reverse=sheet3.col_values(i).reverse()

x1 = range(29, 32, 1)
x2 = range(32, 60, 1)
x3 = range(60, 70, 1)  # 随着日期的改变，只需要更改参数70的值

_x = list(x1) + list(x2) + list(x3)
print(_x)

ax1 = plt.subplot(2, 1, 1)
ax1.set_ylabel(_x, fontsize=15)
ax1.set_title(_x, fontsize=15)

plt.plot(_x[::1], theoratical, label='理论通讯次数', linestyle='-', color='cyan')  # 8A2BE2
plt.plot(_x, excel.sheets()[0].col_values(i)[::-1], label='EDLS00010实际通讯次数', linestyle='--', color='g')
plt.plot(_x, excel.sheets()[1].col_values(i)[::-1], label='EDLS00013实际通讯次数', linestyle='-.', color='r')
plt.plot(_x, excel.sheets()[2].col_values(i)[::-1], label='EDLS00015实际通讯次数', linestyle=':', color='b')
plt.ylim(-2,50)
plt.legend(loc='lower right')

_xticks = ['1月{}日'.format(i) for i in x1]
_xticks += ['2月{}日'.format(j - 31) for j in x2]
_xticks += ['3月{}日'.format(k - 59) for k in x3]
plt.xticks(_x, _xticks, rotation=45, fontsize=15)
plt.yticks(range(0, 49, 4), fontsize=15)

plt.ylabel('通讯次数')
plt.title('移远EC600 CAT.1模组通讯状况图(采样周期5s,上传周期30min)', fontdict={'weight': 'normal', 'size': 15})
ax1.legend(loc='lower right', fontsize=15)
plt.grid(alpha=0.5)

ax2 = plt.subplot(2, 1, 2)
# 绘制误差率的图
edls00010_rate = excel.sheets()[0].col_values(i)[::-1]
edls00013_rate = excel.sheets()[1].col_values(i)[::-1]
edls00015_rate = excel.sheets()[2].col_values(i)[::-1]
plt.plot(_x, [(48 - i) / 48 for i in edls00010_rate], label='EDLS00010误差率', linestyle='--', color='g')
plt.plot(_x, [(48 - j) / 48 for j in edls00013_rate], label='EDLS00013误差率', linestyle='-.', color='r')
plt.plot(_x, [(48 - k) / 48 for k in edls00015_rate], label='EDLS00015误差率', linestyle=':', color='b')
plt.legend(loc='upper right')
_xticks = ['1月{}日'.format(i) for i in x1]
_xticks += ['2月{}日'.format(j - 31) for j in x2]
_xticks += ['3月{}日'.format(k - 59) for k in x3]
plt.xticks(_x, _xticks, color='black', rotation=45)  # xticks([1,2,3,4,5,6,7,8,9....,九个刻度], [labels，每个刻度下的标签，可用字符串格式在列表表示], **kwargs：用于设置标签字体倾斜度和颜色等外观属性)
# plt.yticks(range(0, 1))
ax2.set_xlabel(_x, fontsize=15)
ax2.set_ylabel(_x, fontsize=15)
ax2.legend(fontsize=15)

plt.xlabel('日期')
plt.xticks(fontsize=15)
plt.yticks(fontsize=15)
plt.ylabel('误差率(*100%)')

plt.grid(alpha=0.7)
mpl.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体
plt.show()
