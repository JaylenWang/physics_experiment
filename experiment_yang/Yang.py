import xlrd
import math

# 获取数据
wordbook = xlrd.open_workbook("C:\\Users\\ADMIN\\Desktop\\data.xls")
sheet1_object = wordbook.sheet_by_index(0)
L = sheet1_object.cell_value(rowx = 0, colx = 0) / 1000
H = sheet1_object.cell_value(rowx = 0, colx = 1) / 1000
D = sheet1_object.cell_value(rowx = 0, colx = 2) / 1000
d_line = sheet1_object.col_values(3, 0, sheet1_object.nrows)
d = (d_line[0] + d_line[1]) / 1000
x_plus = sheet1_object.col_values(4, 0, sheet1_object.nrows)
x_minus = sheet1_object.col_values(5, 0, sheet1_object.nrows)
x_average = sheet1_object.col_values(6, 0, sheet1_object.nrows)

# 逐差法求delta_x / m
delta_x_div_m = 0
for i in range(5):
    delta_x_div_m += x_average[i + 5] - x_average[i]
delta_x_div_m /= 25 * 1000
print("delta_x_div_m = {:e}".format(delta_x_div_m))

# 求解E
E_average = 8 * 10 * L * H / (math.pi * d * d * D * delta_x_div_m)
print("E_average = {:e}".format(E_average))

# 求解不确定度

u_m = (1 / 4.5 * 0.005) ** 2
u_L = (1 / L * 8e-4) ** 2
u_H = (1 / H * 8e-4) ** 2
u_d = (-2 / d * 4e-6) ** 2
u_D = (-1 / D * 2e-5) ** 2
e = math.sqrt(u_m + u_L + u_H + u_d + u_D)
print("e = %.2f%%" % (e * 100))
print("U = %.3fx10^11" % (e * E_average * 1e-11))
