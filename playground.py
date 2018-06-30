# coding = utf-8
import functions

# settings
file_name = '组合构成.xlsx'  # 文件名
sheet_name = 'Sheet1'       # 表名
col_par_code = 'F'          # 列号 - 父项物料编码
col_chi_seq = 'I'           # 列号 - 项次
col_chi_code = 'J'          # 列号 - 子项物料编码
col_do_mol = 'M'            # 列号 - 用量分子
col_do_den = 'N'            # 列号 - 用量分母

# initialization
# nodes = []                  # 总节点序列
re_list = []                  # 重新处理序列

# Open file & sheet
sheet = functions.openExcelSheet(file_name, sheet_name)
nodes = functions.importContents(sheet, col_par_code, col_chi_seq, col_chi_code, col_do_mol, col_do_den)

# for i in range(2, len(nodes)):
#     for j in range(0, len(nodes[i]['child_nodes'])):
#         # print(nodes[i]['child_nodes'][j]['code'])
#         for k in range(2, len(nodes)): 
#             if nodes[i]['child_nodes'][j]['code'] == nodes[k]['code']:
#                 nodes[i]['child_nodes'][j]['child_nodes'] = nodes[k]['child_nodes']

a = []
for i in range(2, len(nodes)):
    for child_node in nodes[i]['child_nodes']:
        for node2 in nodes:
            if child_node['code'] == node2['code']:
                child_node['child_nodes'] = node2['child_nodes']
                re_list.append(i)
                a.append(nodes[i]['code'])

print(a)


# print(nodes[1])