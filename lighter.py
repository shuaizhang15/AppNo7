# coding = utf-8
# lighter.py

import time
from functions import *

def launch(options):

    # Settings
    # structure-file's settings
    if '.xlsx' in options[0, 1]:
        str_file_name = options[0, 1]           # 组合结构文件名
    else: str_file_name = options[0, 1] + '.xlsx'
    str_sheet_name = options[1, 1]              # 组合结构文件表名
    col_par_code = options[2, 1]                # 列号 - 父项物料编码
    col_chi_seq = options[3, 1]                 # 列号 - 项次
    col_chi_code = options[4, 1]                # 列号 - 子项物料编码
    col_do_mol = options[5, 1]                  # 列号 - 用量分子
    col_do_den = options[6, 1]                  # 列号 - 用量分母
    # price-file's settings
    if '.xlsx' in options[0, 3]:
        pri_file_name = options[0, 3]           # 价格文件名
    else: pri_file_name = options[0, 3] + '.xlsx'
    pri_sheet_name = options[1, 3]              # 价格文件表名
    col_code = options[2, 3]                    # 列号 - 物料编码
    col_price = options[3, 3]                   # 列号 - 价格
    col_price_tax = options[4, 3]               # 列号 - 含税价格
    # output settings
    if '.xlsx' in options[9, 1]:
        output_file_name = options[9, 1]        # 导出文件名
    else: output_file_name = options[9, 1] + '.xlsx'


    # Initialization
    result_dict = {}        # 结果字典 (key, value) = (code, sum of all children nodes' mol/den) 
    price_dict = {}         # 结果字典 (key, value) = (code, {price, price_tax}) 
    false_mat_set = set()


    print('即将开始读取文件...')
    # input()

    
    # Read files
    # read structure-file
    str_sheet = openExcelSheet(str_file_name, str_sheet_name)
    nodes = readSheet(str_sheet, col_par_code, col_chi_seq, col_chi_code, col_do_mol, col_do_den)
    # read price-file
    pri_sheet = openExcelSheet(pri_file_name, pri_sheet_name)


    print('即将开始处理文件内容...')
    # input()

    
    # Sort file contents
    sortStrNodes(nodes)
    # sort price and price-with-tax contents
    try:
        for i in range(0, pri_sheet.max_row):
            price_dict[pri_sheet[col_code][i].value] = {'price': pri_sheet[col_price][i].value, 'price_tax': pri_sheet[col_price_tax][i].value}
    except Exception as e:
        print('价格文件内容读取出错: '+str(e))


    print('即将开始生成节点内容...')
    # input()

   
    # Generate result dictionary & output
    # calculate sum value of materials & sort code dictionary
    for i in range(2, len(nodes)):
        end_price_dict = {'price': 0.0, 'price_tax': 0.0}
        result_dict[nodes[i]['code']] = calMatPrice(nodes[i]['child_nodes'], price_dict, end_price_dict, false_mat_set)

    print('即将开始写出文件...')
    # input()
    
    # Output results
    writeExcel(output_file_name, result_dict, nodes, false_mat_set)

    print('流程完毕，按回车结束')
    input()
    