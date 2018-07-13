# coding = utf-8
# functions.py
import openpyxl
import xlsxwriter
import time
import os
import traceback


# Import an excel file by openpyxl package.
def openExcelSheet(file, sheet):
    try:
        # make an absolute dir
        file = os.getcwd() + '\\' + file
        data = openpyxl.load_workbook(filename = file)
        sheet_data = data[sheet]
        return sheet_data
    except Exception as e:
        print('打开文件出错: ' + str(e))
        print('按回车结束本程序')
        input()
        exit()



# Put sheet-contents into nodes[].
def readSheet(sheet,
            col_par_code,                       # 列号 - 父项物料编码
            col_chi_seq,                        # 列号 - 项次
            col_chi_code,                       # 列号 - 子项物料编码
            col_con_mol,                        # 列号 - 用量分子
            col_con_den,                        # 列号 - 用量分母
            col_mat_name,                       # 列号 - 物料名称
            col_mat_size                        # 列号 - 规格型号
            ):
    nodes = []                  # 总节点序列
    nodes_code = [0]            # 父节点编码、物料名称及规格型号
    nodes_mat_name = ['']       # 父节点物料名称
    nodes_mat_size = ['']       # 父节点规格型号
    node = {}                   # 父节点对象
    child_nodes = []            # 子节点序列
    child_node = {}             # 子节点对象

    par_code = sheet[col_par_code]
    chi_seq = sheet[col_chi_seq]
    chi_code = sheet[col_chi_code]
    con_mol = sheet[col_con_mol]
    con_den = sheet[col_con_den]
    mat_name = sheet[col_mat_name]
    mat_size = sheet[col_mat_size]
    try: 
        # store node
        for i in range(0, len(par_code)):
            # parent code exists
            if par_code[i].value:
                # generate a new node by storing the last --parent-- node & its child nodes
                node = {'code': nodes_code[len(nodes_code)-1],
                        'name': nodes_mat_name[len(nodes_mat_name)-1],
                        'size': nodes_mat_size[len(nodes_mat_size)-1],
                        'child_nodes': child_nodes}
                nodes.append(node)
                nodes_code.append(par_code[i].value)
                nodes_mat_name.append(mat_name[i].value)
                nodes_mat_size.append(mat_size[i].value)
                # initilize a new child node 
                child_nodes = []
                child_node = {'seq': chi_seq[i].value,
                            'code': chi_code[i].value,
                            'con_mol': con_mol[i].value,
                            'con_den': con_den[i].value}
                child_nodes.append(child_node)
            # parent code does not exist, keep appending child nodes
            else:
                child_node = {'seq': chi_seq[i].value,
                            'code': chi_code[i].value,
                            'con_mol': con_mol[i].value,
                            'con_den': con_den[i].value}
                child_nodes.append(child_node)
    except Exception as e:
        print('文件读取有误: ' + str(e))
        print('按回车结束本程序')
        input()
        exit()

    node = {'code': nodes_code[len(nodes_code)-1],
        'name': nodes_mat_name[len(nodes_mat_name)-1],
        'size': nodes_mat_size[len(nodes_mat_size)-1],
        'child_nodes': child_nodes}
    nodes.append(node)
    return nodes



# Search all children nodes of the root nodes in stucture-file
def sortStrNodes(nodes): 
    try:
        for i in range(2, len(nodes)):
            for child_node in nodes[i]['child_nodes']:
                # 第一层子节点code与根节点code对比
                for j in range(2, len(nodes)):
                    # 若相同，将根节点的子节点枝存入第一层子节点形成第二层子节点
                    if child_node['code'] == nodes[j]['code']:
                        child_node['child_nodes'] = nodes[j]['child_nodes']
    except Exception as e:
        print('结构文件内容读取有误: '+str(e))
        print('按回车结束本程序')
        input()
        exit()



# Set child node list in a node.
# -- no need --
# def setChildNodes(childNodes, codeDict):
#     for child_node in childNodes:
#         if child_node['code'] in codeDict.keys():
#             child_node['child_nodes'] = codeDict[child_node['code']]
#     return childNodes


# Validate if the input is non-None type & if the value is a number.
def valNum(s):
    try:
        if s:
            return float(s)
        else:
            return None
    except:
        return None

# Recursively calculate end price of each node.
def calMatPrice(node_list, price_dict, end_price_dict, false_mat_set):
    try:
        for node in node_list:
            # calculate the division of molecular and denominator
            mol = valNum(node['con_mol'])
            den = valNum(node['con_den'])
            if mol and den:
                if mol>0 and den>0:
                    k = mol / den
                else:
                    # validation failed, adding a big negative number to make notice
                    false_mat_set['con-non-posi'].add(node['code'])
                    end_price_dict['price'] += -100000000
                    end_price_dict['price_tax'] += -100000000
                    continue
            else:
                false_mat_set['con-non-num'].add(node['code'])
                end_price_dict['price'] += -100000000
                end_price_dict['price_tax'] += -100000000
                continue

            # calculate material price
            if 'child_nodes' not in node:
                if node['code'] in price_dict:
                    price = valNum(price_dict[node['code']]['price'])
                    price_tax = valNum(price_dict[node['code']]['price_tax'])
                    # price validation
                    if price:
                        if price>0:
                            end_price_dict['price'] += k * price
                        else:
                            false_mat_set['pri-non-posi'].add(node['code'])
                            end_price_dict['price'] += -100000000
                    else:
                        false_mat_set['pri-non-num'].add(node['code'])
                        end_price_dict['price'] += -100000000
                    # price_tax validation
                    if price_tax:
                        if price_tax>0:
                            end_price_dict['price_tax'] += k * price_tax
                        else:
                            false_mat_set['tpri-non-posi'].add(node['code'])
                            end_price_dict['price_tax'] += -100000000
                    else:
                        false_mat_set['tpri-non-num'].add(node['code'])
                        end_price_dict['price_tax'] += -100000000
                else:
                    false_mat_set['pri-none'].add(node['code'])
                    end_price_dict['price'] += -100000000
                    end_price_dict['price_tax'] += -100000000
            # if this node has children, recursively invoke this functions
            else:
                end_price_dict = calMatPrice(node['child_nodes'],
                                            price_dict,
                                            end_price_dict,
                                            false_mat_set)
                end_price_dict['price'] *= k
                end_price_dict['price_tax'] *= k
    except Exception as e:
        print('计算物料价格出错，表单内容有误: ' + str(e))
        print('按回车结束本程序')
        input()
        exit()
    return end_price_dict



# Recursively get node code.
def writeNodeCode(node_list, worksheet, row, col, cell_format):
    try:
        for node in node_list:
            worksheet.write(row, col, node['code'], cell_format)
            if 'child_nodes' in node:
                row = writeNodeCode(node['child_nodes'], worksheet, row, col+1, cell_format)
            else:
                row += 1
    except Exception as e:
        print('写入物料结构'+str(row)+'行'+str(col)+'列出错，表单内容有误: ' + str(e))
        print('按回车结束本程序')
        input()
        exit()
    return row



# Print outputs in an excel file.
def writeExcel(file_name, output_dict, nodes, false_mat_set):

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(file_name)
    price_sheet = workbook.add_worksheet('物料价格')
    nodes_sheet = workbook.add_worksheet('物料结构')
    remark_sheet = workbook.add_worksheet('备注')
    cell_format = workbook.add_format({'align': 'center'})

    # Start from the first cell. Rows and columns are zero indexed.
    price_sheet.write(0, 0, "物料编码", cell_format)
    price_sheet.write(0, 1, "物料名称", cell_format)
    price_sheet.write(0, 2, "规格型号", cell_format)
    price_sheet.write(0, 3, "用量价格", cell_format)
    price_sheet.write(0, 4, "含税用量价格", cell_format)
    nodes_sheet.write(0, 0, "物料编码", cell_format)
    nodes_sheet.write(0, 1, "子项物料编码", cell_format)
    remark_sheet.write(0, 0, "错误信息", cell_format)

    try:
        # Iterate over the data and write it out row by row.
        row = 1
        col = 0
        for output in output_dict:
            price_sheet.write(row, col, output, cell_format)
            price_sheet.write(row, col+1, output_dict[output]['name'], cell_format)
            price_sheet.write(row, col+2, output_dict[output]['size'] or '', cell_format)
            price_sheet.write_number(row, col+3, output_dict[output]['price'], cell_format)
            price_sheet.write_number(row, col+4, output_dict[output]['price_tax'], cell_format)
            row += 1

        # Create nodes structure in Sheet2.
        row = 1
        col = 0
        cell_format = workbook.add_format({'align': 'center'})
        for i in range(2, len(nodes)):
            nodes_sheet.write(row, col, nodes[i]['code'], cell_format)
            row = 2 + writeNodeCode(nodes[i]['child_nodes'], nodes_sheet, row, col+1, cell_format)

        # Print error. 
        remark_sheet.write(0, 1, '用量非数字', cell_format)
        row = 1
        for value in false_mat_set['con-non-num']:
            remark_sheet.write(row, 1, value, cell_format)
            row += 1

        remark_sheet.write(0, 2, '用量为0或负数', cell_format)
        row = 1
        for value in false_mat_set['con-non-posi']:
            remark_sheet.write(row, 2, value, cell_format)
            row += 1

        remark_sheet.write(0, 3, '无单价及含税单价', cell_format)
        row = 1
        for value in false_mat_set['pri-none']:
            remark_sheet.write(row, 3, value, cell_format)
            row += 1

        remark_sheet.write(0, 4, '单价非数字', cell_format)
        row = 1
        for value in false_mat_set['pri-non-num']:
            remark_sheet.write(row, 4, value, cell_format)
            row += 1

        remark_sheet.write(0, 5, '单价为0或负数', cell_format)
        row = 1
        for value in false_mat_set['pri-non-posi']:
            remark_sheet.write(row, 5, value, cell_format)
            row += 1

        remark_sheet.write(0, 6, '含税单价非数字', cell_format)
        row = 1
        for value in false_mat_set['tpri-non-num']:
            remark_sheet.write(row, 6, value, cell_format)
            row += 1

        remark_sheet.write(0, 7, '含税单价为0或负数', cell_format)
        row = 1
        for value in false_mat_set['tpri-non-posi']:
            remark_sheet.write(row, 7, value, cell_format)
            row += 1
        # Adapt column width.
        price_sheet.set_column('A:E', 15)
        nodes_sheet.set_column('A:K', 15)
        remark_sheet.set_column('A:H', 20)
        workbook.close()

    except Exception as e:
        print('生成文件出错: ' + str(e))
        print(traceback.print_exc())
        print('按回车结束本程序')
        input()
        exit()
    return