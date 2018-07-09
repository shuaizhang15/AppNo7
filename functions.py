# coding = utf-8
# functions.py
import openpyxl
import xlsxwriter



# Import an excel file by openpyxl package.
def openExcelSheet(file, sheet):
    try:
        data = openpyxl.load_workbook(filename = file)
        sheet_data = data[sheet]
        return sheet_data
    except Exception as e:
        print('打开文件出错: ' + str(e))



# Put sheet-contents into nodes[].
def readSheet(sheet, col_par_code, col_chi_seq, col_chi_code, col_do_mol, col_do_den):
    nodes = []                  # 总节点序列
    nodes_code = [0]            # 父节点编码
    node = {}                   # 父节点对象
    child_nodes = []            # 子节点序列
    child_node = {}             # 子节点对象
    par_code = sheet[col_par_code]
    chi_seq = sheet[col_chi_seq]
    chi_code = sheet[col_chi_code]
    do_mol = sheet[col_do_mol]
    do_den = sheet[col_do_den]
    try: 
        # store node
        for i in range(0, len(par_code)):
            # parent code exists
            if par_code[i].value:
                # initilize a new node by storing the last node code & its child nodes
                node = {'code': nodes_code[len(nodes_code)-1], 'child_nodes': child_nodes}
                nodes.append(node)
                nodes_code.append(par_code[i].value)
                # initilize a new child node 
                child_nodes = []
                child_node = {'seq': chi_seq[i].value, 'code': chi_code[i].value, 'do_mol': do_mol[i].value, 'do_den': do_den[i].value}
                child_nodes.append(child_node)
            # parent code does not exist, keep appending child nodes
            else:
                child_node = {'seq': chi_seq[i].value, 'code': chi_code[i].value, 'do_mol': do_mol[i].value, 'do_den': do_den[i].value}
                child_nodes.append(child_node)
    except Exception as e:
        print('文件读取有误: ' + str(e))

    node = {'code': nodes_code[len(nodes_code)-1], 'child_nodes': child_nodes}
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



# Set child node list in a node.
# -- no need --
# def setChildNodes(childNodes, codeDict):
#     for child_node in childNodes:
#         if child_node['code'] in codeDict.keys():
#             child_node['child_nodes'] = codeDict[child_node['code']]
#     return childNodes



# Recursively calculate end price of each node.
def calMatPrice(node_list, price_dict, end_price_dict, false_mat_set):
    try:
        for node in node_list:
            k = float(node['do_mol']) / float(node['do_den'])
            if 'child_nodes' not in node:
                if node['code'] in price_dict:
                    end_price_dict['price'] += k * price_dict[node['code']]['price']
                    end_price_dict['price_tax'] += k * price_dict[node['code']]['price_tax']
                else:
                    false_mat_set.add(node['code'])
                    # print('未查询到编码为' + str(node['code']) + '的物料价格')
            else:
                end_price_dict = calMatPrice(node['child_nodes'], price_dict, end_price_dict, false_mat_set)
                end_price_dict['price'] *= k
                end_price_dict['price_tax'] *= k
    except Exception as e:
        print('计算物料价格出错，表单内容有误: ' + str(e))
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
    price_sheet.write(0, 1, "用量价格", cell_format)
    price_sheet.write(0, 2, "含税用量价格", cell_format)
    nodes_sheet.write(0, 0, "物料编码", cell_format)
    nodes_sheet.write(0, 1, "子项物料编码", cell_format)
    remark_sheet.write(0, 0, "错误信息", cell_format)

    # Iterate over the data and write it out row by row.
    row = 1
    col = 0
    for output in output_dict:
        price_sheet.write(row, col, output, cell_format)
        price_sheet.write_number(row, col+1, output_dict[output]['price'], cell_format)
        price_sheet.write_number(row, col+2, output_dict[output]['price_tax'], cell_format)
        row += 1

    # Create nodes structure in Sheet2.
    row = 1
    col = 0
    cell_format = workbook.add_format({'align': 'center'})
    for i in range(2, len(nodes)):
        nodes_sheet.write(row, col, nodes[i]['code'], cell_format)
        row = 2 + writeNodeCode(nodes[i]['child_nodes'], nodes_sheet, row, col+1, cell_format)

    # Print error. 
    row = 1
    col = 0
    for false_mat_code in false_mat_set:
        remark_sheet.write(row, col, '编码物料价格不存在：', cell_format)
        remark_sheet.write(row, col+1, false_mat_code, cell_format)
        row += 1

    # Adapt column width.
    price_sheet.set_column('A:C', 15)
    nodes_sheet.set_column('A:K', 15)
    remark_sheet.set_column('A:C', 20)

    # Generate file.
    try:
        workbook.close()
    except Exception as e:
        print('生成文件出错: ' + str(e))

    return