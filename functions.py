# coding = utf-8
# functions.py
import openpyxl

# Import an excel file by openpyxl package.
def openExcelSheet(file='file.xlsx', sheet='Sheet1'):
    try:
        data = openpyxl.load_workbook(filename = file)
        sheet_data = data[sheet]
        return sheet_data
    except Exception.e:
        print('打开文件出错: '+e)

# Put sheet-contents into nodes[].
def importContents(sheet, col_par_code, col_chi_seq, col_chi_code, col_do_mol, col_do_den):
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
    node = {'code': nodes_code[len(nodes_code)-1], 'child_nodes': child_nodes}
    nodes.append(node)
    return nodes

