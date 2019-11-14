import xlwt
import xlrd
import configparser
from xlutils.copy import copy


def column_classification(start,end):

    # 读excel
    for i in range(start, end):
        data_list.extend(table.row_values(i))
        item = data_list[index]
        print('name:'+item)
        temp = False
        #  字符串比较
        for j in range(len(type_list)):
            if type_list[j] == item:
                temp = True
        if temp:
            write_excel(item, data_list)
            data_list = []
        else:
            type_list.append(item)
            new_excel(item, title)
            write_excel(item, data_list)
            data_list = []



# 新建excel
def new_excel(name, list):
    try:
        data = xlwt.Workbook()
        table = data.add_sheet(name)
        for i in range(len(list)):
            table.write(0, i, list[i])
        data.save(name + '.xls')
    except Exception as e :
        print(name+':'+e)


# 往已存在的excel写数据
def write_excel(name, list):
    try:
        oldWb = xlrd.open_workbook(name + '.xls')
        newWb = copy(oldWb)
        newWs = newWb.get_sheet(0)
        high = oldWb.sheets()[0].nrows
        for i in range(len(list)):
            newWs.write(high, i, list[i])
        newWb.save(name + '.xls')
    except Exception as e:
        print(name+':'+e)


if __name__ == '__main__':
    conf = configparser.ConfigParser()
    conf.read('conf.ini')
    file_name = conf.get('file', 'file_name')
    index = int(conf.get('file', 'index'))

    data = xlrd.open_workbook(file_name)

    table = data.sheets()[0]
    # 记录当前的行数据
    data_list = []
    # 记录已经生成的excel
    type_list = []
    # 记录第一行
    title = []
    title.extend(table.row_values(0))

    rows = table.nrows
    print('rows:' + str(rows))
    print(type(rows))

    # 多线程读写excel

    thread_num = conf.get('file','thread_num')
    # 余数
    remainder = rows%thread_num
    pos = (rows-remainder)/thread_num

    t = []
    # for i in range(thread_num):


    column_classification()
