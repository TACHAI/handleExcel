import xlwt, threading
import xlrd, time
import configparser
from xlutils.copy import copy


def column_classification(start, end):
    # 读excel
    for i in range(start, end):
        print(i)
        mutex.acquire()
        data_list = []
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
        else:
            # 上锁
            type_list.append(item)
            new_excel(item, title)
            write_excel(item, data_list)
            # 释放锁
        mutex.release()


# 新建excel
def new_excel(name, new_list):
    try:
        # xlwt 支持256行  todo 需要改
        new_data = xlwt.Workbook()
        new_table = new_data.add_sheet(name)
        for n in range(len(new_list)):
            new_table.write(0, n, new_list[n])
        new_data.save(name + '.xls')
    except Exception as e:
        print(e)


# 往已存在的excel写数据
def write_excel(name, write_list):
    try:
        old_wb = xlrd.open_workbook(name + '.xls')
        new_wb = copy(old_wb)
        new_ws = new_wb.get_sheet(0)
        high = old_wb.sheets()[0].nrows
        for w in range(len(write_list)):
            new_ws.write(high, w, write_list[w])
        new_wb.save(name + '.xls')
    except Exception as e:
        print(e)


# 创建一个互斥锁，默认是没有上锁的
mutex = threading.Lock()


if __name__ == '__main__':
    conf = configparser.ConfigParser()
    conf.read('conf.ini')
    file_name = conf.get('file', 'file_name')
    index = int(conf.get('file', 'index'))

    data = xlrd.open_workbook(file_name)
    table = data.sheets()[0]
    # 记录当前的行数据
    # 记录已经生成的excel
    type_list = []
    # 记录第一行
    title = []
    title.extend(table.row_values(0))

    rows = table.nrows
    print('rows:' + str(rows))
    print(type(rows))

    # 多线程读写excel

    thread_num = int(conf.get('file', 'thread_num'))
    # 余数
    remainder = rows % thread_num
    pos = int((rows-remainder)/thread_num)
    print(rows)
    print(pos)
    print(remainder)

    t = []

    start_time = time.time()

    # 加入线程组
    for x in range(thread_num-1):
        t.append(threading.Thread(target=column_classification, args=(x*pos+1, (x+1)*pos)))
    t.append(threading.Thread(target=column_classification, args=((thread_num-1)*pos, (thread_num*pos)+remainder)))

    # 启动线程
    for tt in t:
        # 守护线程
        # tt.setDaemon(True)
        tt.start()

    end_time = time.time()
    print('花费时间是:', round(end_time-start_time, 4))


