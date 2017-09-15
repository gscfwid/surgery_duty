# coding: utf-8

import sys
reload(sys)
sys.setdefaultencoding('utf8')

import random
import xlrd
import xlwt
import itertools
from xlutils.copy import copy

def read_data(a):
#建立7个片区的列表
    pianqu1 = [u'01－',u'02－',u'03－',u'04－',u'05－',u'06－',u'07－']
    pianqu2 = [u'08－',u'09－',u'10－',u'10DSA',u'10MR']
    pianqu3 = [u'11－',u'12－',u'13－',u'15－',u'16－',u'17－',u'18－']
    pianqu4 = [u'19－',u'20－',u'21－',u'22－',u'23－',u'25－',u'26－']
    pianqu5 = [u'27－',u'28－',u'29－',u'30－',u'31－',u'32－']
    pianqu6 = [u'33－',u'34－',u'35－',u'36－',u'37－',u'38－',u'39－',u'40－']
    pianqu7 = [u'A(',u'B-',u'C-',u'D-']
    rooms_persons = {}
    persons = xlrd.open_workbook(u'./persons.xls')
    for i in range(0,7):
        group_name = 'pianqu'+str(i+1) #这里的+1是为了目标窗
        rooms = vars()[group_name]
        if len(persons.sheets()[a].row_values(i+1)) > len(rooms)+1:
            persons_list = persons.sheets()[a].row_values(i+1)[1:len(rooms)+1]
            random.shuffle(persons_list)
        if len(persons.sheets()[a].row_values(i+1)) == len(rooms)+1:
            persons_list = persons.sheets()[a].row_values(i+1)[1:len(rooms)+1] #参数a代表不同的sheet，sheet1是高年资医生，sheet2是低年资医生
            random.shuffle(persons_list)
        if len(persons.sheets()[a].row_values(i+1)) < len(rooms)+1:
            n = len(rooms)+1 - len(persons.sheets()[a].row_values(i+1))
            empties = ['']*n
            persons_list = persons.sheets()[a].row_values(i+1)[1:len(rooms)+1]
            random.shuffle(persons_list) #为了体现随机，打乱了人员列表顺序
            persons_list.extend(empties)
        for room in rooms:
            if len(persons_list) <= 2 and len(persons_list) >0: #建立这几个条件是为了尽量避免连续空余病房太多
                room_person = random.choice(persons_list)
                rooms_persons[room] = room_person
                persons_list.remove(room_person)
            elif len(persons_list) % 2 == 1:
                room_person = random.choice(persons_list[0:2])
                rooms_persons[room] = room_person
                persons_list.remove(room_person)
            elif len(persons_list) % 2 == 0:
                room_person = random.choice(persons_list[-2:])
                rooms_persons[room] = room_person
                persons_list.remove(room_person)
            elif len(persons_list) == 0:
                pass
    return rooms_persons
def write_data(line,mark_col,fill_col): #三个参数的意义分别是：line代表从第几行还是填入数据；mark_col表示手术室编号位于第几列；fill_col代表要填入数据的第一列位于第几列
    duties = xlrd.open_workbook(u'./input.xls',formatting_info=True)
    duties_copy = copy(duties)
    duties_table = duties_copy.get_sheet(0)
    duties_nrows = duties.sheets()[0].nrows
    rooms_skilled = read_data(0)
    rooms_young = read_data(1)
    style = xlwt.easyxf('font:height 220');
    for i in range(line,duties_nrows):
        for key in rooms_skilled.keys():
            if key in duties.sheets()[0].row_values(i)[mark_col]:
                duties_table.write(i,fill_col,rooms_skilled[key],style)
                duties_table.write(i,fill_col+1,rooms_skilled[key],style)
            else:
                pass
        for key in rooms_young.keys():
            if key in duties.sheets()[0].row_values(i)[mark_col]:
                duties_table.write(i,fill_col+2,rooms_young[key],style)
            else:
                pass
    duties_copy.save('./output.xls')
write_data(1,11,8)
