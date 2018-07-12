# coding=utf-8
import os, xlrd, xlwt
from numpy import *


def WriteSheet(sheet, List, j, Len):
    sheet.col(j).width = 256*12
    for i in range(0, Len):
        sheet.write(i + 1, j, List[i])


def ReferDict(List, Dict, Len):
    ListRe = list(range(0, Len))
    for i in range(0, Len):
        ListRe[i] = Dict[List[i]]
    return ListRe


def delEmptyElement(List):
    while List[-1] == '':
        del List[-1]
    return List


def ContactDeps(TheDep, TheDep1, TheDep2, PhoneNum, QQ, Name, Assignment, AssignmentDict):
    TheDep.write(0, 0, '手机'); TheDep.col(0).width = 256*12
    TheDep.write(0, 1, 'QQ'); TheDep.col(1).width = 256*12
    TheDep.write(0, 2, '姓名'); TheDep.col(2).width = 256*10
    TheDep.write(0, 3, '第几志愿')
    TheDep.write(0, 4, '服从调剂？'); TheDep.col(4).width = 256*10
    for i in range(0, len(TheDep1)):
        TheDep.write(i + 1, 0, PhoneNum[int(TheDep1[i])])
        TheDep.write(i + 1, 1, QQ[int(TheDep1[i])])
        TheDep.write(i + 1, 2, Name[int(TheDep1[i])])
        TheDep.write(i + 1, 3, 1)
        TheDep.write(i + 1, 4, AssignmentDict[Assignment[int(TheDep1[i])]])
    for i in range(0, len(TheDep2)):
        TheDep.write(i + 1 + len(TheDep1), 0, PhoneNum[int(TheDep2[i])])
        TheDep.write(i + 1 + len(TheDep1), 1, QQ[int(TheDep2[i])])
        TheDep.write(i + 1 + len(TheDep1), 2, Name[int(TheDep2[i])])
        TheDep.write(i + 1 + len(TheDep1), 3, 2)
        TheDep.write(i + 1 + len(TheDep1), 4, AssignmentDict[Assignment[int(TheDep2[i])]])


# 班级Class和数字1~4的索引
ClassDict = {
    1: '应物(严)',
    2: '应物',
    3: '光电1',
    4: '光电2'
}
# 部门Department和数字1~8的索引
DepDict = {
    1: '外联部',
    2: '学术科创部',
    3: '秘书部',
    4: '新媒体运营部',
    5: '生活权益部',
    6: '组织部',
    7: '体育部',
    8: '文艺部'
}
# 是否服从调剂? 1为'是', 服从调剂, 2为'否', 不服从调剂
AssignmentDict = {
    1: '是',
    2: '否'
}

# 从RawData中pick out出萌新们的个人信息和第一第二志愿. RawData导出自www.wjx.cn问卷星
Dir = os.getcwd() + '/'
print('欢迎使用物理与光电学院团委学生会 招新ScriptV2.0 :)\n<<<<<<<<<<<<<<<<<< 啦啦啦我是分割线 >>>>>>>>>>>>>>>>>>')
print('author:\n马守然 (2014级应用物理学)\n学术科创部\n物理与光电学院团委学生会'
      '\nEmail: 1941688873@qq.com / Ma.Seoyin@gmail.com\n<<<<<<<<<<<<<<<<<< 啦啦啦我是分割线 >>>>>>>>>>>>>>>>>>')
FileName = input('请输入文件名(切勿包含扩展名!):\n')
RawData = xlrd.open_workbook(Dir + FileName + '.xls').sheet_by_index(0)
Len = RawData.nrows - 1 # 一共有Len那么多的人报名团学(可能有人重复填了, 不管)
Num = list(range(1, Len + 1)) # 给名字标号, 用于重整第一第二志愿
Class = ReferDict(list(map(int, RawData.col_values(6)[1:])), ClassDict, Len) # 班级
Name = list(map(str, RawData.col_values(7)[1:]))
PhoneNum = list(map(str, RawData.col_values(8)[1:])) # 手机Num
QQ = list(map(str, list(map(int, RawData.col_values(9)[1:]))))
Volunteer1 = list(map(int, RawData.col_values(10)[1:])) # 第一志愿
Volunteer2 = list(map(int, RawData.col_values(11)[1:])) # 第二志愿
Assignment = list(map(int, RawData.col_values(12)[1:])) # 服从调剂?
Mat_Volunteer = vstack((array([Num]), array([Volunteer1]), array([Volunteer2]), array([Assignment]))).T

WorkBook = xlwt.Workbook()
# 录入部门志愿
DepVol = WorkBook.add_sheet('录入部门志愿')
DepVol.write(0, 0, '手机')
WriteSheet(DepVol, PhoneNum, 0, Len)
DepVol.write(0, 1, 'QQ')
WriteSheet(DepVol, QQ, 1, Len)
DepVol.write(0, 2, '班级')
WriteSheet(DepVol, Class, 2, Len)
DepVol.write(0, 3, '姓名')
WriteSheet(DepVol, Name, 3, Len)
DepVol.write(0, 4, '第一志愿')
WriteSheet(DepVol, ReferDict(Volunteer1, DepDict, Len), 4, Len)
DepVol.write(0, 5, '第二志愿')
WriteSheet(DepVol, ReferDict(Volunteer2, DepDict, Len), 5, Len)
DepVol.write(0, 6, '服从调剂？')
WriteSheet(DepVol, ReferDict(Assignment, AssignmentDict, Len), 6, Len)

# 第一第二志愿名单
Roster = WorkBook.add_sheet('第一第二志愿名单')
Dep1 = 1000 * ones((Len, 8))
Dep2 = 1000 * ones((Len, 8))
Iron = 1000 * ones((Len, 8))
for i in range(0, Len):
    if (Mat_Volunteer[i, 1] == Mat_Volunteer[i, 2]) & (Mat_Volunteer[i, 3] == 2):
        Iron[i, (Mat_Volunteer[i, 1] - 1)] = i # 一个第一第二志愿都填了同一个部门而且不服从调剂的萌新i
    for j in range(0, 8):
        if Mat_Volunteer[i, 1] == (j + 1):
            Dep1[i, j] = i # 第一志愿去部门j的萌新i
        if (Mat_Volunteer[i, 2] == (j + 1)) & (Mat_Volunteer[i, 1] != (j + 1)):
            Dep2[i, j] = i # 第二志愿去部门j的萌新i
Dep1 = sort(Dep1, 0); Dep2 = sort(Dep2, 0)
# 第一第二志愿名单 - 第一志愿名单
col = 0
Roster.write(0, col, '第一志愿名单')
Roster.write(1, col, '序号')
xlwt.add_palette_colour('yellow', 0x22)
color = xlwt.easyxf('pattern: pattern solid, fore_colour yellow')
for j in range(0, 8):
    Roster.write(1, j + 1 + col, DepDict[j + 1])
    Iron_Dep = list(Iron[:, j])
    for i in range(0, Len):
        if (Dep1[i, j] != 1000) & (Dep1[i, j] in Iron):
            # 如果萌新i铁定想去部门j(亦即第一第二志愿填同一个部门, 而且不服从调剂), 就highlight出来
            Roster.write(2 + i, j + 1 + col, Name[int(Dep1[i, j])], color)
        if (Dep1[i, j] != 1000) & ((Dep1[i, j] in Iron) == 0):
            # 如果萌新i第一志愿报了部门j而并不是铁定想去部门j就不highlight. 包括但不限于, 第一第二志愿不同, 或者第一第二志愿相同却不服从调剂, 等情况
            Roster.write(2 + i, j + 1 + col, Name[int(Dep1[i, j])])
WaiLian1 = extract(Dep1[:, 0] < 1000, Dep1[:, 0])
XueChuang1 = extract(Dep1[:, 1] < 1000, Dep1[:, 1])
MiShu1 = extract(Dep1[:, 2] < 1000, Dep1[:, 2])
XinMeiTi1 = extract(Dep1[:, 3] < 1000, Dep1[:, 3])
ShengHuo1 = extract(Dep1[:, 4] < 1000, Dep1[:, 4])
ZuZhi1 = extract(Dep1[:, 5] < 1000, Dep1[:, 5])
TiYu1 = extract(Dep1[:, 6] < 1000, Dep1[:, 6])
WenYi1 = extract(Dep1[:, 7] < 1000, Dep1[:, 7])
No1 = max([len(WaiLian1), len(XueChuang1), len(MiShu1), len(XinMeiTi1), len(ShengHuo1), len(ZuZhi1), len(TiYu1), len(WenYi1)])
for i in range(0, No1):
    Roster.write(i + 2, col, i + 1)
# 第一第二志愿名单 - 第二志愿名单
col = 8 + 1
Roster.write(0, col, '第二志愿名单')
Roster.write(1, col, '序号')
for j in range(0, 8):
    Roster.write(1, j + 1 + col, DepDict[j + 1])
    for i in range(0, Len):
        if Dep2[i, j] != 1000:
            Roster.write(2 + i, j + 1 + col, Name[int(Dep2[i, j])])
WaiLian2 = extract(Dep2[:, 0] < 1000, Dep2[:, 0])
XueChuang2 = extract(Dep2[:, 1] < 1000, Dep2[:, 1])
MiShu2 = extract(Dep2[:, 2] < 1000, Dep2[:, 2])
XinMeiTi2 = extract(Dep2[:, 3] < 1000, Dep2[:, 3])
ShengHuo2 = extract(Dep2[:, 4] < 1000, Dep2[:, 4])
ZuZhi2 = extract(Dep2[:, 5] < 1000, Dep2[:, 5])
TiYu2 = extract(Dep2[:, 6] < 1000, Dep2[:, 6])
WenYi2 = extract(Dep2[:, 7] < 1000, Dep2[:, 7])
No2 = max([len(WaiLian2), len(XueChuang2), len(MiShu2), len(XinMeiTi2), len(ShengHuo2), len(ZuZhi2), len(TiYu2), len(WenYi2)])
for i in range(0, No2):
    Roster.write(i + 2, col, i + 1)

# 外联部
WaiLian = WorkBook.add_sheet('外联部')
ContactDeps(WaiLian, WaiLian1, WaiLian2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 学术科创部
XueChuang = WorkBook.add_sheet('学术科创部')
ContactDeps(XueChuang, XueChuang1, XueChuang2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 秘书部
MiShu = WorkBook.add_sheet('秘书部')
ContactDeps(MiShu, MiShu1, MiShu2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 新媒体运营部
XinMeiTi = WorkBook.add_sheet('新媒体运营部')
ContactDeps(XinMeiTi, XinMeiTi1, XinMeiTi2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 生活权益部
ShengHuo = WorkBook.add_sheet('生活权益部')
ContactDeps(ShengHuo, ShengHuo1, ShengHuo2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 组织部
ZuZhi = WorkBook.add_sheet('组织部')
ContactDeps(ZuZhi, ZuZhi1, ZuZhi2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 体育部
TiYu = WorkBook.add_sheet('体育部')
ContactDeps(TiYu, TiYu1, TiYu2, PhoneNum, QQ, Name, Assignment, AssignmentDict)
# 文艺部
WenYi = WorkBook.add_sheet('文艺部')
ContactDeps(WenYi, WenYi1, WenYi2, PhoneNum, QQ, Name, Assignment, AssignmentDict)

WorkBook.save(Dir + '招新.xls')
input('\'招新.xls\'文件已生成, 请按回车关闭本宝宝 :)')
