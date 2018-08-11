# 学生会自动化招新系统 Recruitment System for SU

这是学生会招新系统, 可以自动化生成各部门名单, 并可突出强调对某部门特别感兴趣的个人. This is the recuritment system of Students' Union, which can generate a xls roster automatically in a short time, and can highlight the person who is especially interested in a particular department.

本系统最初用于华南理工大学物理与光电学院团委学生会的招新, 可推广到一般学生会的招新程序, 将招新工作搬到云端, 更加智能化.

## 使用说明:

1. 在 www.wjx.cn 问卷星上制作好问卷 (对应于本招新系统的问卷请联系我), 到点截止填写, 后台下载xls问卷结果, 名字叫95_92_2.xls (2017年9月招17级新生的时候是叫这名字, 后来问卷星改版, 问卷结果的文件名含有中文字符"16548846_2_17级团学招新报名_97_94.xls")
2. 请将'团学招新.exe'跟95_92_2.xls放在同一个文件夹里.
3. 点击'团学招新.exe', 输入xls文件的文件名, 切勿包含扩展名! 如95_92_2.xls这个文件, 输入95_92_2就好了. (即使文件名含有中文字符以及空格也不影响小程序运行. 但若操作系统提示该exe是未知发布者有风险, 请忽略.)
4. 回车, 一键生成'招新.xls'.

## 亮点

1. 一键生成xls文件仅需0.4秒.
2. 采用学号追踪, 以最后一次的问卷提交为准. 考虑到有人重名 (但学号肯定是不同的), 有人会重复填写 (可能是由于不慎填错或者想要放弃重填或者想用多个手机号QQ号来申请多个部门等等), 以学号追踪重复的问卷提交是可行的.
3. 加入了highlight, 把那些第一第二志愿填了同一个部门而且还不服从调剂的萌新highlight起来, 以便部长和主管可以在面试时给予special care.
4. 后面诸表开列了各applicants的手机, QQ, 志愿信息, 方便部门主管在面试前发送短信通知面试时间.
