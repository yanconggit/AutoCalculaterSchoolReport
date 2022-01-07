import pandas as pd
import numpy as np
import math
import os
import tkinter.filedialog
import pyttsx3


##### 读取所学的主要课程和对应的学分数
# 获取文件
print('提示:请选择您处理后的成绩单，注:输出默认保存到相同路径\n')
#使用TK GUI选择成绩单文件
file_path = tkinter.filedialog.askopenfilename(title='请选择您处理过的成绩单', filetypes=[('所有文件', '.*'), ('xlsx文件', '.xlsx'), ('xls文件', '.xls')])
file_name2 = os.path.split(os.path.splitext(file_path)[0]) # 获取成绩单的文件名
outputfile1 = os.path.dirname(os.path.abspath(file_path))  # 获取成绩单所在路径的父路径
outputfile_path = outputfile1 + "\\" + file_name2[1] + "--output.xlsx" # 得出输出文件的路径
courselist = pd.read_excel(file_path,sheet_name = "Sheet2",usecols="A:B") #读取所学的主要课程和对应的学分数
# 将课程和对应的学分转换位字典 方便快速查找
cs1 = courselist.set_index(['课程名称'])['学分'].to_dict()

###### 读取成绩单
schoolreport = pd.read_excel(file_path,sheet_name = "Sheet1") # 将Sheet1中的原始成绩读入
# 查看主要课程学分列表全不全
a = 8
while schoolreport.columns[a] != "课程、学分、成绩": # 判断所有主要课程名称是否有对应的学分
    if cs1.setdefault(schoolreport.columns[a]) == None:
        print("sheet2页中无\"",schoolreport.columns[a],"\"对应的学分数,请手动输入（每次都要输入）或者在成绩单sheet2页补充（一劳永逸）")
        print("请输出上述课程对应的学分数，输入完毕后按回车键结束")
        cs1[schoolreport.columns[a]] = float(input())
    a += 1

# 新添加三列用来记录结果
schoolreport['所学总学分数（分母）'] = '' 
schoolreport['所修课程成绩*该课程学分（分子）'] = ''
schoolreport['课程成绩分'] = ''

# 将等级制的等级转换为对应的分数
schoolreport.replace(['良好','85'])
for row in range(0,schoolreport.shape[0]):
    for col in range(8,schoolreport.shape[1]-3):
        content = schoolreport.loc[row][col]
        if not content is np.nan:
            if schoolreport.loc[row][col] == "优秀":
                schoolreport.iat[row,col] = 95
            elif schoolreport.loc[row][col] == "良好":
                schoolreport.iat[row,col] = 85
            elif schoolreport.loc[row][col] == "中等":
                schoolreport.iat[row,col] = 75
            elif schoolreport.loc[row][col] == "及格":
                schoolreport.iat[row,col] = 65
            elif schoolreport.loc[row][col] == "不及格":
                schoolreport.iat[row,col] = 0

degree2score = {'优秀 ':95,'良好 ':85,'中等 ':75,'及格 ':65,'不及格 ':0}

###### 计算课程成绩分
# for循环一次一行 即一位同学一位同学的计算
for i in range(0,schoolreport.shape[0]):
    a = 8  # 从第八列开始才是成绩 注：序号列是第 0 列
    sumscore = 0 # 用来存放累加的值
    sumcredit = 0
    while schoolreport.columns[a] != "课程、学分、成绩": # 先将主要课程计算完毕
        if not math.isnan(schoolreport.loc[i][a]):  # 计算非空单元格的成绩，有成绩就表示修了这门课
        #if not (schoolreport.loc[i][a]) is np.nan:
            sumscore += (schoolreport.loc[i][a] * cs1.get(schoolreport.columns[a])) # 累加成绩与学分的乘积
            sumcredit += cs1.get(schoolreport.columns[a]) # 累加所学课程学分数
            a = a + 1 
            if a >= schoolreport.shape[1]-3: # 越界检查
                break
        else: # 单元格为空表示学生没有修这门课
            a = a + 1
            if a >= schoolreport.shape[1]-3:
                break
            continue
    while True: # 计算单列的课程成绩 并不是所有学生都有单列课程（班里没几个同学学的课才会单列为 “课程、学分、成绩”）
        if schoolreport.loc[i][a] is not np.nan: # 判断是否有单列课程成绩
            temp1=schoolreport.loc[i][a].split(',')  # 提取单列课程的
            if degree2score.get(temp1[2],101) != 101 :
                sumscore += float(temp1[1]) * degree2score.get(temp1[2])
            else:
                sumscore += float(temp1[1]) * float(temp1[2]) 
            sumcredit += float(temp1[1])
            a += 1;
            if a >= schoolreport.shape[1]-3:
                break
        else:
            break
##### 保存结果
    schoolreport.iat[i,schoolreport.shape[1]-3] = sumcredit 
    schoolreport.iat[i,schoolreport.shape[1]-2] = sumscore
    schoolreport.iat[i,schoolreport.shape[1]-1] = sumscore / sumcredit

schoolreport.to_excel(outputfile_path,sheet_name = '成绩单',index = False) # 转存为Excel表格
print('Successful Done!') # 操作成功提示
pyttsx3.speak("Successful Done!")
# 注意事项  缓考的成绩要手动剔除

# 文件路径使用GUI 选择   OK
# 等级制的课程计算       OK 
# 课程学分列表缺项 警告  OK
# 大一体育成绩的处理     暂无数据 
# 计算必修课最低分



# git tag 1.0.0 1b2e1d63ff
# 1b2e1d63ff 是你想要标记的提交 ID 的前 10 位字符。可以使用下列命令获取提交 ID：
# git log
# 你也可以使用少一点的提交 ID 前几位，只要它的指向具有唯一性。
# 内建的图形化 git：
# gitk
# 彩色的 git 输出：
# git config color.ui true
# 显示历史记录时，每个提交的信息只显示一行：
# git config format.pretty oneline