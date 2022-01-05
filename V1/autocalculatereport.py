import pandas as pd
import numpy as np
import math

##### 读取所学的主要课程和对应的学分数
# 根据需要自己修改路径
file_path = r'D:/wyc/Mirror/大学/自己学的一些/python/处理成绩单/V1/firstclass.xlsx'
outputfile_path = r'D:/wyc/Mirror/大学/自己学的一些/python/处理成绩单/V1/output.xlsx' # 另建一个表格存放原始数据和结果 避免原始数据丢失
courselist = pd.read_excel(file_path,sheet_name = "Sheet2",usecols="A:B")
# 将课程和对应的学分转换位字典 方便快速查找
cs1 = courselist.set_index(['课程'])['学分'].to_dict()

###### 读取成绩单
schoolreport = pd.read_excel(file_path,sheet_name = "Sheet1") # 将Sheet1中的原始成绩读入
# 新添加三列用来记录结果
schoolreport['所学总学分数（分母）'] = '' 
schoolreport['所修课程成绩*该课程学分（分子）'] = ''
schoolreport['课程成绩分'] = ''

###### 计算课程成绩分
# for循环一次一行 即一位同学一位同学的计算
for i in range(0,schoolreport.shape[0]):
    a = 8  # 从第八列开始才是成绩 注：序号列是第 0 列
    sumscore = 0 # 用来存放累加的值
    sumcredit = 0
    while schoolreport.columns[a] != "课程、学分、成绩": # 先将主要课程计算完毕
        if not math.isnan(schoolreport.loc[i][a]):  # 计算非空单元格的成绩，有成绩就表示修了这门课
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

schoolreport.to_excel(outputfile_path,index = False) # 转存为Excel表格
print('Successful Done!') # 操作成功提示

#注意事项  缓考的成绩要手动剔除
# 等级制的课程计算
# 课程学分列表缺项 警告
# 大一体育成绩的处理

