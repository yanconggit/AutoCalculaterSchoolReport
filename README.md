# AutoCalculaterSchoolReport
The project can calculater course score for optimal pioneer by python.Apply to HPU's school report.



## 使用说明
本程序可以较方便的计算HPU综合评定积分表中的课程成绩分（X1）

前期准备：需要下载本程序exe文件，有学院发的成绩单
使用方法：
1.新建一个空的Excel表格（.xlsx文件），命名为“班级成绩单.xlsx”
2.将专业成绩单中自己班级的成绩复制到“班级成绩单.xlsx”的sheet1页
3.进入HPU综合教务系统查看我的课表，
   将对应两学期的课程清单复制到“班级成绩单.xlsx”的sheet2页，
   粘贴的时候选择“匹配目标格式”，
   删除除课程属性、学分、课程名称外的其他列
   按照课程名称排序
4.保存“班级成绩单.xlsx”文件，然后关闭该文件
5.运行下载好的exe文件，稍等片刻
6.根据提示在弹出的界面中选择“班级成绩单.xlsx”表格
7.根据提示可能需要手动输入部分课程对应的学分数
8.听到或看到 “Successful Done!”即处理完成
9.在“班级成绩单.xlsx”所在的文件夹中可看到“班级成绩单--output.xlsx”文件，计算结果保存在该文件的最后几列

注：因为没有大一大二学年的成绩单，暂时没有添加处理体育课的功能，暂不适合大一大二学年综合评定使用，如果可以提供相应的成绩单可以添加该功能
注：如有同学课程缓考，需要在第2步粘贴成绩后，手动将该同学该成绩所在单元格清空（右键点击该单元格 选择 清除内容或选中该单元格点击键盘上的delete
