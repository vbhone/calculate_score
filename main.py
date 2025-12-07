import pandas as pd
import warnings
import shutil
import os
import math
table_info = ['序号', '班级', '学号', '姓名', '性别', '旷课次数', '请假次数', '迟到早退次数', '考勤总次数','考勤分数', '作业1',
              '作业2', '作业3', '作业成绩总和', '实验1', '实验2', '实验3', '实验成绩总和', '平时成绩总分']

full_marks=[]

# 读取文件
def read_excel(file_path):
    df = pd.read_excel(file_path, sheet_name='作业统计')
    df_check_in = pd.read_excel(file_path, sheet_name='签到详情统计')
    df_class_points = pd.read_excel(file_path, sheet_name='综合成绩（各权重项百分制得分）')
    return df,df_check_in,df_class_points

# 统计考勤信息
def get_check_in(df_check_in):
    sink_status=['事假','病假']
    normal_status=['已签','教师代签']
    absence_status=['未参与','缺勤']
    check_in={}
    for i in range(3, len(df)):
        sink_num = 0
        normal_num = 0
        absence_num = 0
        student_no=df_check_in.iloc[i][1]
        for j in range(len(df_check_in.iloc[i])):
            print(str(df_check_in.iloc[i][j]))
            if str(df_check_in.iloc[i][j]) in sink_status:
                sink_num+=1
                continue
            if str(df_check_in.iloc[i][j]) in normal_status:
                normal_num+=1
                continue
            if str(df_check_in.iloc[i][j]) in absence_status:
                absence_num+=1
                continue
        # 存储了四个数，请假，缺勤，迟到早退，签到总数。我的课没有迟到早退，所以是零
        check_in[student_no]=[sink_num,absence_num,0,sink_num+absence_num+normal_num]
    return check_in



# 确定计分方式，由用户分别输入六个分数分别由哪几项作业构成
def score_method(df):
    score_method_file=open("score_method.txt", 'r', encoding='utf-8')

    col_index = []
    col_name = []
    # 第一行记录了所有作业项
    for col_num, cell_value in enumerate(df.iloc[1], start=1):
        if pd.isna(cell_value):
            pass
        else:
            col_index.append(col_num-1)
            col_name.append(cell_value)
    print("所有作业和实验列表：")
    for i in range(len(col_name)):
        print(col_index[i], col_name[i])
    print("请选择六个分数对应的作业/实验项，输入每一项前面的序号，有多项用空格分隔：")
    score = []
    # with open(file_path, 'r', encoding='utf-8') as f:
    for i in range(1, 7):
        print("第" + str(i) + "个分数对应的作业实验项：")
        temp=score_method_file.readline().split(" ")
        temp=[int(x) for x in temp]
        score.append(temp)
    # score：6*p的二维数组，记录了每个分数需要计算的作业项，在excel对应的列号，col_index:作业项在excel中对应的列号
    return score

# 获取计分时需要用到的作业六个平均分
def get_name_ori_score(df,score):
    ori_score=[]
    # 行号从3开始才是学生的成绩记录
    for i in range(3,len(df)):
        student_info={}
        student_info['姓名']=df.iloc[i][0]
        student_info['学号']=df.iloc[i][1]
        for j in range(6):
            temp_list=[]
            for k in range(len(score[j])):
                # print("df:",df.iloc[i])
                # print("i:",i,score[j][k])
                if pd.isna(df.iloc[i][score[j][k]]):
                    df.iloc[i][score[j][k]]=0
                temp_list.append(df.iloc[i][score[j][k]])
            # 直接计算平均分
            student_info['分数'+str(j+1)]=sum(temp_list)/len(temp_list)
        # list，包括多个学生,每个学生是一个字典，包括学生姓名，学号，六个分数
        ori_score.append(student_info)
    return ori_score

# 把平均分作为中间过程写入excel表格中
def write_ori_score(file_name,check_in,ori_score):
    table_head = ['序号', '班级', '学号', '姓名', '性别', '旷课次数', '请假次数', '迟到早退次数', '考勤总次数', '作业1',
                  '作业2', '作业3', '作业成绩总和', '实验1', '实验2', '实验3', '实验成绩总和', '平时成绩总分']
    record_score=[]
    for i in range(len(ori_score)):
        student_info={}
        student_info["序号"]=i+1
        student_info["学号"]=ori_score[i]['学号']
        student_info["姓名"] = ori_score[i]['姓名']
        student_info["请假次数"] = check_in[ori_score[i]['学号']][0]
        student_info["旷课次数"] = check_in[ori_score[i]['学号']][1]
        student_info["迟到早退次数"] = check_in[ori_score[i]['学号']][2]
        student_info["考勤总次数"] = check_in[ori_score[i]['学号']][3]
        student_info["作业1"] = ori_score[i]["分数1"]
        student_info["作业2"] = ori_score[i]["分数2"]
        student_info["作业3"] = ori_score[i]["分数3"]
        student_info["实验1"] = ori_score[i]["分数4"]
        student_info["实验2"] = ori_score[i]["分数5"]
        student_info["实验3"] = ori_score[i]["分数6"]
        record_score.append(student_info)
    df_scores = pd.DataFrame(record_score)
    df_scores = df_scores.reindex(columns=table_head)
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_scores.to_excel(writer, sheet_name='百分制平均分计算', index=False)

# 计算每个分数的平均分
# def calculate_avg_100(ori_score):
#
#     for i in range(len(ori_score)):
#         for j in range(6):
#             temp_list.append(sum(ori_score[i][j])/len(ori_score[i][j]))
#         # 二维数组，记录了每个学生。每个学生有六个分数（此时的分数是百分制的平均分）
#         avg_100_score.append(temp_list)
#     return avg_100_score

# 计算最终得分，考勤分数，作业分数需要把百分制分数转换成相应满分分数，对于小数部分，采取四舍五入
def final_score(check_in,ori_score):

    fin_grade=[]
    for i in range(len(ori_score)):
        student_grade={}
        student_grade["序号"] = i + 1
        student_grade["学号"] = ori_score[i]['学号']
        student_grade["姓名"] = ori_score[i]['姓名']

        student_grade["请假次数"] = max(check_in[ori_score[i]['学号']][0],0)
        student_grade["旷课次数"] = check_in[ori_score[i]['学号']][1]
        student_grade["迟到早退次数"] = check_in[ori_score[i]['学号']][2]
        sink=student_grade["请假次数"]
        absent=student_grade["旷课次数"]
        # 旷课扣一分，三次请假扣一分，+0.5是因为要四舍五入取整
        student_grade["考勤分数"] = max(0,int(full_marks[0]-absent-sink//3+0.5))
        # 把百分制的分数转化成相应的满分制，四舍五入
        student_grade["作业1"] = int(ori_score[i]["分数1"]/100*full_marks[1]+0.5)
        student_grade["作业2"] = int(ori_score[i]["分数2"]/100*full_marks[2]+0.5)
        student_grade["作业3"] = int(ori_score[i]["分数3"]/100*full_marks[3]+0.5)
        student_grade["作业成绩总和"]=student_grade["作业1"]+student_grade["作业2"]+student_grade["作业3"]
        student_grade["实验1"] = int(ori_score[i]["分数4"]/100*full_marks[4]+0.5)
        student_grade["实验2"] = int(ori_score[i]["分数5"]/100*full_marks[5]+0.5)
        student_grade["实验3"] = int(ori_score[i]["分数6"]/100*full_marks[6]+0.5)
        student_grade["实验成绩总和"] = student_grade["实验1"]+student_grade["实验2"]+student_grade["实验3"]
        student_grade["平时成绩总分"] = student_grade["考勤分数"]+student_grade["作业成绩总和"]+student_grade["实验成绩总和"]
        fin_grade.append(student_grade)
    return fin_grade

def write_fin_score(file_name,fin_score_46):
    table_head= ['序号', '班级', '学号', '姓名', '性别', '旷课次数', '请假次数', '迟到早退次数','考勤分数', '作业1',
              '作业2', '作业3', '作业成绩总和', '实验1', '实验2', '实验3', '实验成绩总和', '平时成绩总分']
    df_scores = pd.DataFrame(fin_score_46)
    df_scores = df_scores.reindex(columns=table_head)
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_scores.to_excel(writer, sheet_name='最终平均成绩', index=False)

# 我看了一下今年大部分人没有加分。最多加分的同学为7分，所以加分方式为除以2向上取整
# def add_point(fin_score_46,df_class_points):
#     class_point={}
#     for i in range(2, len(df_class_points)):
#         class_point[df_class_points.iloc[i][1]]=math.ceil(df_class_points.iloc[i][11]/2)
#     for i in range(len(fin_score_46)):
#         temp_point=class_point[fin_score_46[i]['学号']]
#         while temp_point:
#
#
#             temp_point=temp_point-1
#
#     print(class_point)
    # return class_point


if __name__ == '__main__':
    # 统计平时成绩程序

    stu_class=['班级1','班级2','班级3','班级4']
    for i in stu_class:
        file_path = str(i)+"_统计一键导出.xlsx"
        if not os.path.exists(file_path):
            print("不存在"+str(i)+"班成绩文件")
            continue
        result_path="result"+file_path
        shutil.copy2(file_path, result_path)
        df,df_check_in,df_class_points=read_excel(file_path)
        check_in=get_check_in(df_check_in)
        # print(check_in)
        score=score_method(df)
        # print(score)

        ori_score=get_name_ori_score(df,score)
        # print(ori_score)
        # avg_100_score=calculate_avg_100(ori_score)
        write_ori_score(result_path,check_in,ori_score)
        # 首先需要知道每个分数的满分是多少，创建一个txt，里面输入7个数字，分别对应考勤满分，作业1-6满分。
        f = open("full_mark.txt", 'r', encoding='utf-8')
        full_marks = f.readline().split(" ")
        full_marks = [int(x) for x in full_marks]
        fin_score_46=final_score(check_in,ori_score)
        write_fin_score(result_path,fin_score_46)
        # 根据课堂表现加减分
        # point_score=add_point(fin_score_46,df_class_points)





