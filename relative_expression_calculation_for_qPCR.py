import pandas as pd
import numpy as np
import openpyxl as op
import csv
import argparse

args = argparse.ArgumentParser(usage="relative_expression_calculation_for_qPCR.py Ct contrast [options] output")

args.add_argument('Ct',
                   type=str,
                   help="input file in excel")

args.add_argument('Contrast',
                   type=str,
                   help="input file in excel")

args.add_argument('-o', '--output',
                   type=str,
                   help="input filename in csv including absolute path")




args = args.parse_args()


#读取想要比较的表格
contrast = pd.read_excel(args.Contrast)
#读取Ct值
df = pd.read_excel(args.Ct)

# 1. 创建文件对象
f = open(args.output, 'w', encoding='utf-8')

    # 2. 基于文件对象构建 csv写入对象
csv_writer = csv.writer(f)

#提取比较信息
nrow = contrast.shape[0]
for x in range(nrow):
    Control_sample = contrast.loc[x][0]
    Control_gene_name = contrast.loc[x][1]
    Treatment_sample = contrast.loc[x][2]
    Treatment_gene_name = contrast.loc[x][3]
    print("=============Reading the {} group===========".format(x+1))



    control_reference_gene = []
    control_target_gene = []
    treatment_reference_gene = []
    treatment_target_gene = []
    for (m,n,i) in zip(df["Target"],df["Sample"],df["Cq"]):
        if m == Control_gene_name and n == Control_sample:
            control_reference_gene.append(i)
    for (m, n, i) in zip(df["Target"], df["Sample"], df["Cq"]):
        if m == Treatment_gene_name and n == Control_sample:
            control_target_gene.append(i)
    for (m, n, i) in zip(df["Target"], df["Sample"], df["Cq"]):
        if m == Control_gene_name and n == Treatment_sample:
            treatment_reference_gene.append(i)
    for (m, n, i) in zip(df["Target"], df["Sample"], df["Cq"]):
        if m == Treatment_gene_name and n == Treatment_sample:
            treatment_target_gene.append(i)




    #计算对照和处理Ct的平均值
    control_reference_gene_mean = np.mean(control_reference_gene)
    treatment_reference_gene_mean = np.mean(treatment_reference_gene)


    #计算▲Ct
    control_cCt = []
    treatment_cCt = []

    for i in control_target_gene:
        control_cCt.append(i-control_reference_gene_mean)

    for i in treatment_target_gene:
        treatment_cCt.append(i-treatment_reference_gene_mean)



    control_cCt_mean = np.mean(control_cCt)


    #计算▲▲Ct，并且计算相对表达量
    control_ccCt_expression = []
    treatment_ccCt_expression = []


    for i in control_cCt:
        control_ccCt_expression.append(pow(2,-(i-control_cCt_mean)))
        csv_writer.writerow([Control_sample, Control_gene_name, pow(2,-(i-control_cCt_mean))])
    for i in treatment_cCt:
        treatment_ccCt_expression.append(pow(2,-(i-control_cCt_mean)))
        csv_writer.writerow([Treatment_sample, Treatment_gene_name, pow(2,-(i-control_cCt_mean))])

print("==========The program has finished running=============")
















