from utamethods import UTA, UTASTAR, UTADIS
import pandas as pd
from operator import itemgetter
from termcolor import colored
import xlrd
import matplotlib.pyplot as plt
import xlsxwriter

Global=[]
Global_User_weights=[]

wb2 = xlrd.open_workbook(r'bathmologies_kinhtwn_sunolou_agoras.xlsx')
sheet1 = wb2.sheet_by_index(0)
data=pd.read_excel(r"bathmologies_kinhtwn_sunolou_anaforas.xlsx")
workbook = xlrd.open_workbook(r'dedomena_apokleismou.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
marka = worksheet.cell(0, 0).value
leitourgiko = worksheet.cell(0, 1).value
hmarka=worksheet.cell(0, 2).value
toOs=worksheet.cell(0, 3).value
bestvalues=[4,4,3,3,3,4,4]
worstvalues=[1,1,1,1,1,1,1]
monotonicity=[1,1,1,1,1,1,1]
intervals=[3,3,2,2,2,3,3]
w0=[10,6,8,7,7,6,9,7,6,8,8,6]

w2=[(w0[1]+w0[2]+w0[6]+w0[9])/4,(w0[3]+w0[5]+w0[7])/3,w0[4],w0[8],w0[0],w0[10],w0[11]]


if marka=='Όχι' :
    if leitourgiko=='Όχι':
        del bestvalues[4]
        del worstvalues[4]
        del monotonicity[4]
        del intervals[4]
        del bestvalues[5]
        del worstvalues[5]
        del monotonicity[5]
        del intervals[5]
        del w2[4]
        del w2[5]

    else: #5 leitou, 7 marka
        del bestvalues[6]
        del worstvalues[6]
        del monotonicity[6]
        del intervals[6]
        del w2[6]
else:
     if leitourgiko=='Όχι':
         del bestvalues[4]
         del worstvalues[4]
         del monotonicity[4]
         del intervals[4]
         del w2[4]
w2=[float(i)/sum(w2) for i in w2]

res=UTASTAR(data,bestvalues=bestvalues, worstvalues=worstvalues, intervals=intervals, monotonicity=monotonicity)

figure, axis = plt.subplots(1, len(bestvalues)+1)
Sum=0
temp=0
k=0
we=[]
counter=0
print('intervals',intervals,len(bestvalues))
for i in range(0,len(bestvalues)):
    for l in range(0,intervals[i]):

        temp=res.w[0][counter]+temp
        counter=counter+1


    we.append(temp)
    temp=0


add_utilities=[]
temp1=0
for k in range(0,len(bestvalues)):
    add_utilities1=[]
    for i in range(temp1,temp1+intervals[k]+1):
        add_utilities1.append(res.crit_utilities[i])
    temp1=temp1+intervals[k]+1
    add_utilities.append(add_utilities1)




crit_utilities1=[]
for a in add_utilities:


    xmin = min(a)
    xmax = max(a)
    for i, x in enumerate(a):
        a[i] = (x - xmin) / (xmax - xmin)

    crit_utilities1.extend(a)


for kinhto_row in range(0,25):
    Sum = 0
    SUM1=0
    temp = 0
    k = 0

    for x in sheet1.row_values(kinhto_row):
        if k is not 7:

            Sum=Sum+crit_utilities1[int(x-1+temp)]*we[k]
            SUM1=SUM1+crit_utilities1[int(x-1+temp)]*w2[k]
            temp=temp+intervals[k]+1

            k = k + 1

    Global.append([Sum,kinhto_row])
    Global=(sorted(Global, key=itemgetter(0),reverse=True))
    Global_User_weights.append([SUM1, kinhto_row])
    Global_User_weights = (sorted(Global_User_weights, key=itemgetter(0), reverse=True))


loc2=("xarakthristika_kinhtwn.xlsx")
wb2 = xlrd.open_workbook(loc2)
sheet1 = wb2.sheet_by_index(0)

print(colored("Global Utilites \t Alt# \t Global Utilites \t Alt#",'blue'))
counter3=0
workbook33 = xlsxwriter.Workbook('Results.xlsx')
worksheet = workbook33.add_worksheet("GlobalUtilities")
row1=1
column1=0
cell_format3 = workbook33.add_format({'bold': True, 'font_color': 'blue'})
worksheet.write(row1-1, column1, "Alternative Name",cell_format3)
worksheet.write(row1-1, column1+1, "Global Value",cell_format3)

for ii,ih  in zip(Global,Global_User_weights):
    counter3=counter3+1

    if marka=='Όχι':

        if (sheet1.cell_value(int(ii[1]+1), 12)) !=hmarka:

            continue

    if leitourgiko=='Όχι':

        if (int(sheet1.cell_value(int(ii[1]+1), 9))) !=1 and toOs=='Android':
            continue
        if (int(sheet1.cell_value(int(ii[1]+1), 9))) !=0 and toOs=='Ios':
            continue
    cell_format1 = workbook33.add_format({'bold': True, 'font_color': 'green'})
    print(colored ("%0.13f" %ii[0],'red'),'\t',ii[1]," \t")

    worksheet.write(row1, column1+1, ii[0])
    worksheet.write(row1, column1, sheet1.cell_value(ii[1]+1,0),cell_format1)


    row1=row1+1
worksheet=workbook33.add_worksheet('Marginal_Utilities')
row=1
column=0
counter1=0
cell_format1 = workbook33.add_format({'bold': True, 'font_color': 'green'})
for critirion in res.crit_utilities:
    worksheet.write(row, column, critirion,cell_format1)
    column+=1
    kkk=0
    if float(critirion)==0:
        print(int(critirion))
        cell_format = workbook33.add_format({'bold': True, 'font_color': 'blue'})
        worksheet.write(0, column-1, res.criteria[counter1],cell_format)
        counter1+=1

        for kk in range(1,intervals[counter1-1]+2):
            worksheet.write(2, column-1+kkk,kk)
            kkk+=1
        kkk=0
counter1=0
worksheet=workbook33.add_worksheet('Criterion_Weights')
for m in range(0,7):
    worksheet.write(0, m,res.criteria[m],cell_format)
    worksheet.write(1, m,we[m],cell_format1)
workbook33.close()
print('\n')
print(crit_utilities1)
print(we)
start=0

for y in range(0,len(intervals)):
    list11=list(range(1, intervals[y]+2))

    axis[y].plot(list11,crit_utilities1[start:start+intervals[y]+1],marker='o',color='blue')
    axis[y].set_title(str(y+1))

    start=intervals[y]+start+1
listweights=list(range(0,len(bestvalues)))
axis[len(bestvalues)].plot(listweights,we,marker='o',color='blue')
print(res.crit_utilities)
plt.show()
for ii, hha in zip(res.alternatives,res.alts_utilities):
    print(ii,hha)
print(res.alts_utilities)
print(res.alternatives)