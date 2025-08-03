# Reading an excel file using Python
import xlrd
import csv
import xlsxwriter
loc3 = ("Apaithseis_Xrhsth.xlsx")
wb3 = xlrd.open_workbook(loc3)
user_sheet = wb3.sheet_by_index(0)
from random import randint
# Give the location of the file
loc = ("kritiria.xlsx")
loc2=("xarakthristika_kinhtwn.xlsx")
wb2 = xlrd.open_workbook(loc2)
# To open Workbook
wb = xlrd.open_workbook(loc)
listA=[2,3,5,7]
listB=[1,6,10]
Brands={
     1: "Xiaomi",
     2: "Apple",
     3: "Samsung",
     4: "Huawei",
     5: "OnePlus",
     6: "LG",
     7: "Nokia"
}
titloi=("Apple iPhone 12 Pro",
"Samsung Galaxy A71",
"OnePlus Nord",
"Samsung Galaxy A51",
"Xiaomi Poco X3",
"Xiaomi Redmi Note 9",
"Apple iPhone 12 Mini",
"Samsung Galaxy S20+"
        )
def get_key(val):
    for key, value in Brands.items():
         if val == value:
             return key


sheet1 = wb2.sheet_by_index(0)
ratings1=[]
ratings2=[]
ratingsRanking=[]
for user in range(1,2):
    kritiria3=["Alternative","Epidoseis","Apeikonish","Diarkeia Xrhshs","Emfanish",'Leitourgiko',"Timh","Marka","Rank"]


    AllhMarka=(user_sheet.cell_value(user, 32))
    print(user_sheet.cell_value(user, 31),AllhMarka)
    AlloLeit=(user_sheet.cell_value(user, 33))
    workbook111 = xlsxwriter.Workbook('dedomena_apokleismou.xlsx')
    markes=(user_sheet.cell_value(user, 11))
    for x in range(1,10):
        l=markes.find(str(x))
        print(markes.find(str(x)))
        if l!=-1:
            print(Brands[l+1])
            break


    worksheet5 = workbook111.add_worksheet()
    worksheet5.write('A1', AllhMarka)
    worksheet5.write('B1',AlloLeit)
    worksheet5.write('C1', Brands[l+1])
    worksheet5.write('D1', user_sheet.cell_value(user, 8))
    workbook111.close()
    Multicriteria=[]

    for kinhto in range(1, sheet1.nrows): #gia ka8e kinhto
       kritiria3 = ["Alternative", "Epidoseis", "Apeikonish", "Diarkeia Xrhshs", "Emfanish", 'Leitourgiko', "Timh",
                     "Marka", "Rank"]
       ratings=[]
       ratings2 = []
       sum=0
       sumA = 0
       sumB = 0
       for sheetcounter in range(1,12): #gia ka8e sheet, ara gia ka8e xarakthristiko sto kritiria.xlsx

            k = (sheet1.cell_value(kinhto, sheetcounter))
            print('k=',k)
            l=user_sheet.cell_value(user, sheetcounter-1)
            print('l=',l)

            if sheetcounter==9:
                if k==1.0:
                    if l=='Android':
                        k1=1
                    else:
                        k1=0.4
                else:
                    if l=='iOS':
                        k1=1
                    else:
                        k1=0.2

                ratings.append(k1)
                ratings2.append(b)

            else:

                sheet = wb.sheet_by_index(sheetcounter)
                for i in range(1,sheet.ncols):
                    a=sheet.cell_value(0,i)
                    if isinstance(a,float):

                        if k<=int(a):
                            x1 = i
                            break
                    elif '*' in str(a):
                        temp=a.split("*")
                        if k>int(temp[0]) and k<=int(temp[1]):
                            x1 = i

                            break

                    elif '+' in str(a):

                        temp2=a.split("+")[0]

                        if k>=int(temp2):
                            x1=i

                            break


                for y in range(1,sheet.nrows): #gia ka8e seira me metrhseis xarakthristikwn
                    b = sheet.cell_value(y, 0)

                    if sheetcounter<=3:

                        b= b.split("*")

                        if l>=int(b[0]) and l<=int(b[1]):

                            y1=y
                            break
                    elif sheetcounter>3:

                        if b==l:

                            y1 = y
                if (sheetcounter) in listA:
                    sumA=sumA+ sheet.cell_value(y1,x1)/4
                elif (sheetcounter) in listB:
                    sumB=sumB+sheet.cell_value(y1,x1)/3
                else:

                    ratings2.append(b)

                    ratings.append(sheet.cell_value(y1,x1))
       marka_kin = (sheet1.cell_value(kinhto, 12))
       m=get_key(marka_kin)

       marka_user = str(user_sheet.cell_value(user, 11))
       m1=int(marka_user[m-1])



       if sumA < 0.39: #epidoseis
           k = 1
       elif sumA >= 0.4 and sumA <= 0.65:
           k = 2
       elif sumA>= 0.66 and sumA<= 0.8:
           k = 3
       elif sumA >= 0.81 and sumA<= 1:
           k = 4
       ratings1.append(k) #apeikonisi

       if sumB < 0.39:
           k = 1
       elif sumB >= 0.4 and sumB <= 0.65:
           k = 2
       elif sumB >= 0.66 and sumB <= 0.94:
           k = 3
       elif sumB >= 0.95 and sumB <= 1:
           k = 4
       ratings1.append(k)

       if ratings[0] < 0.3: #4 mpataria
           k = 1
       elif ratings[0]  >= 0.31 and ratings[0]  <= 0.6:
           k = 2
       elif ratings[0]  >= 0.66 and ratings[0]  <= 1:
           k = 3
       ratings1.append(k)

       if ratings[1] < 0.3:#8 emfanish

           k = 1
       elif ratings[1]  >= 0.31 and ratings[1]  <= 0.79:
           k = 2
       elif ratings[1]  >= 0.8 and ratings[1]  <= 1:
           k = 3
       ratings1.append(k)

       if ratings[2] < 0.21:#9 leitourgiko
           k = 1
       elif ratings[2]  >= 0.22 and ratings[2]  <= 0.4:
           k = 2
       elif ratings[2]  >= 0.41 and ratings[2]  <= 1:
           k = 3

       if (AlloLeit=='Ναι'):
            ratings1.append(k)
       else:
            for a in kritiria3:
                print(a)
            print(len(kritiria3))
            index=kritiria3.index('Leitourgiko')
            kritiria3.pop(index)



       if ratings[3] < 0.11: #10 timh
           k = 1
       elif ratings[3] >= 0.11 and ratings[3] <= 0.2:
           k = 2
       #elif ratings[3] >= 0.21and ratings[3] <= 0.31:
       #    k = 3
       elif ratings[3] >= 0.21 and ratings[3] <= 0.4:
           k = 3
       elif ratings[3] >= 0.41 and ratings[3] <= 1:
           k = 4

       ratings1.append(k)
       if m1== 1:
           marka_timh = 1

       elif m1 == 2:
           marka_timh = 0.7
       elif m1 == 3:
           marka_timh = 0.4
       elif m1 == 4:
           marka_timh = 0.1
       else:
           marka_timh = 0
       if marka_timh < 0.12: #10 timh
           k = 1
       elif marka_timh >= 0.13 and marka_timh <= 0.4:
           k = 2
       elif marka_timh >= 0.41and marka_timh <= 0.7:
           k = 3
       elif marka_timh >= 0.71 and marka_timh <= 1:
           k = 4

       if (AllhMarka=='Ναι'):
            ratings1.append(k)
       else:
            kritiria3.remove("Marka")
       Multicriteria.append(ratings1)
       ratings1 = []
       counter=0
    for j,x in zip(range(24,32),Multicriteria):
           ratings1.append(int(user_sheet.cell_value(user,j)))
           #print('multi=',x)
           Multicriteria[counter].extend([int(user_sheet.cell_value(user,j))])
           counter=counter+1



    ratings1=[]
    print("Multicriteria Array")
    x=0
    for a in Multicriteria:
        x=x+1
        if x<=8:
            print(a)
    counter=0
    for j, x in zip(range(0, 8), Multicriteria):
        Multicriteria[counter].insert(0,titloi[j])
        counter = counter + 1
    with xlsxwriter.Workbook('bathmologies_kinhtwn_sunolou_anaforas.xlsx') as workbook:
        worksheet = workbook.add_worksheet()
        worksheet.write_row(0,0,kritiria3)

        for row_num, data in enumerate(Multicriteria):
            if row_num<8:
                worksheet.write_row(row_num+1, 0, data)
    with xlsxwriter.Workbook('bathmologies_kinhtwn_sunolou_agoras.xlsx') as workbook:
        worksheet = workbook.add_worksheet()

        for row_num, data in enumerate(Multicriteria):
            if row_num>=8:
                worksheet.write_row(row_num-8, 0, data)
    Multicriteria = []

