import pandas as pd
import numpy as np
import os 
os.chdir(r"C:\Users\USER\Downloads\Бие даалт - 1.xlsx")

a1=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Ирц')
a1['Нийт -10']=a1.apply(lambda row: sum(row=='и')/32*10,axis=1)
# a1.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaaltt1.xlsx',sheet_name='Ирц')
a2=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Семинар')
a2.fillna(0)
b2=a2.drop('No',axis=1)
b2['Нийт -25']=b2.apply(lambda row: sum(row==1)/19*25,axis=1)
# b2.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaaltt1.xlsx',sheet_name='Семинар')
# b2
a3=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Шалгалт')
b3=a3.drop('No',axis=1)
b3['Нийт']=b3.sum(axis=1)
# b3.fillna(0)
a4=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Бие даалт')
b4=a4.drop('No',axis=1)
b4['Нийт']=b4.sum(axis=1)
# b4.fillna(0)
a5=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Хүмүүжил хандлага')
b5=a5.drop('No',axis=1)
b5['Нийт']=b5.sum(axis=1)
# b5.fillna(0)
a6=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Нэмэлт оноо')
b6=a6.drop('No',axis=1)
b6['Нийт']=b6.sum(axis=1)
# b6.fillna(0)
a7=pd.read_excel('Бие Даалт - 1.xlsx', sheet_name='Нийт')
a7['Ирц-10']=pd.Series.replace(a1['Нийт -10'])
a7['Семинар-25']=pd.Series.replace(b2['Нийт -25'])
a7['Шалгалт-30']=pd.Series.replace(b3['Нийт'])
a7['Бие даалт-25']=pd.Series.replace(b4['Нийт'])
a7['ХХ-10']=pd.Series.replace(b5['Нийт'])
a7['Нэмэлт оноо']=pd.Series.replace(b6['Нийт'])
b7=a7.drop('No',axis=1)
b7.drop(b7.columns[[0,10,11,12,13,14,15,16,17,18,19,20,21,22,23]], axis = 1, inplace = True)
b7['Нийт']=b7.sum(axis=1)
b8=b7.drop(28,axis=0)
x1=0
x2=0
x3=0
x4=0
x5=0
for i in b8['Нийт']:
    if i>90:
        x1+=1
    elif 90>i and i>80:
        x2+=1
    elif 80>i and i>70:
        x3+=1 
    elif 70>i and i>60:
        x4+=1
    else:
        x5+=1
b8.at[0,'Үсгэн дүнгүүд']='A'
b8.at[1,'Үсгэн дүнгүүд']='B'
b8.at[2,'Үсгэн дүнгүүд']='C'
b8.at[3,'Үсгэн дүнгүүд']='D'
b8.at[4,'Үсгэн дүнгүүд']='F'
b8.at[5,'Үсгэн дүнгүүд']='Нийт хүүхдүүд'
b8.at[0,'Ангийн дүн']=x1
b8.at[1,'Ангийн дүн']=x2
b8.at[2,'Ангийн дүн']=x3
b8.at[3,'Ангийн дүн']=x4
b8.at[4,'Ангийн дүн']=x5
b8.at[5,'Ангийн дүн']=x1+x2+x3+x4+x5
b8.at[28,'Нийт']=b8['Нийт'].mean()
ratio = {96:"A",90:'A-',87:'B+',83:'B',80:"B-",77:'C+',73:"C",70:'C-',67:'D+',63:'D',60:'D-',0:'F'}
def c(value):
    for key, letter in ratio.items():
        if value >= key:
            return letter
useg_dun= b8['Нийт'].map(c)
b8['Үсгэн дүн'] = pd.Categorical(useg_dun, categories=ratio.values(), ordered=True)
b8['temp'] = b8['Үсгэн дүнгүүд']
b8['Үсгэн дүнгүүд'] = b8['Үсгэн дүн']
b8['Үсгэн дүн']= b8['temp']
b8.drop(columns=['temp'], inplace=True)
passed=(b8['Нийт'] >= 60)
p=passed.sum()
Амжилт=p/28*100
print('Амжилт:',Амжилт)
AaBb=(b8['Нийт'] >= 80)
AB=AaBb.sum()
Чанар= AB/28*100
print('Чанар:',Чанар)
print('Ирц:Хамгийн өндөр оноо:',a1['Нийт -10'].max(), 'Хамгийн бага оноо:',a1['Нийт -10'].min(),'Дундаж оноо:',a1['Нийт -10'].mean())
print('Семинар:Хамгийн өндөр оноо:',b2['Нийт -25'].max(), 'Хамгийн бага оноо:',b2['Нийт -25'].min(),'Дундаж оноо:',b2['Нийт -25'].mean())
print('Шалгалты:Хамгийн өндөр оноо:',b3['Нийт'].max(), 'Хамгийн бага оноо:',b3['Нийт'].min(),'Дундаж оноо:',b3['Нийт'].mean())
print('Бие даалт:Хамгийн өндөр оноо:',b4['Нийт'].max(), 'Хамгийн бага оноо:',b4['Нийт'].min(),'Дундаж оноо:',b4['Нийт'].mean())
print('Хүмүүжил хандлага:Хамгийн өндөр оноо:',b5['Нийт'].max(), 'Хамгийн бага оноо:',b5['Нийт'].min(),'Дундаж оноо:',b5['Нийт'].mean())
print('Нэмэлт оноо:Хамгийн өндөр оноо:',b6['Нийт'].max(), 'Хамгийн бага оноо:',b6['Нийт'].min(),'Дундаж оноо:',b6['Нийт'].mean())
b8

a1.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Ирц')
b2.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Семинар')
b3.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Шалгалт')
b4.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Бие Даалт')
b5.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Хүмүүжил Хандлага')
b6.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Нэмэлт оноо')
b8.to_excel(r'C:\Users\laptop\Documents\DBP221\biydaalt1-3.xlsx',sheet_name='Нийт')



