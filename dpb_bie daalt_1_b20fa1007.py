Python 3.11.3 (tags/v3.11.3:f3909b8, Apr  4 2023, 23:49:59) [MSC v.1934 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license()" for more information.
>>> import pandas as pd
... data = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 0)
... dat = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 0)
... dt = data.replace(['и', 'т'], [1,0])
... sumdt = dt.iloc[:,4:36].sum(axis=1)
... dat['Нийт -10'] = sumdt*0.3125
... max1 = max(dat['Нийт -10'])
... min1 = min(dat['Нийт -10'])
... mean1 = dat['Нийт -10'].mean()
... attendance = dat.drop(dat.columns[1], axis = 1)
... data1 = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 1)
... data1['Нийт -25'] = (data1.iloc[:,4:23].sum(axis=1))*25/19
... max2 = max(data1['Нийт -25'])
... min2 = min(data1['Нийт -25'])
... mean2 = data1['Нийт -25'].mean()
... seminar = data1.drop(data1.columns[1], axis = 1)
... data2 = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 2)
... data2['Нийт'] = data2.iloc[:,4:7].sum(axis=1)
... max3 = max(data2['Нийт'])
... min3 = min(data2['Нийт'])
... mean3 = data2['Нийт'].mean()
... exam = data2.drop(data2.columns[1], axis = 1)
... data3 = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 3)
... data3['Нийт'] = data3.iloc[:,4:6].sum(axis=1)
... max4 = max(data3['Нийт'])
... min4 = min(data3['Нийт'])
... mean4 = data3['Нийт'].mean()
... cwork = data3.drop(data3.columns[1], axis = 1)
... data4 = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 4)
... data4['Нийт'] = data4.iloc[:,4:8].sum(axis=1)
... max5 = max(data4['Нийт'])
... min5 = min(data4['Нийт'])
... mean5 = data4['Нийт'].mean()
... behavior = data4.drop(data4.columns[1], axis = 1)
... data5 = pd.read_excel(r"C:\Users\Dell\Downloads\bie daalt-1.xlsx", sheet_name = 5)
... data5['Нийт'] = data5.iloc[:,4:5].sum(axis=1)
... max6 = max(data5['Нийт'])
min6 = min(data5['Нийт'])
mean6 = data5['Нийт'].mean()
additional = data5.drop(data5.columns[1], axis = 1)
tot = (sumdt*0.3125) + (data1.iloc[:,4:23].sum(axis=1))*25/19 + data2.iloc[:,4:7].sum(axis=1) + data3.iloc[:,4:6].sum(axis=1) + data4.iloc[:,4:8].sum(axis=1) + data5.iloc[:,4:5].sum(axis=1)
max7 = max(tot)
min7 = min(tot)
mean7 = tot.mean()
points = dt.iloc[:,0:4]
points['Ирц'] = dt['Нийт -10']
points['Семинар'] = data1['Нийт -25']
points['Шалгалт'] = data2['Нийт']
points['Хүмүүжил хандлага'] = data3['Нийт']
points['Нэмэлт оноо'] = data4['Нийт']
points['Дүн'] = tot
points['Үнэлгээ'] = ''
points['Үнэлгээ'] = points['Дүн'].map(lambda x: 'F' if 0 <= x <= 59 else 
                                      'D-' if 60 <= x <= 62 else 
                                      'D ' if 63 <= x <= 66 else 
                                      'D+' if 67 <= x <= 69 else 
                                      'C-' if 70 <= x <= 72 else
                                      'C ' if 72.5 <= x <= 76.5 else
                                      'C+' if 77 <= x <= 79 else
                                      'B-' if 80 <= x <= 82 else
                                      'B ' if 83 <= x <= 86.4 else
                                      'B+' if 86.5 <= x <= 89 else
                                      'A-' if 90 <= x <= 95 else
                                      'A ' if 96 <= x <= 100 else None)
total = points.drop(points.columns[1], axis = 1)
gradecount = points['Үнэлгээ'].value_counts().sort_index().to_frame()
gcount = gradecount.rename(columns = {'Үнэлгээ': 'Давтамж'})
points['Үнэлгээ'].map(lambda x: 1 if x == 'A' or x == 'A-'else
                                1 if x == 'B' or x == 'B+' or x == 'B-' else 
                                1 if x == 'C' or x == 'C+' or x == 'C-' else
                                1 if x == 'D' or x == 'D+' or x == 'D-' else
                                0 if x == 'F' else None)
success = gradecount.drop('F', axis=0).sum(axis=0)
stud = gradecount.sum(axis = 0)
s = (success / stud)
srate = s.rename(index = {'Үнэлгээ': 'Амжилт'})
ab = gradecount.iloc[0:5].sum()
qualrate = ab/stud
qualrate = qualrate.rename(index = {'Үнэлгээ' : 'Чанар'})
rates = pd.concat([srate, qualrate]).to_frame()
rates = rates.rename(columns = {0:'Хувь'}).style.format('{:.2%}')
#Max, min, mean
df = {'Ирц': [max1,min1,mean1],'Семинар': [max2,min2,mean2],'Шалгалт':[max3,min3,mean3],'Бие даалт':[max4,min4,mean4],'Хүмүүжил хандлага':[max5,min5,mean5],'Нэмэлт оноо':[max6,min6,mean6],'Нийт':[max7,min7,mean7]}
stats = pd.DataFrame(df).rename(index = {0 : 'Хамгийн их', 1 : 'Хамгийн бага',2 :'Дундаж'})
with pd.ExcelWriter('Бие даалт-1.xlsx') as writer:
    attendance.to_excel(writer, sheet_name='Ирц')
    seminar.to_excel(writer, sheet_name='Семинар')    
    exam.to_excel(writer, sheet_name='Шалгалт')
    cwork.to_excel(writer, sheet_name='Бие даалт')
    behavior.to_excel(writer, sheet_name='Хүмүүжил хандлага')
    additional.to_excel(writer, sheet_name='Нэмэлт оноо')
    total.to_excel(writer, sheet_name='Нийт')
    gcount.to_excel(writer, sheet_name ='Үнэлгээний давтамж')
    rates.to_excel(writer, sheet_name ='Амжилт, чанар')
    stats.to_excel(writer, sheet_name ='Их, бага, дундаж')
writer.save()

