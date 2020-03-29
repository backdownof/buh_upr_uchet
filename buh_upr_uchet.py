# -*- coding: utf-8 -*-
import json
from pprint import pprint
import pandas as pd
import os, os.path

def go_through_df_keys(fbuhpd):
    #l = []
    i = 0
    for item in fbuhpd[fbuhpd.columns[0]]:
        #print(fbuhpd[item])
        if item == '№':
            return i
        i += 1
    i = 0
    for item in fbuhpd[fbuhpd.columns[1]]:
        if item == '№':
            return i
        i += 1
    
sved = {}
objSet = []
    
for name in os.listdir('сводить/'):
    if os.path.isfile(os.path.join('сводить/', name)):
        objSet.append(str(name))
        #print(name)
        fbuhpd = pd.read_excel('сводить/' + name)
        fbuhpd.reset_index(inplace=True)
        numHeadCol = go_through_df_keys(fbuhpd)
        #print(numHeadCol)
        fbuhpd.columns = fbuhpd.iloc[numHeadCol]
        fbuhpd = fbuhpd[fbuhpd.index != numHeadCol]
        fbuhpd = fbuhpd.fillna(0)
        summname = ''
        kUder = ''
        zpBuh = ''
        if 'СУММА 1-15' in fbuhpd.columns:
            summname = 'СУММА 1-15'
        elif 'З/п 1-15 ' in fbuhpd.columns:
            summname = 'З/п 1-15 '
        if 'К удержанию' in fbuhpd.columns:
            kUder = 'К удержанию'
        if 'ЗП Бухгалтерия' in fbuhpd.columns:
            zpBuh = 'ЗП Бухгалтерия'
        
        for index, row in fbuhpd.iterrows():
            if row['Ф.И.О.'] != 0 and row['Ф.И.О.'] != 'ИТОГО:':
                #print(row)
                if row['Ф.И.О.'] not in sved.keys():
                    sved[row['Ф.И.О.']] = {'Сумма денег': row[summname]}
                else:
                    sved[row['Ф.И.О.']]['Сумма денег'] += row[summname]
                
                if kUder != '' and row[kUder].values[0] != 0:
                    if 'К удержанию' not in sved[row['Ф.И.О.']].keys():
                        sved[row['Ф.И.О.']]['К удержанию'] = row[kUder].values[0]
                    else:
                        sved[row['Ф.И.О.']]['К удержанию'] += row[kUder].values[0]
                
                if zpBuh != '' and row[zpBuh].values[0] != 0:
                    #print(type(row[zpBuh]))
                    #print(row[zpBuh])
                    if 'ЗП Бухгалтерия' not in sved[row['Ф.И.О.']].keys():
                        sved[row['Ф.И.О.']]['ЗП Бухгалтерия'] = row[zpBuh].values[0]
                    else:
                        sved[row['Ф.И.О.']]['ЗП Бухгалтерия'] += row[zpBuh].values[0]
            
                if 'Объект' not in sved[row['Ф.И.О.']]:
                    sved[row['Ф.И.О.']]['Объект'] = {name: row[summname]}
                else:
                    sved[row['Ф.И.О.']]['Объект'][name] = row[summname]
                        

buhpd = pd.read_excel('buh_uch.xls').fillna(0)
uprpd = pd.read_excel('fin_uch.xls').fillna(0)


def nameIsOk(name):
    if name != '<...>' and not name.startswith('пп.') and name != 70 and name != 'Вид начислений оплаты труда':
        return True
    else:
        return False
    
# for index, row in buhpd.iterrows():
#     rab = row['Работники организаций']
#     if nameIsOk(rab):
#         print(rab)



for index, row in buhpd.iterrows():
    rab = row['Работники организаций']
    if nameIsOk(rab):
        excelname = ''
        buhname = ''
        for name in sved.keys():
            fullname = rab.split(' ')
            abbreviat = name.split(' ')

            #print(fullname)
            if fullname[0] == abbreviat[0]:
                if len(abbreviat) > 1:
                    secondab = abbreviat[1].split('.')
                    if fullname[1].startswith(secondab[0]):
                        excelname = name
                        buhname = rab
                    
                else:
                    excelname = name
                    buhname = rab
            
            buhname = rab

        #if buhNameIsInExcelName:
        credit = 0
        debet = row['Дебет.1']
        if row['Кредит.1'] != 0:
            credit = row['Кредит.1']
        #print(excelname)
        #print(sved[excelname])
        if excelname != '' and buhname != '':
            sved[excelname]['Бух Дебет'] = (debet-credit)
        elif buhname != '' and excelname == '':
            sved[buhname] = {'Бух Дебет': (debet-credit)}
        #else:
        #    sved[buhname]['Бух Дебет'] = (debet-credit)


for index, row in uprpd.iterrows():    
    if row['Дебет'] == 70:
        deb = 9
        rab = row['Аналитика Дт']
        col = 'Упр Дебет'
    elif row['Дебет'] == 51:
        deb = 14
        rab = row['Аналитика Кт']
        col = 'Упр Кредит'
    else:
        continue

    excelname = ''
    buhname = ''
    uprname = ''
    for name in sved.keys():
        fullname = rab.split(' ')
        abbreviat = name.split(' ')

        if fullname[0] == abbreviat[0]:
            if len(abbreviat) > 1:
                secondab = abbreviat[1].split('.')
                if fullname[1].startswith(secondab[0]):
                    excelname = name
                    #uprname = ' '.join(rab.split(' ')[:3])

            else:
                excelname = name
                #uprname = ' '.join(rab.split(' ')[:3])

        uprname = ' '.join(rab.split(' ')[:2])
        
    if excelname != '' and uprname != '':
        sved[excelname][col] = row[(f"Unnamed: {int(deb)}")]
    elif uprname != '' and excelname == '':
        sved[uprname] = {col: row[(f"Unnamed: {int(deb)}")]}





for item in sved.keys():
    if 'Упр Дебет' in sved[item] and 'Упр Кредит' in sved[item]:
        if sved[item]['Упр Дебет'] == sved[item]['Упр Кредит']:
            sved[item]['Упр Дебет'] = '-'
        else:
            sved[item]['Упр Дебет'] -= sved[item]['Упр Кредит']
    if 'Сумма денег' in sved[item]:
        if 'К удержанию' in sved[item] and 'Бух Дебет' not in sved[item]:
            sved[item]['На руки'] = sved[item]['Сумма денег'] + sved[item]['К удержанию']
        elif 'К удержанию' not in sved[item] and 'Бух Дебет' in sved[item]:
            sved[item]['На руки'] = sved[item]['Сумма денег'] - sved[item]['Бух Дебет']
        elif 'К удержанию' in sved[item] and 'Бух Дебет' in sved[item]:
            sved[item]['На руки'] = sved[item]['Сумма денег'] - sved[item]['Бух Дебет'] + sved[item]['К удержанию']
        else:
            sved[item]['На руки'] = sved[item]['Сумма денег']
            
        for key, value in sved[item]['Объект'].items():
            if (sved[item]['Сумма денег'] != 0):
                sved[item][key] = (value / sved[item]['Сумма денег']) * sved[item]['На руки']


        
#        print(row['Unnamed: 9'])
finaldf = pd.DataFrame.from_dict(sved, orient='index').fillna('-') 
columns=['Сумма денег', 'К удержанию', 'ЗП Бухгалтерия','Бух Дебет', 'Упр Дебет', 'На руки']
columns += objSet
finaldf = finaldf[columns]
#for name, values in sved.items():
#    for state, summ in sved[name].items():
#        finaldf.append({'Имя': name, state: state})
writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
finaldf.to_excel(writer,'Sheet1', encoding='utf8')
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
format1 = workbook.add_format({'num_format': '# ##0.00'})
format2 = workbook.add_format({'bold': True, 'num_format': '# ##0.00'})
worksheet.set_column('A:A', 25)
worksheet.set_column('B:N', 12, format1)
worksheet.set_column('B:B', 12, format2)
worksheet.set_column('E:E', 12, format2)
worksheet.set_column('G:G', 12, format2)
writer.save()
