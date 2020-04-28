# -*- coding: utf-8 -*-
"""
Created on Mon Apr 27 19:03:54 2020

@author: Ankit Raj
"""
import os
import tabula
import pandas as pd

os.chdir(r'C:\Users\sriva\OneDrive\Desktop\ISM6930_Statistical_Programming_For_BA\Project')

# Extracted table A1-a
df1 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 1)
df1 = pd.DataFrame(df1[0])
df1 = df1.drop([0])
df1['Unnamed: 0'].update(df1.pop('Unnamed: 1'))
df1.columns = ['Selected characteristic', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
sex = df1[:4]
age = df1[4:9]
race = df1[9:25]
education = df1[25:30]
current_employment = df1[30:36]

writer = pd.ExcelWriter('SHS_table1.xlsx')
sex.to_excel(writer, 'sex')
age.to_excel(writer, 'age')
race.to_excel(writer, 'race')
education.to_excel(writer, 'education')
current_employment.to_excel(writer, 'current_employment')
writer.save()

df2 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 2)
df2 = pd.DataFrame(df2[0])

df2.columns = ['Selected characteristic', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
df2 = df2.drop([0])
family_income = df2[:7]
poverty_status = df2[7:11]
health_insurance_coverage = df2[11:24]
marital_status = df2[24:30]
place_of_residence = df2[30:]

writer = pd.ExcelWriter('SHS_table2.xlsx')
family_income.to_excel(writer, 'family_income')
poverty_status.to_excel(writer, 'poverty_status')
health_insurance_coverage.to_excel(writer, 'health_insurance_coverage')
marital_status.to_excel(writer, 'marital_status')
place_of_residence.to_excel(writer, 'place_of_residence')
writer.save()

df3 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 3)
df3 = pd.DataFrame(df3[0])

df3.columns = ['Selected characteristic', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
df3 = df3.drop([0])

writer = pd.ExcelWriter('SHS_table3.xlsx')
df3.to_excel(writer, 'hispanic_or_latino_origin')
writer.save()

# Extracted table A1-b
df4 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 4)
df4 = pd.DataFrame(df4[0])
df4 = df4.drop([0])
df4['Unnamed: 0'].update(df4.pop('Unnamed: 1'))
df4.columns = ['Selected characteristic', 'All adults aged 18 and over', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
sex = df4[:4]
age = df4[4:9]
race = df4[9:25]
education = df4[25:30]
current_employment = df4[30:36]

writer = pd.ExcelWriter('SHS_table4.xlsx')
sex.to_excel(writer, 'sex')
age.to_excel(writer, 'age')
race.to_excel(writer, 'race')
education.to_excel(writer, 'education')
current_employment.to_excel(writer, 'current_employment')
writer.save()

df5 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 5)
df5 = pd.DataFrame(df5[0])

df5.columns = ['Selected characteristic', 'All adults aged 18 and over', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
df5 = df5.drop([0])
family_income = df5[:7]
poverty_status = df5[7:11]
health_insurance_coverage = df5[11:24]
marital_status = df5[24:30]
place_of_residence = df5[30:]

writer = pd.ExcelWriter('SHS_table5.xlsx')
family_income.to_excel(writer, 'family_income')
poverty_status.to_excel(writer, 'poverty_status')
health_insurance_coverage.to_excel(writer, 'health_insurance_coverage')
marital_status.to_excel(writer, 'marital_status')
place_of_residence.to_excel(writer, 'place_of_residence')
writer.save()

df6 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 6)
df6 = pd.DataFrame(df6[0])

df6.columns = ['Selected characteristic', 'All adults aged 18 and over', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
df6 = df6.drop([0])

writer = pd.ExcelWriter('SHS_table6.xlsx')
df6.to_excel(writer, 'hispanic_or_latino_origin')
writer.save()

# Extracted table A1-c
df7 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 7)
df7 = pd.DataFrame(df7[0])
df7 = df7.drop([0])
df7['Unnamed: 0'].update(df7.pop('Unnamed: 1'))
df7.columns = ['Selected characteristic', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
sex = df7[:4]
age = df7[4:9]
race = df7[9:25]
education = df7[25:30]
current_employment = df7[30:36]

writer = pd.ExcelWriter('SHS_table7.xlsx')
sex.to_excel(writer, 'sex')
age.to_excel(writer, 'age')
race.to_excel(writer, 'race')
education.to_excel(writer, 'education')
current_employment.to_excel(writer, 'current_employment')
writer.save()

df8 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 8)
df8 = pd.DataFrame(df8[0])

df8.columns = ['Selected characteristic', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
df8 = df8.drop([0])
family_income = df8[:7]
poverty_status = df8[7:11]
health_insurance_coverage = df8[11:24]
marital_status = df8[24:30]
place_of_residence = df8[30:]

writer = pd.ExcelWriter('SHS_table8.xlsx')
family_income.to_excel(writer, 'family_income')
poverty_status.to_excel(writer, 'poverty_status')
health_insurance_coverage.to_excel(writer, 'health_insurance_coverage')
marital_status.to_excel(writer, 'marital_status')
place_of_residence.to_excel(writer, 'place_of_residence')
writer.save()

df9 = tabula.read_pdf('2018_SHS_Table_A-1.pdf', multiple_tables = True, pages = 9)
df9 = pd.DataFrame(df9[0])

df9.columns = ['Selected characteristic', 'All types of heart disease', 'Coronary heart disease', 'Hypertension', 'Stroke']
df9 = df9.drop([0])

writer = pd.ExcelWriter('SHS_table9.xlsx')
df9.to_excel(writer, 'hispanic_or_latino_origin')
writer.save()
