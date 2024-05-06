import os
import pandas as pd
import openpyxl
import xlwings as xw



workbook = openpyxl.load_workbook('domain.xlsx')

sheet_name = 'User Creation Form'
sheet = workbook[sheet_name]

print('-----------------Intelligent Image Management Limited-------------------')
print('Start row 15.')
print('End row 15.')
print('OU Name - IIM3-HW, IIM3-MAP, IIM3-MEDICAL, IIM3-Permision, IIM3-Servers')
print('Member of - IIM3-HW-DA, IIM3-MAP-Sup, IIM3-MAP-Research, IIM3-Map-Locator, IIM3-Map-DA')
print('Office Name - IIM3')
print('Department Name - HW, MAP, MEDICAL')
print('-----------------Build By Limon-------------------')

start_row = int(input("Start row : "))
end_row = int(input("End row : "))
ouname = input("Enter Your OU Name : ")
memberof = input("Enter Your Member of Name : ")
officename = input("Enter Your Office Name : ")
department = input("Enter Your Department Name : ")

# start_row, end_row = 15, 24
fullname = 'B'
employeeid = 'C'
designation_col = 'D'

for row in range(start_row, end_row + 1):
    username = sheet[fullname + str(row)].value
    emid = sheet[employeeid + str(row)].value
    desig = sheet[designation_col + str(row)].value
    try:
        os.system(f'dsadd user "cn={username},ou={ouname},dc=iiml,dc=local" -fn "{username}" -samid {emid} -Title "{desig}" -upn {emid}@iiml.local -pwd m@123456 -mustchpwd yes -memberof "cn={memberof},ou={ouname},dc=iiml,dc=local" -company "IIML" -dept "{department}" -office "{officename}"') 
        print(f"{username} {emid} {desig} Domain created. \n")
    except:
        print(f"{username} {emid} {desig} Not created. \n")
workbook.close()
input("Press any key to close.....")