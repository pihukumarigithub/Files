# -*- coding: utf-8 -*-
"""
Created on Mon Feb 27 12:22:55 2023

@author: priyakgu
"""

import pandas as pd
fileload = pd.read_excel("C:/Users/PRIYAKGU/Downloads/Dashboard_Data_AA.xlsx")
fileload2 = pd.DataFrame(fileload)           
fileload2 = fileload2[fileload2.Final_Status != 'In Progress']
print(fileload2) 

Emails_find = fileload2['User Email'].tolist()


unique_Email1 =(set(Emails_find ))

data = pd.read_excel("C:/Users/PRIYAKGU/Downloads/Dashboard_Data_AA.xlsx", sheet_name ='Assessments',skiprows=1)
data_Assessments = pd.DataFrame(data)
improper_data = data_Assessments.dropna(how = 'any')
specific_column = data_Assessments['Email-id']

duplicate_Rows = specific_column[specific_column.duplicated()].drop_duplicates()
print(duplicate_Rows)

collst = duplicate_Rows.tolist()

unique_Email2=set(collst)
print(len(unique_Email2))


final_Emails=unique_Email1 & unique_Email2
print(len(final_Emails))

final_output = fileload.loc[ fileload['User Email'].isin(final_Emails)]

import smtplib

sender = 'kumaripriya2996@gmail.com'
receivers = ['priyakrigupta96@gmail.com']

message = 'you have completed course successfully'


try:
    connection = smtplib.SMTP('smtp.gmail.com',587)
    connection.starttls()
    connection.login ('kumaripriya2996@gmail.com','kodomqansdfpasre')
    connection.sendmail(sender,receivers,message)
    print ("Successfully sent email")
except Exception as e :
        print ("Error: unable to send email",e)

import sqlite3
conn= sqlite3.connect('AA_Foundation')
c = conn.cursor()
c.execute('CREATE TABLE IF NOT EXISTS AF (Email)')

final_output.to_sql('AF', conn, if_exists = 'replace',index = False)
c.execute('SELECT * FROM AF')
for i in c.fetchall():
    print(i)