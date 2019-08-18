import openpyxl,pandas as pd
import numpy as np
import os,xlsxwriter
from datetime import datetime,timedelta
from pandas import DataFrame

pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', 5)
pd.set_option('display.width', 1000)



data = pd.read_excel(r'D:\NFR\data.xlsx',sheet_name=0,index_col=0)

writer = pd.ExcelWriter('D:\\NFR\\Bi-Weekly Network Faults Report.xlsx',engine='xlsxwriter')

workbook=writer.book

# workbook = xlsxwriter.Workbook(r'C:\Users\Huawei03\Desktop\NFR\Bi-Weekly Network Faults Report.xlsx')

worksheet = workbook.add_worksheet('To Region')
worksheet1 = workbook.add_worksheet('To District')
worksheet2 = workbook.add_worksheet('By R-Cause')
worksheet3 = workbook.add_worksheet('By RC - Power')
worksheet4 = workbook.add_worksheet('By RC T.Media')
worksheet5 = workbook.add_worksheet('By RC - Equip')
worksheet6 = workbook.add_worksheet('Outages - By RC')
worksheet7 = workbook.add_worksheet('Outages By RC - Power')
worksheet8 = workbook.add_worksheet('Outages By RC - T.Media')
worksheet9 = workbook.add_worksheet('Outages By RC - Equip')
worksheet10 = workbook.add_worksheet('Raw-Data')



bold =workbook.add_format({'bold': True,'bg_color': '#4babc6','border':1})
bg_color = workbook.add_format({'bg_color': '#b7dee8','border':1})

worksheet.set_column('A:D',40)
worksheet1.set_column('A:D',40)
worksheet2.set_column('A:D',40)
worksheet3.set_column('A:D',40)
worksheet4.set_column('A:D',40)
worksheet5.set_column('A:D',40)
worksheet6.set_column('A:D',40)
worksheet7.set_column('A:D',40)
worksheet8.set_column('A:D',40)
worksheet9.set_column('A:D',40)
worksheet10.set_column('A:AH',20)

# data.insert(loc=0,column='S.No',value=np.arange(len(data)))
# data.insert(loc=1,column='Time in sec',value= datetime.datetime.now().date()- data['Alarm Start Time'])
            
# print(data)

# print(type(data))

minimum_date = datetime.date(min(data['Ticket Closed Date']))
maximum_date = datetime.date(max(data['Ticket Closed Date']))


current_week = maximum_date - timedelta(14)
sec_last_week = current_week - timedelta(14)
third_last_week = minimum_date + timedelta(14)


current_week_label = current_week + timedelta(1)


# all PRs
third_last_week_data = data[(pd.to_datetime(data['Ticket Closed Date']) >= minimum_date) & (pd.to_datetime(data['Ticket Closed Date']) <= third_last_week)]
second_last_week_data = data[(pd.to_datetime(data['Ticket Closed Date']) >= sec_last_week) & (pd.to_datetime(data['Ticket Closed Date']) <= current_week)]
current_week_data = data[(pd.to_datetime(data['Ticket Closed Date']) >= current_week) & (pd.to_datetime(data['Ticket Closed Date']) <= maximum_date)]





# print(third_last_week_data)
# print(os.getcwd())

# current_week_data.to_csv(r'C:\Users\Huawei03\Desktop\NFR\current_week_data.csv')
# second_last_week_data.to_csv(r'C:\Users\Huawei03\Desktop\NFR\second_last_week_data.csv')
# third_last_week_data.to_csv(r'C:\Users\Huawei03\Desktop\NFR\third_last_week_data.csv')
# 
# print("min date "+ str(minimum_date))
# 
# print("max date " + str(maximum_date))
# print("current week " + str(current_week))
# print("second last week " + str(sec_last_week))
# print("third last week " + str(third_last_week))
# print(third_last_week_data)



############################################################################
#                                                                          #  
#                             By Region                                    #  
#                                                                          #
############################################################################

#current week
central_current = current_week_data.Region.str.count("Central").sum()
west_current = current_week_data.Region.str.count("West").sum()
south_current = current_week_data.Region.str.count("South").sum()
east_current = current_week_data.Region.str.count("East").sum()
total_current = central_current + west_current + south_current + east_current



#second last weekt
central_second = second_last_week_data.Region.str.count("Central").sum()
west_second = second_last_week_data.Region.str.count("West").sum()
south_second = second_last_week_data.Region.str.count("South").sum()
east_second = second_last_week_data.Region.str.count("East").sum()
total_second = central_second + west_second + south_second + east_second



#third last weekt
central_third = third_last_week_data.Region.str.count("Central").sum()
west_third = third_last_week_data.Region.str.count("West").sum()
south_third = third_last_week_data.Region.str.count("South").sum()
east_third = third_last_week_data.Region.str.count("East").sum()
total_third = central_third + west_third + south_third + east_third



worksheet.write_string(0,0, 'Region',bold)
worksheet.write_string(1,0, 'Central',bg_color)
worksheet.write_string(2,0, 'West',bg_color)
worksheet.write_string(3,0, 'South',bg_color)
worksheet.write_string(4,0, 'East',bg_color)
worksheet.write_string(5,0, 'Grand Total',bold)

worksheet.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet.write_number(1,1, central_third,bg_color)
worksheet.write_number(2,1, west_third,bg_color)
worksheet.write_number(3,1, south_third,bg_color)
worksheet.write_number(4,1, east_third,bg_color)
worksheet.write_number(5,1, central_third+west_third+south_third+east_third,bold)

worksheet.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet.write_number(1,2,central_second,bg_color)
worksheet.write_number(2,2,west_second,bg_color)
worksheet.write_number(3,2,south_second,bg_color)
worksheet.write_number(4,2,east_second,bg_color)
worksheet.write_number(5,2,central_second+west_second+south_second+east_second,bold)


worksheet.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet.write_number(1,3,central_current,bg_color)
worksheet.write_number(2,3,west_current,bg_color)
worksheet.write_number(3,3,south_current,bg_color)
worksheet.write_number(4,3,east_current,bg_color)
worksheet.write_number(5,3,central_current+west_current+south_current+east_current,bold)

chart = workbook.add_chart({'type': 'column'})
chart.add_series({
  'name': '=To Region!$B$1',
  'categories': '= '"'To Region'"'!$A$2:$A$5',
  'values': '='"'To Region'"'!$B$2:$B$5',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart.add_series({
  'name': '='"'To Region'"'!$C$1',
  'categories': '= '"'To Region'"'!$A$2:$A$5',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'To Region'"'!$C$2:$C$5',
  'data_labels': {'value': True}})
  
chart.add_series({
  'name': '=To Region!$D$1',
  'categories': '= '"'To Region'"'!$A$2:$A$5',
  'values': '= '"'To Region'"'!$D$2:$D$5',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart.set_y_axis({
    'name': 'Total  Network  Faults  - By Region',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart.set_legend({'position': 'bottom'})
  
chart.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart.set_chartarea({
    'border': {'none': True}
    
})
# chart.set_x_axis({
#     'major_gridlines': {
#         'visible': True,
#         'line': {'width': 1.25, 'dash_type': 'dash'}
#     },
# })
  
  
worksheet.insert_chart('A9',chart,{'x_scale': 2, 'y_scale': 1})


# insert callout


row = 9
col = 2
text = "Total PRs = " + str(central_current+west_current+south_current+east_current)
options = {
    'width': 180,
    'height': 50,
     'border': {'color': 'blue',
               'width': 1,
               'dash_type': 'round_dot'},
     'align': {'vertical': 'middle',
               'horizontal': 'center'},
    'font': {'bold': True,
             'italic': False,
             
             'name': 'Arial',
             'color': 'black',
             'size': 12},
     # 'fill': {'color': '#ccd9d9'},
             
}

worksheet.insert_textbox(row, col, text, options)



# 
# worksheet1 = workbook.add_worksheet('To District')
# worksheet1.write(0,0, 'Region')


############################################################################
#                                                                          #  
#                             By District                                  #  
#                                                                          #
############################################################################

#current week

Riyadh_City_current = current_week_data.District.str.count("Riyadh City").sum()
Asir_current = current_week_data.District.str.count("Asir").sum()
Jeddah_current = current_week_data.District.str.count("Jeddah").sum()
Riyadh_District_current = current_week_data.District.str.count("Riyadh District").sum()
Najran_current = current_week_data.District.str.count("Najran").sum()
Madinah_current = current_week_data.District.str.count("Madinah").sum()
Makkah_current = current_week_data.District.str.count("Makkah").sum()
Dammam_current = current_week_data.District.str.count("Dammam").sum()
Jizan_current = current_week_data.District.str.count("Jizan").sum()
Ahsa_current = current_week_data.District.str.count("Ahsa").sum()
Taif_current = current_week_data.District.str.count("Taif").sum()
Jubail_current = current_week_data.District.str.count("Jubail").sum()
Tabuk_current = current_week_data.District.str.count("Tabuk").sum()
Baha_current = current_week_data.District.str.count("Baha").sum()
Qassim_current = current_week_data.District.str.count("Qassim").sum()
Northern_Border_current = current_week_data.District.str.count("Northern Border").sum()
Hail_current = current_week_data.District.str.count("Hail").sum()
Jouf_current = current_week_data.District.str.count("Jouf").sum()
Yanbu_current = current_week_data.District.str.count("Yanbu").sum()



#second last weekt
Riyadh_City_second = second_last_week_data.District.str.count("Riyadh City").sum()
Asir_second = second_last_week_data.District.str.count("Asir").sum()
Jeddah_second = second_last_week_data.District.str.count("Jeddah").sum()
Riyadh_District_second = second_last_week_data.District.str.count("Riyadh District").sum()
Najran_second = second_last_week_data.District.str.count("Najran").sum()
Madinah_second = second_last_week_data.District.str.count("Madinah").sum()
Makkah_second = second_last_week_data.District.str.count("Makkah").sum()
Dammam_second = second_last_week_data.District.str.count("Dammam").sum()
Jizan_second = second_last_week_data.District.str.count("Jizan").sum()
Ahsa_second = second_last_week_data.District.str.count("Ahsa").sum()
Taif_second = second_last_week_data.District.str.count("Taif").sum()
Jubail_second = second_last_week_data.District.str.count("Jubail").sum()
Tabuk_second = second_last_week_data.District.str.count("Tabuk").sum()
Baha_second = second_last_week_data.District.str.count("Baha").sum()
Qassim_second = second_last_week_data.District.str.count("Qassim").sum()
Northern_Border_second = second_last_week_data.District.str.count("Northern Border").sum()
Hail_second = second_last_week_data.District.str.count("Hail").sum()
Jouf_second = second_last_week_data.District.str.count("Jouf").sum()
Yanbu_second = second_last_week_data.District.str.count("Yanbu").sum()


#third last week
Riyadh_City_third = third_last_week_data.District.str.count("Riyadh City").sum()
Asir_third = third_last_week_data.District.str.count("Asir").sum()
Jeddah_third = third_last_week_data.District.str.count("Jeddah").sum()
Riyadh_District_third = third_last_week_data.District.str.count("Riyadh District").sum()
Najran_third = third_last_week_data.District.str.count("Najran").sum()
Madinah_third = third_last_week_data.District.str.count("Madinah").sum()
Makkah_third = third_last_week_data.District.str.count("Makkah").sum()
Dammam_third = third_last_week_data.District.str.count("Dammam").sum()
Jizan_third = third_last_week_data.District.str.count("Jizan").sum()
Ahsa_third = third_last_week_data.District.str.count("Ahsa").sum()
Taif_third = third_last_week_data.District.str.count("Taif").sum()
Jubail_third = third_last_week_data.District.str.count("Jubail").sum()
Tabuk_third = third_last_week_data.District.str.count("Tabuk").sum()
Baha_third = third_last_week_data.District.str.count("Baha").sum()
Qassim_third = third_last_week_data.District.str.count("Qassim").sum()
Northern_Border_third = third_last_week_data.District.str.count("Northern Border").sum()
Hail_third = third_last_week_data.District.str.count("Hail").sum()
Jouf_third = third_last_week_data.District.str.count("Jouf").sum()
Yanbu_third = third_last_week_data.District.str.count("Yanbu").sum()




worksheet1.write_string(0,0, 'District',bold)
worksheet1.write_string(1,0, 'Riyadh City',bg_color)
worksheet1.write_string(2,0, 'Jeddah',bg_color)
worksheet1.write_string(3,0, 'Asir',bg_color)
worksheet1.write_string(4,0, 'Makkah',bg_color)
worksheet1.write_string(5,0, 'Riyadh District',bg_color)
worksheet1.write_string(6,0, 'Madinah',bg_color)
worksheet1.write_string(7,0, 'Qassim',bg_color)
worksheet1.write_string(8,0, 'Jizan',bg_color)
worksheet1.write_string(9,0, 'Dammam',bg_color)
worksheet1.write_string(10,0, 'Taif',bg_color)
worksheet1.write_string(11,0, 'Jubail',bg_color)
worksheet1.write_string(12,0, 'Tabuk',bg_color)
worksheet1.write_string(13,0, 'Northern Border',bg_color)
worksheet1.write_string(14,0, 'Najran',bg_color)
worksheet1.write_string(15,0, 'Hail',bg_color)
worksheet1.write_string(16,0, 'Ahsa',bg_color)
worksheet1.write_string(17,0, 'Baha',bg_color)
worksheet1.write_string(18,0, 'Jouf',bg_color)
worksheet1.write_string(19,0, 'Yanbu',bg_color)



worksheet1.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet1.write_number(1,1,Riyadh_City_third,bg_color)
worksheet1.write_number(2,1,Jeddah_third,bg_color)
worksheet1.write_number(3,1, Asir_third,bg_color)
worksheet1.write_number(4,1, Makkah_third,bg_color)
worksheet1.write_number(5,1, Riyadh_District_third,bg_color)
worksheet1.write_number(6,1, Madinah_third,bg_color)
worksheet1.write_number(7,1, Qassim_third,bg_color)
worksheet1.write_number(8,1, Jizan_third,bg_color)
worksheet1.write_number(9,1, Dammam_third,bg_color)
worksheet1.write_number(10,1, Taif_third,bg_color)
worksheet1.write_number(11,1, Jubail_third,bg_color)
worksheet1.write_number(12,1, Tabuk_third,bg_color)
worksheet1.write_number(13,1, Northern_Border_third,bg_color)
worksheet1.write_number(14,1, Najran_third,bg_color)
worksheet1.write_number(15,1, Hail_third,bg_color)
worksheet1.write_number(16,1, Ahsa_third,bg_color)
worksheet1.write_number(17,1, Baha_third,bg_color)
worksheet1.write_number(18,1, Jouf_third,bg_color)
worksheet1.write_number(19,1, Yanbu_third,bg_color)


worksheet1.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet1.write_number(1,2, Riyadh_City_second,bg_color)
worksheet1.write_number(2,2, Jeddah_second,bg_color)
worksheet1.write_number(3,2, Asir_second,bg_color)
worksheet1.write_number(4,2, Makkah_second,bg_color)
worksheet1.write_number(5,2, Riyadh_District_second,bg_color)
worksheet1.write_number(6,2, Madinah_second,bg_color)
worksheet1.write_number(7,2, Qassim_second,bg_color)
worksheet1.write_number(8,2, Jizan_second,bg_color)
worksheet1.write_number(9,2, Dammam_second,bg_color)
worksheet1.write_number(10,2, Taif_second,bg_color)
worksheet1.write_number(11,2, Jubail_second,bg_color)
worksheet1.write_number(12,2, Tabuk_second,bg_color)
worksheet1.write_number(13,2, Northern_Border_second,bg_color)
worksheet1.write_number(14,2, Najran_second,bg_color)
worksheet1.write_number(15,2, Hail_second,bg_color)
worksheet1.write_number(16,2, Ahsa_second,bg_color)
worksheet1.write_number(17,2, Baha_second,bg_color)
worksheet1.write_number(18,2, Jouf_second,bg_color)
worksheet1.write_number(19,2, Yanbu_second,bg_color)





worksheet1.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet1.write_number(1,3, Riyadh_City_current,bg_color)
worksheet1.write_number(2,3, Jeddah_current,bg_color)
worksheet1.write_number(3,3, Asir_current,bg_color)
worksheet1.write_number(4,3, Makkah_current,bg_color)
worksheet1.write_number(5,3, Riyadh_District_current,bg_color)
worksheet1.write_number(6,3, Madinah_current,bg_color)
worksheet1.write_number(7,3, Qassim_current,bg_color)
worksheet1.write_number(8,3, Jizan_current,bg_color)
worksheet1.write_number(9,3, Dammam_current,bg_color)
worksheet1.write_number(10,3, Taif_current,bg_color)
worksheet1.write_number(11,3, Jubail_current,bg_color)
worksheet1.write_number(12,3, Tabuk_current,bg_color)
worksheet1.write_number(13,3, Northern_Border_current,bg_color)
worksheet1.write_number(14,3, Najran_current,bg_color)
worksheet1.write_number(15,3, Hail_current,bg_color)
worksheet1.write_number(16,3, Ahsa_current,bg_color)
worksheet1.write_number(17,3, Baha_current,bg_color)
worksheet1.write_number(18,3, Jouf_current,bg_color)
worksheet1.write_number(19,3, Yanbu_current,bg_color)



chart1 = workbook.add_chart({'type': 'column'})
chart1.add_series({
  'name': '=To District!$B$1',
  'categories': '= '"'To District'"'!$A$2:$A$20',
  'values': '=To District!$B$2:$B$20',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart1.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'To District'"'!$A$2:$A$20',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'To District'"'!$C$2:$C$20',
  'data_labels': {'value': True}})
  
chart1.add_series({
  'name': '=To District!$D$1',
  'categories': '= '"'To Riyadh_District_second'"'!$A$2:$A$20',
  'values': '=To District!$D$2:$D$20',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart1.set_y_axis({
    'name': 'Total  Network  Faults  - By District',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart1.set_legend({'position': 'bottom'})
  
chart1.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart1.set_chartarea({
    'border': {'none': True}
    
})

worksheet1.insert_chart('A22',chart1,{'x_scale': 2.5, 'y_scale': 1})





# insert callout


row = 24
col = 2
text = "Total PRs = " + str(central_current+west_current+south_current+east_current)
options = {
    'width': 180,
    'height': 50,
     'border': {'color': 'blue',
               'width': 1,
               'dash_type': 'round_dot'},
     'align': {'vertical': 'middle',
               'horizontal': 'center'},
    'font': {'bold': True,
             'italic': False,
             
             'name': 'Arial',
             'color': 'black',
             'size': 12},
     # 'fill': {'color': '#ccd9d9'},
             
}

worksheet1.insert_textbox(row, col, text, options)



############################################################################
#                                                                          #  
#                             By R-Cause                                   #  
#                                                                          #
############################################################################

#current week

DCNSynchronization_current_week = current_week_data['Root Cause'].str.count("DCN/Synchronization").sum()
Environment_current_week = current_week_data['Root Cause'].str.count("Environment").sum()
Equipment_current_week = current_week_data['Root Cause'].str.count("Equipment").sum()
Power_current_week = current_week_data['Root Cause'].str.count("Power").sum()
TransMedia_current_week = current_week_data['Root Cause'].str.count("Trans Media").sum()
# ManMade_current_week = current_week_data['Root Cause'].str.count("Man Made").sum()
# Transmission_current_week = current_week_data['Root Cause'].str.count("Transmission").sum()
total_current_week = current_week_data["Root Cause"].count().sum()
Miscellaneous_current_week = total_current_week - (DCNSynchronization_current_week+Environment_current_week+Equipment_current_week+Power_current_week+TransMedia_current_week)


#second last week

DCNSynchronization_second_week = second_last_week_data['Root Cause'].str.count("DCN/Synchronization").sum()
Environment_second_week = second_last_week_data['Root Cause'].str.count("Environment").sum()
Equipment_second_week = second_last_week_data['Root Cause'].str.count("Equipment").sum()
Power_second_week = second_last_week_data['Root Cause'].str.count("Power").sum()
TransMedia_second_week = second_last_week_data['Root Cause'].str.count("Trans Media").sum()
# ManMade_second_week = second_last_week_data['Root Cause'].str.count("Man Made").sum()
# Transmission_second_week = second_last_week_data['Root Cause'].str.count("Transmission").sum()
total_second_week = second_last_week_data["Root Cause"].count().sum()
Miscellaneous_second_week = total_second_week - (DCNSynchronization_second_week+Environment_second_week+Equipment_second_week+Power_second_week+TransMedia_second_week)

# print(Miscellaneous_second)

#third last week

DCNSynchronization_third_week = third_last_week_data['Root Cause'].str.count("DCN/Synchronization").sum()
Environment_third_week = third_last_week_data['Root Cause'].str.count("Environment").sum()
Equipment_third_week = third_last_week_data['Root Cause'].str.count("Equipment").sum()
Power_third_week = third_last_week_data['Root Cause'].str.count("Power").sum()
TransMedia_third_week = third_last_week_data['Root Cause'].str.count("Trans Media").sum()
# ManMade_third_week = third_last_week_data['Root Cause'].str.count("Man Made").sum()
# Transmission_third_week = third_last_week_data['Root Cause'].str.count("Transmission").sum()
total_third_week = third_last_week_data["Root Cause"].count().sum()
Miscellaneous_third_week = total_third_week - (DCNSynchronization_third_week+ Environment_third_week + Equipment_third_week + Power_third_week+TransMedia_third_week)



worksheet2.write_string(0,0, 'Total Cases/ Causes',bold)
worksheet2.write_string(1,0, "Power",bg_color)
worksheet2.write_string(2,0, "Equipment",bg_color)
worksheet2.write_string(3,0, "Trans Media",bg_color)
worksheet2.write_string(4,0, "DCN/Synchronization",bg_color)
worksheet2.write_string(5,0, "Environment",bg_color)
worksheet2.write_string(6,0, "Miscellaneous",bg_color)
worksheet2.write_string(7,0, "Total Cases",bold)


worksheet2.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet2.write_number(1,1,Riyadh_City_third,bg_color)
worksheet2.write_number(1,1, Power_third_week,bg_color)
worksheet2.write_number(2,1, Equipment_third_week,bg_color)
worksheet2.write_number(3,1, TransMedia_third_week,bg_color)
worksheet2.write_number(4,1, DCNSynchronization_third_week,bg_color)
worksheet2.write_number(5,1, Environment_third_week,bg_color)
worksheet2.write_number(6,1, Miscellaneous_third_week,bg_color)
worksheet2.write_number(7,1, total_third_week,bold)

worksheet2.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet2.write_number(1,2, Power_second_week,bg_color)
worksheet2.write_number(2,2, Equipment_second_week,bg_color)
worksheet2.write_number(3,2, TransMedia_second_week,bg_color)
worksheet2.write_number(4,2, DCNSynchronization_second_week,bg_color)
worksheet2.write_number(5,2, Environment_second_week,bg_color)
worksheet2.write_number(6,2, Miscellaneous_second_week,bg_color)
worksheet2.write_number(7,2, total_second_week,bold)

worksheet2.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet2.write_number(1,3, Power_current_week,bg_color)
worksheet2.write_number(2,3, Equipment_current_week,bg_color)
worksheet2.write_number(3,3, TransMedia_current_week,bg_color)
worksheet2.write_number(4,3, DCNSynchronization_current_week,bg_color)
worksheet2.write_number(5,3, Environment_current_week,bg_color)
worksheet2.write_number(6,3, Miscellaneous_current_week,bg_color)
worksheet2.write_number(7,3, total_current_week,bold)


chart2 = workbook.add_chart({'type': 'column'})
chart2.add_series({
  'name': '=By R-Cause!$B$1',
  'categories': '= '"'By R-Cause'"'!$A$2:$A$7',
  'values': '=By R-Cause!$B$2:$B$7',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart2.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'By R-Cause'"'!$A$2:$A$7',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'By R-Cause'"'!$C$2:$C$7',
  'data_labels': {'value': True}})
  
chart2.add_series({
  'name': '=By R-Cause!$D$1',
  'categories': '= '"'By R-Cause'"'!$A$2:$A$7',
  'values': '=By R-Cause!$D$2:$D$7',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart2.set_y_axis({
    'name': 'Total  Network Faults  - By R/Cause',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart2.set_legend({'position': 'bottom'})
  
chart2.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart2.set_chartarea({
    'border': {'none': True}
    
})

worksheet2.insert_chart('A11',chart2,{'x_scale': 2.5, 'y_scale': 1})



# insert callout


row = 12
col = 2
text = "Total PRs = " + str(central_current+west_current+south_current+east_current)
options = {
    'width': 160,
    'height': 50,
     'border': {'color': 'blue',
               'width': 1,
               'dash_type': 'round_dot'},
     'align': {'vertical': 'middle',
               'horizontal': 'center'},
    'font': {'bold': True,
             'italic': False,
             
             'name': 'Arial',
             'color': 'black',
             'size': 12},
     # 'fill': {'color': '#ccd9d9'},
             
}

worksheet2.insert_textbox(row, col, text, options)

############################################################################
#                                                                          #  
#                             By RC - Power                                #  
#                                                                          #
############################################################################


#current week

SCECOCommercialPower_current_week = current_week_data['Root Cause Detail'].str.count("SCECO - Commercial Power").sum()
Generator_current_week = current_week_data['Root Cause Detail'].str.count("Generator").sum()
Airconditioning_current_week = current_week_data['Root Cause Detail'].str.count("Air Conditioning").sum()
Rectifier_current_week = current_week_data['Root Cause Detail'].str.count("Rectifier").sum()
Circuitbreaker_current_week = current_week_data['Root Cause Detail'].str.count("Circuit breaker").sum()
Battery_current_week = current_week_data['Root Cause Detail'].str.count("Battery").sum()
power_cable_current_week = current_week_data['Root Cause Detail'].str.count("Power Cable").sum()
CustomerPower_current_week = current_week_data['Root Cause Detail'].str.count("Customer Power").sum()
total_current_week = current_week_data['Root Cause'].str.count("Power").sum()
other_current_week = total_current_week - (SCECOCommercialPower_current_week + Generator_current_week + Airconditioning_current_week + Rectifier_current_week + Circuitbreaker_current_week + Battery_current_week + power_cable_current_week + CustomerPower_current_week)



#second last week

SCECOCommercialPower_second_week = second_last_week_data['Root Cause Detail'].str.count("SCECO - Commercial Power").sum()
Generator_second_week = second_last_week_data['Root Cause Detail'].str.count("Generator").sum()
Airconditioning_second_week = second_last_week_data['Root Cause Detail'].str.count("Air Conditioning").sum()
Rectifier_second_week = second_last_week_data['Root Cause Detail'].str.count("Rectifier").sum()
Circuitbreaker_second_week = second_last_week_data['Root Cause Detail'].str.count("Circuit breaker").sum()
Battery_second_week = second_last_week_data['Root Cause Detail'].str.count("Battery").sum()
PowerCable_second_week = second_last_week_data['Root Cause Detail'].str.count("Power Cable").sum()
CustomerPower_second_week = second_last_week_data['Root Cause Detail'].str.count("Customer Power").sum()
total_second_week = second_last_week_data['Root Cause'].str.count("Power").sum()
other_second_week = total_second_week - (PowerCable_second_week + SCECOCommercialPower_second_week + Generator_second_week + Airconditioning_second_week + Rectifier_second_week + Circuitbreaker_second_week + Battery_second_week + CustomerPower_second_week)

#third last week

SCECECommercialPower_third_week = third_last_week_data['Root Cause Detail'].str.count("SCECO - Commercial Power").sum()
Generator_third_week = third_last_week_data['Root Cause Detail'].str.count("Generator").sum()
Airconditioning_third_week = third_last_week_data['Root Cause Detail'].str.count("Air Conditioning").sum()
Rectifier_third_week = third_last_week_data['Root Cause Detail'].str.count("Rectifier").sum()
Circuitbreaker_third_week = third_last_week_data['Root Cause Detail'].str.count("Circuit breaker").sum()
powercable_third_week = third_last_week_data['Root Cause Detail'].str.count("Power Cable").sum()
Battery_third_week = third_last_week_data['Root Cause Detail'].str.count("Battery").sum()
CustomerPower_third_week = third_last_week_data['Root Cause Detail'].str.count("Customer Power").sum()
total_third_week = third_last_week_data['Root Cause'].str.count("Power").sum()
other_third_week = total_third_week - (powercable_third_week + SCECECommercialPower_third_week + Generator_third_week + Airconditioning_third_week + Rectifier_third_week + Circuitbreaker_third_week + Battery_third_week + CustomerPower_third_week )


worksheet3.write_string(0,0, 'Power Type',bold)
worksheet3.write_string(1,0, "SCECO - Commercial Power",bg_color)
worksheet3.write_string(2,0, "Generator",bg_color)
worksheet3.write_string(3,0, "Air Conditioning",bg_color)
worksheet3.write_string(4,0, "Rectifier",bg_color)
worksheet3.write_string(5,0, "Circuit Breaker",bg_color)
worksheet3.write_string(6,0, "Battery",bg_color)
worksheet3.write_string(7,0, "Power Cable",bg_color)
worksheet3.write_string(8,0, "Customer Power",bg_color)
worksheet3.write_string(9,0, "Others",bg_color)
worksheet3.write_string(10,0, "Total Cases (Power)",bold)


worksheet3.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet3.write_number(1,1, SCECECommercialPower_third_week,bg_color)
worksheet3.write_number(2,1, Generator_third_week,bg_color)
worksheet3.write_number(3,1, Airconditioning_third_week,bg_color)
worksheet3.write_number(4,1, Rectifier_third_week,bg_color)
worksheet3.write_number(5,1, Circuitbreaker_third_week,bg_color)
worksheet3.write_number(6,1, Battery_third_week,bg_color)
worksheet3.write_number(7,1, powercable_third_week,bg_color)
worksheet3.write_number(8,1, CustomerPower_third_week,bg_color)
worksheet3.write_number(9,1, other_third_week,bg_color)
worksheet3.write_number(10,1, total_third_week,bold)


worksheet3.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet3.write_number(1,2, SCECOCommercialPower_second_week,bg_color)
worksheet3.write_number(2,2, Generator_second_week,bg_color)
worksheet3.write_number(3,2, Airconditioning_second_week,bg_color)
worksheet3.write_number(4,2, Rectifier_second_week,bg_color)
worksheet3.write_number(5,2, Circuitbreaker_second_week,bg_color)
worksheet3.write_number(6,2, Battery_second_week,bg_color)
worksheet3.write_number(7,2, PowerCable_second_week,bg_color)
worksheet3.write_number(8,2, CustomerPower_second_week,bg_color)
worksheet3.write_number(9,2, other_second_week,bg_color)
worksheet3.write_number(10,2, total_second_week,bold)
  

worksheet3.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet3.write_number(1,3, SCECOCommercialPower_current_week,bg_color)
worksheet3.write_number(2,3, Generator_current_week,bg_color)
worksheet3.write_number(3,3, Airconditioning_current_week,bg_color)
worksheet3.write_number(4,3, Rectifier_current_week,bg_color)
worksheet3.write_number(5,3, Circuitbreaker_current_week,bg_color)
worksheet3.write_number(6,3, Battery_current_week,bg_color)
worksheet3.write_number(7,3, power_cable_current_week,bg_color)
worksheet3.write_number(8,3, CustomerPower_current_week,bg_color)
worksheet3.write_number(9,3, other_current_week,bg_color)
worksheet3.write_number(10,3, total_current_week,bold)




chart3 = workbook.add_chart({'type': 'column'})
chart3.add_series({
  'name': '=By RC - Power!$B$1',
  'categories': '= '"'By RC - Power'"'!$A$2:$A$10',
  'values': '=By RC - Power!$B$2:$B$10',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart3.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'By RC - Power'"'!$A$2:$A$10',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'By RC - Power'"'!$C$2:$C$10',
  'data_labels': {'value': True}})
  
chart3.add_series({
  'name': '=By RC - Power!$D$1',
  'categories': '= '"'By RC - Power'"'!$A$2:$A$10',
  'values': '=By RC - Power!$D$2:$D$10',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart3.set_y_axis({
    'name': 'Breakdown of  Power Cases',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart3.set_legend({'position': 'bottom'})
  
chart3.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart3.set_chartarea({
    'border': {'none': True}
    
})

worksheet3.insert_chart('A13',chart3,{'x_scale': 2.5, 'y_scale': 1})



############################################################################
#                                                                          #  
#                             By RC T.Media                                #  
#                                                                          #
############################################################################

#current week

OSPOFCABLECUT_current_week = current_week_data['Root Cause Detail'].str.count("OSP OF Cable Cut").sum()
OpticalElectricalindoorCables_current_week = current_week_data['Root Cause Detail'].str.count("Optical/Electrical indoor Cables").sum()
OpticalElectricalConnectors_current_week = current_week_data['Root Cause Detail'].str.count("Optical/Electrical Connectors").sum()
OpticalElectricalOutdoorCables_current_week = current_week_data['Root Cause Detail'].str.count("Optical/Electrical Outdoor Cables").sum()
Connections_current_week = current_week_data['Root Cause Detail'].str.count("Connections").sum()
total_current_week = current_week_data['Root Cause'].str.count("Trans Media").sum()
other_current_week = total_current_week - (OSPOFCABLECUT_current_week + OpticalElectricalindoorCables_current_week + OpticalElectricalConnectors_current_week + OpticalElectricalOutdoorCables_current_week)



#second last week

OSPOFCABLECUT_second_week = second_last_week_data['Root Cause Detail'].str.count("OSP OF Cable Cut").sum()
OpticalElectricalindoorCables_second_week = second_last_week_data['Root Cause Detail'].str.count("Optical/Electrical indoor Cables").sum()
OpticalElectricalConnectors_second_week = second_last_week_data['Root Cause Detail'].str.count("Optical/Electrical Connectors").sum()
OpticalElectricalOutdoorCables_second_week = second_last_week_data['Root Cause Detail'].str.count("Optical/Electrical Outdoor Cables").sum()
Connections_second_week = second_last_week_data['Root Cause Detail'].str.count("Connections").sum()
total_second_week = second_last_week_data['Root Cause'].str.count("Trans Media").sum()
other_second_week = total_second_week - (OSPOFCABLECUT_second_week + OpticalElectricalindoorCables_second_week + OpticalElectricalConnectors_second_week + OpticalElectricalOutdoorCables_second_week)



#third last week


OSPOFCABLECUT_third_week = third_last_week_data['Root Cause Detail'].str.count("OSP OF Cable Cut").sum()
OpticalElectricalindoorCables_third_week = third_last_week_data['Root Cause Detail'].str.count("Optical/Electrical indoor Cables").sum()
OpticalElectricalConnectors_third_week = third_last_week_data['Root Cause Detail'].str.count("Optical/Electrical Connectors").sum()
OpticalElectricalOutdoorCables_third_week = third_last_week_data['Root Cause Detail'].str.count("Optical/Electrical Outdoor Cables").sum()
total_third_week = third_last_week_data['Root Cause'].str.count("Trans Media").sum()
other_third_week = total_third_week - (OSPOFCABLECUT_third_week + OpticalElectricalindoorCables_third_week + OpticalElectricalConnectors_third_week + OpticalElectricalOutdoorCables_third_week)


worksheet4.write_string(0,0, 'TransMedia Type',bold)
worksheet4.write_string(1,0, "OSP OF Cable Cut",bg_color)
worksheet4.write_string(2,0, "Optical/Electrical indoor Cables",bg_color)
worksheet4.write_string(3,0, "Optical/Electrical Connectors",bg_color)
worksheet4.write_string(4,0, "Optical/Electrical Outdoor Cables",bg_color)
worksheet4.write_string(5,0, "Others",bg_color)
worksheet4.write_string(6,0, "Total Cases (Trans Media)",bold)




worksheet4.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet4.write_number(1,1, OSPOFCABLECUT_third_week,bg_color)
worksheet4.write_number(2,1, OpticalElectricalindoorCables_third_week,bg_color)
worksheet4.write_number(3,1, OpticalElectricalConnectors_third_week,bg_color)
worksheet4.write_number(4,1, OpticalElectricalOutdoorCables_third_week,bg_color)
worksheet4.write_number(5,1, other_third_week,bg_color)
worksheet4.write_number(6,1, total_third_week,bold)

worksheet4.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet4.write_number(1,2, OSPOFCABLECUT_second_week,bg_color)
worksheet4.write_number(2,2, OpticalElectricalindoorCables_second_week,bg_color)
worksheet4.write_number(3,2, OpticalElectricalConnectors_second_week,bg_color)
worksheet4.write_number(4,2, OpticalElectricalOutdoorCables_second_week,bg_color)
worksheet4.write_number(5,2, other_second_week,bg_color)
worksheet4.write_number(6,2, total_second_week,bold)


worksheet4.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet4.write_number(1,3, OSPOFCABLECUT_current_week,bg_color)
worksheet4.write_number(2,3, OpticalElectricalindoorCables_current_week,bg_color)
worksheet4.write_number(3,3, OpticalElectricalConnectors_current_week,bg_color)
worksheet4.write_number(4,3, OpticalElectricalOutdoorCables_current_week,bg_color)
worksheet4.write_number(5,3, other_current_week,bg_color)
worksheet4.write_number(6,3, total_current_week,bold)



chart4 = workbook.add_chart({'type': 'column'})
chart4.add_series({
  'name': '=By RC T.Media!$B$1',
  'categories': '= '"'By RC T.Media'"'!$A$2:$A$6',
  'values': '=By RC T.Media!$B$2:$B$6',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart4.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'By RC T.Media'"'!$A$2:$A$6',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'By RC T.Media'"'!$C$2:$C$6',
  'data_labels': {'value': True}})
  
chart4.add_series({
  'name': '=By RC T.Media!$D$1',
  'categories': '= '"'By RC T.Media'"'!$A$2:$A$6',
  'values': '=By RC T.Media!$D$2:$D$6',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart4.set_y_axis({
    'name': 'Breakdown of  TransMedia Cases',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart4.set_legend({'position': 'bottom'})
  
chart4.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart4.set_chartarea({
    'border': {'none': True}
    
})

worksheet4.insert_chart('A13',chart4,{'x_scale': 2.5, 'y_scale': 1})






############################################################################
#                                                                          #  
#                             By RC - Equip.                               #  
#                                                                          #
############################################################################

current_week_data_equipment = current_week_data[current_week_data['Root Cause'] == 'Equipment']

#current week
Hardware_current_week = current_week_data_equipment['Root Cause Detail'].str.count("Hardware").sum()
Software_current_week = current_week_data_equipment['Root Cause Detail'].str.count("Software").sum()
Cleaned_current_week = current_week_data_equipment['Root Cause Detail'].str.count("Cleaned").sum()
Configuration_current_week = current_week_data_equipment['Root Cause Detail'].str.count("Configuration").sum()
SelfCleared_current_week = current_week_data_equipment['Root Cause Detail'].str.count("Self-Cleared").sum()
NRF_current_week = current_week_data_equipment['Root Cause Detail'].str.count("NRF").sum()
Synchronization_current_week = current_week_data_equipment['Root Cause Detail'].str.count("Synchronization").sum()
total_current_week = current_week_data['Root Cause'].str.count("Equipment").sum()

other_current_week = total_current_week - (Hardware_current_week + Software_current_week + Cleaned_current_week + Configuration_current_week + SelfCleared_current_week + NRF_current_week + Synchronization_current_week)



second_week_data_equipment = second_last_week_data[second_last_week_data['Root Cause']== 'Equipment']

#second week


Hardware_second_week = second_week_data_equipment['Root Cause Detail'].str.count("Hardware").sum()
Software_second_week = second_week_data_equipment['Root Cause Detail'].str.count("Software").sum()
Cleaned_second_week = second_week_data_equipment['Root Cause Detail'].str.count("Cleaned").sum()
Configuration_second_week = second_week_data_equipment['Root Cause Detail'].str.count("Configuration").sum()
SelfCleared_second_week = second_week_data_equipment['Root Cause Detail'].str.count("Self-Cleared").sum()
NRF_second_week = second_week_data_equipment['Root Cause Detail'].str.count("NRF").sum()
Synchronization_second_week = second_week_data_equipment['Root Cause Detail'].str.count("Synchronization").sum()
total_second_week = second_last_week_data['Root Cause'].str.count("Equipment").sum()
other_second_week = total_second_week - (Hardware_second_week + Software_second_week + Cleaned_second_week + Configuration_second_week + SelfCleared_second_week + NRF_second_week + Synchronization_second_week)



third_week_data_equipment = third_last_week_data[third_last_week_data['Root Cause'] == 'Equipment']

# third week

Hardware_third_week = third_week_data_equipment['Root Cause Detail'].str.count("Hardware").sum()
Software_third_week = third_week_data_equipment['Root Cause Detail'].str.count("Software").sum()
Cleaned_third_week = third_week_data_equipment['Root Cause Detail'].str.count("Cleaned").sum()
Configuration_third_week = third_week_data_equipment['Root Cause Detail'].str.count("Configuration").sum()
SelfCleared_third_week = third_week_data_equipment['Root Cause Detail'].str.count("Self-Cleared").sum()
NRF_third_week = third_week_data_equipment['Root Cause Detail'].str.count("NRF").sum()
Synchronization_third_week = third_week_data_equipment['Root Cause Detail'].str.count("Synchronization").sum()
total_third_week = third_last_week_data['Root Cause'].str.count("Equipment").sum()
other_third_week = total_third_week - (Hardware_third_week + Software_third_week + Cleaned_third_week + Configuration_third_week + SelfCleared_third_week + NRF_third_week + Synchronization_third_week)



worksheet5.write_string(0,0, 'By RC - Equip.',bold)
worksheet5.write_string(1,0, "Hardware",bg_color)
worksheet5.write_string(2,0, "Software",bg_color)
worksheet5.write_string(3,0, "Cleaned",bg_color)
worksheet5.write_string(4,0, "Configuration",bg_color)
worksheet5.write_string(5,0, "Self-Cleared",bg_color)
worksheet5.write_string(6,0, "NRF",bg_color)
worksheet5.write_string(7,0, "Synchronization",bg_color)
worksheet5.write_string(8,0, "Others",bg_color)
worksheet5.write_string(9,0, "Total Cases (Equipment)",bold)





worksheet5.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet5.write_number(1,1, Hardware_third_week,bg_color)
worksheet5.write_number(2,1, Software_third_week,bg_color)
worksheet5.write_number(3,1, Cleaned_third_week,bg_color)
worksheet5.write_number(4,1, Configuration_third_week,bg_color)
worksheet5.write_number(5,1, SelfCleared_third_week,bg_color)
worksheet5.write_number(6,1, NRF_third_week,bg_color)
worksheet5.write_number(7,1, Synchronization_third_week,bg_color)
worksheet5.write_number(8,1, other_third_week,bg_color)
worksheet5.write_number(9,1, total_third_week,bold)


worksheet5.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet5.write_number(1,2, Hardware_second_week,bg_color)
worksheet5.write_number(2,2, Software_second_week,bg_color)
worksheet5.write_number(3,2, Cleaned_second_week,bg_color)
worksheet5.write_number(4,2, Configuration_second_week,bg_color)
worksheet5.write_number(5,2, SelfCleared_second_week,bg_color)
worksheet5.write_number(6,2, NRF_second_week,bg_color)
worksheet5.write_number(7,2, Synchronization_second_week,bg_color)
worksheet5.write_number(8,2, other_second_week,bg_color)
worksheet5.write_number(9,2, total_second_week,bold)





worksheet5.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet5.write_number(1,3, Hardware_current_week,bg_color)
worksheet5.write_number(2,3, Software_current_week,bg_color)
worksheet5.write_number(3,3, Cleaned_current_week,bg_color)
worksheet5.write_number(4,3, Configuration_current_week,bg_color)
worksheet5.write_number(5,3, SelfCleared_current_week,bg_color)
worksheet5.write_number(6,3, NRF_current_week,bg_color)
worksheet5.write_number(7,3, Synchronization_current_week,bg_color)
worksheet5.write_number(8,3, other_current_week,bg_color)
worksheet5.write_number(9,3, total_current_week,bold)




chart5 = workbook.add_chart({'type': 'column'})
chart5.add_series({
  'name': '=By RC - Equip!$B$1',
  'categories': '= '"'By RC - Equip'"'!$A$2:$A$9',
  'values': '=By RC - Equip!$B$2:$B$9',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart5.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'By RC - Equip'"'!$A$2:$A$9',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'By RC - Equip'"'!$C$2:$C$9',
  'data_labels': {'value': True}})
  
chart5.add_series({
  'name': '=By RC - Equip!$D$1',
  'categories': '= '"'By RC - Equip'"'!$A$2:$A$9',
  'values': '=By RC - Equip!$D$2:$D$9',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart5.set_y_axis({
    'name': 'Breakdown of Equipment Cases',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart5.set_legend({'position': 'bottom'})
  
chart5.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart5.set_chartarea({
    'border': {'none': True}
    
})

worksheet5.insert_chart('A13',chart5,{'x_scale': 2.5, 'y_scale': 1})





# insert callout


row = 12
col = 2
text = "others = international fault"
options = {
    'width': 160,
    'height': 50,
     'border': {'color': 'blue',
               'width': 1,
               'dash_type': 'round_dot'},
     'align': {'vertical': 'middle',
               'horizontal': 'center'},
    'font': {'bold': True,
             'italic': False,
             
             'name': 'Arial',
             'color': 'black',
             'size': 12},
     # 'fill': {'color': '#ccd9d9'},
             
}

worksheet5.insert_textbox(row, col, text, options)




#outage related

current_week_data_outage = current_week_data[current_week_data['Service Impacted'] == 'Outage']
second_last_week_data_ouatge = second_last_week_data[second_last_week_data['Service Impacted'] == 'Outage']
third_last_week_data_outage = third_last_week_data[third_last_week_data['Service Impacted'] == 'Outage']





############################################################################
#                                                                          #  
#                             Outages - By RC                              #  
#                                                                          #
############################################################################


#current week

Power_current_week_outage = current_week_data_outage['Root Cause'].str.count("Power").sum()
Equipment_current_week_outage = current_week_data_outage['Root Cause'].str.count("Equipment").sum()
TransMedia_current_week_outage = current_week_data_outage['Root Cause'].str.count("Trans Media").sum()
DCNSynchronization_current_week_ouatge = current_week_data_outage['Root Cause'].str.count("DCN/Synchronization").sum()
Environment_current_week_outage = current_week_data_outage['Root Cause'].str.count("Environment").sum()
# ManMade_current_week_outage = current_week_data_outage['Root Cause'].str.count("Man Made").sum()
# Transmission_current_week_outage = current_week_data_outage['Root Cause'].str.count("Transmission").sum()
total_current_week_outage = current_week_data_outage["Root Cause"].count().sum()
Miscellaneous_current_outage = total_current_week_outage - (Power_current_week_outage + Equipment_current_week_outage + TransMedia_current_week_outage + DCNSynchronization_current_week_ouatge + Environment_current_week_outage )




#second week

Power_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("Power").sum()
Equipment_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("Equipment").sum()
TransMedia_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("Trans Media").sum()
DCNSynchronization_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("DCN/Synchronization").sum()
Environment_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("Environment").sum()
# ManMade_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("Man Made").sum()
# Transmission_second_week_outage = second_last_week_data_ouatge['Root Cause'].str.count("Transmission").sum()
total_second_week_outage = second_last_week_data_ouatge["Root Cause"].count().sum()
Miscellaneous_second_outage = total_second_week_outage - (Power_second_week_outage + Equipment_second_week_outage + TransMedia_second_week_outage + DCNSynchronization_second_week_outage + Environment_second_week_outage )





#third week


Power_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("Power").sum()
Equipment_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("Equipment").sum()
TransMedia_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("Trans Media").sum()
DCNSynchronization_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("DCN/Synchronization").sum()
Environment_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("Environment").sum()
# ManMade_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("Man Made").sum()
# Transmission_third_week_outage = third_last_week_data_outage['Root Cause'].str.count("Transmission").sum()
total_third_week_outage = third_last_week_data_outage["Root Cause"].count().sum()
Miscellaneous_third_outage = total_third_week_outage - (Power_third_week_outage + Equipment_third_week_outage + TransMedia_third_week_outage + DCNSynchronization_third_week_outage + Environment_third_week_outage )




worksheet6.write_string(0,0, 'Equipment Type',bold)
worksheet6.write_string(1,0, "Power",bg_color)
worksheet6.write_string(2,0, "Equipment",bg_color)
worksheet6.write_string(3,0, "Trans Media",bg_color)
worksheet6.write_string(4,0, "DCN/Synchronization",bg_color)
worksheet6.write_string(5,0, "Environment",bg_color)
worksheet6.write_string(6,0, "Miscellaneous",bg_color)
worksheet6.write_string(7,0, "Total Outages Cases",bold)


worksheet6.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet6.write_number(1,1, Power_third_week_outage,bg_color)
worksheet6.write_number(2,1, Equipment_third_week_outage,bg_color)
worksheet6.write_number(3,1, TransMedia_third_week_outage,bg_color)
worksheet6.write_number(4,1, DCNSynchronization_third_week_outage,bg_color)
worksheet6.write_number(5,1, Environment_third_week_outage,bg_color)
worksheet6.write_number(6,1, Miscellaneous_third_outage,bg_color)
worksheet6.write_number(7,1, total_third_week_outage,bold)


worksheet6.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet6.write_number(1,2, Power_second_week_outage,bg_color)
worksheet6.write_number(2,2, Equipment_second_week_outage,bg_color)
worksheet6.write_number(3,2, TransMedia_second_week_outage,bg_color)
worksheet6.write_number(4,2, DCNSynchronization_second_week_outage,bg_color)
worksheet6.write_number(5,2, Environment_second_week_outage,bg_color)
worksheet6.write_number(6,2, Miscellaneous_second_outage,bg_color)
worksheet6.write_number(7,2, total_second_week_outage,bold)
  
  
worksheet6.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet6.write_number(1,3, Power_current_week_outage,bg_color)
worksheet6.write_number(2,3, Equipment_current_week_outage,bg_color)
worksheet6.write_number(3,3, TransMedia_current_week_outage,bg_color)
worksheet6.write_number(4,3, DCNSynchronization_current_week_ouatge,bg_color)
worksheet6.write_number(5,3, Environment_current_week_outage,bg_color)
worksheet6.write_number(6,3, Miscellaneous_current_outage,bg_color)
worksheet6.write_number(7,3, total_current_week_outage,bold)



chart6 = workbook.add_chart({'type': 'column'})
chart6.add_series({
  'name': '=Outages - By RC!$B$1',
  'categories': '= '"'Outages - By RC'"'!$A$2:$A$7',
  'values': '=Outages - By RC!$B$2:$B$7',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart6.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'Outages - By RC'"'!$A$2:$A$7',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'Outages - By RC'"'!$C$2:$C$7',
  'data_labels': {'value': True}})
  
chart6.add_series({
  'name': '=Outages - By RC!$D$1',
  'categories': '= '"'Outages - By RC'"'!$A$2:$A$7',
  'values': '=Outages - By RC!$D$2:$D$7',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart6.set_y_axis({
    'name': 'Total  Network Faults  - Outages - By RC',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart6.set_legend({'position': 'bottom'})
  
chart6.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart6.set_chartarea({
    'border': {'none': True}
    
})

worksheet6.insert_chart('A11',chart6,{'x_scale': 2.5, 'y_scale': 1})


  
  
# insert callout


row = 12
col = 2
text = "others = Man Made"
options = {
    'width': 160,
    'height': 50,
     'border': {'color': 'blue',
               'width': 1,
               'dash_type': 'round_dot'},
     'align': {'vertical': 'middle',
               'horizontal': 'center'},
    'font': {'bold': True,
             'italic': False,
             
             'name': 'Arial',
             'color': 'black',
             'size': 12},
     # 'fill': {'color': '#ccd9d9'},
             
}

worksheet6.insert_textbox(row, col, text, options)



############################################################################
#                                                                          #  
#                             Outages By RC - Power                        #  
#                                                                          #
############################################################################

# current week

SCECOCommercialPower_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("SCECO - Commercial Power").sum()
Generator_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Generator").sum()
Airconditioning_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Air Conditioning").sum()
Rectifier_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Rectifier").sum()
Circuitbreaker_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Circuit breaker").sum()
Battery_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Battery").sum()
Power_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Power Cable").sum()
CustomerPower_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Customer Power").sum()
total_current_week_outage = current_week_data_outage["Root Cause"].str.count("Power").sum()
other_current_week = total_current_week_outage - (SCECOCommercialPower_current_week_outage + Generator_current_week_outage +Airconditioning_current_week_outage + Rectifier_current_week_outage + Circuitbreaker_current_week_outage + Battery_current_week_outage +  Power_current_week_outage + CustomerPower_current_week_outage)
#second week


SCECOCommercialPower_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("SCECO - Commercial Power").sum()
Generator_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Generator").sum()
Airconditioning_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Air Conditioning").sum()
Rectifier_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Rectifier").sum()
Circuitbreaker_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Circuit breaker").sum()
Battery_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Battery").sum()
CustomerPower_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Customer Power").sum()
Power_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Power Cable").sum()
total_second_week_outage = second_last_week_data_ouatge["Root Cause"].str.count("Power").sum()
other_second_week = total_second_week_outage - (SCECOCommercialPower_second_week_outage + Generator_second_week_outage +Airconditioning_second_week_outage + Rectifier_second_week_outage + Circuitbreaker_second_week_outage + Battery_second_week_outage +Power_second_week_outage + CustomerPower_second_week_outage)


#third week


SCECOCommercialPower_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("SCECO - Commercial Power").sum()
Generator_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Generator").sum()
Airconditioning_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Air Conditioning").sum()
Rectifier_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Rectifier").sum()
Circuitbreaker_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Circuit breaker").sum()
Battery_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Battery").sum()
CustomerPower_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Customer Power").sum()
Power_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Power Cable").sum()
total_third_week_outage = third_last_week_data_outage["Root Cause"].str.count("Power").sum()
other_third_week = total_third_week_outage - (SCECOCommercialPower_third_week_outage + Generator_third_week_outage +Airconditioning_third_week_outage + Rectifier_third_week_outage + Circuitbreaker_third_week_outage + Battery_third_week_outage +  Power_third_week_outage + CustomerPower_third_week_outage)


worksheet7.write_string(0,0, "Power Type",bold)
worksheet7.write_string(1,0, "SCECO - Commercial Power",bg_color)
worksheet7.write_string(2,0, "Generator",bg_color)
worksheet7.write_string(3,0, "Air Conditioning",bg_color)
worksheet7.write_string(4,0, "Rectifier",bg_color)
worksheet7.write_string(5,0, "Circuit Breaker",bg_color)
worksheet7.write_string(6,0, "Battery",bg_color)
worksheet7.write_string(7,0, "Power Cable",bg_color)
worksheet7.write_string(8,0, "Customer Power",bg_color)
worksheet7.write_string(9,0, "Others",bg_color)
worksheet7.write_string(10,0, "Outages Cases (Power)",bold)


worksheet7.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet7.write_number(1,1, SCECOCommercialPower_third_week_outage,bg_color)
worksheet7.write_number(2,1, Generator_third_week_outage,bg_color)
worksheet7.write_number(3,1, Airconditioning_third_week_outage,bg_color)
worksheet7.write_number(4,1, Rectifier_third_week_outage,bg_color)
worksheet7.write_number(5,1, Circuitbreaker_third_week_outage,bg_color)
worksheet7.write_number(6,1, Battery_third_week_outage,bg_color)
worksheet7.write_number(7,1, Power_third_week_outage,bg_color)
worksheet7.write_number(8,1, CustomerPower_third_week_outage,bg_color)
worksheet7.write_number(9,1, other_third_week,bg_color)
worksheet7.write_number(10,1, total_third_week_outage,bold)

worksheet7.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet7.write_number(1,2, SCECOCommercialPower_second_week_outage,bg_color)
worksheet7.write_number(2,2, Generator_second_week_outage,bg_color)
worksheet7.write_number(3,2, Airconditioning_second_week_outage,bg_color)
worksheet7.write_number(4,2, Rectifier_second_week_outage,bg_color)
worksheet7.write_number(5,2, Circuitbreaker_second_week_outage,bg_color)
worksheet7.write_number(6,2, Battery_second_week_outage,bg_color)
worksheet7.write_number(7,2, Power_second_week_outage,bg_color)
worksheet7.write_number(8,2, CustomerPower_second_week_outage,bg_color)
worksheet7.write_number(9,2, other_second_week,bg_color)
worksheet7.write_number(10,2, total_second_week_outage,bold)

worksheet7.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet7.write_number(1,3, SCECOCommercialPower_current_week_outage,bg_color)
worksheet7.write_number(2,3, Generator_current_week_outage,bg_color)
worksheet7.write_number(3,3, Airconditioning_current_week_outage,bg_color)
worksheet7.write_number(4,3, Rectifier_current_week_outage,bg_color)
worksheet7.write_number(5,3, Circuitbreaker_current_week_outage,bg_color)
worksheet7.write_number(6,3, Battery_current_week_outage,bg_color)
worksheet7.write_number(7,3, Power_current_week_outage,bg_color)
worksheet7.write_number(8,3, CustomerPower_current_week_outage,bg_color)
worksheet7.write_number(9,3, other_current_week,bg_color)
worksheet7.write_number(10,3, total_current_week_outage,bold)


chart7 = workbook.add_chart({'type': 'column'})
chart7.add_series({
  'name': '=Outages By RC - Power!$B$1',
  'categories': '= '"'Outages By RC - Power'"'!$A$2:$A$9',
  'values': '=Outages By RC - Power!$B$2:$B$9',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart7.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'Outages By RC - Power'"'!$A$2:$A$9',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'Outages By RC - Power'"'!$C$2:$C$9',
  'data_labels': {'value': True}})
  
chart7.add_series({
  'name': '=Outages - By RC!$D$1',
  'categories': '= '"'Outages By RC - Power'"'!$A$2:$A$9',
  'values': '=Outages By RC - Power!$D$2:$D$9',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart7.set_y_axis({
    'name': 'Breakdown of  Power Cases',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart7.set_legend({'position': 'bottom'})
  
chart7.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart7.set_chartarea({
    'border': {'none': True}
    
})

worksheet7.insert_chart('A14',chart7,{'x_scale': 2.5, 'y_scale': 1})




############################################################################
#                                                                          #  
#                             Outages By RC - T.Media                      #  
#                                                                          #
############################################################################


#current week

OSPOFCABLECUT_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("OSP OF Cable Cut").sum()
OpticalElectricalindoorCables_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Optical/Electrical indoor Cables").sum()
OpticalElectricalConnectors_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Optical/Electrical Connectors").sum()
OpticalElectricalOutdoorCables_current_week_outage = current_week_data_outage['Root Cause Detail'].str.count("Optical/Electrical Outdoor Cables").sum()
total_current_week_outage = current_week_data_outage["Root Cause"].str.count("Trans Media").sum()
Other_current_week_outage = total_current_week_outage - (OSPOFCABLECUT_current_week_outage+OpticalElectricalindoorCables_current_week_outage+OpticalElectricalConnectors_current_week_outage+OpticalElectricalOutdoorCables_current_week_outage)


#second week

OSPOFCABLECUT_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("OSP OF Cable Cut").sum()
OpticalElectricalindoorCables_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Optical/Electrical indoor Cables").sum()
OpticalElectricalConnectors_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Optical/Electrical Connectors").sum()
OpticalElectricalOutdoorCables_second_week_outage = second_last_week_data_ouatge['Root Cause Detail'].str.count("Optical/Electrical Outdoor Cables").sum()
total_second_week_outage = second_last_week_data_ouatge["Root Cause"].str.count("Trans Media").sum()
other_second_week_outage = total_second_week_outage - (OSPOFCABLECUT_second_week_outage+OpticalElectricalindoorCables_second_week_outage+OpticalElectricalConnectors_second_week_outage+OpticalElectricalOutdoorCables_second_week_outage)


#third week

OSPOFCABLECUT_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("OSP OF Cable Cut").sum()
OpticalElectricalindoorCables_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Optical/Electrical indoor Cables").sum()
OpticalElectricalConnectors_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Optical/Electrical Connectors").sum()
OpticalElectricalOutdoorCables_third_week_outage = third_last_week_data_outage['Root Cause Detail'].str.count("Optical/Electrical Outdoor Cables").sum()
total_third_week_outage = third_last_week_data_outage["Root Cause"].str.count("Trans Media").sum()
other_third_week_outage = total_third_week_outage - (OSPOFCABLECUT_third_week_outage+OpticalElectricalindoorCables_third_week_outage+OpticalElectricalConnectors_third_week_outage+OpticalElectricalOutdoorCables_third_week_outage)


worksheet8.write_string(0,0, "TransMedia Type",bold)
worksheet8.write_string(1,0, "OSP OF Cable Cut",bg_color)
worksheet8.write_string(2,0, "Optical/Electrical indoor Cables",bg_color)
worksheet8.write_string(3,0, "Optical/Electrical Connectors",bg_color)
worksheet8.write_string(4,0, "Optical/Electrical Outdoor Cables",bg_color)
worksheet8.write_string(5,0, "Others",bg_color)
worksheet8.write_string(6,0, "Outages Cases (Trans Media)",bold)



worksheet8.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet8.write_number(1,1, OSPOFCABLECUT_third_week_outage,bg_color)
worksheet8.write_number(2,1, OpticalElectricalindoorCables_third_week_outage,bg_color)
worksheet8.write_number(3,1, OpticalElectricalConnectors_third_week_outage,bg_color)
worksheet8.write_number(4,1, OpticalElectricalOutdoorCables_third_week_outage,bg_color)
worksheet8.write_number(5,1, other_third_week_outage,bg_color)
worksheet8.write_number(6,1, total_third_week_outage,bold)



worksheet8.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet8.write_number(1,2, OSPOFCABLECUT_second_week_outage,bg_color)
worksheet8.write_number(2,2, OpticalElectricalindoorCables_second_week_outage,bg_color)
worksheet8.write_number(3,2, OpticalElectricalConnectors_second_week_outage,bg_color)
worksheet8.write_number(4,2, OpticalElectricalOutdoorCables_second_week_outage,bg_color)
worksheet8.write_number(5,2, other_second_week_outage,bg_color)
worksheet8.write_number(6,2, total_second_week_outage,bold)

worksheet8.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet8.write_number(1,3, OSPOFCABLECUT_current_week_outage,bg_color)
worksheet8.write_number(2,3, OpticalElectricalindoorCables_current_week_outage,bg_color)
worksheet8.write_number(3,3, OpticalElectricalConnectors_current_week_outage,bg_color)
worksheet8.write_number(4,3, OpticalElectricalOutdoorCables_current_week_outage,bg_color)
worksheet8.write_number(5,3, Other_current_week_outage,bg_color)
worksheet8.write_number(6,3, total_current_week_outage,bold)



chart8 = workbook.add_chart({'type': 'column'})
chart8.add_series({
  'name': '=Outages By RC - T.Media!$B$1',
  'categories': '= '"'Outages By RC - T.Media'"'!$A$2:$A$6',
  'values': '=Outages By RC - T.Media!$B$2:$B$6',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart8.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'Outages By RC - T.Media'"'!$A$2:$A$6',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'Outages By RC - T.Media'"'!$C$2:$C$6',
  'data_labels': {'value': True}})
  
chart8.add_series({
  'name': '=Outages - By RC!$D$1',
  'categories': '= '"'Outages By RC - T.Media'"'!$A$2:$A$6',
  'values': '=Outages By RC - T.Media!$D$2:$D$6',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart8.set_y_axis({
    'name': 'Breakdown of  TransMedia Cases',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart8.set_legend({'position': 'bottom'})
  
chart8.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart8.set_chartarea({
    'border': {'none': True}
    
})

worksheet8.insert_chart('A10',chart8,{'x_scale': 2.5, 'y_scale': 1})



############################################################################
#                                                                          #  
#                             Outages By RC - Equip.                       #  
#                                                                          #
############################################################################


current_week_data_outage_equipment = current_week_data_outage[current_week_data_outage['Root Cause'] == 'Equipment']

#current week

Hardware_current_week_outage = current_week_data_outage_equipment['Root Cause Detail'].str.count("Hardware").sum()
Software_current_week_outage = current_week_data_outage_equipment['Root Cause Detail'].str.count("Software").sum()
Cleaned_current_week_outage = current_week_data_outage_equipment['Root Cause Detail'].str.count("Cleaned").sum()
Configuration_current_week_outage = current_week_data_outage_equipment['Root Cause Detail'].str.count("Configuration").sum()
NRF_current_week_outage = current_week_data_outage_equipment['Root Cause Detail'].str.count("NRF").sum()
total_current_week_outage = current_week_data_outage_equipment["Root Cause"].str.count("Equipment").sum()
Other_current_week_outage = total_current_week_outage - (Hardware_current_week_outage+Software_current_week_outage+Cleaned_current_week_outage+Configuration_current_week_outage+NRF_current_week_outage)


second_week_data_outage_equipment = second_last_week_data_ouatge[second_last_week_data_ouatge['Root Cause'] == 'Equipment']
#second week

Hardware_second_week_outage = second_week_data_outage_equipment['Root Cause Detail'].str.count("Hardware").sum()
Software_second_week_outage = second_week_data_outage_equipment['Root Cause Detail'].str.count("Software").sum()
Cleaned_second_week_outage = second_week_data_outage_equipment['Root Cause Detail'].str.count("Cleaned").sum()
Configuration_second_week_outage = second_week_data_outage_equipment['Root Cause Detail'].str.count("Configuration").sum()
NRF_second_week_outage = second_week_data_outage_equipment['Root Cause Detail'].str.count("NRF").sum()
total_second_week_outage = second_week_data_outage_equipment["Root Cause"].str.count("Equipment").sum()
Other_second_week_outage = total_second_week_outage - (Hardware_second_week_outage+Software_second_week_outage+Cleaned_second_week_outage+Configuration_second_week_outage+NRF_second_week_outage)


third_week_data_outage_equipment = third_last_week_data_outage[third_last_week_data_outage['Root Cause'] == 'Equipment']
#third week

Hardware_third_week_outage = third_week_data_outage_equipment['Root Cause Detail'].str.count("Hardware").sum()
Software_third_week_outage = third_week_data_outage_equipment['Root Cause Detail'].str.count("Software").sum()
Cleaned_third_week_outage = third_week_data_outage_equipment['Root Cause Detail'].str.count("Cleaned").sum()
Configuration_third_week_outage = third_week_data_outage_equipment['Root Cause Detail'].str.count("Configuration").sum()
NRF_third_week_outage = third_week_data_outage_equipment['Root Cause Detail'].str.count("NRF").sum()
total_third_week_outage = third_week_data_outage_equipment["Root Cause"].str.count("Equipment").sum()
Other_third_week_outage = total_third_week_outage - (Hardware_third_week_outage+Software_third_week_outage+Cleaned_third_week_outage+Configuration_third_week_outage+NRF_third_week_outage)



worksheet9.write_string(0,0, "Equipment Type",bold)
worksheet9.write_string(1,0, "Hardware",bg_color)
worksheet9.write_string(2,0, "Software",bg_color)
worksheet9.write_string(3,0, "Cleaned",bg_color)
worksheet9.write_string(4,0, "Configuration",bg_color)
worksheet9.write_string(5,0, "NRF",bg_color)
worksheet9.write_string(6,0, "Others",bg_color)
worksheet9.write_string(7,0, "Outages Cases (Equipment)",bold)


worksheet9.write_string(0,1, '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')',bold)
worksheet9.write_number(1,1, Hardware_third_week_outage,bg_color)
worksheet9.write_number(2,1, Software_third_week_outage,bg_color)
worksheet9.write_number(3,1, Cleaned_third_week_outage,bg_color)
worksheet9.write_number(4,1, Configuration_third_week_outage,bg_color)
worksheet9.write_number(5,1, NRF_third_week_outage,bg_color)
worksheet9.write_number(6,1, Other_third_week_outage,bg_color)
worksheet9.write_number(7,1, total_third_week_outage,bold)



worksheet9.write_string(0,2, 'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',bold)
worksheet9.write_number(1,2, Hardware_second_week_outage,bg_color)
worksheet9.write_number(2,2, Software_second_week_outage,bg_color)
worksheet9.write_number(3,2, Cleaned_second_week_outage,bg_color)
worksheet9.write_number(4,2, Configuration_second_week_outage,bg_color)
worksheet9.write_number(5,2, NRF_second_week_outage,bg_color)
worksheet9.write_number(6,2, Other_second_week_outage,bg_color)
worksheet9.write_number(7,2, total_second_week_outage,bold)


worksheet9.write_string(0,3, 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')',bold)
worksheet9.write_number(1,3, Hardware_current_week_outage,bg_color)
worksheet9.write_number(2,3, Software_current_week_outage,bg_color)
worksheet9.write_number(3,3, Cleaned_current_week_outage,bg_color)
worksheet9.write_number(4,3, Configuration_current_week_outage,bg_color)
worksheet9.write_number(5,3, NRF_current_week_outage,bg_color)
worksheet9.write_number(6,3, Other_current_week_outage,bg_color)
worksheet9.write_number(7,3, total_current_week_outage,bold)

chart9 = workbook.add_chart({'type': 'column'})
chart9.add_series({
  'name': '=Outages By RC - Equip!$B$1',
  'categories': '= '"'Outages By RC - Equip'"'!$A$2:$A$7',
  'values': '=Outages By RC - Equip!$B$2:$B$7',
  'fill':   {'color': '#31859c'},
  'data_labels': {'value': True}})
chart9.add_series({
  'name': '='"'To District'"'!$C$1',
  'categories': '= '"'Outages By RC - Equip'"'!$A$2:$A$7',
  'fill':   {'color': '#e46c0a'},
  'values': '='"'Outages By RC - Equip'"'!$C$2:$C$7',
  'data_labels': {'value': True}})
  
chart9.add_series({
  'name': '=Outages - By RC!$D$1',
  'categories': '= '"'Outages By RC - Equip'"'!$A$2:$A$7',
  'values': '=Outages By RC - Equip!$D$2:$D$7',
   'fill':   {'color': '#8064a2'},
  'data_labels': {'value': True}}) 
  
chart9.set_y_axis({
    'name': 'Breakdown of Equipment Cases',
    'major_gridlines': {
        'visible': False,
        'line': {'width': 1.25, 'dash_type': 'dash'}}
    })

chart9.set_legend({'position': 'bottom'})
  
chart9.set_plotarea({
   
   
     'layout': {
        'x':      0.13,
        'y':      0.26,
        'width':  5.73,
        'height': 0.57,
    }
   
})  
  
chart9.set_chartarea({
    'border': {'none': True}
    
})

worksheet9.insert_chart('A10',chart9,{'x_scale': 2.5, 'y_scale': 1})



############################################################################
#                                                                          #  
#                             Generate excel file                          #  
#                                                                          #
############################################################################



# By Region 

# region_ = {'Region': ["Central","West","South","East","Grand Total"], '2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')': [central_third,west_third,south_third,east_third,central_third+west_third+south_third+east_third]
# ,'Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')': [central_second,west_second,south_second,east_second,central_second+west_second+south_second+east_second],
# 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')': [central_current,west_current,south_current,east_current,central_current+west_current+south_current+east_current]}
# region_df = DataFrame(region_,columns= ['Region','2nd Last Two Weeks (' + str(minimum_date.strftime("%d-%m-%y")) + ' to ' + str(sec_last_week.strftime("%d-%m-%y")) + ')','Last Two Weeks (' + str(third_last_week.strftime("%d-%m-%y")) + ' to ' + str(current_week.strftime("%d-%m-%y")) + ')',
# 'Present Two Weeks (' + str(current_week_label.strftime("%d-%m-%y")) + ' to ' + str(maximum_date.strftime("%d-%m-%y")) + ')'])

writer.sheets['Raw-Data'] = worksheet10
# current_week_data.to_excel(writer,sheet_name='Raw-Data',startrow=0 , startcol=0)
current_week_data.to_excel(writer,sheet_name='Raw-Data',startrow=0 , startcol=0)



# region_df.to_excel(r"C:\Users\Huawei03\Desktop\NFR\Bi-Weekly Network Faults Report.xlsx",index= False,sheet_name='To Region')
# region_df.to_excel(r"C:\Users\Huawei03\Desktop\NFR\Bi-Weekly Network Faults Report.xlsx",index= False,sheet_name='To District')



workbook.close()


comparison_data = current_week_data[['Root Cause','Root Cause Detail','District']]
comparison_data_1 = second_last_week_data[['Root Cause','Root Cause Detail','District']]
comparison_data_outage = current_week_data_outage[['Root Cause','Root Cause Detail','District']]
comparison_data_outage_ = second_last_week_data_ouatge[['Root Cause','Root Cause Detail','District']]


table = comparison_data.pivot_table(index=["Root Cause","Root Cause Detail"], columns='District', 
                        aggfunc=len, fill_value=0)

table_ = comparison_data_1.pivot_table(index=["Root Cause","Root Cause Detail"], columns='District', 
                        aggfunc=len, fill_value=0)
table_outage = comparison_data_outage.pivot_table(index=["Root Cause","Root Cause Detail"], columns='District', 
                        aggfunc=len, fill_value=0)
table_outage_ = comparison_data_outage_.pivot_table(index=["Root Cause","Root Cause Detail"], columns='District', 
                        aggfunc=len, fill_value=0)






writer1 = pd.ExcelWriter('D:\\NFR\\Bi-Weekly Comparison Table.xlsx',engine='xlsxwriter')

workbook_1= writer1.book

#format
bold =workbook_1.add_format({'bold': True,'bg_color': '#4babc6','border':1, 'valign': 'top'})
bg_color = workbook_1.add_format({'bg_color': '#b7dee8','border':1,'text_wrap': True, 'valign': 'top'})


merge_format = workbook_1.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#4babc6'})


table.to_excel(writer1,sheet_name = 'Comparison Table',startrow=1 , startcol=0)
table_.to_excel(writer1,sheet_name = 'Comparison Table',startrow=1 , startcol=24)
table_outage.to_excel(writer1,sheet_name = 'Table(SA-NSA)',startrow=1 , startcol=0)
table_outage_.to_excel(writer1,sheet_name = 'Table(SA-NSA)',startrow=1 , startcol=24)


# Get the xlsxwriter workbook and worksheet objects.

worksheet_ = writer1.sheets['Comparison Table']
worksheet_.merge_range('A1:F1', "Present Two Week", merge_format)
worksheet_.merge_range('Y1:AC1', "Last Two Week", merge_format)
worksheet_.set_column('A:B',20)   
worksheet_.set_column('Y:Z',20)



worksheet_outage_ = writer1.sheets['Table(SA-NSA)']
worksheet_outage_.merge_range('A1:F1', "Present Two Week", merge_format)
worksheet_outage_.merge_range('Y1:AC1', "Last Two Week", merge_format)
worksheet_outage_.set_column('A:B',20)   
worksheet_outage_.set_column('Y:Z',20)
# worksheet_.write_string(0,5, "Comparison Table",bold)    


workbook_1.close()

################################################################
#                                                              #
#                           pptx                               #
#                                                              #
################################################################

# 
# from pptx import Presentation
# from pptx.chart.data import CategoryChartData
# from pptx.enum.chart import XL_CHART_TYPE
# from pptx.util import Inches
# from pptx.dml.color import RGBColor
# from pptx.enum.chart import XL_LABEL_POSITION
# from pptx.util import Pt
# 
# SLD_LAYOUT_TITLE_AND_CONTENT = 5
# 
# 
# 
# 
# # create presentation with 1 slide ------
# prs = Presentation()
# slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_AND_CONTENT]
# slide = prs.slides.add_slide(slide_layout)
# 
# 
# 
# # define chart data ---------------------
# chart_data = CategoryChartData()
# chart_data.categories = ['Central', 'West', 'South','East']
# chart_data.add_series('Series 1', (central_current,west_current,south_current,east_current))
# chart_data.add_series('Series 2', (central_second,west_second,south_second,east_second))
# chart_data.add_series('Series 3', (central_third,west_third,south_third,east_third))
# # add chart to slide --------------------
# x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
# frame = slide.shapes.add_chart(
#     XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
# )
# #label
# chart =frame.chart
# plot = chart.plots[0]
# plot.has_data_labels = True
# data_labels = plot.data_labels
# 
# data_labels.font.size = Pt(13)
# data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
# data_labels.position = XL_LABEL_POSITION.INSIDE_END
# 
# category_axis = chart.category_axis
# category_axis.has_major_gridlines = False
# category_axis.has_minor_gridlines = False
# value_axis = chart.value_axis
# value_axis.format.line.width = 0
# 
# from pptx.enum.chart import XL_LEGEND_POSITION
# 
# chart.has_legend = True
# chart.legend.position = XL_LEGEND_POSITION.BOTTOM
# chart.legend.include_in_layout = False
# 
# 


# prs = Presentation('C:\\Users\\Huawei03\\Desktop\\NFR\\Template.pptx')
# # title_slide_layout = prs.slide_layouts[2]
# # slide = prs.slides.add_slide(title_slide_layout)
# # title = slide.shapes.title
# # subtitle = slide.placeholders[1]
# #
# # title.text = "Hello, World!"
# # subtitle.text = "python-pptx was here!"
# slide = prs.slide[2]
# title = slide.shapes.title
# title.text= "my slide"
# prs.save('C:\\Users\\Huawei03\\Desktop\\NFR\\test.pptx')




