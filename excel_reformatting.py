# -*- coding: utf-8 -*-
"""
Created on Wed Feb 22 10:35:40 2023
@author: Yuwen.Fang
@Email: yuwen.fang24@gmail.com
@Project: Excel re-formatting
"""

import xlsxwriter 
import pandas as pd


workbook = xlsxwriter.Workbook(r'C:\Users\yuwfang\OneDrive - Stantec\文件\Yuwen.Fang\Documents\GitHub\月報格式調整\new_format.xlsx')

file_name = r'C:\Users\yuwfang\OneDrive - Stantec\文件\Yuwen.Fang\Documents\GitHub\月報格式調整\raw_data.xlsx'
sheet_name = 'table'

#------- write data --------------------
df = pd.read_excel(file_name,
                   sheet_name)
columns = list(df.columns)

for i in range(len(df)):
    worksheet = workbook.add_worksheet()
    print("正在寫入第" + str(i+1) + "筆資料....")
          
#----- table name -----------------------
    column_name = ('單元名稱', '項目', '數值', '單元名稱', '項目', '數值', '單元名稱', '項目', '數值')

    column_name_1 = ('進流抽水站', '流量計累計讀數', 
                     '初沉池', '初沉污泥累計讀數', 
                     '鼓風機室', '生物池風量累計讀數 (FE-421B)', '生物池風量瞬時讀數 (FE-421B)', '生物池風量累計讀數 (FE-422B)', '生物池風量瞬時讀數 (FE-422B)', '好氧消化池風量累計讀數 (FE-644)',
                     '膜濾池(一)', '累計產水流量(FE-441A)', '廢棄污泥累計流量(FE-454)',
                     '膜濾池(二)', '累計產水流量(FE-442A)', '廢棄污泥累計流量(FE-454)',
                     '迴流污泥泵', '迴流污泥累計流量(FE-451A)', '迴流污泥累計流量(FE-451B)',
                     '其他巡廠事項')

    column_name_2 = ('生物處理單元', 
                     '第一缺氧池(TK-411A)', 'pH (AE-411A)', 'DO (AE-411B)', 'ORP (AE-411C)',
                     '第二好氧池(TK-421B)', 'DO (AE-421A)', 'pH (AE-421B)', 'SS (AE-421C)',
                     '缺氧池(TK-412)', 'pH (AE-412A)', 'DO (AE-412B)', 'ORP (AE-412C)',
                     '第二好氧池(TK-422B)', 'DO (AE-422A)', 'pH (AE-422B)', 'SS (AE-422C)',
                     '工作日誌')

    column_name_3 = ('回收加壓及放流系統', 
                     '放流水2累計流量(FE-516)', '放流水1累計流量(FE-53B)', '滯洪池累計流量(FE-53A)',
                     '濃縮機', '進流污泥累計流量(FE-613)', 
                     '脫水機', '進流污泥累計流量(FE-663)',
                     '除臭設備(LS-713)', '第一段-pH (AE-713)', '第二段-pH (AE-714)', '第一段-ORP (AE-714)')

    title_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_name':'標楷體',
        'font_size':'14',
        'bold': True, 
        'border':2

    })

    title_format_1 = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_name':'標楷體',
        'border':2
    })

    title_format_2 = workbook.add_format({
        'valign': 'vcenter',
        'font_name':'標楷體',
        'border':2
    })


    worksheet.write('A3', column_name[0], title_format_1)
    worksheet.write('B3', column_name[1], title_format_1)
    worksheet.write('C3', column_name[2], title_format_1)
    worksheet.write('D3', column_name[3], title_format_1)
    worksheet.write('E3', column_name[4], title_format_1)
    worksheet.write('F3', column_name[5], title_format_1)
    worksheet.write('G3', column_name[6], title_format_1)
    worksheet.write('H3', column_name[7], title_format_1)
    worksheet.write('I3', column_name[8], title_format_1)

    worksheet.write('A4', column_name_1[0], title_format_1)
    worksheet.write('B4', column_name_1[1], title_format_2)
    worksheet.write('A5', column_name_1[2], title_format_1)
    worksheet.write('B5', column_name_1[3], title_format_2)
    worksheet.write('A6', column_name_1[4], title_format_2)
    worksheet.write('B6', column_name_1[5], title_format_2)
    worksheet.write('B7', column_name_1[6], title_format_2)
    worksheet.write('B8', column_name_1[7], title_format_2)
    worksheet.write('B9', column_name_1[8], title_format_2)
    worksheet.write('B10', column_name_1[9], title_format_2)
    worksheet.write('A11', column_name_1[10])
    worksheet.write('B11', column_name_1[11], title_format_2)
    worksheet.write('B12', column_name_1[12], title_format_2)
    worksheet.write('A13', column_name_1[13])
    worksheet.write('B13', column_name_1[14], title_format_2)
    worksheet.write('B14', column_name_1[15], title_format_2)
    worksheet.write('A15', column_name_1[16])
    worksheet.write('B15', column_name_1[17], title_format_2)
    worksheet.write('B16', column_name_1[18], title_format_2)
    worksheet.write('A17', column_name_1[19])


    worksheet.write('D4', column_name_2[0])
    worksheet.write('D5', column_name_2[1])
    worksheet.write('E5', column_name_2[2], title_format_2)
    worksheet.write('E6', column_name_2[3], title_format_2)
    worksheet.write('E7', column_name_2[4], title_format_2)
    worksheet.write('D8', column_name_2[5])
    worksheet.write('E8', column_name_2[6], title_format_2)
    worksheet.write('E9', column_name_2[7], title_format_2)
    worksheet.write('E10', column_name_2[8], title_format_2)
    worksheet.write('D11', column_name_2[9])
    worksheet.write('E11', column_name_2[10], title_format_2)
    worksheet.write('E12', column_name_2[11], title_format_2)
    worksheet.write('E13', column_name_2[12], title_format_2)
    worksheet.write('D14', column_name_2[13])
    worksheet.write('E14', column_name_2[14], title_format_2)
    worksheet.write('E15', column_name_2[15], title_format_2)
    worksheet.write('E16', column_name_2[16], title_format_2)
    worksheet.write('E17', column_name_2[17], title_format_2)

    worksheet.write('G4', column_name_3[0])
    worksheet.write('H4', column_name_3[1], title_format_2)
    worksheet.write('H5', column_name_3[2], title_format_2)
    worksheet.write('H6', column_name_3[3], title_format_2)
    worksheet.write('G7', column_name_3[4], title_format_1)
    worksheet.write('H7', column_name_3[5], title_format_2)
    worksheet.write('G8', column_name_3[6], title_format_1)
    worksheet.write('H8', column_name_3[7], title_format_2)
    worksheet.write('G9', column_name_3[8])
    worksheet.write('H9', column_name_3[9], title_format_2)
    worksheet.write('H10', column_name_3[10], title_format_2)
    worksheet.write('H11', column_name_3[11], title_format_2)


    #----- merge cells -------------------------------------------
    worksheet.merge_range('A1:I1', '污水處理廠巡查紀錄表', title_format)
    worksheet.merge_range('C2:I2', ' ')
    worksheet.merge_range('A6:A10', '鼓風機室', title_format_1)
    worksheet.merge_range('A11:A12', '膜濾池(一)', title_format_1)
    worksheet.merge_range('A13:A14', '膜濾池(二)', title_format_1)
    worksheet.merge_range('A15:A16', '迴流污泥泵', title_format_1)
    worksheet.merge_range('A17:D17', '其他巡廠事項', title_format_1)
    worksheet.merge_range('D4:F4', '生物處理單元', title_format_1)
    worksheet.merge_range('D5:D7', '第一缺氧池(TK-411A)', title_format_1)
    worksheet.merge_range('D8:D10', '第二好氧池(TK-421B)', title_format_1)
    worksheet.merge_range('D11:D13', '缺氧池(TK-412)', title_format_1)
    worksheet.merge_range('D14:D16', '第二好氧池(TK-422B)', title_format_1)
    worksheet.merge_range('G4:G6', '回收加壓及放流系統', title_format_1)
    worksheet.merge_range('G9:G11', '除臭設備(LS-713)', title_format_1)
    worksheet.merge_range('E17:I17', '工作日誌', title_format_1)
    worksheet.merge_range('A18:D18', ' ')
    worksheet.merge_range('E18:I18', ' ')

    #------ expand cells ----------------------------
    worksheet.set_row(17, 100)  
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 36)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 25)
    worksheet.set_column('E:E', 28)
    worksheet.set_column('F:F', 11)
    worksheet.set_column('G:G', 25)
    worksheet.set_column('H:H', 30)
    worksheet.set_column('I:I', 15)



    format1 = workbook.add_format({'num_format': 'yyyy/m/d',
                                   'font_name':'Times New Roman',
                                   'border':2})
    worksheet.write('A2', df['巡查日期'][i].date(), format1)

    format2 = workbook.add_format({'num_format': 'hh:mm',
                                   'font_name':'Times New Roman',
                                   'align': 'Left',
                                   'border':2})
    worksheet.write('B2', df['巡查時間'][i], format2)

    format_number_di = workbook.add_format({'font_name':'Times New Roman',
                                            'num_format': '#,##0.00',
                                            'border':2})
    format_number = workbook.add_format({'font_name':'Times New Roman',
                                         'num_format': '#,##0',
                                         'border':2})
    format_word = workbook.add_format({'font_name':'標楷體',
                                       'align': 'Top',
                                       'text_wrap': True,
                                       'border':2})

    worksheet.write('C2', "巡查人員: "+df['巡查人員'][i], title_format_2)
    worksheet.write('C4', df['流量計累計讀數 (m3)'][i], format_number)
    worksheet.write('C5', df['初沉污泥累計讀數 (m3)'][i], format_number)
    worksheet.write('C6', df['生物池風量累計讀數 (FE-421B) (m3)'][i], format_number)
    worksheet.write('C7', df['生物池風量瞬時讀數 (FE-421B) (CMM)'][i], format_number_di)
    worksheet.write('C8', df['生物池風量累計讀數 (FE-422B) (m3)'][i], format_number)
    worksheet.write('C9', df['生物池風量瞬時讀數 (FE-422B) (CMM)'][i], format_number_di)
    worksheet.write('C10', df['好氧消化池風量累計讀數 (FE-644) (m3)'][i], format_number)
    worksheet.write('C11', df['累計產水流量(FE-441A) (m3)'][i], format_number)
    worksheet.write('C12', df['廢棄污泥累計流量(FE-454) (m3)'][i], format_number)
    worksheet.write('C13', df['累計產水流量(FE-442A) (m3)'][i], format_number)
    worksheet.write('C14', df['廢棄污泥累計流量2(FE-454) (m3)'][i], format_number)
    worksheet.write('C15', df['迴流污泥累計流量(FE-451A) (m3)'][i], format_number)
    worksheet.write('C16', df['迴流污泥累計流量(FE-451B) (m3)'][i], format_number)
    worksheet.write('C16', df['迴流污泥累計流量(FE-451B) (m3)'][i], format_number)
    worksheet.write('F5', df['pH (AE-411A)'][i], format_number_di)
    worksheet.write('F6', df['DO (AE-411B) (mg/L)'][i], format_number_di)
    worksheet.write('F7', df['ORP (AE-411C) (mV)'][i], format_number)
    worksheet.write('F8', df['DO (AE-421A) (mg/L)'][i], format_number_di)
    worksheet.write('F9', df['pH (AE-421B)'][i], format_number_di)
    worksheet.write('F10', df['SS (AE-421C) (mg/L)'][i], format_number)
    worksheet.write('F11', df['pH (AE-412A)'][i], format_number_di)
    worksheet.write('F12', df['DO (AE-412B) (mg/L)'][i], format_number_di)
    worksheet.write('F13', df['ORP (AE-412C) (mV)'][i], format_number)
    worksheet.write('F14', df['DO (AE-422A) (mg/L)'][i], format_number_di)
    worksheet.write('F15', df['pH (AE-422B)'][i], format_number_di)
    worksheet.write('F16', df['SS (AE-422C) (mg/L)'][i], format_number)
    worksheet.write('I4', df['放流水2累計流量(FE-516) (m3)'][i], format_number)
    worksheet.write('I5', df['放流水1累計流量(FE-53B) (m3)'][i], format_number)
    worksheet.write('I6', df['滯洪池累計流量(FE-53A) (m3)'][i], format_number)
    worksheet.write('I7', df['進流污泥累計流量(FE-613) (m3)'][i], format_number)
    worksheet.write('I8', df['進流污泥累計流量(FE-663) (m3)'][i], format_number)
    worksheet.write('I9', df['第一段-pH (AE-713)'][i], format_number_di)
    worksheet.write('I10', df['第二段-pH (AE-714)'][i], format_number_di)
    worksheet.write('I11', df['第一段-ORP (AE-714) (mV)'][i], format_number)
    worksheet.write('A18', df['其他巡廠事項'][i], format_word)
    worksheet.write('E18', df['工作日誌'][i], format_word)

    border_format=workbook.add_format({'border':1})

    worksheet.conditional_format( 'A1:I18', {'type':'no_blanks', 'format':border_format} )
          
    print("第" + str(i+1) + "筆資料寫入成功!")

workbook.close()