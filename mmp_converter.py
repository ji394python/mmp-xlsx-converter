import xlsxwriter
from xlsxwriter.utility import xl_cell_to_rowcol_abs
import pandas as pd 
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
import json
import os
window = tk.Tk()

window.title('Author: Denver - SAS Project Converter')
window.geometry('350x300') 
canvas = tk.Canvas(window, width=400, height=135, bg='#DDE8B9')
image_file = tk.PhotoImage(file=r'ddr/sas.png')
image = canvas.create_image(175, 15, anchor='n', image=image_file)
canvas.pack(side='top')
tk.Label(window, text='Wellcome',font=('Arial', 16)).pack()

# 第5步，用户信息
tk.Label(window, text='作者名稱：', font=('Arial', 14)).place(x=10, y=170)
tk.Label(window, text='創造日期', font=('Arial', 14)).place(x=10, y=210)

# 用户名
var_usr_name = tk.StringVar()
var_usr_name.set('denver1072@gmail.com')
entry_usr_name = tk.Entry(window, textvariable=var_usr_name, font=('Arial', 14))
entry_usr_name.place(x=120,y=175)
# 用户密码
var_usr_pwd = tk.StringVar()
var_usr_pwd.set("2021/05/21")
entry_usr_pwd = tk.Entry(window, textvariable=var_usr_pwd, font=('Arial', 14))
entry_usr_pwd.place(x=120,y=215)

def UploadAction(event=None):
    file = filedialog.askopenfilename()
    fileName = os.path.basename(file)
    fileName = fileName[:fileName.find('.xlsx')]
    try:
        df = pd.read_excel(file,engine='openpyxl').fillna('').drop(['作用中','任務模式'],axis=1)
        globals()['df'] = df
        next = True
    except:
        next = False
        tkinter.messagebox.showinfo(title = '檔案錯誤', # 視窗標題
                                    message = '請重新上傳符合格式之xlsx檔案')   
    if next:
        main(df,fileName)

## columnNames
colNames = ['識別碼','完成百分比','項目名稱','工期','開始時間','完成時間','資源名稱','前置任務','大綱階層']
subColNames = ['CUB-PM','CUB-IT','CUB-AF','CUB-TF','CUB-TFM','SAS-PM','SAS-AF','SAS-TF','SAS-TFM']

#Date Control
month = {'一':1,'二':2,'三':y ,'四':4,'五':5,'六':6,
        '七':7,'八':8,'九':9,'十':10,'十一':11,'十二':12}

#Milestone Color
milestoneColor = {'AF':'#ffff00','TF':'#FABF94','TFM':'#92D050'}

def main(df,fileName):
    try:
        #Dataframe handle
        #df = pd.read_excel('before.xlsx',engine='openpyxl').fillna('').drop(['作用中','任務模式'],axis=1)
        # df = pd.read_excel('CUB SFM_M Project_20210524_v1.3.xlsx',engine='openpyxl').fillna('').drop(['作用中','任務模式'],axis=1)
        beforeColNames = df.columns.values.tolist()
        beforeColNames[1] = '項目名稱'
        df.columns = beforeColNames
        df['開始時間'] = df['開始時間'].apply(lambda x: x[:x.find(',')+6]).apply(lambda x: str(month[x[:x.find('月')]])+ x[x.find('月'):])
        df['完成時間'] = df['完成時間'].apply(lambda x: x[:x.find(',')+6]).apply(lambda x:  str(month[x[:x.find('月')]])+ x[x.find('月'):])

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(f'[轉換]{fileName}.xlsx')
        worksheet = workbook.add_worksheet(name='任務_表格')

        # Add a bold format to use to highlight cells.
        header = workbook.add_format(
            {'bold':True,'bg_color':'#C4BD97','font_name':'微軟正黑體',
            'align':'center','valign':'vcenter','font_size':11,
            'border':1
            }
            )
        tick = {'font_name':'Wingdings','bold':True,
            'align':'center','valign':'vcenter','font_size':10
            }

        # Add a bold format to use to highlight cells.
        number_format = {'font_name':'微軟正黑體',
            'align':'center','valign':'vcenter','font_size':10
            }

        # Create a format for the date or time.
        date_format = {'num_format': 'yyyy-mm-dd','align': 'right',
            'font_name':'微軟正黑體','valign':'vbuttom',
            'font_size':10}

        # create name format
        name_format = {'font_name':'微軟正黑體',
            'align':'left','valign':'vbuttom','font_size':10}
            

        dict_format = {'識別碼':number_format,'完成百分比':number_format,'項目名稱':name_format,
                        '工期':name_format,'開始時間':date_format,'完成時間':date_format,'前置任務':number_format,'大綱階層':number_format}


        for i in range(len(colNames)):
            cell = chr(65+i)
            if (cell == 'G'):
                main = 'G1:O1'
                worksheet.merge_range(main, colNames[i], header)
                for r in range(len(subColNames)):
                    sub = chr(65+i+r) + '2'
                    worksheet.write(sub, subColNames[r], header)
            else:
                if (cell >= 'G'):
                    cell = chr(65+i+8)
                main = cell + '1:' + cell + '2'
                worksheet.merge_range(main, colNames[i], header)

                
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 60)
        worksheet.set_column('D:D', 10)
        worksheet.set_column('E:F', 12)
        worksheet.set_column('G:O', 10)
        worksheet.set_column('P:Q', 11)

        worksheet.set_zoom(77)
        #worksheet.write_datetime(2,2,datetime.strptime(df['開始時間'].values[0],'%m月 %d, %Y'),date_format)

        row = 2
        for item in df.iterrows():
            index = item[0]
            item = item[1]
            boldControl = True if index==len(df)-1 else True if item['大綱階層'] < df.iloc[index+1,]['大綱階層'] else True if  item['大綱階層'] == 1 else False

            indent = item['大綱階層'] - 1
            bgColor =  milestoneColor[item['項目名稱'][:3].strip()] if ( (item['工期'] == '0 工作日') & (indent == 0) ) else ''
            worksheet.set_row(row,None,workbook.add_format({'bg_color':bgColor}),{'level':indent} ) if bgColor != '' else worksheet.set_row(row,None,None,{'level':indent} )

            for k,v in item.items():
                if k in ['附註','資源名稱']: continue
                col = colNames.index(k)
                col = col+8 if col in [7,8] else col
                temp_format = dict_format[k].copy()
                temp_format['bold'] = boldControl
                temp_format['indent'] = indent if k=='項目名稱' else 0
                if bgColor != '':
                    temp_format['bg_color'] = bgColor
                else:
                    pass
                

                temp = workbook.add_format(temp_format)
                if k.find('時間') != -1:
                    worksheet.write(row, col,datetime.strptime(v,'%m月 %d, %Y'),temp)
                else:
                    worksheet.write(row, col,v,temp)
            row+=1

            # try:
            #     v = task[item['項目名稱']]
            #     temp_tick = tick.copy()
            #     if bgColor != '':
            #         temp_tick['bg_color'] = bgColor
            #     else:
            #         pass
            #     temp = workbook.add_format(temp_tick)
            #     for col in range(len(v)):
            #         worksheet.write(row-1, col+6,v[col],temp)
            # except:
            #     continue
            try:
                temp_tick = tick.copy()
                teams_temp = item['資源名稱'].split(',')

                if ( (len(teams_temp) != 1) & (teams_temp[0] != '')):
                    teams = teams_temp
                    teams_order = item['大綱階層']
                    teams_index = index
                
                if ((teams_index != index) & (teams_order >= item['大綱階層']) ):
                    teams = ['']

                if bgColor != '':
                    temp_tick['bg_color'] = bgColor
                else:
                    pass
                temp = workbook.add_format(temp_tick)
                for team in teams:
                    if team == '': continue
                    col = subColNames.index(team)
                    worksheet.write(row-1, col+6,'ü',temp)
            except:
                continue

            
        worksheet.autofilter('A1:Q1')
        worksheet.freeze_panes(2, 0)

        workbook.close()
        tkinter.messagebox.showinfo(title = '程序完成', # 視窗標題
                                    message = '完成-可直接關閉視窗離開')   # 訊息內容
    except:
        tkinter.messagebox.showinfo(title = '程序錯誤', # 視窗標題
                                    message = '請重新再操作一次')   #


# 第7步，login and sign up 按钮
btn_login = tk.Button(window, text='上傳檔案', command=UploadAction)
btn_login.place(x=70, y=255)
btn_sign_up = tk.Button(window, text='完成離開', command=window.destroy)
btn_sign_up.place(x=195, y=255)

file = tk.filedialog.askopenfile(parent=window,mode='rb',title='Choose a xlsx file')
fileName = os.path.basename(file.name)
globals()['fileName'] = fileName[:fileName.find('.xlsx')]   
if file != None:
    try:
        df = pd.read_excel(file,engine='openpyxl').fillna('').drop(['作用中','任務模式'],axis=1)
        next = True
    except:
        next = False
        tkinter.messagebox.showinfo(title = '檔案錯誤', # 視窗標題
                                    message = '請重新上傳符合格式之xlsx檔案')   
if next:
    main(df,fileName)

window.mainloop()