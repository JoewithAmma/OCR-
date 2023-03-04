# -*- coding: utf-8 -*-
"""
Created on Tue Jan 31 09:51:17 2023

@author: User
"""

import os
import glob
import time
import PySimpleGUI as sg
import pandas as pd
from requests import get, post


layout = [[sg.Text("您好，歡迎使用法務小工具",size=(50,1))],
          [sg.Text('請選擇PDF資料夾路徑：',  size=(50, 1))],
          [sg.Input('  ', key="_FILE_", readonly=True,size=(60, 1)),sg.FolderBrowse('選擇路徑')],
          [sg.Text('請選擇TXT檔案輸出路徑(圖片轉文字結果)：',  size=(50, 1))],
          [sg.Input('  ', key="_FILE_", readonly=True,size=(60, 1)),sg.FolderBrowse('選擇路徑')],
          [sg.Text('請選擇EXCEL檔案輸出路徑(欄位擷取結果)：',  size=(50, 1))],
          [sg.Input('  ', key="_FILE_", readonly=True,size=(60, 1)),sg.FolderBrowse('選擇路徑')],
          #[sg.Text('請選擇資料庫CSV檔案(進行新舊案比對)：',  size=(50, 1))],
          #[sg.Input('  ', key="_FILE_", readonly=True,size=(60, 1)),sg.FileBrowse('選擇檔案')],
          [sg.Button('開始執行'),sg.Exit("離開")]]

# 建立窗口
window = sg.Window('法務文件小工具', layout,finalize=True)  # Part 3 - 窗口定义

#讀取窗口數值與窗口事件
event1,values = window.read()
window.close()

#點選離開見的執行事件
if(event1=='離開'):
    os._exit(0)
    



pdf_list=[]
#PDF、圖片檔案、文字檔案儲存路徑
pdf_path = values["選擇路徑"]
txt_path = values['選擇路徑1']
csv_path=values['選擇路徑3']
#database_csv_path=values['選擇檔案']

pdf_list = glob.glob(os.path.join(pdf_path, '*.pdf'))

layout = [[sg.Text('OCR完成進度')],
          [sg.ProgressBar(len(pdf_list), orientation='h', size=(20, 20), key='progressbar')]]

window = sg.Window('完成進度', layout)
progress_bar = window['progressbar']

event, values = window.read(timeout=10)

portnumbner=""
with open("localhostport.txt",'r',encoding="utf-8") as localhostport:
    portnumbner=localhostport.read()
    
situation_number=0
for i in pdf_list:
    if(situation_number!=len(pdf_list)):
        progress_bar.UpdateBar(situation_number)
        name = i.split('/')[-1].split('\\')[-1].split('.')[0]
        print(name)
        with open(i, "rb") as fd:
            document = fd.read()    
        endpoint = r"XXXX"
        apim_key = "XXXX"
        post_url = "http://"+str(portnumbner)+"/formrecognizer/v2.1/layout/syncAnalyze"
        #中間的localhost要改成可以用外面文字檔讀取(如果port變就可以在外面更改)
        
        params = {
            'language': 'zh-Hant',
            'pages': '1-',
            'readingOrder': 'natural',
            "includeTextDetails": True
            }
        headers = {
            # Request headers
            'Content-Type': 'application/octet-stream',
            }
        resp = post(url = post_url, data = document, headers = headers, params = params)
        final_result=resp.json()
        whole_text=""
        for i in range(0,len(final_result["analyzeResult"]["readResults"])):
            print(i)
            for j in range(0,len(final_result["analyzeResult"]["readResults"][i]["lines"])):
                print(j)
                whole_text=whole_text+final_result["analyzeResult"]["readResults"][i]["lines"][j]['text']
        
        txtfile = os.path.join(txt_path,name +'.txt')
        f = open(txtfile, 'w',encoding='utf-8-sig')
        f.write(whole_text)
        f.close()
    situation_number+=1 
window.close()


txt_file = txt_path
total_txt_file = glob.glob(os.path.join(txt_file, '*.txt'))

#這裡是取發文機關會用到的關鍵字
target = "執行"
txt_target = '署'
txt_target2 = '處'
txt_target3 = '院'
txt_target4='民事'

#這裡是取發文日期會用到的關鍵字
date_target = "中華民國"
date_target2="日"

#這裡是取身分證字號會用到的關鍵字
ID_target="統一編號"
ID_target2="身分證字號"
ID_target3="身分證號"

#這裡是取債務人會用到的關鍵字
name_target1="債務人"
name_target2="義務人"
name_target3="說明"
name_target4="主旨"


#這裡是股別會用到的關鍵字
division_target="股)"
division_target2="股書記官"
division_target3="股  "
division_target4="電話"
division_target5="聯絡方式"
division_target6="電話:"
division_target7="承辦股及電話:"

#這裡是發文字號會用到的關鍵字
wordnumber_target="發文字號:"

#這裡是分類會用到的關鍵字
category_target1="應予撤銷"
category_target1_2="撤銷"
category_target2="予以扣押"
category_target2_1="陳報扣押情形"
category_target2_2="陳明扣押情形"
category_target2_3="陳明扣押結果"
category_target2_4="規定扣押之"
category_target2_5="陳報執行扣押情形"
category_target3="代為變賣"
category_target3_1="代為拍賣"

category_target4_1="移送機關"
category_target4_2="陳報收取情形"


fin_data = []
flat = True
num = 0


agency_list=[]
date_list=[]
ID_list=[]
name_list=[]
division_list=[]
wordnumber_list=[]


#這邊進行發文機關欄位擷取
x=0
whole_txt=" "
for x,data in enumerate(total_txt_file):
    agency = ''
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            txt = i.split()
            for y in txt:
                if (flat == True) and (target in y) and (len(y)>8) and (len(y)<1000):
                    #print(y)
                    start1=y.find("法")
                    start2=y.find("臺")
                    if txt_target in y:
                        agency = y[start1:start1+12]
                        agency_list.append(agency)
                        #print(agency)
                        break
                    if txt_target2 in y:
                        agency = y[start2:start2+13]
                        agency_list.append(agency)
                        #print(agency)
                        break
                    if txt_target3 in y and txt_target4 not in y:
                        agency = y[start2:start2+8]
                        agency_list.append(agency)
                        #print(agency)
                        break
                    flat = False
                    break

#如果沒有進到任何一個條件，表示沒有比對出來，就給none
    if (len(agency_list) != x+1):
        agency_list.append("none")


#這邊進行發文機關欄位的後處理
new_agency_list=[]
for agency_word in agency_list:
    agency_word=agency_word.replace(" ","")
    agency_word=agency_word.replace(".","")
    if(len(agency_word)==0):
        agency_word="none"
        new_agency_list.append(agency_word)
    else:
        new_agency_list.append(agency_word)
agency_list=[]
agency_list=new_agency_list
        


#這邊進行發文日期欄位萃取
x=0
whole_txt=""            
for x,data in enumerate(total_txt_file):
    count=0
    agency = ''
    name = ''
    ID_number = ''
    No = ''
    issue_date = ''
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            flat=True
            whole_txt=i.replace(" ", "")
            #print(whole_txt)
            if flat == True and date_target in whole_txt and len(whole_txt)>1 and len(whole_txt)<2000:
                start=whole_txt.find('中華民國')
                #print(start)
                new_txt=whole_txt[start:]
                #print(new_txt)
                end=new_txt.find("日")
                issue_date=new_txt[:end+1]
                issue_date.replace(".","")
                issue_date.replace(" ","")
                date_list.append(issue_date)
                #print(issue_date)
                print()
                flat=False
                break

    if (len(date_list) != x+1):
        date_list.append("none")        
        
#這邊進行發文日期欄位的後處理
new_date_list=[]
for date_word in date_list:
    date_word=date_word.replace(" ","")
    date_word=date_word.replace(".","")
    if(len(date_word)==0):
        date_word="none"
        new_date_list.append(date_word)
    else:
        new_date_list.append(date_word)
date_list=[]
date_list=new_date_list

            
#這邊進行身分證字號欄位萃取 
whole_txt=""            
for x,data in enumerate(total_txt_file):
    count=0
    agency = ''
    name = ''
    ID_number = ''
    No = ''
    issue_date = ''
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            flat=True
            whole_txt=i.replace(" ", "")
            if flat == True and (ID_target in whole_txt or ID_target2 in whole_txt or ID_target3 in whole_txt) and len(whole_txt)>1 and len(whole_txt)<2000:
                count=count+1
                start=whole_txt.find('統一編號:')
                ID_number=whole_txt[start+5:start+15]
                if(start==-1):
                    start=whole_txt.find('身分證字號:')
                    print(start)
                    ID_number=whole_txt[start+6:start+16]
                if(start==-1):
                    start=whole_txt.find('身分證號:')
                    print(start)
                    ID_number=whole_txt[start+5:start+15]
                if(ID_number.find("號")!=-1):
                    start=whole_txt.find('身分證字號:')
                    ID_number=whole_txt[start+6:start+16]
                if ("*" in ID_number):
                    ID_number="none"
                print(ID_number[0])
                if(ID_number[0]!="A" and ID_number[0]!="B"
                   and ID_number[0]!="C" and ID_number[0]!="D"
                   and ID_number[0]!="E" and ID_number[0]!="F"
                   and ID_number[0]!="G" and ID_number[0]!="H"
                   and ID_number[0]!="I" and ID_number[0]!="J"
                   and ID_number[0]!="K" and ID_number[0]!="L"
                   and ID_number[0]!="M" and ID_number[0]!="N"
                   and ID_number[0]!="O" and ID_number[0]!="P"
                   and ID_number[0]!="Q" and ID_number[0]!="R"
                   and ID_number[0]!="S" and ID_number[0]!="T"
                   and ID_number[0]!="U" and ID_number[0]!="V"
                   and ID_number[0]!="W" and ID_number[0]!="X"
                   and ID_number[0]!="Y" and ID_number[0]!="Z"):
                    ID_number="none"
                ID_list.append(ID_number)
                print(ID_number)
                print(data)
                print()
                flat=False
                break
#如果沒有進到任何一個條件，表示沒有比對出來，就給none
    if (len(ID_list) != x+1):
        ID_list.append("none")
        
        
#這邊進行發文日期欄位的後處理
new_ID_list=[]
for ID_word in ID_list:
    ID_word=ID_word.replace(" ","")
    ID_word=ID_word.replace(".","")
    if(len(ID_word)==0):
        ID_word="none"
        new_ID_list.append(ID_word)
    else:
        new_ID_list.append(ID_word)
ID_list=[]
ID_list=new_ID_list
            
            

#這邊進行義務人欄位擷取
count=0
whole_txt=""            
for x,data in enumerate(total_txt_file):
    agency = ''
    name = ''
    ID_number = ''
    No = ''
    issue_date = ''
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            flat=True
            whole_txt=i.replace(" ", "")
            #print(whole_txt)
            if flat == True and (name_target1 in whole_txt) or(name_target2 in whole_txt) and len(whole_txt)>1 and len(whole_txt)<2000:
                check_txt=whole_txt
                #print(check_txt)
                start=check_txt.find('義務人')
                if(start==-1):
                    start=check_txt.find('債務人')
                    name=check_txt[start+3:start+6]
                    if(name[0]=="所" or name[0]=="現" or name[0]=="之"):
                        start=check_txt.find('債務人', check_txt.find('債務人') + 1)
                        name=check_txt[start+3:start+6]
                        name_list.append(name)
                    else:
                        name_list.append(name)
                    flat=False
                    break
                if(start!=1):
                    name=check_txt[start+3:start+6]
                    if(name[0]=="所" or name[0]=="現" or name[0]=="之"):
                        start=check_txt.find('義務人', check_txt.find('義務人') + 1)
                        name=check_txt[start+3:start+6]
                        name_list.append(name)
                    else:
                        name_list.append(name)
                    flat=False
                    break
#如果沒有進到任何一個條件，表示沒有比對出來，就給none
    if (len(name_list) != x+1):
        name_list.append("none")
        
        
#這邊進行義務人欄位的後處理
new_name_list=[]
for name_word in name_list:
    name_word=name_word.replace(" ","")
    name_word=name_word.replace(".","")
    if(len(name_word)==0):
        name_word="none"
        new_name_list.append(name_word)
    else:
        new_name_list.append(name_word)
        
name_list=[]
name_list=new_name_list
   


#這邊是判斷股別            
whole_txt=" "
for x,data in enumerate(total_txt_file):
    y=" "
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            txt = i.split()
            for y in txt: 
                if flat == True and (division_target6 in y ) and len(y)>0 and len(y)<200:
                    start=y.find('股')
                    if(y[start-2]==":"):
                        #print(check_txt[start+3])
                        division=y[start-1:start+1]
                        division_list.append(division)
                        flat = False
                        break
                if flat == True and (division_target5 in y ) and len(y)>0 and len(y)<200:
                    start=y.find('股')
                    if(y[start-2]==":"):
                        division=y[start-1:start+1]
                        division_list.append(division)
                        flat = False
                        break
                if flat == True and (division_target in y or division_target2 in y or division_target3 in y) or (division_target3 in y and division_target4 in y) and len(y)>0 and len(y)<100:
                    start=y.find('股')
                    if(y[start-1]!=")"):
                        #print(check_txt[start+3])
                        division=y[start-1:start+1]
                        division_list.append(division)
                        flat = False
                        break
                    if(y[start-3]==":"):
                        division=y[start-1:start+1]
                        division_list.append(division)
                        flat = False
                        break
                    else:
                        division_list.append("none")
                        flat = False
                        break
                if flat == True and (division_target7 in y) and len(y)>0 and len(y)<300:
                    start=y.find('承辦股及電話:')
                    division=y[start+7:start+9]
                    division_list.append(division)
                    flat = False
                    break
#如果沒有進到任何一個條件，表示沒有比對出來，就給none
    if (len(division_list) != x+1):
        division_list.append("none")


#這邊進行股別欄位的後處理
new_divison_list=[]
for division_word in division_list:
    if(len(division_word)==0 or division_word[0]=="押"):
        division_word="none"
        new_divison_list.append(division_word)
    elif(division_word[0]=="西"):
        division_word=division_word.replace("西", "酉")
        new_divison_list.append(division_word)
    else:
        new_divison_list.append(division_word)
division_list=[]
division_list=new_divison_list
        


count=0
#這邊是判斷發文字號            
whole_txt=""            
for x,data in enumerate(total_txt_file):
    word_number = ''
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            flat=True
            whole_txt=i.replace(" ", "")
            if flat == True and wordnumber_target in whole_txt and len(whole_txt)>1 and len(whole_txt)<2000:
                count+=1
                start=whole_txt.find('發文字號:')
                new_txt=whole_txt[start:]
                end=new_txt.find("速")
                wordnumer=new_txt[:end]
                wordnumer = wordnumer.replace("發文字號:","")
                wordnumber_list.append(wordnumer)
                print()
                flat=False
                break
#如果沒有進到任何一個條件，表示沒有比對出來，就給none
    if (len(wordnumber_list) != x+1):
        wordnumber_list.append("none")
            

#這邊進行發文字號欄位的後處理
new_wordnumber_list=[]
for wordnumber_word in wordnumber_list:
    wordnumber_word=wordnumber_word.replace(" ","")
    wordnumber_word=wordnumber_word.replace(".","")
    if(len(wordnumber_word)==0):
        name_word="none"
        new_wordnumber_list.append(wordnumber_word)
    else:
        new_wordnumber_list.append(wordnumber_word)
wordnumber_list=[]
wordnumber_list=new_wordnumber_list            



#這邊是判斷分類
cat_list=[]    
whole_txt=""            
for cat_x,data in enumerate(total_txt_file):
    print(data)
    flat = True
    with open(data,'r',newline=(''),encoding="utf-8") as txt_read:
        for n,i in enumerate(txt_read.readlines()):
            flat=True
            whole_txt=i.replace(" ", "")
            #print(whole_txt)
            if flat == True and (category_target1 in whole_txt or category_target1_2 in whole_txt) and len(whole_txt)>1 and len(whole_txt)<2000:
                cat_list.append("撤銷")
                flat=False
                break
            if flat == True and (category_target2 in whole_txt or category_target2_1 in whole_txt or category_target2_2 in whole_txt or category_target2_3 in whole_txt or category_target2_4 in whole_txt or category_target2_5 in whole_txt) and len(whole_txt)>1 and len(whole_txt)<2000:
                cat_list.append("扣押")
                flat=False
                break
            if flat == True and (category_target3 in whole_txt or category_target3_1 in whole_txt) and len(whole_txt)>1 and len(whole_txt)<2000:
                cat_list.append("拍賣")
                flat=False
                break
            if flat == True and (category_target4_2 in whole_txt or category_target4_1 in whole_txt) and len(whole_txt)>1 and len(whole_txt)<2000:
                cat_list.append("移轉")
                flat=False
                break
    if (len(cat_list) != cat_x+1):
        cat_list.append("none")
        
        
#開始讀取系統日期
systemtime_list=[]
collect_wordnumber_list=[]
localtime = time.localtime()
year=time.strftime("%Y",localtime)
year=int(year)-1911#西元年轉換成民國年
month_day = time.strftime(".%m.%d", localtime)
system_time=str(year)+month_day
re_word_definition=((round(year/10))*(10**4))


for i in range(0,len(division_list)):
    re_word_definition+=1
    systemtime_list.append(system_time)
    collect_wordnumber_list.append(re_word_definition)


#
df_reword = pd.DataFrame(collect_wordnumber_list, columns = ['收文字號'])                
df_redate = pd.DataFrame(systemtime_list, columns = ['收文日期'])
df_godate = pd.DataFrame(date_list, columns = ['發文日期'])
df_agency = pd.DataFrame(agency_list, columns = ['發文機關'])
df_wordnumber = pd.DataFrame(wordnumber_list, columns = ['發文字號'])
df_division=pd.DataFrame(division_list, columns = ['股別'])
df_name=pd.DataFrame(name_list, columns = ['義務人/債務人'])
df_ID=pd.DataFrame(ID_list, columns = ['身分證字號'])
df_cat=pd.DataFrame(cat_list, columns = ['文件分類'])


#進行所有資料整併
final_df=pd.DataFrame()
final_df = df_reword.merge(df_redate, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_godate, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_agency, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_wordnumber, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_division, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_name, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_ID, how='inner', left_index=True, right_index=True)
final_df = final_df.merge(df_cat, how='inner', left_index=True, right_index=True)

final_df.to_excel(csv_path+"/"+system_time+"法務命令.xlsx",encoding="utf-8",index=False)


#把所有資料夾內的文字檔案刪除，後面才不會又重新讀取
# =============================================================================
# for z in total_txt_file:
#     os.remove(z)
# =============================================================================

sg.popup('恭喜您，資料分析完成，資料已存放至'+csv_path)

layout = [[sg.Text("excel檔案以存放至")],
          [sg.Text(csv_path)]]

window = sg.Window('解析完成', layout)
