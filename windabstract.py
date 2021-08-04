# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

print('go')
import time
import uiautomation as auto
import re
import xlwt

#locate myreport and open myreport page
wind = auto.WindowControl(searchDepth=1, ClassName="TMasterForm")
#report = wind.WindowControl(ClassName='TfrmMyIEPageContainer')
#table = report.TableControl(foundIndex=8)
#goals = table.GetChildren()

#goalcount = 0

all_abstract = []
all_information = []
all_title = [] 
while True:
    try:
        report = wind.WindowControl(ClassName='TfrmMyIEPageContainer')
        table = report.TableControl(foundIndex=8)
        goals = table.GetChildren()
        
        goalcount = 0
        for goal in goals:
            #get title
            p0 = goal.DocumentControl()
            p1 = p0.GetChildren()
            p2 = p1[1].GetChildren()
            title = ''.join([item.Name for item in p2])
            all_title.append(title)
            #open a specific report
            link = goal.HyperlinkControl()
            link.Click()
            
            time.sleep(1)
            #locate abstract text in myreport
            myreport = wind.WindowControl(ClassName='TfrmWebBrowser')
            mygoal = myreport.DocumentControl(foundIndex=1)
            mygoal1 = mygoal.GetChildren()
            mygoal2 = mygoal1[0].GetChildren()
            mygoal3 = mygoal2[1].GetChildren()
            mygoal4 = mygoal3[0].GetChildren()
            mygoal5 = mygoal4[0].GetChildren()
            mygoal_abstract = mygoal5[2].GetChildren()
            mygoal_information = mygoal5[1].GetChildren()
            
            #get information
            info = []
            for item in mygoal_information[:-2]:
                if  len(item.Name) > 0 :
                    info.append(item.Name)
                else:
                    tmp = item.GetChildren()
                    info.append(tmp[0].Name)
            
            info = ''.join(info)
            all_information.append(info)
            
            #get abstract
            lines = []
            for item in mygoal_abstract:
                texts = item.GetChildren()
                line = []
                for text in texts:
                    line.append(text.Name)
                line = ''.join(line)
                #print(line)
                lines.append(line)
                
            abstract = '\n'.join(lines)
            abstract = re.sub('\u3000','',abstract)
            all_abstract.append(abstract)
            #save in a txt file
            #filename = 'report' + str(goalcount)
            #path = 'E:/Windreport/' + filename + '.txt'
            #with open(path,'w') as f:   
                 #f.write(abstract)
                 
            #close myreport page
            auto.SendKeys("{Ctrl}w")
            goalcount += 1
            
            if goalcount %5 == 0:
                auto.WheelDown(wheelTimes=3,interval=0.01)
                
        nextarea = report.ListControl(foundIndex = 2)
        nextitems = nextarea.GetChildren()
        nextpage = nextitems[-2]
        
        if nextpage.IsEnabled:
            nextpage.Click()
        else:
            break
    except:
       break


#save in an xls file    
f = xlwt.Workbook() #创建工作
sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet

j = 0
for title in all_title:
    sheet1.write(j,0,title)
    j=j+1
     
j = 0
for information in all_information:
    sheet1.write(j,1,information)
    j=j+1

j = 0
for abstract in all_abstract:
    sheet1.write(j,2,abstract)
    j=j+1

f.save('E:/Windreport/test.xls')

print('finish')     
