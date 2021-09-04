# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#需要手动下载一次改变默认下载路径
print('go')
import time
import uiautomation as auto

#locate myreport and open myreport page
wind = auto.WindowControl(searchDepth=1, ClassName="TMasterForm")
pages = 5
weekcheck = '第30周' ########################
for page in range(pages):
    
    try:
        report = wind.WindowControl(ClassName='TfrmMyIEPageContainer')
        table = report.TableControl(foundIndex=8)
        goals = table.GetChildren()
        
        goalcount = 0
        for goal in goals:
            
            #open a specific report
            link = goal.HyperlinkControl()
            linkname = link.Name
            print(linkname)
            if weekcheck in linkname:###################
                link.Click() ##################
            else:
                continue
            
            time.sleep(1)
            #locate abstract text in myreport
            myreport = wind.WindowControl(ClassName='TfrmWebBrowser')
            mygoal = myreport.DocumentControl(foundIndex=1)
            mygoal1 = mygoal.GetChildren()
            mygoal2 = mygoal1[0].GetChildren()
            mygoal3 = mygoal2[1].GetChildren()
            mygoal4 = mygoal3[0].GetChildren()
            mygoal5 = mygoal4[0].GetChildren()
            mygoal6 = mygoal5[1].GetChildren()
            mygoal_download_area = mygoal6[-2].GetChildren()
            mygoal_download_button = mygoal_download_area[-1].GetChildren()
            print('locate')
            #download report
            if len(mygoal_download_button) > 0:    
                time.sleep(0.2)
                mygoal_download_button[0].Click()
                time.sleep(0.2)
                auto.SendKeys("{Right}")
                time.sleep(0.2)
                auto.SendKeys("{Enter}")
                auto.SendKeys("{Enter}")
                time.sleep(1)
                auto.SendKeys("{Space}")
                auto.SendKeys("{Ctrl}w")
            else:
                print('no download')
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



print('finish')     
