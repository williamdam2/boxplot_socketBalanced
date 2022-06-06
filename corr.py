from cmath import log
from doctest import master
from operator import le
import pandas as pd
from matplotlib import pyplot as plt
import seaborn
import statistics
import xlsxwriter
from datetime import datetime
import os
import argparse
import win32com.client as win32

pd.option_context('display.max_rows', None, 'display.max_columns', None)
pd.options.mode.chained_assignment = None 
    

sheetList = []
summaryRowList = []

ap = argparse.ArgumentParser()
ap.add_argument('-s', '--setting', required=True)
ap.add_argument('-d', '--data', required=True)

args = ap.parse_args()

thisPath = os.path.dirname(os.path.abspath(__file__))
print(thisPath)

def remove_header(dataframe):
    dataframe=dataframe[dataframe['station'] != "station"]
    dataframe.reset_index(drop=True, inplace=True)
    dataframe.to_csv("temp.CSV",index=False)
    dataframe = pd.read_csv("temp.CSV")
    return dataframe

def makeReportSheet(workbook,setting,masterMachine):
    global sheetList

    greenCell = workbook.add_format()
    greenCell.set_bg_color("#00dd00")
    greenCell.set_border(1)
    grayCellCenter = workbook.add_format()
    grayCellCenter.set_border(1)
    grayCellCenter.set_bg_color("#cccccc")
    grayCellCenter.set_align("Center")
    grayCellCenter.set_bold()
    redCell = workbook.add_format()
    redCell.set_bg_color("#ffb300")
    redCell.set_border(1)
    grayCell = workbook.add_format()
    grayCell.set_bg_color("#cccccc")
    grayCell.set_bold()
    grayCell.set_border(1)
    yellowCell = workbook.add_format()
    yellowCell.set_bg_color("#ffff1a")
    yellowCell.set_border(1)
    normalCell = workbook.add_format()
    normalCell.set_border(1)
    whiteCell = workbook.add_format()
    whiteCell.set_bold()
    whiteCell.set_border(1)
    bigWhiteCell = workbook.add_format()
    bigWhiteCell.set_bold()
    bigWhiteCell.set_border(1)
    bigWhiteCell.set_font_size(15)

    

    logFileName = setting.log
    sheetName = ""
    stringTemp = ""
    if(len(logFileName[0]) <= 30):
        stringTemp = "Make Sheet: {} ".format(logFileName[0])
        sheetName = logFileName[0]
    elif(len(logFileName[0]) > 30):
        stringTemp = "Make Sheet: {} ".format(logFileName[0][:30])
        sheetName = logFileName[0][:30]

    print(stringTemp)
    
    if(os.path.exists(thisPath+"\\image\\"+logFileName[0]) == False):
        os.makedirs(thisPath+"\\image\\"+logFileName[0])
    worksheet = workbook.add_worksheet(sheetName)

    sheetList.append(sheetName)

    worksheet.write(0,0,"CORR Summary Report",bigWhiteCell)
    dt = datetime.now()
    str_date = dt.strftime("%d %B, %Y")
    worksheet.write(1,0,"Date: "+str_date,bigWhiteCell)
    worksheet.write(2,0,"List of Machine: ",bigWhiteCell)
    listCategory = setting.category
    countCate = 0
    for cate in listCategory:
        if str(cate) == "nan":
            break
        else:
            countCate += 1
    summaryRowList.append(countCate)

    #dataspec
    listDeltaSpecRed = setting.redSpec
    listDeltaSpecYellow = setting.yellowSpec

    listDeltaSpecRed = setting.redSpec
    listDeltaSpecYellow = setting.yellowSpec

    Alldata = pd.read_csv(args.data+"/"+str(logFileName[0])+".CSV")
    Alldata = remove_header(Alldata)

    listMachine = Alldata["station"].unique()
    
    
    masterStationName = ""
    for machine in listMachine:
        if(machine[5:9]==masterMachine):
            masterStationName = machine
    machineCount = 0
    if(masterStationName != ""):
        masterData = Alldata[Alldata["station"]==masterStationName]
        print("Master machine is",masterStationName)
        
        for machine in listMachine:
            stepColumn = 15
            offsetColumnImage=6   + stepColumn*machineCount
            offsetColumntable = 6 + stepColumn*machineCount
            offsetChileTable = 26
            offsetRow = 12  
            stepRow = 40

            if(os.path.exists(thisPath+"\\image\\"+logFileName[0]+"\\"+machine[5:9]) == False):
                os.makedirs(thisPath+"\\image\\"+logFileName[0]+"\\"+machine[5:9])
            
            imageMachinePath = thisPath+"\\image\\"+logFileName[0]+"\\"+machine[5:9]+"\\"
            
            worksheet.write(3,offsetColumntable,"Machine:",bigWhiteCell)
            worksheet.write(3,offsetColumntable+1,"Master",grayCell)
            worksheet.write(2,offsetColumntable+1,masterStationName[5:9],yellowCell)
            
            worksheet.write(3,offsetColumntable+2,"Tester",grayCell)
            worksheet.write(2,offsetColumntable+2,machine[5:9],grayCell)

            worksheet.write(3,offsetColumntable+3,"Median Master - Median Tester",grayCell)
            count = 0
            for cate in listCategory:
                if(str(cate) == "nan"):
                    break
                else:
                    columnLatter = xlsxwriter.utility.xl_col_to_name(offsetColumnImage+11)
                    worksheet.write_url(4+count,offsetColumntable,"internal:"+str(logFileName[0])+"!"+columnLatter+str(count*stepRow+offsetRow+stepRow+1))
                    worksheet.write(4+count,offsetColumntable,cate,whiteCell)
                    worksheet.set_column(offsetColumntable,offsetColumntable,35)
                    count = count + 1
            
            #print("so luong cate: ",count)
            offsetRow +=count

            if(machine[5:9]!=masterMachine):
                corrData = Alldata[Alldata["station"]==machine]
                mergeData = Alldata[ (Alldata["station"]==masterStationName) | (Alldata["station"]==machine) ]
                
                for i in range(0,mergeData.shape[0]):
                    if(mergeData["station"][i]!=masterStationName):
                        #print(mergeData["site"][i]+8)
                        newSite = mergeData["site"][i]+8
                        mergeData["site"][i] = newSite
                
                # boxplot = plt.figure(figsize=(10, 5))
                # boxplot = seaborn.boxplot(x="site",y=setting.category[0],data=mergeData,width=0.8,hue="station").set(xlabel="site",ylabel="value",title=str(setting.category[0]))
                # plt.subplots_adjust(right=1,left=0.1)
                # #plt.show()
                count = 0
                for cate in listCategory:
                    if(str(cate) == "nan"):
                        break
                    print("---> ",cate)
                    deltaSpecRed = listDeltaSpecRed[count]
                    deltaSpecYellow = listDeltaSpecYellow[count]
                    medianOfMaster = statistics.median(masterData[cate])

                    boxplot = plt.figure(figsize=(10, 5))
                    boxplot = seaborn.boxplot(x="site",y=cate,data=mergeData,width=0.8,hue="station").set(xlabel="site",ylabel="value",title=str(cate))
                    plt.subplots_adjust(right=1,left=0.1)

                    plt.plot([0,16],[medianOfMaster,medianOfMaster],color="green",linestyle='--')
                    plt.text(x=15.5,y=medianOfMaster,s=str(round(medianOfMaster,4)),color="green",size="small" )
                    plt.savefig(imageMachinePath+"image" + str(count+1) + ".png")
                    plt.close()
                    #plt.show()

                    
                    worksheet.insert_image(count*stepRow+offsetRow-2,offsetColumnImage,imageMachinePath+"image" + str(count+1) + ".png")
                    worksheet.write(count*stepRow+offsetRow-3+offsetChileTable,offsetColumntable, logFileName[0], grayCell)
                    worksheet.write(count*stepRow+offsetRow-2+offsetChileTable,offsetColumntable, cate, whiteCell)
                    worksheet.write(count*stepRow+offsetRow-1+offsetChileTable,offsetColumntable, "Machine", grayCellCenter)
                    worksheet.write(count*stepRow+offsetRow+offsetChileTable,offsetColumntable, "Median",grayCell)
                    worksheet.write(count*stepRow+offsetRow+offsetChileTable-1,offsetColumntable+3, "Delta = Master Median - Tester Median",grayCell)
                    worksheet.set_column(offsetColumntable+3,offsetColumntable+3,40)
                    worksheet.write(count*stepRow+offsetRow+1+offsetChileTable,offsetColumntable, "Median of Master = "+ str(round(medianOfMaster,4)), greenCell)
                    worksheet.write(count*stepRow+offsetRow+2+offsetChileTable,offsetColumntable, "Delta Spec Yellow = " + str(deltaSpecYellow), yellowCell)
                    worksheet.write(count*stepRow+offsetRow+3+offsetChileTable,offsetColumntable, "Delta Spec Red = "+str(deltaSpecRed),redCell)
                    columnLatter = xlsxwriter.utility.xl_col_to_name(offsetColumntable)
                    worksheet.write_url(count*stepRow+offsetRow+offsetChileTable-4,offsetColumntable,"internal:"+str(logFileName[0])+"!"+columnLatter+"1")
                    worksheet.write(count*stepRow+offsetRow+offsetChileTable-4,offsetColumntable, "Back to TOP",grayCell)
                    
                    summaryRow = 4+count
                    worksheet.write(summaryRow ,offsetColumntable+1,round(medianOfMaster,4),whiteCell)
                    
                    medianOfTester = statistics.median(corrData[cate])
                    delta = medianOfMaster - medianOfTester
                    absDelta = abs(delta)
                    columnTable = 2 + offsetColumntable
                    worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable,columnTable-1,"Master",yellowCell)
                    worksheet.write((count*stepRow+offsetRow-2)+offsetChileTable,columnTable-1,masterStationName[5:9],grayCellCenter)

                    worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable,columnTable,"Tester",grayCellCenter)
                    worksheet.write((count*stepRow+offsetRow-2)+offsetChileTable,columnTable,machine[5:9],grayCellCenter)
                    worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable+1,columnTable-1,round(medianOfMaster,4),whiteCell)
                    
                    sign = ""
                    if(delta > 0):
                        sign = "+"
                    ##############Phân định Đỏ Vàng Xanh########################
                    worksheet.write(summaryRow,columnTable,round(medianOfTester,4),whiteCell)
                    if(absDelta > deltaSpecRed):
                        worksheet.write((count * stepRow + offsetRow)+offsetChileTable, columnTable+1,sign+ str(round(delta,4)), redCell)
                        worksheet.write(summaryRow,columnTable+1,sign+ str(round(delta,4)),redCell)
                        
                        worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable+1,columnTable,medianOfTester,redCell)
                    #Vàng
                    elif(absDelta < deltaSpecRed and absDelta > deltaSpecYellow):
                        worksheet.write((count * stepRow + offsetRow)+offsetChileTable, columnTable+1, sign+ str(round(delta,4)), yellowCell)
                        worksheet.write(summaryRow,columnTable+1,sign+ str(round(delta,4)),yellowCell)
                        worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable+1,columnTable,medianOfTester,yellowCell)
                    #Xanh
                    elif(absDelta<deltaSpecYellow):
                        worksheet.write((count * stepRow + offsetRow)+offsetChileTable, columnTable+1, sign+ str(round(delta,4)), greenCell)
                        worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable+1,columnTable,medianOfTester,greenCell)
                        worksheet.write(summaryRow,columnTable+1,sign+ str(round(delta,4)),greenCell)
                    for j in range(0,1000):
                            worksheet.write(count*stepRow-4+offsetRow,j,"",grayCell)                  
                    count +=1
                
                machineCount+=1
                


def main():
    returnString = "Success"
    seaborn.set(style='whitegrid')
    workbook = xlsxwriter.Workbook('.\\report.xlsx')

    settingCount = 0
    rawSetting = pd.read_csv(args.setting)
    dataFrameColumn = ["log","category","redSpec","yellowSpec","master"]
    setting = pd.DataFrame(columns = dataFrameColumn)
    if(os.path.exists(thisPath+"\\image") == False):
        os.makedirs(thisPath+"\\image")
    
    
    while(True):
        try:
            settingCount += 1
            setting.log = rawSetting["log"+str(settingCount)]
            setting.category = rawSetting["category"+str(settingCount)]
            setting.redSpec = rawSetting["redSpec"+str(settingCount)]
            setting.yellowSpec = rawSetting["yellowSpec"+str(settingCount)]
            masterMachine = str(int(rawSetting["master"][0]))

            if(str(setting.log[0]) == "nan"):
                break
            makeReportSheet(workbook,setting,masterMachine) 
        except Exception as e:   
            print ("Kiểm tra lại file setting")
            returnString = "Failure"
            
    workbook.close()  
    return returnString



result = ""
if __name__ == "__main__":
    result = main()
input("Report "+result+", press ENTER to exit...")   
