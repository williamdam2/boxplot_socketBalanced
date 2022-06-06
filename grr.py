from sre_constants import SUCCESS
import pandas as pd
from matplotlib import pyplot as plt
import seaborn
import statistics
import xlsxwriter
from datetime import datetime
import os
import argparse
import win32com.client as win32
import pathlib
import os

resultText = "GRR Report\r\n"

thisPath = os.path.dirname(os.path.abspath(__file__))
print(thisPath)

ap = argparse.ArgumentParser()
ap.add_argument('-s', '--setting', required=True)
ap.add_argument('-d', '--data',required=True)

args = ap.parse_args()


summaryColStep = 10
sheetList = []
summaryRowList = []
numberOfMachines = 0


def remove_header(dataframe):
    dataframe=dataframe[dataframe['station'] != "station"]
    dataframe.reset_index(drop=True, inplace=True)
    dataframe.to_csv("temp.CSV",index=False)
    dataframe = pd.read_csv("temp.CSV")
    return dataframe

def makeReportSheet(workbook,setting,summarySheet):
    global resultText
    global sheetList
    global summaryRowList

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


    #lấy file Log cần phân tích
    logFileName = setting.log
    
    stringTemp = ""
    sheetName = ""
    print(len(logFileName[0]))
    
    if(len(logFileName[0]) <= 30):
        stringTemp = "Make Sheet: {} ".format(logFileName[0])
        sheetName = logFileName[0]
    elif(len(logFileName[0]) > 30):
        stringTemp = "Make Sheet: {} ".format(logFileName[0][:30])
        sheetName = logFileName[0][:30]

    worksheet = workbook.add_worksheet(sheetName)    
    print(stringTemp)
    resultText += stringTemp + "\r\n"
    

    if(os.path.exists(thisPath+"\\image\\"+logFileName[0]) == False):
        os.makedirs(thisPath+"\\image\\"+logFileName[0])
    
    
    sheetList.append(sheetName)
    worksheet.write(0,0,"GRR Summary Report",bigWhiteCell)
    dt = datetime.now()
    str_date = dt.strftime("%d %B, %Y")
    worksheet.write(1,0,"Date: "+str_date,bigWhiteCell)
    worksheet.write(2,0,"List of Machine: ",bigWhiteCell)
    
    
    #lấy các chỉ số quan tâm
    listCategory = setting.category
    countCate = 0
    for cate in listCategory:
        if str(cate) == "nan":
            break
        else:
            countCate += 1
    summaryRowList.append(countCate)
    
    #các delta spec
    listDeltaSpecRed = setting.redSpec
    listDeltaSpecYellow = setting.yellowSpec
    #lấy data
    Alldata = pd.read_csv(args.data+"/"+str(logFileName[0])+".CSV")
    #print(listCategory)
    Alldata = remove_header(Alldata)
    
    listMachine = Alldata["station"].unique()
    machineCount = 0
    for machine in listMachine:
        if(os.path.exists(thisPath+"\\image\\"+logFileName[0]+"\\"+machine[5:9]) == False):
            os.makedirs(thisPath+"\\image\\"+logFileName[0]+"\\"+machine[5:9])
        stringtemp = "machine number "+str(machineCount+1)+": "+machine[5:9]
        print(stringtemp)
        resultText += stringtemp + "\r\n"

        imageMachinePath = thisPath+"\\image\\"+logFileName[0]+"\\"+machine[5:9]+"\\"
        data = Alldata[Alldata["station"]==machine]

        stepColumn = 15
        offsetColumnImage=6   + stepColumn*machineCount
        offsetColumntable = 6 + stepColumn*machineCount
        offsetChileTable = 26
        offsetRow = 12  
        stepRow = 35

        columnLatter = xlsxwriter.utility.xl_col_to_name(offsetColumntable+10)
        worksheet.write_url(3+machineCount,0,"internal:"+str(logFileName[0])+"!"+columnLatter+str(1))
        worksheet.write(3+machineCount,0,machine[5:9],normalCell)
        #summary table
        tempMachineCountPre = 0
        if(machineCount==0):
            tempMachineCountPre = 0
        if(machineCount > 0):
            tempMachineCountPre = machineCount-1
        columnLatter = xlsxwriter.utility.xl_col_to_name(tempMachineCountPre*stepColumn)
        worksheet.write_url(1,offsetColumntable-1,"internal:"+str(logFileName[0])+"!"+columnLatter+str(1))
        worksheet.write(1,offsetColumntable-1,"<<Pre ",grayCell)
        columnLatter = xlsxwriter.utility.xl_col_to_name((machineCount+1)*stepColumn+stepColumn-1)
        worksheet.write_url(1,offsetColumntable+1,"internal:"+str(logFileName[0])+"!"+columnLatter+str(1))
        worksheet.write(1,offsetColumntable+1,"Next>> ",grayCell)
        worksheet.write_url(0,offsetColumntable,"internal:"+str(logFileName[0])+"!A1")
        worksheet.write(0,offsetColumntable,"Back to First",whiteCell)
        dt = datetime.now()
        str_date = dt.strftime(" (%d-%B)")
        worksheet.write(1,offsetColumntable,"Machine: "+ machine[5:9] + str_date,bigWhiteCell)
        summarySheet.write(1,machineCount*summaryColStep,"Machine: "+ machine[5:9] + str_date,bigWhiteCell)
        summarySheet.set_column(machineCount*summaryColStep,machineCount*summaryColStep,32)
        worksheet.write(2,offsetColumntable,"<<- Summary table "+logFileName[0]+" ->>",whiteCell)
        worksheet.set_column(offsetColumntable,offsetColumntable,40)
        worksheet.write(3,offsetColumntable,"Socket",grayCellCenter)
        for i in range(1,9):
            worksheet.write(3,offsetColumntable+i,i,grayCellCenter)
        worksheet.write(3,offsetColumntable+i+2,"Medians All",grayCell)
        worksheet.set_column(offsetColumntable+i+2,offsetColumntable+i+2,15)
        
        for j in range(0,1000):
            worksheet.write(j,offsetColumntable+i+4,"",grayCell)

        count = 0
        for cate in listCategory:
            if(str(cate) == "nan"):
                break
            else:
                columnLatter = xlsxwriter.utility.xl_col_to_name(offsetColumnImage+11)
                worksheet.write_url(4+count,offsetColumntable,"internal:"+str(logFileName[0])+"!"+columnLatter+str(count*stepRow+offsetRow+stepRow+1))
                worksheet.write(4+count,offsetColumntable,cate,whiteCell)

                count = count + 1
        
        #print("so luong cate: ",count)
        offsetRow +=count


        count = 0
        for cate in listCategory:
            if(str(cate) == "nan"):
                break
            stringtemp = "---> {}".format(cate)
            print(stringtemp)
            resultText += stringtemp +"\r\n"
            #vẽ boxplot
            boxplot = seaborn.boxplot(x="site",y=cate,data=data,color="skyblue",width=0.5).set(xlabel="site",ylabel="value",title=str(cate))
            #vẽ đường median
            medianOfAllSocket = statistics.median(data[cate])
            plt.plot([-1, 9], [medianOfAllSocket, medianOfAllSocket], color="green", )
            plt.text(x=9, y=medianOfAllSocket, s=str(round(medianOfAllSocket,6)), color="green", size="x-small")
            deltaSpecRed = listDeltaSpecRed[count]
            deltaSpecYellow = listDeltaSpecYellow[count]
            #vẽ đường delta median
            plt.plot([-1, 9], [medianOfAllSocket+deltaSpecYellow,medianOfAllSocket+deltaSpecYellow], color="yellow",linestyle='--' )
            plt.plot([-1, 9], [medianOfAllSocket-deltaSpecYellow,medianOfAllSocket-deltaSpecYellow], color="yellow",linestyle='--' )
            plt.plot([-1, 9], [medianOfAllSocket + deltaSpecRed, medianOfAllSocket + deltaSpecRed], color="#ffb300", linestyle='--')
            plt.plot([-1, 9], [medianOfAllSocket - deltaSpecRed, medianOfAllSocket - deltaSpecRed], color="#ffb300", linestyle='--')
            #danh sách các median
            medians = data.groupby(['site'])[cate].median()
            #print(medians)
            #lưu plot
            plt.savefig(imageMachinePath+"image" + str(count+1) + ".png")


            #insert wooksheet
            ###this table
            worksheet.insert_image(count*stepRow+offsetRow-2,offsetColumnImage,imageMachinePath+"image" + str(count+1) + ".png")
            worksheet.write(count*stepRow+offsetRow-3+offsetChileTable,offsetColumntable, logFileName[0], grayCell)
            worksheet.write(count*stepRow+offsetRow-2+offsetChileTable,offsetColumntable, cate, whiteCell)
            worksheet.write(count*stepRow+offsetRow-1+offsetChileTable,offsetColumntable, "Socket", grayCellCenter)
            worksheet.write(count*stepRow+offsetRow+offsetChileTable,offsetColumntable, "Median",grayCell)
            worksheet.write(count*stepRow+offsetRow+1+offsetChileTable,offsetColumntable, "Delta = |Socket Median - Median All|",grayCell)
            worksheet.write(count*stepRow+offsetRow+2+offsetChileTable,offsetColumntable, "Median of all socket = "+ str(round(medianOfAllSocket,6)), greenCell)
            worksheet.write(count*stepRow+offsetRow+3+offsetChileTable,offsetColumntable, "Delta Spec Yellow = " + str(deltaSpecYellow), yellowCell)
            worksheet.write(count*stepRow+offsetRow+4+offsetChileTable,offsetColumntable, "Delta Spec Red = "+str(deltaSpecRed),redCell)
            columnLatter = xlsxwriter.utility.xl_col_to_name(offsetColumntable)
            worksheet.write_url(count*stepRow+offsetRow+offsetChileTable-4,offsetColumntable,"internal:"+str(logFileName[0])+"!"+columnLatter+"1")
            worksheet.write(count*stepRow+offsetRow+offsetChileTable-4,offsetColumntable, "Back to TOP",grayCell)
            ###summary table
            summaryRow = 4+count
            worksheet.write(summaryRow ,offsetColumntable+10,round(medianOfAllSocket,6),greenCell)

##############Phân định Đỏ Vàng Xanh########################
            for i in range(1,9):
                columnTable = i + offsetColumntable
                worksheet.write((count*stepRow+offsetRow-1)+offsetChileTable,columnTable,i,grayCellCenter)
                worksheet.write((count * stepRow + offsetRow)+offsetChileTable, columnTable, round(medians[i],6), normalCell)
                delta = medians[i]-medianOfAllSocket
                absDelta = abs(delta)

                sign=""
                if(delta>0):
                    sign = "+"
                
                #Đỏ
                if(absDelta > deltaSpecRed):
                    worksheet.write((count * stepRow + offsetRow+1)+offsetChileTable, columnTable,sign+ str(round(delta,6)), redCell)
                    worksheet.write(summaryRow,columnTable,sign + str(round(delta,3)),redCell)
                #Vàng
                elif(absDelta < deltaSpecRed and absDelta > deltaSpecYellow):
                    worksheet.write((count * stepRow + offsetRow+1)+offsetChileTable, columnTable, sign+ str(round(delta,6)), yellowCell)
                    worksheet.write(summaryRow,columnTable,sign+str(round(delta,3)),yellowCell)
                #Xanh
                elif(absDelta<deltaSpecYellow):
                    worksheet.write((count * stepRow + offsetRow+1)+offsetChileTable, columnTable, sign + str(round(delta,6)), greenCell)

                for j in range(0,1000):
                    worksheet.write(count*stepRow-4+offsetRow,j,"",grayCell)


            plt.close()
            count+=1
        machineCount+=1
    
    global numberOfMachines
    numberOfMachines = machineCount
    


def main():
    global resultText
    global sheetList
    global summaryRow
    seaborn.set(style='whitegrid')
    workbook = xlsxwriter.Workbook(thisPath+'\\report.xlsx')

    grayCell = workbook.add_format()
    grayCell.set_bg_color("#cccccc")
    grayCell.set_bold()
    #grayCell.set_border(1)
    grayCellCenter = grayCell
    grayCellCenter.set_align("Center")

    settingCount = 0
    rawSetting = pd.read_csv(args.setting)
    dataFrameColumn = ["log","category","redSpec","yellowSpec"]
    setting = pd.DataFrame(columns = dataFrameColumn)
    #print(rawSetting)
    if(os.path.exists(thisPath+"\\image") == False):
        os.makedirs(thisPath+"\\image")
    summarySheet = workbook.add_worksheet("Summary") 
    while True:
        settingCount += 1
        setting.log = rawSetting["log"+str(settingCount)]
        setting.category = rawSetting["category"+str(settingCount)]
        setting.redSpec = rawSetting["redSpec"+str(settingCount)]
        setting.yellowSpec = rawSetting["yellowSpec"+str(settingCount)]
        if(str(setting.log[0]) == "nan"):
            break
        #print(setting)
        
        if(os.path.exists(args.data+"/"+str(setting.log[0])+".CSV")):
            makeReportSheet(workbook,setting,summarySheet)
        else:
            stringtemp = "Log file {} does not exist".format(str(setting.log[0])+".CSV")
            print(stringtemp)
            resultText += stringtemp +"\r\n"


    for i in range(1,numberOfMachines+1):
        summarySheet.write(2,(i-1)*summaryColStep,"Socket",grayCell)
        for j in range(1,9):
            summarySheet.write(2,(i-1)*summaryColStep+j,str(j),grayCellCenter)
        for j in range(1,100):
            summarySheet.write(j,(i-1)*summaryColStep+9,"",grayCellCenter)
    #print("Close")
    workbook.close()  

    stringtemp = "================================================"+"\r\n"+"Making summary sheet"+"\r\n"
    print(stringtemp)
    resultText += stringtemp
    excel = win32.dynamic.Dispatch('Excel.Application')
    excel.Visible = True

    wb = excel.Workbooks.Open(thisPath+"\\report.xlsx")
    
    
    count = 0
    summaryNewRow = 3
    for sheet in sheetList:
        for j in range(0,numberOfMachines):
            startColumnLatter = xlsxwriter.utility.xl_col_to_name(6+15*j)
            stopColumnLatter = xlsxwriter.utility.xl_col_to_name(6+15*j+8)   
            startColumnLatterSum = xlsxwriter.utility.xl_col_to_name(0+10*j)

            wb.Worksheets(sheet).Range(startColumnLatter+"5:"+stopColumnLatter+str(5+summaryRowList[count])).Copy(Destination=wb.Worksheets("Summary").Range(startColumnLatterSum+str(summaryNewRow+1)))
        summaryNewRow += summaryRowList[count]+1 
        count += 1
    
    sheet = wb.Worksheets("Summary")

    for j in range(0,numberOfMachines):
        startColumnLatterSum = xlsxwriter.utility.xl_col_to_name(0+10*j)
        stopColumnLatterSum = xlsxwriter.utility.xl_col_to_name(8+10*j)
        rangeStr = startColumnLatterSum+"2:"+stopColumnLatterSum+str(summaryNewRow-1)
        sheet.Range(rangeStr).Borders(12).Color = 0
        sheet.Range(rangeStr).Borders(11).Color = 0
        sheet.Range(rangeStr).Borders(7).Color = 0
        sheet.Range(rangeStr).Borders(8).Color = 0
        sheet.Range(rangeStr).Borders(9).Color = 0
        sheet.Range(rangeStr).Borders(10).Color = 0
    wb.Save()

    

    
    

successfully = True
if __name__ == "__main__":
    
        main()
        resultText +"\r\n Report Success"
    
        resultText +="\n\r"
        resultText += "\r\nReport Failure"
        successfully = False
    
with open("report_log.txt", "w") as report_log:
    report_log.write(resultText)
if successfully:
    input("Report Success, press ENTER to exit...")
else:
    input("Report Failure, press ENTER to exit...")
