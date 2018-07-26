import os
import xlrd
import xlwt
from appJar import gui
import datetime

global exp_rows
global con_names

fileList=[]#保存excel文件列表，含路径

def collect_files(mydir):
    for root, dirs, files in os.walk(mydir):       
        for file in files:
            if(con_names in file):
                final_file = os.path.join(root, file)
                if os.path.splitext(final_file)[1] in {'.xls', '.xlsx','.csv'}:
                    fileList.append(final_file)
                    print(final_file)

dataList=[]#保存读取数据

def read_file(file, is_the_first_file):#读取数据
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    wk_row = sheet.nrows
    wk_col = sheet.ncols
    if(is_the_first_file):
        for i in range(wk_row):
            dataList.append(sheet.row_values(i))
    else:
        for i in range(exp_rows,wk_row):
            dataList.append(sheet.row_values(i))
    
def write_file():#写数据
    nowTime=datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    print("\nCreated Excel:Total"+nowTime+".xls\n")

    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('sheet 1')
    for i in range(len(dataList)):
        for j in range(len(dataList[i])):
            sheet.write(i,j, dataList[i][j])#第i行第j列写入内容
    wbk.save('Total'+nowTime+'.xls')                    

def press_select(button):
    if button=="button1":
        temp = app.directoryBox("Select a path")
        if temp:
            root_dir = temp
            app.setEntry("Source Folder Dir", root_dir)

def press_action(button):
    if button=="Start":        
        action()

    if button=="Clear":
        app.clearEntry("Source Folder Dir")

def action():
    global exp_rows
    global con_names
    exp_rows = int(app.getEntry("Remove Rows"))
    con_names = app.getEntry("Words In File Name")
    mydir= app.getEntry("Source Folder Dir")
    collect_files(mydir)
    if(len(fileList)<1):
        print("No File To Be Dealed")
    elif(len(fileList)==1):
        read_file(fileList[0],True)
    elif(len(fileList)>1):
        read_file(fileList[0],True)
        for i in range(1, len(fileList)):
            read_file(fileList[i],False)
    if(len(fileList)>0):
        write_file()    
    fileList.clear()
    dataList.clear()
    print('Done!!!!!!!!')

root_dir = os.getcwd()
app = gui("Merge Excel","450x140")
app.setResizable(False)

app.addLabelEntry("Remove Rows",0,0)
app.setEntry("Remove Rows",1)
app.addLabelEntry("Words In File Name",0,1)
app.setEntry("Words In File Name","Total")

app.addLabelEntry("Source Folder Dir",1,0)
app.setEntry("Source Folder Dir",root_dir)
app.addNamedButton("Select","button1",press_select,1,1)

app.addButtons(["Start", "Clear"], press_action)
app.go()
