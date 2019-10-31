# encoding: utf-8
import xlrd
import json
import unicodecsv as ucsv
import configparser
import os
from itertools import combinations
from collections import Counter;

#0 配置参数的加载
WRITE_FILE_COLUMNS_NAME = ["货品组合","次数"]
KEY_COL = 2
DATA_COL = 3
GROUP_COUNT = 2
INPUT_FORDER = 'in'
OUTPUT_FORDER = 'out'

def loadConfig():
    global KEY_COL
    global DATA_COL
    global GROUP_COUNT
    global INPUT_FORDER
    global OUTPUT_FORDER
    cf = configparser.ConfigParser()
    cf.read("conf.ini")
    KEY_COL = int(cf.get("reckon", "key_col"))
    DATA_COL = int(cf.get("reckon", "data_col"))
    GROUP_COUNT = int(cf.get("reckon", "group_count"))
    INPUT_FORDER = cf.get("reckon", "input_forder")
    OUTPUT_FORDER = cf.get("reckon", "output_forder")

def listFiles(rootDir,fileList):
    f_list = os.listdir(rootDir)
    for i in f_list:
      if os.path.splitext(i)[1] == '.xlsx' or os.path.splitext(i)[1] == '.xls':
        fileList.append(i)
    #print json.dumps(FILE_LIST).decode("unicode-escape")

def handleFileList(fileList):
    for file in fileList:
        name = os.path.basename(file);
        infile = os.path.join(INPUT_FORDER,name)
        outfile = os.path.join(OUTPUT_FORDER,os.path.splitext(name)[0]+'.csv')
        handleFile(infile,outfile)

def handleFile(infile,outfile):
    # 读取的初始数据
    row_list=[]
    # 处理后的结果数据
    returnList=[]
    print "="*20
    print "1、开始读取文件%s" % (infile.encode('gbk'))
    readExcel(infile,row_list)
    print "2、读取文件%s完毕" % (infile.encode('gbk'))
    #print json.dumps(row_list).decode("unicode-escape")
    print "3、开始处理数据..."
    returnList = handleList(row_list)
    print "4、处理数据完毕..."
    #print json.dumps(returnList).decode("unicode-escape")
    print "5、开始写入文件%s..." % (outfile.encode('gbk'))
    writeCsv(outfile,returnList)
    print "6、写入文件%s完毕..." % (outfile.encode('gbk'))

def readExcel(infile,row_list):
    wb = xlrd.open_workbook(filename=infile)
    sh = wb.sheet_by_index(0)
    nrows = sh.nrows
    ncols = sh.ncols
    print "EXCEL[%-15s]:行数 %d, 列数 %d" % (infile.encode('gbk'),nrows,ncols)
    
    #row_list = []
    for i in range(1,nrows):
      row_data = sh.row_values(i)
      row_list.append(row_data)
      #print json.dumps(row_data).decode("unicode-escape")
    row_list.sort(key=takeKey)
    #print json.dumps(row_list).decode("unicode-escape")

def takeKey(elem):
    return elem[KEY_COL]
def handleList(row_list):
     
    combined_list =[]
    index = -1;
    current_key = '';
    # 合并相同key的另一字段
    for i in range(0,len(row_list)):
      if current_key == row_list[i][KEY_COL]:
        combined_list[index][DATA_COL].append(row_list[i][DATA_COL])
      else:
        current_key = row_list[i][KEY_COL]
        combined_list.append(row_list[i])
        index += 1
        combined_list[index][DATA_COL] = [row_list[i][DATA_COL]]
    #print "============合并相同key的另一字段==========="
    #print json.dumps(combined_list).decode("unicode-escape")
    # 去重 排序 相同key的另一字段
    for i in range(0,len(combined_list)):
        duplicatesList = list(set(combined_list[i][DATA_COL]))
        duplicatesList.sort()
        combined_list[i][DATA_COL] = duplicatesList
    #print "============去重 排序 相同key的另一字段==========="
    #print json.dumps(combined_list).decode("unicode-escape")

    # 分组排列组合 （对相同key的另一字段）
    for i in range(0,len(combined_list)):
        combinations_list = list(combinations(combined_list[i][DATA_COL], GROUP_COUNT));
        combined_list[i][DATA_COL] = combinations_list
    #print "============分组排列组合==========="
    #print json.dumps(combined_list, indent=2).decode("unicode-escape")
    
    # 关键数据扁平化数组
    resultList = []
    for i in range(0,len(combined_list)):
        for j in range(0,len(combined_list[i][DATA_COL])):
            keys = ','.join(combined_list[i][DATA_COL][j])
            resultList.append(keys)
    resultList.sort()
    #print "============关键数据扁平化数组==========="
    #print json.dumps(resultList, indent=2).decode("unicode-escape")
    
    resultDict = dict(Counter(resultList))
    #print json.dumps(resultDict, indent=2).decode("unicode-escape")
    
    returnList = sorted(resultDict.items(), key=lambda d: d[1], reverse=True)
    #print json.dumps(returnList, indent=2).decode("unicode-escape")
    
    return returnList

def writeCsv(outfile,list):
    dirname = os.path.dirname(outfile)
    if not os.path.exists(dirname):
       os.makedirs(dirname)
    csvFile = open(outfile, "wb+")
    writer = ucsv.writer(csvFile,encoding = 'gbk')
    #先写入columns_name
    writer.writerow(WRITE_FILE_COLUMNS_NAME)
    #写入多行用writerows
    writer.writerows(list)
    csvFile.close()


# 获取要处理文件列表
FILE_LIST=[]

loadConfig()
listFiles(INPUT_FORDER, FILE_LIST)
handleFileList(FILE_LIST)




