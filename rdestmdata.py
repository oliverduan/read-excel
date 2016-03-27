#coding=utf-8
import os
import xlrd
import xlwt
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
reload(sys)

strdeffilename = "strdef.xls"
construct_cell = "construct_cell.xls"

xlwt.__VERSION__

from tempfile import TemporaryFile
filetyp = [".xls",".xlsx"]
tender_estfile_char = [u"投标估算表",u"投标项目估算表"]
module_estfile_char = u"模块"
#                                    
outputname = []
sheetdef = []
inputcell = []
trender_estfile_cell_list_a = [[u"项目编号",1,3,2],[u"项目名称",1,4,2],[u"投标项目执行规模",1,5,2],[u"软件预计投标金额",1,8,5],[u"中标服务费",1,9,5],
                        [u"预计采购成本",1,12,5],[u"预计税费",1,13,5],[u"预计现场费用驻场人员费用",1,14,5],[u"项目预计价差",1,15,5],
                        [u"项目预计差价率",1,16,5],[u"预计直接人工费",1,17,5],[u"预计差旅费",1,18,5],[u"预计其他实施费用",1,19,5],
                        [u"实施费用合计",1,20,5],[u"项目预计毛利",1,21,5],[u"项目预计毛利率",1,22,5],[u"预计综合运营费用率",1,23,5],
                        [u"项目预计纯利率",1,24,5]]
# 


trender_estfile_cell_list_b = [[u"项目编号",1,2,2],[u"项目名称",1,3,2],[u"投标项目资金规模",1,4,2],[u"软件预计投标金额",1,7,5],[u"中标服务费",1,8,5],
                        [u"预计采购成本",1,9,5],[u"预计税费",1,10,5],[u"预计现场费用驻场人员费用",1,11,5],[u"项目预计价差",1,12,5],
                        [u"项目预计差价率",1,13,5],[u"预计直接人工费",1,14,5],[u"预计差旅费",1,15,5],[u"预计其他实施费用",1,16,5],
                        [u"实施费用合计",1,17,5],[u"项目预计毛利",1,18,5],[u"项目预计毛利率",1,19,5],[u"预计综合运营费用率",1,20,5],
                        [u"项目预计纯利率",1,21,5]]

trender_estfile_cell_list_c = [[u"项目编号",1,2,2],[u"项目名称",1,3,2],[u"投标项目执行规模",1,4,2],[u"软件预计投标金额",1,7,5],[u"中标服务费",1,8,5],
                        [u"预计采购成本",1,9,5],[u"预计税费",1,10,5],[u"预计现场费用驻场人员费用",1,11,5],[u"项目预计毛利",1,12,5],
                        [u"项目预计毛利率",1,13,5],[u"预计直接人工费",1,2,3],[u"预计差旅费",1,2,3],[u"预计其他实施费用",1,2,3],
                        [u"预计实施费用",1,15,5],[u"项目预计利润",1,16,5],[u"项目预计利润率",1,17,5],[u"预计综合运营费用率",1,2,3],
                        [u"项目预计纯利率",1,2,3]]

trender_estfile_cell_list_d = [[u"项目编号",1,2,2],[u"项目名称",1,3,2],[u"投标项目执行规模",1,4,2],[u"软件预计投标金额",1,7,5],[u"中标服务费",1,8,5],
                        [u"预计采购成本",1,9,5],[u"预计税费",1,10,5],[u"预计现场费用驻场人员费用",1,11,5],[u"项目预计毛利",1,12,5],
                        [u"项目预计毛利率",1,13,5],[u"预计直接人工费",1,2,3],[u"预计差旅费",1,2,3],[u"预计其他实施费用",1,2,3],
                        [u"预计实施费用",1,15,5],[u"项目预计利润",1,16,5],[u"项目预计利润率",1,17,5],[u"预计综合运营费用率",1,2,3],
                        [u"项目预计纯利率",1,2,3]]


trender_estfile_cell_list = trender_estfile_cell_list_a

module_estfile_cell_list = [[u"工作量（人月）",u"项目按模块估算表",u"小计",u"工作量（人月）"],
                            [u"成本估算-人工（万元）",u"项目按模块估算表",u"小计",u"成本估算（万元）"],
                            [u"差旅费",u"项目按模块估算表",u"差旅费",u"成本估算（万元）"],
                            [u"合计工程实施成本",u"项目按模块估算表",u"合  计",u"成本估算（万元）"]]

def construct_cell_write(sh,row,cell_def_list,def_type):
    for cell_def in cell_def_list:
        for col in range(len(cell_def)):
            sh.write(row,col,cell_def[col])
        col = col + 1
        sh.write(row,col,def_type)
        row = row + 1
    return row,sh

def construct_cell_def(deffilename):
    bk = xlwt.Workbook(encoding='utf-8')
    sheet1 = bk.add_sheet(u"inputcell")
    row = 0
    new_row,sh = construct_cell_write(sheet1, row, trender_estfile_cell_list_a,u"a")
    row = new_row
    sheet1 = sh
    new_row,sh = construct_cell_write(sheet1, row, trender_estfile_cell_list_b,u"b")
    row = new_row
    sheet1 = sh
    new_row,sh = construct_cell_write(sheet1, row, trender_estfile_cell_list_c,u"c")
    row = new_row
    sheet1 = sh
    new_row,sh = construct_cell_write(sheet1, row, trender_estfile_cell_list_d,u"d")
    
    bk.save(deffilename)
    bk.save(TemporaryFile())
"""
    for cell_def in trender_estfile_cell_list_a:
        for col in range(len(cell_def)):
            sheet1.write(row,col,cell_def[col])
        col = col + 1
        sheet1.write(row,col,u"a")
        row = row + 1
"""    
    
def get_struct_def(deffilename):
    wk=xlrd.open_workbook(deffilename)
    outputname_sh = wk.sheet_by_name("outputname")
    for row in range(outputname_sh.nrows - 1):
        outputname.append(outputname_sh.cell(row+1,0).value)
    sheetdef_sh = wk.sheet_by_name(u"sheetdef")
    for row in range(sheetdef_sh.nrows -1 ):
        sheetdef_cell = []
        for col in range(sheetdef_sh.ncols):
            sheetdef_cell.append(sheetdef_sh.cell(row,col))
        sheetdef.append(sheetdef_cell)
            
    
    
def readestdata(rootdir,datafilename):
    subdirlist = os.listdir(rootdir)
    book = xlwt.Workbook(encoding='utf-8')
    sheet1 = book.add_sheet(u"项目与文件完整性")
    sheet2 = book.add_sheet(u"估算数据")
    project_num  = 0
    cell_col = 0
    sheet1.write(0,cell_col,u"序号")
    cell_col = cell_col + 1
    sheet1.write(0,cell_col,u"项目目录")
    cell_col = cell_col + 1
    
    sheet1.write(0,cell_col,u"投标估算文件")
    cell_col = cell_col + 1
    trender_estfile_cell_list = trender_estfile_cell_list_a
    for tender_estfile_cell in trender_estfile_cell_list:
        sheet1.write(0,cell_col,tender_estfile_cell[0])
        cell_col = cell_col + 1
        
    sheet1.write(0,cell_col,u"模块估算文件")
    cell_col = cell_col + 1
    for module_estfile_cell in module_estfile_cell_list:
        sheet1.write(0,cell_col,module_estfile_cell[0])
        cell_col = cell_col + 1

    for subdir in subdirlist:
        if os.path.isdir(os.path.join(rootdir,subdir)):
            datadir = os.path.join(rootdir,subdir)
            project_num = project_num + 1
            filelist = os.listdir(datadir)
            tender_estfile_name = []
            module_estfile_name = []  
            for filename  in filelist:
            # find all of tender estifile and find which one is newer.
                if os.path.isfile(os.path.join(datadir,filename)):
#                    print tender_estfile_char,filename,filetyp
                    if  ((tender_estfile_char[0] in os.path.splitext(filename)[0]) or (tender_estfile_char[1] in os.path.splitext(filename)[0])) and (os.path.splitext(filename)[1] in filetyp):
                        if u"新" in os.path.splitext(filename)[0]:
                            tender_estfile_name.append([os.path.join(datadir,filename),u"新"])
                        else:
                            tender_estfile_name.append([os.path.join(datadir,filename),u"原"])                           
                    if  (module_estfile_char in os.path.splitext(filename)[0]) and (os.path.splitext(filename)[1] in filetyp):
                        if u"新" in os.path.splitext(filename)[0]:
                            module_estfile_name.append([ os.path.join(datadir,filename),u"新"])
                        else:
                            module_estfile_name.append([ os.path.join(datadir,filename),u"旧"])                            
            sheet1.write(project_num,0,project_num)
            sheet1.write(project_num,1,datadir)

            cell_col = 2
            if len(tender_estfile_name) == 0:
                print u"项目： ",datadir,u"投标估算文件：不存在"
                sheet1.write(project_num,cell_col,u"无")
                cell_col = cell_col + 1
                cell_col = cell_col + len(trender_estfile_cell_list)
            elif len(tender_estfile_name) == 1:
                print u"项目：",datadir,u"投标估算文件(只有一个)是 ：",tender_estfile_name[0][0]
                sheet1.write(project_num,cell_col,tender_estfile_name[0][0])
                cell_col = cell_col + 1 
                wb = xlrd.open_workbook(tender_estfile_name[0][0])
                
                if wb.sheets()[0].cell(1,0).value == u"项目编号":
                    if wb.sheets()[0].cell(12,0).value == u"项目预计毛利率":
                        trender_estfile_cell_list = trender_estfile_cell_list_c
                    else:
                        trender_estfile_cell_list = trender_estfile_cell_list_b
                else:
                    trender_estfile_cell_list = trender_estfile_cell_list_a                                
                for tender_estfile_cell in trender_estfile_cell_list:
                    content = wb.sheets()[tender_estfile_cell[1] - 1].cell(tender_estfile_cell[2] -1,tender_estfile_cell[3] -1).value
                    print tender_estfile_cell[0]," = ",content
                    sheet1.write(project_num,cell_col,content)
                    cell_col = cell_col + 1
            else:
                #多个文件,找到有“新”，否则取第一个
                num_tender_estfile = 0
                tender_estfile_name_new = ""
                for tender_estfile_name_element in tender_estfile_name:
                    if u"新" in tender_estfile_name_element[1]:
                        tender_estfile_name_new = tender_estfile_name_element[0]
                     #   break
                if len(tender_estfile_name_new) == 0:
                    tender_estfile_name_new = tender_estfile_name[0][0]
                sheet1.write(project_num,cell_col,tender_estfile_name_new)
                cell_col = cell_col + 1
                wb = xlrd.open_workbook(tender_estfile_name_new)
                if wb.sheets()[0].cell(1,0).value == u"项目编号":
                    if wb.sheets()[0].cell(12,0).value == u"项目预计毛利率":
                        trender_estfile_cell_list = trender_estfile_cell_list_c
                    else:
                        trender_estfile_cell_list = trender_estfile_cell_list_b
                else:
                    trender_estfile_cell_list = trender_estfile_cell_list_a                                
                
                for tender_estfile_cell in trender_estfile_cell_list:
                    content = wb.sheets()[tender_estfile_cell[1] - 1].cell(tender_estfile_cell[2] -1,tender_estfile_cell[3] -1).value
                    print tender_estfile_cell[0]," = ",content
                    sheet1.write(project_num,cell_col,content)
                    cell_col = cell_col + 1                    
                    
 
            if len(module_estfile_name) > 0:
                print u"项目：",datadir,u"模块估算文件是 ：",module_estfile_name[0][0]
                sheet1.write(project_num,cell_col,module_estfile_name[0][0])
                cell_col = cell_col + 1
                
                wb = xlrd.open_workbook(module_estfile_name[0][0])
                for module_estfile_cell in module_estfile_cell_list:
                    try:
                        s = wb.sheet_by_name(module_estfile_cell[1])
                    except xlrd.XLRDError:
                        s = wb.sheet_by_name(u"项目按模块估算记录表")
                    
                    data_cell_row = 0
                    data_cell_col = 0
                    for row_index in range(s.nrows):
                        if (s.cell(row_index,0).value == module_estfile_cell[2]) or (s.cell(row_index,1).value == module_estfile_cell[2]):
                            data_cell_row = row_index
                            break
                    for col_index in range(s.ncols):
                        if s.cell(2,col_index).value == module_estfile_cell[3]:
                            data_cell_col = col_index
                            break
                    content = s.cell(data_cell_row,data_cell_col).value                   
#                    content = wb.sheets()[module_estfile_cell[1] -1].cell(module_estfile_cell[2] -1,module_estfile_cell[3]-1)
                    print module_estfile_cell[0], " = ",content
                    sheet1.write(project_num,cell_col,content)
                    cell_col = cell_col +1
            
            else:
                print u"项目：",datadir,u"模块估算文件 ：不存在"
                sheet1.write(project_num,cell_col,u"无")
                cell_col + 1
                
    book.save(datafilename)
    book.save(TemporaryFile())

def main():
    os.chdir(u"D:\软件质量保障中心\合同评审\数据导出3\数据导出3")
    rootdir = os.getcwd()
    rootdir = unicode(rootdir,"gbk")
    
    strdeffilename1= os.path.join(rootdir,strdeffilename)
    construct_cell_file = os.path.join(rootdir,construct_cell)
    construct_cell_def(construct_cell_file)
    
    get_struct_def(strdeffilename1)
    
    datafilename = os.path.join(rootdir,u"data.xls")
#   rootdir1 = u"D:\\软件质量保障中心\合同评审\\投标项目估算审计"
#  datafilename1 = u"D:\\软件质量保障中心\合同评审\\投标项目估算审计\\data.xls"

    print rootdir,datafilename
# print rootdir1,datafilename1

 #   if (rootdir1 == rootdir) and (datafilename == datafilename1):
 #   readestdata(rootdir1,datafilename1)
  #      readestdata(rootdir,datafilename)
        
#    readestdata(rootdir,datafilename)
#    readestdata(rootdir=u"D:\\软件质量保障中心\合同评审\\投标项目估算审计",datafilename=u"D:\\软件质量保障中心\合同评审\\投标项目估算审计\\data.xls")

    
if __name__ == "__main__":
    main()