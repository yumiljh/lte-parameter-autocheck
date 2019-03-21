# -*- coding: gbk -*-
import sys, os
from pyExcelerator import *
from myExcelerator import *
from inspection import *

reload(sys)
sys.setdefaultencoding("gbk")


writeBook=open("LTE参数检查结果.csv","w")
content='"参数英文名","参数中文名","MOI","SubNetwork","ENBID","网元名称","现网设置值","合理设置值","重要性"\n'
writeBook.write(content)

rulename=os.getcwd()+"\\"+"检查规范.xls"
cwd=os.getcwd()+"\\压缩文件"
cwdfilelist=os.listdir(cwd)

for x in cwdfilelist:
    if x.split('.')[-1]=='xls':
        bookname=cwd+"\\"+x

        readBook=myExcelerator(bookname)
        ruleBook=myExcelerator(rulename)

        e=inspection(readBook,writeBook,ruleBook)
        e.run()

writeBook.close()
raw_input(">> 参数检查完毕，请按任意键结束，欢迎使用！")
