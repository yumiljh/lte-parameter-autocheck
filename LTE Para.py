# -*- coding: gbk -*-
import sys, os
from pyExcelerator import *
from myExcelerator import *
from inspection import *

reload(sys)
sys.setdefaultencoding("gbk")


writeBook=open("LTE���������.csv","w")
content='"����Ӣ����","����������","MOI","SubNetwork","ENBID","��Ԫ����","��������ֵ","��������ֵ","��Ҫ��"\n'
writeBook.write(content)

rulename=os.getcwd()+"\\"+"���淶.xls"
cwd=os.getcwd()+"\\ѹ���ļ�"
cwdfilelist=os.listdir(cwd)

for x in cwdfilelist:
    if x.split('.')[-1]=='xls':
        bookname=cwd+"\\"+x

        readBook=myExcelerator(bookname)
        ruleBook=myExcelerator(rulename)

        e=inspection(readBook,writeBook,ruleBook)
        e.run()

writeBook.close()
raw_input(">> ���������ϣ��밴�������������ӭʹ�ã�")
