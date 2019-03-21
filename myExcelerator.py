# -*- coding: gbk -*-
import sys, os
from pyExcelerator import *

reload(sys)
sys.setdefaultencoding("gbk")

class myExcelerator:
    __book=""

    def __init__(self, bookname):
        c=bookname.split('\\')[-1]
        print ">> 正在打开"+c+"，请等候..."
        try:
            self.__book=parse_xls(bookname)
            print ">> "+c+"已经成功打开..."
        except:
            print "!! 打开文件出错！"
            sys.exit(0)

    def getSheet(self, sheet):
        nsheet=0
        for k,v in self.__book:
            if k==sheet:
                return nsheet
            nsheet=nsheet+1
            if nsheet>len(self.__book):
                print "!! 没找到相应表格，程序退出！"
                sys.exit(0)
            
    def getCell(self, sheet, row, col):
        n=self.getSheet(sheet)
        try:
            return self.__book[n][1][(row,col)]
        except:
            return ""

    def getLen(self, sheet):
        n=self.getSheet(sheet)
        rowLen=0
        colLen=0
        for k,v in self.__book[n][1]:
            if k>rowLen:
                rowLen=k
            if v>colLen:
                colLen=v
        return (rowLen+1, colLen+1)

    def getLineR(self, sheet, row):
        n=self.getSheet(sheet)
        ncol=self.getLen(sheet)[1]
        line=[0 for i in xrange(ncol)]
        for i in xrange(ncol):
            try:
                line[i]=self.__book[n][1][(row,i)]
            except:
                line[i]=""
        return line

    def getLineC(self, sheet, col):
        n=self.getSheet(sheet)
        nrow=self.getLen(sheet)[0]
        line=[0 for i in xrange(nrow)]
        for i in xrange(nrow):
            try:
                line[i]=self.__book[n][1][(i,col)]
            except:
                line[i]=""
        return line
