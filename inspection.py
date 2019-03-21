# -*- coding: gbk -*-
import sys, os
from pyExcelerator import *
from myExcelerator import *

reload(sys)
sys.setdefaultencoding("gbk")
    
class inspection:
    __rs=""
    __ws=""
    __rule=""
    __meid=""
    __userLabel=""

    def __init__(self, rs, ws, rule):
        self.__rs=rs
        self.__ws=ws
        self.__rule=rule
        self.__meid=rs.getLineC('ENBFunctionTDD',2)
        self.__userLabel=rs.getLineC('ENBFunctionTDD',3)

    def writeContent(self,content):
        c=""
        for x in content:
            c=c+"\""+str(x)+"\","
        c=c+"\n"
        self.__ws.write(c)

    def getCol(self, sheet, searchStr):
        header=self.__rs.getLineR(sheet,0)
        for x in header:
            if x==searchStr:
                return header.index(x)
        return 0

    def getLabel(self,searchMeid):
        for x in self.__meid:
            if x==searchMeid:
                return self.__userLabel[self.__meid.index(x)]
        return ""

    def getCond(self,a,b,c,searcha,searchb):
        for i in xrange(5,len(a)):
            if float(a[i])==float(searcha) and float(b[i])==float(searchb):
                return c[i]

    def check1(self,sheet,eng,ch,smallval,bigval,level):
        col=self.getCol(sheet,eng)
        line=self.__rs.getLineC(sheet,col)
        
        for i in xrange(5,len(line)):
            temp=""
            try:
                temp=float(line[i])
            except:
                temp=line[i]
            
            if temp<smallval or temp>bigval:
                moi=self.__rs.getCell(sheet,i,0)
                sub=self.__rs.getCell(sheet,i,1)
                enb=self.__rs.getCell(sheet,i,2)
                label=self.getLabel(enb)
                correctVal=""
                try:
                    correctVal='['+str(int(smallval))+','+str(int(bigval))+']'
                except:
                    correctVal=str(smallval)
                content=[eng,ch,moi,sub,enb,label,line[i],correctVal,level]
                self.writeContent(content)

    def check2(self,sheet,eng,ch,smallval,bigval,level,cond_sheet,cond_eng,cond_small,cond_big):#type2>type1
        col=self.getCol(sheet,eng)
        line=self.__rs.getLineC(sheet,col)
        cond_col=self.getCol(cond_sheet,cond_eng)#
        cond_line=self.__rs.getLineC(cond_sheet,cond_col)#
        
        for i in xrange(5,len(line)):
            temp=""
            cond=""#
            try:
                temp=float(line[i])
                cond=float(cond_line[i])#
            except:
                temp=line[i]
                cond=cond_line[i]#
            
            if (cond>=cond_small and cond<=cond_big) and (temp<smallval or temp>bigval):#
                moi=self.__rs.getCell(sheet,i,0)
                sub=self.__rs.getCell(sheet,i,1)
                enb=self.__rs.getCell(sheet,i,2)
                label=self.getLabel(enb)
                correctVal=""
                try:
                    correctVal='['+str(int(smallval))+','+str(int(bigval))+']'
                except:
                    correctVal=str(smallval)
                content=[eng,ch,moi,sub,enb,label,line[i],correctVal,level]
                self.writeContent(content)

    def check3(self,sheet,eng,ch,smallval,bigval,level,cond_sheet,cond_eng,cond_small,cond_big):#type3>type2
        col=self.getCol(sheet,eng)
        meid=self.__rs.getLineC(sheet,2)#
        cellid=self.__rs.getLineC(sheet,3)#
        line=self.__rs.getLineC(sheet,col)
        cond_col=self.getCol(cond_sheet,cond_eng)
        cond_meid=self.__rs.getLineC(cond_sheet,2)#
        cond_cellid=self.__rs.getLineC(cond_sheet,3)#
        cond_line=self.__rs.getLineC(cond_sheet,cond_col)
        
        for i in xrange(5,len(line)):
            temp=""
            try:
                temp=float(line[i])
            except:
                temp=line[i]
            enb=meid[i]#
            cell=cellid[i]#
            cond=float(self.getCond(cond_meid,cond_cellid,cond_line,enb,cell))#
            
            if (cond>=cond_small and cond<=cond_big) and (temp<smallval or temp>bigval):
                moi=self.__rs.getCell(sheet,i,0)
                sub=self.__rs.getCell(sheet,i,1)
                #enb=self.__rs.getCell(sheet,i,2)
                label=self.getLabel(enb)
                correctVal=""
                try:
                    correctVal='['+str(int(smallval))+','+str(int(bigval))+']'
                except:
                    correctVal=str(smallval)
                content=[eng,ch,moi,sub,enb,label,line[i],correctVal,level]
                self.writeContent(content)

    def check4(self,sheet,eng,ch,valList,level,cond_sheet,cond_eng,cond_small,cond_big):#type4>type2
        col=self.getCol(sheet,eng)
        line=self.__rs.getLineC(sheet,col)
        cond_col=self.getCol(cond_sheet,cond_eng)
        cond_line=self.__rs.getLineC(cond_sheet,cond_col)

        val=valList.split(',')#
        
        for i in xrange(5,len(line)):
            temp=""
            cond=""
            try:
                temp=float(line[i])
                cond=float(cond_line[i])
            except:
                temp=line[i]
                cond=cond_line[i]

            flag=0#<mod>#0代表不通过，1代表通过
            if cond>=cond_small and cond<=cond_big:
                for x in val:
                    if temp==float(x):
                        flag=1
                if flag==0:
                    moi=self.__rs.getCell(sheet,i,0)
                    sub=self.__rs.getCell(sheet,i,1)
                    enb=self.__rs.getCell(sheet,i,2)
                    label=self.getLabel(enb)
                    content=[eng,ch,moi,sub,enb,label,line[i],valList,level]
                    self.writeContent(content)#</mod>

    def check5(self,sheet,eng,ch,smallval,bigval,level,isnull):
        col=self.getCol(sheet,eng)
        line=self.__rs.getLineC(sheet,col)
        
        for i in xrange(5,len(line)):
            flag=1#0代表不通过，1代表通过
            if isnull=='N' and line[i]=="":
                flag=0
            elif isnull=='Y' and line[i]=="":
                continue
            else:
                temp=line[i].split(';')
                for x in temp:
                    if float(x)<smallval or float(x)>bigval:
                        flag=0
            if flag==0:
                moi=self.__rs.getCell(sheet,i,0)
                sub=self.__rs.getCell(sheet,i,1)
                enb=self.__rs.getCell(sheet,i,2)
                label=self.getLabel(enb)
                correctVal=""
                try:
                    correctVal='['+str(int(smallval))+','+str(int(bigval))+']'
                except:
                    correctVal=str(smallval)
                content=[eng,ch,moi,sub,enb,label,line[i],correctVal,level]
                self.writeContent(content)

    def check6(self,sheet,eng,ch,valList,level,isnull):#check6>check5
        col=self.getCol(sheet,eng)
        line=self.__rs.getLineC(sheet,col)

        val=valList.split(',')#
        
        for i in xrange(5,len(line)):
            flag=1#0代表不通过，1代表通过
            if isnull=='N' and line[i]=="":
                flag=0
            elif isnull=='Y' and line[i]=="":
                continue
            else:
                temp=line[i].split(';')
                for x in temp:
                    flag2=0
                    for y in val:
                        if float(x)==float(y):
                            flag2=1
                    flag=flag & flag2
            if flag==0:
                moi=self.__rs.getCell(sheet,i,0)
                sub=self.__rs.getCell(sheet,i,1)
                enb=self.__rs.getCell(sheet,i,2)
                label=self.getLabel(enb)
                content=[eng,ch,moi,sub,enb,label,line[i],valList,level]
                self.writeContent(content)

    def check7(self,sheet,eng,ch,val,level,cond_sheet,cond_eng,cond1,cond2):
        col=self.getCol(sheet,eng)
        line=self.__rs.getLineC(sheet,col)
        cond_col=self.getCol(cond_sheet,cond_eng)
        cond_line=self.__rs.getLineC(cond_sheet,cond_col)

        meid=self.__rs.getLineC(sheet,2)#
        
        for i in xrange(5,len(line)):
            temp=""
            cond=""
            try:
                temp=float(line[i])
                cond=float(cond_line[i])
            except:
                temp=line[i]
                cond=cond_line[i]

            enb=self.__rs.getCell(sheet,i,2)
            
            if cond==cond1:
                temp2=float(self.getCond(meid,cond_line,line,enb,cond2))
                if temp-temp2<val:
                    moi=self.__rs.getCell(sheet,i,0)
                    sub=self.__rs.getCell(sheet,i,1)
                    label=self.getLabel(enb)
                    content=[eng,ch,moi,sub,enb,label,line[i],'A1-A2>3',level]
                    self.writeContent(content)

    def run(self):
        nrow=self.__rule.getLen('RULE')[0]
        lastParaClass=""
        for i in xrange(1,nrow):
            r=self.__rule.getLineR('RULE',i)
            if lastParaClass!=r[1]:
                print ">> 正在检查"+str(r[1])+"类参数..."
                lastParaClass=r[1]
            if r[0]==1:
                self.check1(r[2], r[3], r[4], r[5], r[6], r[7])
            elif r[0]==2:
                self.check2(r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11])
            elif r[0]==3:
                self.check3(r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11])
            elif r[0]==4:
                self.check4(r[2], r[3], r[4], r[5], r[7], r[8], r[9], r[10], r[11])
            elif r[0]==5:
                self.check5(r[2], r[3], r[4], r[5], r[6], r[7], r[12])
            elif r[0]==6:
                self.check6(r[2], r[3], r[4], r[5], r[7], r[12])
            elif r[0]==7:
                self.check7(r[2], r[3], r[4], r[5], r[7], r[8], r[9], r[10], r[11])
            else:
                print "!! 检查规范输入有误，出错行数："+str(i+1)

