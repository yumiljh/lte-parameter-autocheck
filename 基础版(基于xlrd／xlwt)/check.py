# -*- coding: gbk -*-
import xlrd
import sys
import os

reload(sys)
sys.setdefaultencoding("gbk")

#找出rs里搜索字段searchStr的列号
def getCol(rs,searchStr):
    colNum=0
    for x in rs.row_values(0):
        if x==searchStr:
            return colNum
        colNum=colNum+1
    return 0

#找出searchMeid的网元名称
def getLabel(meidList,userLabel,searchMeid):
    for x in meidList:
        if x==searchMeid:
            return userLabel[meidList.index(x)]
    return 0

#找出rs里searchStr字段的MEID号为searchMeid的单元格的值
def getCompanion(rs,searchStr,searchMeid):
    col_sstr=getCol(rs,searchStr)
    enbCol=getCol(rs,"ENBFunctionTDD")
    rslen=rs.nrows
    for i in range(5,rslen):
        if rs.cell_value(i,enbCol)==searchMeid:
            return rs.cell_value(i,col_sstr)

#找出rs里searchStr字段的MEID号为searchMeid、CELLID号位searchCid的单元格的值
def getCompanion2(rs,searchStr,searchMeid,searchCid):
    col_sstr=getCol(rs,searchStr)
    enbCol=getCol(rs,"ENBFunctionTDD")
    cellCol=getCol(rs,"EUtranCellTDD")
    rslen=rs.nrows
    for i in range(5,rslen):
        if rs.cell_value(i,enbCol)==searchMeid and rs.cell_value(i,cellCol)==searchCid:
            return rs.cell_value(i,col_sstr)

#找出rs里searchStr字段的MEID号为searchMeid、meas号位searchMeas的单元格的值
def getCompanion3(rs,searchStr,searchMeid,searchMeas):
    col_sstr=getCol(rs,searchStr)
    enbCol=getCol(rs,"ENBFunctionTDD")
    measCol=getCol(rs,"measCfgIdx")
    rslen=rs.nrows
    for i in range(5,rslen):
        if rs.cell_value(i,enbCol)==searchMeid and int(rs.cell_value(i,measCol))==searchMeas:
            return rs.cell_value(i,col_sstr)

#创建rs里listHead列的列表
def createList(rs,listHead):
    return [x for x in rs.col_values(getCol(rs,listHead))]

#写入一段content到ws
def writeContent(ws,content,cur):
    tempStr=""
    for x in content:
        tempStr=tempStr+"\""+str(x)+"\","
    tempStr=tempStr+"\n"
    ws.write(tempStr)
    cur[0]+=1
    pass

#适用于单值型
def pmcheck1(rs,ws,eng,ch,enbid,userLabel,corectValue,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
        if temp!=corectValue:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),corectValue,level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于单值型
def pmcheck1a(rs,ws,eng,ch,enbid,userLabel,corectValue,corectValue2,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
        if temp!=corectValue and temp!=corectValue2:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(corectValue)+","+str(corectValue2),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于带单个条件的单值型
def pmcheck2(rs,ws,eng,ch,enbid,userLabel,corectValue,level,condName,condValue,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    condCol=getCol(rs,condName)
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
        if float(rs.cell_value(i,condCol))==condValue and temp!=corectValue:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),corectValue,level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于带单个条件的单值型
def pmcheck2a(rs,ws,eng,ch,enbid,userLabel,corectValue,level,condName,condValue,condValue2,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    condCol=getCol(rs,condName)
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
	temp_cond=float(rs.cell_value(i,condCol))
        if (temp_cond>=condValue and temp_cond<=condValue2)and temp!=corectValue:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),corectValue,level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于带单个条件的单值型
def pmcheck2b(rs,ws,eng,ch,enbid,userLabel,corectValue,corectValue2,level,condName,condValue,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    condCol=getCol(rs,condName)
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
	temp_cond=float(rs.cell_value(i,condCol))
        if temp_cond==condValue and (temp!=corectValue and temp!=corectValue2):
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(corectValue)+","+str(corectValue2),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于多值型
def pmcheck3(rs,ws,eng,ch,enbid,userLabel,smallVal,bigVal,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
        if temp<smallVal or temp>bigVal:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(smallVal)+"~"+str(bigVal),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于带单个条件的多值型
def pmcheck4(rs,ws,eng,ch,enbid,userLabel,smallVal,bigVal,level,condName,condValue,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    condCol=getCol(rs,condName)
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
        if float(rs.cell_value(i,condCol))==condValue and (temp<smallVal or temp>bigVal):
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(smallVal)+"~"+str(bigVal),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于带单个条件的多值型
def pmcheck4a(rs,ws,eng,ch,enbid,userLabel,smallVal,bigVal,level,rs2,condName,condValue,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
        temp=0
        try:
	    temp=float(rs.cell_value(i,col))
        except:
            temp=rs.cell_value(i,col)
        if float(getCompanion(rs2,condName,rs.cell_value(i,enbCol)))==condValue and (temp<smallVal or temp>bigVal):
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(smallVal)+"~"+str(bigVal),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于比较型
def pmcheck5(rs,ws,eng,ch,enbid,userLabel,checkVal,level,condName,condValue,condValue2,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    condCol=getCol(rs,condName)
    for i in range(5,rslen):
	temp=float(rs.cell_value(i,col))
	if float(rs.cell_value(i,condCol))==condValue:
	    temp2=float(getCompanion3(rs,eng,rs.cell_value(i,enbCol),condValue2))
	    if temp-temp2<checkVal:
		subnet=rs.cell_value(i,subCol)
		curMeid=rs.cell_value(i,enbCol)
		curMoi=rs.cell_value(i,moiCol)
		curLabel=getLabel(enbid,userLabel,curMeid)
		content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),"A1-A2>=3",level]
		writeContent(ws,content,cur)
		del content
    pass

#适用于循环型单值型
def pmchecklp(rs,ws,eng,ch,enbid,userLabel,corectValue,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
	if rs.cell_value(i,col)=="":
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),corectValue,level]
            writeContent(ws,content,cur)
            del content
            continue
	temp=rs.cell_value(i,col).split(';')
	sumTemp=0
	for j in range(len(temp)):
	    sumTemp=sumTemp+int(temp[j])
        if sumTemp!=corectValue*len(temp):
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),corectValue,level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于循环型单值型，允许空值
def pmchecklp2(rs,ws,eng,ch,enbid,userLabel,corectValue,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
	if rs.cell_value(i,col)=="":
            continue
	temp=rs.cell_value(i,col).split(';')
	sumTemp=0
	for j in range(len(temp)):
	    sumTemp=sumTemp+int(temp[j])
        if sumTemp!=corectValue*len(temp):
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),corectValue,level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于循环型多值型
def pmcheckmul(rs,ws,eng,ch,enbid,userLabel,corectValue,corectValue2,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
	if rs.cell_value(i,col)=="":
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(corectValue)+","+str(corectValue2),level]
            writeContent(ws,content,cur)
            del content
            continue
	temp=rs.cell_value(i,col).split(';')
        flag=0
	for j in range(len(temp)):
	    if int(temp[j])!=corectValue and int(temp[j])!=corectValue2:
                flag=1
        if flag==1:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(corectValue)+","+str(corectValue2),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于循环型多值型
def pmcheckmul3(rs,ws,eng,ch,enbid,userLabel,corectValue,corectValue2,corectValue3,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
	if rs.cell_value(i,col)=="":
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(corectValue)+","+str(corectValue2)+","+str(corectValue3),level]
            writeContent(ws,content,cur)
            del content
            continue
	temp=rs.cell_value(i,col).split(';')
        flag=0
	for j in range(len(temp)):
	    if int(temp[j])!=corectValue and int(temp[j])!=corectValue2 and int(temp[j])!=corectValue3:
                flag=1
        if flag==1:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),str(corectValue)+","+str(corectValue2)+","+str(corectValue3),level]
            writeContent(ws,content,cur)
            del content
    pass

#适用于序列单值型
def pmcheckpr(rs,ws,eng,ch,enbid,userLabel,corectValue,showValue,level,cur):
    rslen=rs.nrows
    col=getCol(rs,eng)
    subCol=getCol(rs,"SubNetwork")
    enbCol=getCol(rs,"ENBFunctionTDD")
    moiCol=getCol(rs,"MOI")
    for i in range(5,rslen):
	temp=rs.cell_value(i,col).split(';')
        if int(temp[0])!=corectValue:
            subnet=rs.cell_value(i,subCol)
            curMeid=rs.cell_value(i,enbCol)
            curMoi=rs.cell_value(i,moiCol)
            curLabel=getLabel(enbid,userLabel,curMeid)
            content=[eng,ch,curMoi,subnet,curMeid,curLabel,rs.cell_value(i,col),showValue,level]
            writeContent(ws,content,cur)
            del content
    pass

#--------------------------------------------------------------------------------
#创建输出文件
ws=open("LTE参数检查结果.csv","w")

#表头
content=["参数英文名","参数中文名","MOI","SubNetwork","ENBID","网元名称","现网设置值","合理设置值","重要性"]

#[0,1]表示输出文件的第1个.csv文件的第0行
cur=[0,1]

#写入输出文件表头
writeContent(ws,content,cur)
del content

#打开输入文件
cwd=os.getcwd()
cwd=cwd+"\\合并文件"
listfile=os.listdir(cwd)
infile=cwd+"\\"+listfile[0]
print "正在打开要检查的源文件，请等候..."
rb=xlrd.open_workbook(infile)

#--------------------------------------------------------------------------------
#取ENBID和网元名称
rs=rb.sheet_by_name("ENBFunctionTDD")
enbid=createList(rs,"ENBFunctionTDD")
userLabel=createList(rs,"userLabel")
del rs

#--------------------------------------------------------------------------------
'''定时器'''
print "正在检查定时器参数..."
rs=rb.sheet_by_name("UeTimerTDD")
pmcheck1(rs,ws,"t300","UE等待RRC连接响应的定时器(T300)",enbid,userLabel,5,"A",cur)
pmcheck1(rs,ws,"t301","UE等待RRC重建响应的定时器(T301)",enbid,userLabel,4,"A",cur)
pmcheck1(rs,ws,"t302","UE等待RRC连接重试请求的定时器 (T302)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"t304","UE等待切换成功的定时器(T304)",enbid,userLabel,4,"A",cur)
pmcheck1(rs,ws,"t310_Ue","UE监测无线链路失败的定时器(T310_UE)",enbid,userLabel,5,"A",cur)
pmcheck1(rs,ws,"t311_Ue","UE监测无线链路失败转入空闲状态的定时器(T311_UE)",enbid,userLabel,0,"A",cur)
pmcheck1(rs,ws,"n310","UE接收下行失步指示的最大个数(N310_UE)",enbid,userLabel,7,"A",cur)
pmcheck1(rs,ws,"n311","UE接收下行同步指示的最大个数(N311)",enbid,userLabel,0,"A",cur)

#--------------------------------------------------------------------------------
'''DXR参数'''
print "正在检查DRX参数..."
pmcheck1(rs,ws,"tUserInac","控制面user-inactivity定时器",enbid,userLabel,5,"B",cur)
del rs

rs=rb.sheet_by_name("ServiceDRXTDD")
pmcheck2a(rs,ws,"drxInactTimer","DRX非激活定时器(psf)",enbid,userLabel,12,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"drxRetranTimer","DRX的HARQ重传定时器(psf)",enbid,userLabel,2,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"longDrxCyc","长不连续接收循环周期长度(sf)",enbid,userLabel,7,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"onDuratTimer","在DRX循环周期中UE苏醒的时间长度(psf)",enbid,userLabel,6,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"shortDrxCyc","短不连续接收循环周期长度(sf)",enbid,userLabel,5,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"shortDrxCycT","DRX短不连续循环周期定期器长度",enbid,userLabel,2,"A","qCI",8,9,cur)
del rs

rs=rb.sheet_by_name("GlobleSwitchInformationTDD")
pmcheck1(rs,ws,"switchForUserInactivity","User-Inactivity使能",enbid,userLabel,1,"未定级",cur)
del rs

rs=rb.sheet_by_name("PagingTDD")
pmcheck1(rs,ws,"defaultPagingCycle","UE监听寻呼场合的DRX循环周期",enbid,userLabel,2,"B",cur)
del rs

rs=rb.sheet_by_name("EUtranCellTDD")
pmcheck1(rs,ws,"switchForGbrDrx","GBR业务DRX使能开关",enbid,userLabel,1,"B",cur)
pmcheck1(rs,ws,"switchForNGbrDrx","非GBR业务DRX使能开关",enbid,userLabel,1,"B",cur)

#--------------------------------------------------------------------------------
'''下行功率'''
#基带资源参考信号功率（cp）
#PDSCH与小区RS的功率偏差（pa）
#天线端口信号功率比（pb）
print "正在检查下行功率参数..."

rs_cp=rb.sheet_by_name("ECellEquipmentFunctionTDD")
pmcheck3(rs_cp,ws,"cpSpeRefSigPwr","基带资源参考信号功率",enbid,userLabel,6,19,"C",cur)
del rs_cp

rs_pa=rb.sheet_by_name("PowerControlDLTDD")
pmcheck4a(rs_pa,ws,"paForDTCH","PDSCH与小区RS的功率偏差(P_A_DTCH)",enbid,userLabel,4,4,"B",rs,"cellRSPortNum",0,cur)
pmcheck4a(rs_pa,ws,"paForDTCH","PDSCH与小区RS的功率偏差(P_A_DTCH)",enbid,userLabel,2,2,"B",rs,"cellRSPortNum",1,cur)
del rs_pa

pmcheck2(rs,ws,"pb","天线端口信号功率比",enbid,userLabel,0,"B","cellRSPortNum",0,cur)
pmcheck2(rs,ws,"pb","天线端口信号功率比",enbid,userLabel,1,"B","cellRSPortNum",1,cur)

#--------------------------------------------------------------------------------
'''上行功控'''
print "正在检查上行功控参数..."
rs=rb.sheet_by_name("PowerControlULTDD")
pmcheck1(rs,ws,"alpha","PUSCH发射功率时路损弥补因子",enbid,userLabel,5,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat1","PUCCH Format1物理信道功率弥补量(dB)",enbid,userLabel,1,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat1b","PUCCH Format1b物理信道功率弥补量(dB)",enbid,userLabel,1,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat2","PUCCH Format2物理信道功率弥补量(dB)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat2a","PUCCH Format2a物理信道功率弥补量(dB)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat2b","PUCCH Format2b物理信道功率弥补量(dB)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"ks","用于弥补调制和码率对上行物理信道功率偏差值的影响",enbid,userLabel,0,"B",cur)
pmcheck1(rs,ws,"p0NominalPUSCH","PUSCH半静态调度授权方式发送数据所需小区名义功率",enbid,userLabel,-87,"B",cur)
pmcheck3(rs,ws,"poNominalPUCCH","PUCCH物理信道使用的小区相关名义功率",enbid,userLabel,-105,-100,"B",cur)
pmcheck1(rs,ws,"switchForCLPCofPUCCH","PUCCH闭环功控开关",enbid,userLabel,1,"A",cur)
pmcheck1(rs,ws,"switchForCLPCofPUSCH","PUSCH闭环功控开关",enbid,userLabel,1,"A",cur)
del rs

rs=rb.sheet_by_name("PrachTDD")
pmcheck3(rs,ws,"powerRampingStep","PRACH的功率攀升步长(dB)",enbid,userLabel,1,2,"B",cur)
pmcheck3(rs,ws,"preambleTransMax","PRACH前缀最大发送次数",enbid,userLabel,5,6,"B",cur)
pmcheck3(rs,ws,"preambleIniReceivedPower","PRACH初始前缀接收功率(dBm)",enbid,userLabel,8,10,"B",cur)
del rs

rs=rb.sheet_by_name("EUtranReselectionTDD")
pmcheck1(rs,ws,"intraPmax","UE发射功率最大值(dBm)",enbid,userLabel,23,"A",cur)

#--------------------------------------------------------------------------------
'''其他重点参数'''
print "正在检查其他重点参数..."
pmcheck3(rs,ws,"cellBarred","小区禁止接入指示",enbid,userLabel,0,1,"B",cur)
del rs

rs=rb.sheet_by_name("LoadManagementTDD")
pmcheck1(rs,ws,"lbSwch","负荷均衡算法开关",enbid,userLabel,0,"未定级",cur)
pmcheck1(rs,ws,"lcSwch","负荷控制算法开关",enbid,userLabel,0,"未定级",cur)
del rs

rs=rb.sheet_by_name("SecurityManagementTDD")
pmcheckpr(rs,ws,"encrypAlgPriority","加密算法优先级序列",enbid,userLabel,1,"1;4;4;4","未定级",cur)
pmcheckpr(rs,ws,"integProtAlgPriority","完整性保护算法优先级序列",enbid,userLabel,1,"1;4;4;4","未定级",cur)
del rs

rs=rb.sheet_by_name("EUtranCellTDD")
pmcheck1a(rs,ws,"bandWidth","小区系统频域带宽(MHz)",enbid,userLabel,5,3,"C",cur)
pmcheck3(rs,ws,"cFI","CFI选择",enbid,userLabel,0,3,"C",cur)
pmcheck2(rs,ws,"flagSwiMode","切换模式选择",enbid,userLabel,1,"C","cellRSPortNum",0,cur)
pmcheck2b(rs,ws,"flagSwiMode","切换模式选择",enbid,userLabel,3,23,"C","cellRSPortNum",1,cur)
pmcheck2b(rs,ws,"flagSwiMode","切换模式选择",enbid,userLabel,3,23,"C","cellRSPortNum",2,cur)
pmcheck1(rs,ws,"maxUeRbNumDl","下行UE最大分配RB数",enbid,userLabel,100,"未定级",cur)
pmcheck1(rs,ws,"maxUeRbNumUl","上行UE最大分配RB数",enbid,userLabel,33,"未定级",cur)
pmcheck1(rs,ws,"sfAssignment","上下行子帧分配配置",enbid,userLabel,2,"A",cur)
pmcheck2(rs,ws,"specialSfPatterns","特殊子帧配置",enbid,userLabel,7,"A","bandIndicator",38,cur)
pmcheck2(rs,ws,"specialSfPatterns","特殊子帧配置",enbid,userLabel,7,"A","bandIndicator",40,cur)
pmcheck2(rs,ws,"specialSfPatterns","特殊子帧配置",enbid,userLabel,6,"A","bandIndicator",39,cur)
pmcheck1(rs,ws,"rd4ForCoverage","基于覆盖的重定向算法启动开关",enbid,userLabel,1,"A",cur)
del rs

rs=rb.sheet_by_name("PhyChannelTDD")
pmcheck1(rs,ws,"maxUserPucchfmt1","每个RB内PUCCH format1可复用的最大用户数",enbid,userLabel,3,"未定级",cur)
pmcheck1(rs,ws,"pucchAckRepNum","Ack/Nack重复PUCCH信道个数",enbid,userLabel,0,"未定级",cur)
del rs

rs=rb.sheet_by_name("PrachTDD")
pmcheck3(rs,ws,"maxHarqMsg3Tx","Message 3最大发送次数",enbid,userLabel,1,5,"未定级",cur)
pmcheck1(rs,ws,"raResponseWindowSize","UE对随机接入前缀响应接收的搜索窗口(毫秒)",enbid,userLabel,7,"未定级",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-2G空闲态'''
print "正在检查4-2空闲态参数..."
rs=rb.sheet_by_name("GsmReselectionTDD")
#pmcheck3(rs,ws,"geranFreqNum","GERAN载频数目",enbid,userLabel,0,1,"B",cur)
pmchecklp(rs,ws,"gsmRslPara_geranReselectionPriority","GERAN小区重选优先级",enbid,userLabel,2,"B",cur)
pmchecklp(rs,ws,"gsmRslPara_geranThreshXLow","重选到低优先级GERAN小区的RSRP门限(dB)",enbid,userLabel,16,"C",cur)
pmchecklp(rs,ws,"gsmRslPara_qRxLevMin","GERAN小区重选所需要的最小接收电平(dBm)",enbid,userLabel,-109,"C",cur)
pmcheck1(rs,ws,"tReselectionGERAN","重选到GERAN小区判决定时器长度(秒)",enbid,userLabel,2,"B",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-3G空闲态'''
print "正在检查4-3空闲态参数..."
rs=rb.sheet_by_name("UtranTReselectionTDD")
pmchecklp2(rs,ws,"utranTDDRslPara_qRxLevMinTDD","UTRAN TDD小区重选所需要的最小接收电平(dBm)",enbid,userLabel,-115,"B",cur)
pmchecklp2(rs,ws,"utranTDDRslPara_threshXLowTDD","重选到低优先级UTRAN TDD小区的RSRP门限(dB)",enbid,userLabel,22,"B",cur)
pmchecklp2(rs,ws,"utranTDDRslPara_utranTDDReselPriority","UTRAN TDD小区重选优先级",enbid,userLabel,3,"B",cur)
del rs

rs=rb.sheet_by_name("UtranCellReselectionTDD")
pmcheck1(rs,ws,"reselUtran","重选到UTRAN小区判决定时器长度(秒)",enbid,userLabel,2,"B",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-4G空闲态'''
print "正在检查4-4空闲态参数..."
rs=rb.sheet_by_name("EUtranRelationTDD")
pmcheck3(rs,ws,"qofStCell","重选时相邻小区对服务小区偏差(dB)",enbid,userLabel,12,18,"C",cur)
del rs

rs=rb.sheet_by_name("EUtranReselectionTDD")
pmcheck3(rs,ws,"cellReselectionPriority","频内小区重选优先级",enbid,userLabel,4,6,"B",cur)
pmcheck3(rs,ws,"selQrxLevMin","小区选择所需的最小RSRP接收水平",enbid,userLabel,-126,-120,"B",cur)
pmcheck1a(rs,ws,"snonintrasearch","同/低优先级RSRP测量判决门限",enbid,userLabel,28,10,"B",cur)
pmcheck3(rs,ws,"qhyst","服务小区重选迟滞(dB)",enbid,userLabel,3,4,"B",cur)
pmcheck3(rs,ws,"qrxLevMinOfst","小区选择所需的最小RSRP接收电平偏移(dB)",enbid,userLabel,0,2,"B",cur)
pmcheck1a(rs,ws,"sIntraSearch","同频测量RSRP判决门限(db)",enbid,userLabel,40,44,"B",cur)
pmcheck3(rs,ws,"tReselectionIntraEUTRA","频内重选判决定时器时长（秒）",enbid,userLabel,1,2,"B",cur)
pmcheck3(rs,ws,"threshSvrLow","服务载频低门限(dB)",enbid,userLabel,0,4,"B",cur)

pmcheckmul(rs,ws,"eutranRslPara_interReselPrio","频间小区重选优先级",enbid,userLabel,4,6,"B",cur)
pmcheckmul3(rs,ws,"eutranRslPara_qOffsetFreq","频间频率偏移值",enbid,userLabel,12,15,18,"B",cur)
pmcheckmul(rs,ws,"eutranRslPara_interThrdXHigh","异频频点高优先级重选门限",enbid,userLabel,30,8,"B",cur)
pmcheckmul3(rs,ws,"eutranRslPara_interThrdXLow","异频频点低优先级重选门限",enbid,userLabel,16,20,10,"B",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-2G CSFB'''
print "正在检查4-2 CSFB参数..."
rs=rb.sheet_by_name("EUtranCellMeasurementTDD")
pmcheck1(rs,ws,"csfbMethodofGSM","CS Fallback到GERAN时，优先采用的方式",enbid,userLabel,2,"C",cur)
pmcheck1(rs,ws,"geranCarriFreqNum","GERAN载频数目",enbid,userLabel,1,"未定级",cur)
pmcheck1(rs,ws,"geranMeasParas_geranFreqCsfbPriority","GERAN频点CSFB优先级",enbid,userLabel,255,"C",cur)
pmcheck1(rs,ws,"geranMeasParas_geranFreqRdPriority","GERAN频点重定向优先级",enbid,userLabel,254,"C",cur)
pmcheck1(rs,ws,"ratPriCnPara_ratPriCnCSFB1","GERAN系统连接态用户CS Fallback目标系统优先级",enbid,userLabel,100,"未定级",cur)
pmcheck1(rs,ws,"ratPriIdPara_ratPriIdleCSFB1","GERAN系统空闲态用户CS Fallback目标系统优先级",enbid,userLabel,100,"未定级",cur)
pmcheck3(rs,ws,"geranMeasParas_startARFCN","起始ARFCN",enbid,userLabel,0,661,"未定级",cur)

#--------------------------------------------------------------------------------
'''4G-3G连接态'''
print "正在检查4-3连接态参数..."
pmchecklp(rs,ws,"utranMeasParas_utranOffFreq","UTRAN系统间频率偏移值(dB)",enbid,userLabel,0,"C",cur)
del rs

rs=rb.sheet_by_name("UeRATMeasurementTDD")
pmcheck2(rs,ws,"eventId","系统间测量事件标识",enbid,userLabel,1,"未定级","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"hysterisis","B2判决迟滞范围(dB)",enbid,userLabel,0,2,"B","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"rscpSysNbrTrd","RSCP测量UTRAN系统判决绝对门限(dBm)",enbid,userLabel,-93,-90,"B","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"trigTime","B2事件发生到上报时间差",enbid,userLabel,8,11,"B","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"rsrpSrvTrd","RSRP测量时E-UTRAN系统服务小区判决的绝对门限",enbid,userLabel,-140,-116,"B","ratMeasCfgIdx",1110,cur)
del rs

rs=rb.sheet_by_name("UeEUtranMeasurementTDD")
pmcheck4(rs,ws,"hysteresis","A2(测量重定向)判决迟滞范围(dB)",enbid,userLabel,0,2,"B","measCfgIdx",30,cur)
pmcheck4(rs,ws,"hysteresis","A2(盲重定向)判决迟滞范围(dB)",enbid,userLabel,0,2,"B","measCfgIdx",40,cur)
pmcheck4(rs,ws,"timeToTrigger","A2(测量重定向)事件发生到上报时间差",enbid,userLabel,8,11,"B","measCfgIdx",30,cur)
pmcheck4(rs,ws,"timeToTrigger","A2(盲重定向)事件发生到上报时间差",enbid,userLabel,8,11,"B","measCfgIdx",40,cur)
pmcheck4(rs,ws,"thresholdOfRSRP","A2事件(测量重定向)判决的RSRP门限",enbid,userLabel,-119,-105,"B","measCfgIdx",30,cur)
pmcheck4(rs,ws,"thresholdOfRSRP","A2事件(盲重定向)判决的RSRP门限",enbid,userLabel,-127,-121,"B","measCfgIdx",40,cur)

#--------------------------------------------------------------------------------
'''4G-4G连接态'''
print "正在检查4-4连接态参数..."
pmcheck4(rs,ws,"a3Offset","同频A3事件偏移",enbid,userLabel,1,3,"B","measCfgIdx",50,cur)
pmcheck4(rs,ws,"hysteresis","同频A3判决迟滞范围(dB)",enbid,userLabel,1,2,"B","measCfgIdx",50,cur)
pmcheck4(rs,ws,"timeToTrigger","A2事件(频间)发生到上报时间差",enbid,userLabel,8,11,"B","measCfgIdx",20,cur)
pmcheck4(rs,ws,"timeToTrigger","同频A3事件发生到上报时间差",enbid,userLabel,8,11,"B","measCfgIdx",50,cur)
pmcheck4(rs,ws,"timeToTrigger","异频A3事件发生到上报时间差",enbid,userLabel,8,11,"B","measCfgIdx",70,cur)

pmcheck4(rs,ws,"thresholdOfRSRP","A2事件(频间)判决的RSRP门限",enbid,userLabel,-118,-82,"B","measCfgIdx",20,cur)
pmcheck5(rs,ws,"thresholdOfRSRP","A1事件判决的RSRP门限",enbid,userLabel,3,"B","measCfgIdx",10,20,cur)
pmcheck4(rs,ws,"a3Offset","异频A3事件偏移",enbid,userLabel,0,6,"B","measCfgIdx",70,cur)
pmcheck4(rs,ws,"hysteresis","异频A3判决迟滞范围(dB)",enbid,userLabel,0,3,"B","measCfgIdx",70,cur)
pmcheck4(rs,ws,"thresholdOfRSRP","A4事件判决的RSRP门限",enbid,userLabel,-100,-92,"B","measCfgIdx",288,cur)

del rs
#rs_cm=rb.sheet_by_name("CellMeasGroupTDD")#interFHOMeasCfg
#del rs_cm
