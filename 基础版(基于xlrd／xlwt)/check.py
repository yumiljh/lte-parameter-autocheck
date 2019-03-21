# -*- coding: gbk -*-
import xlrd
import sys
import os

reload(sys)
sys.setdefaultencoding("gbk")

#�ҳ�rs�������ֶ�searchStr���к�
def getCol(rs,searchStr):
    colNum=0
    for x in rs.row_values(0):
        if x==searchStr:
            return colNum
        colNum=colNum+1
    return 0

#�ҳ�searchMeid����Ԫ����
def getLabel(meidList,userLabel,searchMeid):
    for x in meidList:
        if x==searchMeid:
            return userLabel[meidList.index(x)]
    return 0

#�ҳ�rs��searchStr�ֶε�MEID��ΪsearchMeid�ĵ�Ԫ���ֵ
def getCompanion(rs,searchStr,searchMeid):
    col_sstr=getCol(rs,searchStr)
    enbCol=getCol(rs,"ENBFunctionTDD")
    rslen=rs.nrows
    for i in range(5,rslen):
        if rs.cell_value(i,enbCol)==searchMeid:
            return rs.cell_value(i,col_sstr)

#�ҳ�rs��searchStr�ֶε�MEID��ΪsearchMeid��CELLID��λsearchCid�ĵ�Ԫ���ֵ
def getCompanion2(rs,searchStr,searchMeid,searchCid):
    col_sstr=getCol(rs,searchStr)
    enbCol=getCol(rs,"ENBFunctionTDD")
    cellCol=getCol(rs,"EUtranCellTDD")
    rslen=rs.nrows
    for i in range(5,rslen):
        if rs.cell_value(i,enbCol)==searchMeid and rs.cell_value(i,cellCol)==searchCid:
            return rs.cell_value(i,col_sstr)

#�ҳ�rs��searchStr�ֶε�MEID��ΪsearchMeid��meas��λsearchMeas�ĵ�Ԫ���ֵ
def getCompanion3(rs,searchStr,searchMeid,searchMeas):
    col_sstr=getCol(rs,searchStr)
    enbCol=getCol(rs,"ENBFunctionTDD")
    measCol=getCol(rs,"measCfgIdx")
    rslen=rs.nrows
    for i in range(5,rslen):
        if rs.cell_value(i,enbCol)==searchMeid and int(rs.cell_value(i,measCol))==searchMeas:
            return rs.cell_value(i,col_sstr)

#����rs��listHead�е��б�
def createList(rs,listHead):
    return [x for x in rs.col_values(getCol(rs,listHead))]

#д��һ��content��ws
def writeContent(ws,content,cur):
    tempStr=""
    for x in content:
        tempStr=tempStr+"\""+str(x)+"\","
    tempStr=tempStr+"\n"
    ws.write(tempStr)
    cur[0]+=1
    pass

#�����ڵ�ֵ��
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

#�����ڵ�ֵ��
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

#�����ڴ����������ĵ�ֵ��
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

#�����ڴ����������ĵ�ֵ��
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

#�����ڴ����������ĵ�ֵ��
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

#�����ڶ�ֵ��
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

#�����ڴ����������Ķ�ֵ��
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

#�����ڴ����������Ķ�ֵ��
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

#�����ڱȽ���
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

#������ѭ���͵�ֵ��
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

#������ѭ���͵�ֵ�ͣ������ֵ
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

#������ѭ���Ͷ�ֵ��
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

#������ѭ���Ͷ�ֵ��
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

#���������е�ֵ��
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
#��������ļ�
ws=open("LTE���������.csv","w")

#��ͷ
content=["����Ӣ����","����������","MOI","SubNetwork","ENBID","��Ԫ����","��������ֵ","��������ֵ","��Ҫ��"]

#[0,1]��ʾ����ļ��ĵ�1��.csv�ļ��ĵ�0��
cur=[0,1]

#д������ļ���ͷ
writeContent(ws,content,cur)
del content

#�������ļ�
cwd=os.getcwd()
cwd=cwd+"\\�ϲ��ļ�"
listfile=os.listdir(cwd)
infile=cwd+"\\"+listfile[0]
print "���ڴ�Ҫ����Դ�ļ�����Ⱥ�..."
rb=xlrd.open_workbook(infile)

#--------------------------------------------------------------------------------
#ȡENBID����Ԫ����
rs=rb.sheet_by_name("ENBFunctionTDD")
enbid=createList(rs,"ENBFunctionTDD")
userLabel=createList(rs,"userLabel")
del rs

#--------------------------------------------------------------------------------
'''��ʱ��'''
print "���ڼ�鶨ʱ������..."
rs=rb.sheet_by_name("UeTimerTDD")
pmcheck1(rs,ws,"t300","UE�ȴ�RRC������Ӧ�Ķ�ʱ��(T300)",enbid,userLabel,5,"A",cur)
pmcheck1(rs,ws,"t301","UE�ȴ�RRC�ؽ���Ӧ�Ķ�ʱ��(T301)",enbid,userLabel,4,"A",cur)
pmcheck1(rs,ws,"t302","UE�ȴ�RRC������������Ķ�ʱ�� (T302)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"t304","UE�ȴ��л��ɹ��Ķ�ʱ��(T304)",enbid,userLabel,4,"A",cur)
pmcheck1(rs,ws,"t310_Ue","UE���������·ʧ�ܵĶ�ʱ��(T310_UE)",enbid,userLabel,5,"A",cur)
pmcheck1(rs,ws,"t311_Ue","UE���������·ʧ��ת�����״̬�Ķ�ʱ��(T311_UE)",enbid,userLabel,0,"A",cur)
pmcheck1(rs,ws,"n310","UE��������ʧ��ָʾ��������(N310_UE)",enbid,userLabel,7,"A",cur)
pmcheck1(rs,ws,"n311","UE��������ͬ��ָʾ��������(N311)",enbid,userLabel,0,"A",cur)

#--------------------------------------------------------------------------------
'''DXR����'''
print "���ڼ��DRX����..."
pmcheck1(rs,ws,"tUserInac","������user-inactivity��ʱ��",enbid,userLabel,5,"B",cur)
del rs

rs=rb.sheet_by_name("ServiceDRXTDD")
pmcheck2a(rs,ws,"drxInactTimer","DRX�Ǽ��ʱ��(psf)",enbid,userLabel,12,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"drxRetranTimer","DRX��HARQ�ش���ʱ��(psf)",enbid,userLabel,2,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"longDrxCyc","������������ѭ�����ڳ���(sf)",enbid,userLabel,7,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"onDuratTimer","��DRXѭ��������UE���ѵ�ʱ�䳤��(psf)",enbid,userLabel,6,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"shortDrxCyc","�̲���������ѭ�����ڳ���(sf)",enbid,userLabel,5,"A","qCI",8,9,cur)
pmcheck2a(rs,ws,"shortDrxCycT","DRX�̲�����ѭ�����ڶ���������",enbid,userLabel,2,"A","qCI",8,9,cur)
del rs

rs=rb.sheet_by_name("GlobleSwitchInformationTDD")
pmcheck1(rs,ws,"switchForUserInactivity","User-Inactivityʹ��",enbid,userLabel,1,"δ����",cur)
del rs

rs=rb.sheet_by_name("PagingTDD")
pmcheck1(rs,ws,"defaultPagingCycle","UE����Ѱ�����ϵ�DRXѭ������",enbid,userLabel,2,"B",cur)
del rs

rs=rb.sheet_by_name("EUtranCellTDD")
pmcheck1(rs,ws,"switchForGbrDrx","GBRҵ��DRXʹ�ܿ���",enbid,userLabel,1,"B",cur)
pmcheck1(rs,ws,"switchForNGbrDrx","��GBRҵ��DRXʹ�ܿ���",enbid,userLabel,1,"B",cur)

#--------------------------------------------------------------------------------
'''���й���'''
#������Դ�ο��źŹ��ʣ�cp��
#PDSCH��С��RS�Ĺ���ƫ�pa��
#���߶˿��źŹ��ʱȣ�pb��
print "���ڼ�����й��ʲ���..."

rs_cp=rb.sheet_by_name("ECellEquipmentFunctionTDD")
pmcheck3(rs_cp,ws,"cpSpeRefSigPwr","������Դ�ο��źŹ���",enbid,userLabel,6,19,"C",cur)
del rs_cp

rs_pa=rb.sheet_by_name("PowerControlDLTDD")
pmcheck4a(rs_pa,ws,"paForDTCH","PDSCH��С��RS�Ĺ���ƫ��(P_A_DTCH)",enbid,userLabel,4,4,"B",rs,"cellRSPortNum",0,cur)
pmcheck4a(rs_pa,ws,"paForDTCH","PDSCH��С��RS�Ĺ���ƫ��(P_A_DTCH)",enbid,userLabel,2,2,"B",rs,"cellRSPortNum",1,cur)
del rs_pa

pmcheck2(rs,ws,"pb","���߶˿��źŹ��ʱ�",enbid,userLabel,0,"B","cellRSPortNum",0,cur)
pmcheck2(rs,ws,"pb","���߶˿��źŹ��ʱ�",enbid,userLabel,1,"B","cellRSPortNum",1,cur)

#--------------------------------------------------------------------------------
'''���й���'''
print "���ڼ�����й��ز���..."
rs=rb.sheet_by_name("PowerControlULTDD")
pmcheck1(rs,ws,"alpha","PUSCH���书��ʱ·���ֲ�����",enbid,userLabel,5,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat1","PUCCH Format1�����ŵ������ֲ���(dB)",enbid,userLabel,1,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat1b","PUCCH Format1b�����ŵ������ֲ���(dB)",enbid,userLabel,1,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat2","PUCCH Format2�����ŵ������ֲ���(dB)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat2a","PUCCH Format2a�����ŵ������ֲ���(dB)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"deltaFPucchFormat2b","PUCCH Format2b�����ŵ������ֲ���(dB)",enbid,userLabel,2,"B",cur)
pmcheck1(rs,ws,"ks","�����ֲ����ƺ����ʶ����������ŵ�����ƫ��ֵ��Ӱ��",enbid,userLabel,0,"B",cur)
pmcheck1(rs,ws,"p0NominalPUSCH","PUSCH�뾲̬������Ȩ��ʽ������������С�����幦��",enbid,userLabel,-87,"B",cur)
pmcheck3(rs,ws,"poNominalPUCCH","PUCCH�����ŵ�ʹ�õ�С��������幦��",enbid,userLabel,-105,-100,"B",cur)
pmcheck1(rs,ws,"switchForCLPCofPUCCH","PUCCH�ջ����ؿ���",enbid,userLabel,1,"A",cur)
pmcheck1(rs,ws,"switchForCLPCofPUSCH","PUSCH�ջ����ؿ���",enbid,userLabel,1,"A",cur)
del rs

rs=rb.sheet_by_name("PrachTDD")
pmcheck3(rs,ws,"powerRampingStep","PRACH�Ĺ�����������(dB)",enbid,userLabel,1,2,"B",cur)
pmcheck3(rs,ws,"preambleTransMax","PRACHǰ׺����ʹ���",enbid,userLabel,5,6,"B",cur)
pmcheck3(rs,ws,"preambleIniReceivedPower","PRACH��ʼǰ׺���չ���(dBm)",enbid,userLabel,8,10,"B",cur)
del rs

rs=rb.sheet_by_name("EUtranReselectionTDD")
pmcheck1(rs,ws,"intraPmax","UE���书�����ֵ(dBm)",enbid,userLabel,23,"A",cur)

#--------------------------------------------------------------------------------
'''�����ص����'''
print "���ڼ�������ص����..."
pmcheck3(rs,ws,"cellBarred","С����ֹ����ָʾ",enbid,userLabel,0,1,"B",cur)
del rs

rs=rb.sheet_by_name("LoadManagementTDD")
pmcheck1(rs,ws,"lbSwch","���ɾ����㷨����",enbid,userLabel,0,"δ����",cur)
pmcheck1(rs,ws,"lcSwch","���ɿ����㷨����",enbid,userLabel,0,"δ����",cur)
del rs

rs=rb.sheet_by_name("SecurityManagementTDD")
pmcheckpr(rs,ws,"encrypAlgPriority","�����㷨���ȼ�����",enbid,userLabel,1,"1;4;4;4","δ����",cur)
pmcheckpr(rs,ws,"integProtAlgPriority","�����Ա����㷨���ȼ�����",enbid,userLabel,1,"1;4;4;4","δ����",cur)
del rs

rs=rb.sheet_by_name("EUtranCellTDD")
pmcheck1a(rs,ws,"bandWidth","С��ϵͳƵ�����(MHz)",enbid,userLabel,5,3,"C",cur)
pmcheck3(rs,ws,"cFI","CFIѡ��",enbid,userLabel,0,3,"C",cur)
pmcheck2(rs,ws,"flagSwiMode","�л�ģʽѡ��",enbid,userLabel,1,"C","cellRSPortNum",0,cur)
pmcheck2b(rs,ws,"flagSwiMode","�л�ģʽѡ��",enbid,userLabel,3,23,"C","cellRSPortNum",1,cur)
pmcheck2b(rs,ws,"flagSwiMode","�л�ģʽѡ��",enbid,userLabel,3,23,"C","cellRSPortNum",2,cur)
pmcheck1(rs,ws,"maxUeRbNumDl","����UE������RB��",enbid,userLabel,100,"δ����",cur)
pmcheck1(rs,ws,"maxUeRbNumUl","����UE������RB��",enbid,userLabel,33,"δ����",cur)
pmcheck1(rs,ws,"sfAssignment","��������֡��������",enbid,userLabel,2,"A",cur)
pmcheck2(rs,ws,"specialSfPatterns","������֡����",enbid,userLabel,7,"A","bandIndicator",38,cur)
pmcheck2(rs,ws,"specialSfPatterns","������֡����",enbid,userLabel,7,"A","bandIndicator",40,cur)
pmcheck2(rs,ws,"specialSfPatterns","������֡����",enbid,userLabel,6,"A","bandIndicator",39,cur)
pmcheck1(rs,ws,"rd4ForCoverage","���ڸ��ǵ��ض����㷨��������",enbid,userLabel,1,"A",cur)
del rs

rs=rb.sheet_by_name("PhyChannelTDD")
pmcheck1(rs,ws,"maxUserPucchfmt1","ÿ��RB��PUCCH format1�ɸ��õ�����û���",enbid,userLabel,3,"δ����",cur)
pmcheck1(rs,ws,"pucchAckRepNum","Ack/Nack�ظ�PUCCH�ŵ�����",enbid,userLabel,0,"δ����",cur)
del rs

rs=rb.sheet_by_name("PrachTDD")
pmcheck3(rs,ws,"maxHarqMsg3Tx","Message 3����ʹ���",enbid,userLabel,1,5,"δ����",cur)
pmcheck1(rs,ws,"raResponseWindowSize","UE���������ǰ׺��Ӧ���յ���������(����)",enbid,userLabel,7,"δ����",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-2G����̬'''
print "���ڼ��4-2����̬����..."
rs=rb.sheet_by_name("GsmReselectionTDD")
#pmcheck3(rs,ws,"geranFreqNum","GERAN��Ƶ��Ŀ",enbid,userLabel,0,1,"B",cur)
pmchecklp(rs,ws,"gsmRslPara_geranReselectionPriority","GERANС����ѡ���ȼ�",enbid,userLabel,2,"B",cur)
pmchecklp(rs,ws,"gsmRslPara_geranThreshXLow","��ѡ�������ȼ�GERANС����RSRP����(dB)",enbid,userLabel,16,"C",cur)
pmchecklp(rs,ws,"gsmRslPara_qRxLevMin","GERANС����ѡ����Ҫ����С���յ�ƽ(dBm)",enbid,userLabel,-109,"C",cur)
pmcheck1(rs,ws,"tReselectionGERAN","��ѡ��GERANС���о���ʱ������(��)",enbid,userLabel,2,"B",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-3G����̬'''
print "���ڼ��4-3����̬����..."
rs=rb.sheet_by_name("UtranTReselectionTDD")
pmchecklp2(rs,ws,"utranTDDRslPara_qRxLevMinTDD","UTRAN TDDС����ѡ����Ҫ����С���յ�ƽ(dBm)",enbid,userLabel,-115,"B",cur)
pmchecklp2(rs,ws,"utranTDDRslPara_threshXLowTDD","��ѡ�������ȼ�UTRAN TDDС����RSRP����(dB)",enbid,userLabel,22,"B",cur)
pmchecklp2(rs,ws,"utranTDDRslPara_utranTDDReselPriority","UTRAN TDDС����ѡ���ȼ�",enbid,userLabel,3,"B",cur)
del rs

rs=rb.sheet_by_name("UtranCellReselectionTDD")
pmcheck1(rs,ws,"reselUtran","��ѡ��UTRANС���о���ʱ������(��)",enbid,userLabel,2,"B",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-4G����̬'''
print "���ڼ��4-4����̬����..."
rs=rb.sheet_by_name("EUtranRelationTDD")
pmcheck3(rs,ws,"qofStCell","��ѡʱ����С���Է���С��ƫ��(dB)",enbid,userLabel,12,18,"C",cur)
del rs

rs=rb.sheet_by_name("EUtranReselectionTDD")
pmcheck3(rs,ws,"cellReselectionPriority","Ƶ��С����ѡ���ȼ�",enbid,userLabel,4,6,"B",cur)
pmcheck3(rs,ws,"selQrxLevMin","С��ѡ���������СRSRP����ˮƽ",enbid,userLabel,-126,-120,"B",cur)
pmcheck1a(rs,ws,"snonintrasearch","ͬ/�����ȼ�RSRP�����о�����",enbid,userLabel,28,10,"B",cur)
pmcheck3(rs,ws,"qhyst","����С����ѡ����(dB)",enbid,userLabel,3,4,"B",cur)
pmcheck3(rs,ws,"qrxLevMinOfst","С��ѡ���������СRSRP���յ�ƽƫ��(dB)",enbid,userLabel,0,2,"B",cur)
pmcheck1a(rs,ws,"sIntraSearch","ͬƵ����RSRP�о�����(db)",enbid,userLabel,40,44,"B",cur)
pmcheck3(rs,ws,"tReselectionIntraEUTRA","Ƶ����ѡ�о���ʱ��ʱ�����룩",enbid,userLabel,1,2,"B",cur)
pmcheck3(rs,ws,"threshSvrLow","������Ƶ������(dB)",enbid,userLabel,0,4,"B",cur)

pmcheckmul(rs,ws,"eutranRslPara_interReselPrio","Ƶ��С����ѡ���ȼ�",enbid,userLabel,4,6,"B",cur)
pmcheckmul3(rs,ws,"eutranRslPara_qOffsetFreq","Ƶ��Ƶ��ƫ��ֵ",enbid,userLabel,12,15,18,"B",cur)
pmcheckmul(rs,ws,"eutranRslPara_interThrdXHigh","��ƵƵ������ȼ���ѡ����",enbid,userLabel,30,8,"B",cur)
pmcheckmul3(rs,ws,"eutranRslPara_interThrdXLow","��ƵƵ������ȼ���ѡ����",enbid,userLabel,16,20,10,"B",cur)
del rs

#--------------------------------------------------------------------------------
'''4G-2G CSFB'''
print "���ڼ��4-2 CSFB����..."
rs=rb.sheet_by_name("EUtranCellMeasurementTDD")
pmcheck1(rs,ws,"csfbMethodofGSM","CS Fallback��GERANʱ�����Ȳ��õķ�ʽ",enbid,userLabel,2,"C",cur)
pmcheck1(rs,ws,"geranCarriFreqNum","GERAN��Ƶ��Ŀ",enbid,userLabel,1,"δ����",cur)
pmcheck1(rs,ws,"geranMeasParas_geranFreqCsfbPriority","GERANƵ��CSFB���ȼ�",enbid,userLabel,255,"C",cur)
pmcheck1(rs,ws,"geranMeasParas_geranFreqRdPriority","GERANƵ���ض������ȼ�",enbid,userLabel,254,"C",cur)
pmcheck1(rs,ws,"ratPriCnPara_ratPriCnCSFB1","GERANϵͳ����̬�û�CS FallbackĿ��ϵͳ���ȼ�",enbid,userLabel,100,"δ����",cur)
pmcheck1(rs,ws,"ratPriIdPara_ratPriIdleCSFB1","GERANϵͳ����̬�û�CS FallbackĿ��ϵͳ���ȼ�",enbid,userLabel,100,"δ����",cur)
pmcheck3(rs,ws,"geranMeasParas_startARFCN","��ʼARFCN",enbid,userLabel,0,661,"δ����",cur)

#--------------------------------------------------------------------------------
'''4G-3G����̬'''
print "���ڼ��4-3����̬����..."
pmchecklp(rs,ws,"utranMeasParas_utranOffFreq","UTRANϵͳ��Ƶ��ƫ��ֵ(dB)",enbid,userLabel,0,"C",cur)
del rs

rs=rb.sheet_by_name("UeRATMeasurementTDD")
pmcheck2(rs,ws,"eventId","ϵͳ������¼���ʶ",enbid,userLabel,1,"δ����","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"hysterisis","B2�о����ͷ�Χ(dB)",enbid,userLabel,0,2,"B","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"rscpSysNbrTrd","RSCP����UTRANϵͳ�о���������(dBm)",enbid,userLabel,-93,-90,"B","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"trigTime","B2�¼��������ϱ�ʱ���",enbid,userLabel,8,11,"B","ratMeasCfgIdx",1110,cur)
pmcheck4(rs,ws,"rsrpSrvTrd","RSRP����ʱE-UTRANϵͳ����С���о��ľ�������",enbid,userLabel,-140,-116,"B","ratMeasCfgIdx",1110,cur)
del rs

rs=rb.sheet_by_name("UeEUtranMeasurementTDD")
pmcheck4(rs,ws,"hysteresis","A2(�����ض���)�о����ͷ�Χ(dB)",enbid,userLabel,0,2,"B","measCfgIdx",30,cur)
pmcheck4(rs,ws,"hysteresis","A2(ä�ض���)�о����ͷ�Χ(dB)",enbid,userLabel,0,2,"B","measCfgIdx",40,cur)
pmcheck4(rs,ws,"timeToTrigger","A2(�����ض���)�¼��������ϱ�ʱ���",enbid,userLabel,8,11,"B","measCfgIdx",30,cur)
pmcheck4(rs,ws,"timeToTrigger","A2(ä�ض���)�¼��������ϱ�ʱ���",enbid,userLabel,8,11,"B","measCfgIdx",40,cur)
pmcheck4(rs,ws,"thresholdOfRSRP","A2�¼�(�����ض���)�о���RSRP����",enbid,userLabel,-119,-105,"B","measCfgIdx",30,cur)
pmcheck4(rs,ws,"thresholdOfRSRP","A2�¼�(ä�ض���)�о���RSRP����",enbid,userLabel,-127,-121,"B","measCfgIdx",40,cur)

#--------------------------------------------------------------------------------
'''4G-4G����̬'''
print "���ڼ��4-4����̬����..."
pmcheck4(rs,ws,"a3Offset","ͬƵA3�¼�ƫ��",enbid,userLabel,1,3,"B","measCfgIdx",50,cur)
pmcheck4(rs,ws,"hysteresis","ͬƵA3�о����ͷ�Χ(dB)",enbid,userLabel,1,2,"B","measCfgIdx",50,cur)
pmcheck4(rs,ws,"timeToTrigger","A2�¼�(Ƶ��)�������ϱ�ʱ���",enbid,userLabel,8,11,"B","measCfgIdx",20,cur)
pmcheck4(rs,ws,"timeToTrigger","ͬƵA3�¼��������ϱ�ʱ���",enbid,userLabel,8,11,"B","measCfgIdx",50,cur)
pmcheck4(rs,ws,"timeToTrigger","��ƵA3�¼��������ϱ�ʱ���",enbid,userLabel,8,11,"B","measCfgIdx",70,cur)

pmcheck4(rs,ws,"thresholdOfRSRP","A2�¼�(Ƶ��)�о���RSRP����",enbid,userLabel,-118,-82,"B","measCfgIdx",20,cur)
pmcheck5(rs,ws,"thresholdOfRSRP","A1�¼��о���RSRP����",enbid,userLabel,3,"B","measCfgIdx",10,20,cur)
pmcheck4(rs,ws,"a3Offset","��ƵA3�¼�ƫ��",enbid,userLabel,0,6,"B","measCfgIdx",70,cur)
pmcheck4(rs,ws,"hysteresis","��ƵA3�о����ͷ�Χ(dB)",enbid,userLabel,0,3,"B","measCfgIdx",70,cur)
pmcheck4(rs,ws,"thresholdOfRSRP","A4�¼��о���RSRP����",enbid,userLabel,-100,-92,"B","measCfgIdx",288,cur)

del rs
#rs_cm=rb.sheet_by_name("CellMeasGroupTDD")#interFHOMeasCfg
#del rs_cm
