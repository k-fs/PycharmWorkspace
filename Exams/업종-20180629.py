import sys
from PyQt5.QtWidgets import *
from enum import Enum
import win32com.client
import time
import pythoncom
 
import pyodbc
from datetime import date

g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objElwMgr = win32com.client.Dispatch("CpUtil.CpElwCode")
g_objFutureMgr = win32com.client.Dispatch("CpUtil.CpFutureCode")
g_objOptionMgr = win32com.client.Dispatch("CpUtil.CpOptionCode")


# ���� ���� ���簡 ���� ���� ����ü
class stockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "����ȣ��/���� �ƴ�", ord('1'): "����ȣ��", ord('2'): "����"}
        self.code = ""
        self.name = ""
        self.cur = 0        # ���簡
        self.open = self.high = self.low = 0 # ��/��/��
        self.diff = 0
        self.diffp = 0
        self.objCur = None
        self.objBid = None
        self.vol = 0 # �ŷ���
        self.tradeAek = 0  # �ŷ���� #
        self.listedStock = 0  # �����ֽļ� #
        self.junga = 0  # ���ϰ��� #
        self.index = 0  # ���� #
 
    # ���� ��� ���
    def makediffp(self, baseprice):
        lastday = 0
        if baseprice :
            lastday =baseprice
        else:
            lastday = self.cur - self.diff
        if lastday:
            self.diffp = (self.diff / lastday) * 100
        else:
            self.diffp = 0
 
    def debugPrint(self, type):
        if type == 0  :
            print("[%s], %s, %s %s, ���簡 %d ��� %d, (%.2f), �ŷ���� %s, �����ֽļ� %s, ���ϰ��� %s"
                  %(self.index, self.dicEx.get(self.exFlag), self.code, self.name, self.cur, self.diff, self.diffp,
                    self.tradeAek, self.listedStock, self.junga ) )
        else :
            print("[%s], %s %s, ���簡 %.2f ��� %.2f, (%.2f), �ŷ���� %s, �����ֽļ� %s, ���ϰ��� %s"
                  % (self.index, self.code, self.name, self.cur, self.diff, self.diffp,
                    self.tradeAek, self.listedStock, self.junga ) )
 
 
class CpMarketEye:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS�� ���������� ������� ����. ")
            return False
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0
 
    def Request(self, codes, dicCodes):
        #rqField= [�ڵ�, ����ȣ, ���, ���簡, �ð�, ��, ����, �ŵ�ȣ��, �ż�ȣ��, �ŷ���, �屸��, �ŵ��ܷ�,�ż��ܷ�,
        rqField = [0, 2, 3, 4, 5, 6, 7, 10, 11, 12, 20, 23]  # ��û �ʵ�
 
        self.objRq.SetInputValue(0, rqField)  # ��û �ʵ�
        self.objRq.SetInputValue(1, codes)  # �����ڵ� or �����ڵ� ����Ʈ
        self.objRq.BlockRequest()
 
        # ���簡 ��� �� ��� ���� ó��
        rqStatus = self.objRq.GetDibStatus()
        print("��Ż���", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(2)
 
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # �ڵ�
            record = None
            if code in dicCodes:
                record = dicCodes.get(code)
            else:
                record = stockPricedData()
 
            record.code = code
            record.name = g_objCodeMgr.CodeToName(code)
 
            record.diff = self.objRq.GetDataValue(2, i)  # ���ϴ��
            record.cur = self.objRq.GetDataValue(3, i)  # ���簡
            record.open = self.objRq.GetDataValue(4, i)  # �ð�
            record.high = self.objRq.GetDataValue(5, i)  # ��
            record.low = self.objRq.GetDataValue(6, i)  # ����
            record.vol = self.objRq.GetDataValue(7, i)  # �ŷ���
            record.tradeAek = self.objRq.GetDataValue(8, i)  # �ŷ���� #
            record.exFlag = self.objRq.GetDataValue(9, i)  # �屸��
            record.listedStock = self.objRq.GetDataValue(10, i)  # �����ֽļ� #
            record.junga = self.objRq.GetDataValue(11, i)  # ���ϰ��� #
            record.index = i  # ���� #
 
            record.makediffp(0)
            dicCodes[code] = record
 
        return True
 
 
 
 
# ���� �ڵ�  ���� Ŭ����
class testMain():
    def __init__(self):
        self.dicStockCodes = dict()  # �ֽ� �� ���� �ü� 
        self.dicElwCodes = dict()  # ELW ������ �ü� 
        self.dicFutreCodes = dict()  # ���� ������ �ü� 
        self.dicOptionCodes = dict()  # �ɼ� ������ �ü� 
        self.dicUpjongCodes = dict()  # ���� ������ �ü� 
        self.obj = CpMarketEye()
        return
 
    def ReqeustStockMst(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # �ŷ���
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # �ڽ���
        codeList3 = g_objCodeMgr.GetStockListByMarket(3)  # K-OTC
        codeList4 = g_objCodeMgr.GetStockListByMarket(4)  # KRX
        codeList5 = g_objCodeMgr.GetStockListByMarket(5)  # KONEX

        allcodelist = codeList + codeList2 + codeList3 + codeList4 + codeList5
        print('�� ���� �ڵ� %d, �ŷ��� %d, �ڽ��� %d, K-OTC %d, KRX %d, KONEX %d' % (len(allcodelist), len(codeList), len(codeList2), len(codeList3), len(codeList4), len(codeList5)))
 
        # Some other example server values are
        server = 'tcp:172.20.200.252,8629' 
        cnxn = pyodbc.connect('DSN=TDB;UID=icam;PWD=kfam1801')
        curs = cnxn.cursor()
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicStockCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicStockCodes)
 
        print("��ü ����� : ", len(self.dicStockCodes))
        for key in self.dicStockCodes :
            self.dicStockCodes[key].debugPrint(0)
 
        curs.close()
        cnxn.close()
 
    def ReqeustElwMst(self):
 
        allcodelist = []
        for i in range(g_objElwMgr.GetCount()) :
            allcodelist.append(g_objElwMgr.GetData(0,i))
 
        print("��ü ���� �ڵ� #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicElwCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicElwCodes)
 
        print("ELW ������", len(self.dicElwCodes))
        for key in self.dicElwCodes :
            self.dicElwCodes[key].debugPrint(0)
 
    def ReqeustFutreMst(self):
        allcodelist = []
        for i in range(g_objFutureMgr.GetCount()) :
            allcodelist.append(g_objFutureMgr.GetData(0,i))
 
        print("��ü ���� �ڵ� #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicFutreCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicFutreCodes)
 
        print("���� ������ ", len(self.dicFutreCodes))
        for key in self.dicFutreCodes :
            self.dicFutreCodes[key].debugPrint(1)
 
    def ReqeustOptionMst(self):
        allcodelist = []
        for i in range(g_objOptionMgr.GetCount()) :
            allcodelist.append(g_objOptionMgr.GetData(0,i))
 
        print("��ü ���� �ڵ� #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicOptionCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicOptionCodes)
 
        print("�ɼ� ������ ", len(self.dicOptionCodes))
        for key in self.dicOptionCodes :
            self.dicOptionCodes[key].debugPrint(1)
 
    def ReqeustUpjongMst(self):
        codeList = g_objCodeMgr.GetIndustryList() # ���� ��� ���� ����Ʈ
 
        allcodelist = codeList
        print("��ü ���� �ڵ� #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            code2 = "U" + code
            rqCodeList.append(code2)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicUpjongCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicUpjongCodes)
 
        print("���ǻ������ ����Ʈ", len(self.dicUpjongCodes))
        for key in self.dicUpjongCodes :
            self.dicUpjongCodes[key].debugPrint(1)
 
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.main = testMain()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 500,  300, 500)
 
        nHeight = 20
        btnStock = QPushButton("�ֽ� �� ����", self)
        btnStock.move(20, nHeight)
        btnStock.clicked.connect(self.btnStock_clicked)
 
        nHeight += 50
        btnElw = QPushButton("ELW �� ����", self)
        btnElw.move(20, nHeight)
        btnElw.clicked.connect(self.btnElw_clicked)
 
        nHeight += 50
        btnFuture = QPushButton("���� �� ����", self)
        btnFuture.move(20, nHeight)
        btnFuture.clicked.connect(self.btnFuture_clicked)
 
        nHeight += 50
        btnOption= QPushButton("�ɼ� �� ����", self)
        btnOption.move(20, nHeight)
        btnOption.clicked.connect(self.btnOption_clicked)
 
        nHeight += 50
        btnUpjong= QPushButton("���� �� ����", self)
        btnUpjong.move(20, nHeight)
        btnUpjong.clicked.connect(self.btnUpjong_clicked)
 
        nHeight += 50
        btnExit = QPushButton("����", self)
        btnExit.move(20, nHeight)
        btnExit.clicked.connect(self.btnExit_clicked)
 
    def btnStock_clicked(self):
        self.main.ReqeustStockMst()
        return
 
    def btnElw_clicked(self):
        self.main.ReqeustElwMst()
        return
 
    def btnFuture_clicked(self):
        self.main.ReqeustFutreMst()
        return
 
    def btnOption_clicked(self):
        self.main.ReqeustOptionMst()
        return
 
    def btnUpjong_clicked(self):
        self.main.ReqeustUpjongMst()
        return
 
 
    def btnExit_clicked(self):
        exit()
        return
 
 
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()

