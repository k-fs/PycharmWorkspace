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

 
# 감시 중인 현재가 정보 저장 구조체
class stockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "동시호가/장중 아님", ord('1'): "동시호가", ord('2'): "장중"}
        self.code = ""
        self.name = ""
        self.cur = 0        # 현재가
        self.open = self.high = self.low = 0 # 시/고/저
        self.diff = 0
        self.diffp = 0
        self.objCur = None
        self.objBid = None
        self.vol = 0 # 거래량
        self.tradeAek = 0  # 거래대금 #
        self.listedStock = 0  # 상장주식수 #
        self.junga = 0  # 전일가격 #
        self.index = 0  # 순번 #
 
    # 전일 대비 계산
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

     # 출력 : 각 라인해당부분
    def debugPrint(self, type):
        if type == 0  :
            print("[%s], %s, %s %s, 현재가 %d 대비 %d, (%.2f), 거래대금 %s, 상장주식수 %s, 전일가격 %s"
                  %(self.index, self.dicEx.get(self.exFlag), self.code, self.name, self.cur, self.diff, self.diffp,
                    self.tradeAek, self.listedStock, self.junga ) )
        else :
            print("[%s], %s %s, 현재가 %.2f 대비 %.2f, (%.2f), 거래대금 %s, 상장주식수 %s, 전일가격 %s"
                  % (self.index, self.code, self.name, self.cur, self.diff, self.diffp,
                    self.tradeAek, self.listedStock, self.junga ) )
            
#            print('[%s] %s [%s] 거래량: %s, 거래대금 %s 상장주식수 %s 현재가 %s 전일가격: %s, 전일대비 %s'
#                    % (today, self.dicStockCodes.code, self.dicStockCodes.name, self.dicStockCodes.vol, self.dicStockCodes.tradeAek, self.dicStockCodes.listedStock, self.dicStockCodes.cur, self.dicStockCodes.junga, self.dicStockCodes.diffp ))
 
 
class CpMarketEye:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0
 
    def Request(self, codes, dicCodes):
        #rqField= [코드, 대비부호, 대비, 현재가, 시가, 고가, 저가, 매도호가, 매수호가, 거래량, 장구분, 매도잔량,매수잔량,
        rqField = [0, 2, 3, 4, 5, 6, 7, 10, 11, 12, 20, 23]  # 요청 필드
 
        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()
 
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False
 
        today = date.today()

        cnt = self.objRq.GetHeaderValue(2)
 
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            record = None
            if code in dicCodes:
                record = dicCodes.get(code)
            else:
                record = stockPricedData()
 
            record.code = code[1:]
            record.name = g_objCodeMgr.CodeToName(code)
 
            record.diff = self.objRq.GetDataValue(2, i)  # 전일대비
            record.cur = self.objRq.GetDataValue(3, i)  # 현재가
            record.open = self.objRq.GetDataValue(4, i)  # 시가
            record.high = self.objRq.GetDataValue(5, i)  # 고가
            record.low = self.objRq.GetDataValue(6, i)  # 저가
            record.vol = self.objRq.GetDataValue(7, i)  # 거래량
            record.tradeAek = self.objRq.GetDataValue(8, i)  # 거래대금 #
            record.exFlag = self.objRq.GetDataValue(9, i)  # 장구분
            record.listedStock = self.objRq.GetDataValue(10, i)  # 상장주식수 #
            record.junga = self.objRq.GetDataValue(11, i)  # 전일가격 #
            record.index = i  # 순번 #

#        try:
            print('[%s] %s [%s] 거래량: %s, 거래대금 %s 상장주식수 %s 현재가 %s 전일가격: %s, 전일대비 %s'
                    % (today, record.code, record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga, record.diffp ))
#            curs.execute("insert into oms9jj (eToDate, eStockCode, eStockname, eNujucksu, eNujuckaek, eSangJangSu, eHyunJaeGa, eJunilGa, eJunilDaebi ) values ('%s', '%s', trim('%s'), %s, %s, %s, %s, %s, %s)"
#                    % (today, record.code, record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga, record.diffp ))

            record.makediffp(0)
            dicCodes[code] = record
            cnxn.commit()

#        except:
#            print('Error : [%s] %s [%s] 거래량: %s, 거래대금 %s 상장주식수 %s 현재가 %s 전일가격: %s, 전일대비 %s'
#                    % (today, record.code, record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga, record.diffp ))
#            cnxn.rollback()
#        try:
#            curs.execute("update oms9jj set eStockname = trim('%s'), eNujucksu = %s, eNujuckaek = %s, eSangJangSu = %s, eHyunJaeGa = %s , eJunilGa = %s where eToDate = '%s' and eStockCode = '%s'"
#                         % (record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga, today, record.code))


#        finally:
#            curs.close()
#            cnxn.close()
        return True
 
 
# 샘플 코드  메인 클래스
class testMain():
    def __init__(self):
        self.dicStockCodes = dict()  # 주식 전 종목 시세 
        self.dicElwCodes = dict()  # ELW 전종목 시세 
        self.dicFutreCodes = dict()  # 선물 전종목 시세 
        self.dicOptionCodes = dict()  # 옵션 전종목 시세 
        self.dicUpjongCodes = dict()  # 업종 전종목 시세 
        self.obj = CpMarketEye()
        return
 
    def ReqeustStockMst(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        codeList3 = g_objCodeMgr.GetStockListByMarket(3)  # K-OTC
        codeList4 = g_objCodeMgr.GetStockListByMarket(4)  # KRX
        codeList5 = g_objCodeMgr.GetStockListByMarket(5)  # KONEX

        allcodelist = codeList + codeList2 + codeList3 + codeList4 + codeList5
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d, K-OTC %d, KRX %d, KONEX %d' % (len(allcodelist), len(codeList), len(codeList2), len(codeList3), len(codeList4), len(codeList5)))
 
        # Some other example server values are
        server = 'tcp:172.20.200.252,8629' 
        cnxn = pyodbc.connect('DSN=TDB;UID=icam;PWD=kfam1801')
        curs = cnxn.cursor()
 
        today = date.today()

        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicStockCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicStockCodes)
 
        print("전체 종목수 : ", len(self.dicStockCodes))
        for key in self.dicStockCodes :
            self.dicStockCodes[key].debugPrint(0)

#            print('오늘 %s' %(today))
#        try:
#            curs.execute("insert into oms9jj (eToDate, eStockCode, eStockname, eNujucksu, eNujuckaek, eSangJangSu, eHyunJaeGa, eJunilGa, eJunilDaebi ) values (to_date(%s), %s, trim(%s), %s, %s, %s, %s, %s, %s)"
#                    % (to_date(today,'yyyy-mm-dd'), record.code, record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga, record.diff ))
#            cnxn.commit()

#        except:
#            print('오늘 : [%s] %s [%s] 거래량: %s, 거래대금 %s 상장주식수 %s 현재가 %s 전일가격: %s, 전일대비 %s'
#                    % (today, self.dicStockCodes.code, self.dicStockCodes.name, self.dicStockCodes.vol, self.dicStockCodes.tradeAek, self.dicStockCodes.listedStock, self.dicStockCodes.cur, self.dicStockCodes.junga, self.dicStockCodes.diff ))
#        try:
#            curs.execute("update oms9jj set eStockname = trim('%s'), eNujucksu = %s, eNujuckaek = %s, eSangJangSu = %s, eHyunJaeGa = %s , eJunilGa = %s where eToDate = '%s' and eStockCode = '%s'"
#                         % (record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga, today, record.code))


#           print('[%s] %s [%s] 현재가: %s, 전일가격: %s, 상장주식수: %s, 시가총액 %s' %(today, code, name, format(cur, ','), format(junga, ','), format(listed, ','), format(markettot, ',')))
#           curs.execute("update oms9jj set eStockname = trim('%s'), eNujucksu = %s, eNujuckaek = %s, eSangJangSu = %s, eHyunJaeGa = %s , eJunilGa = %s where eToDate = '%s' and eStockCode = '%s'" % (name, lvolume, lvalue, listed, cur, junga, today, code))
#           curs.execute("update sjt1tg set close = %s, preclose = %s, volume = %s, value = %s, sangj_jusu = %s, cybos_dt = sysdate where ymd = '%s' and koscom_cd = '%s'" % (cur, junga, lvolume, lvalue, listed, today, code))

#        finally:
#            curs.close()
#            cnxn.close()

#        cnxn.commit()

    def ReqeustElwMst(self):
 
        allcodelist = []
        for i in range(g_objElwMgr.GetCount()) :
            allcodelist.append(g_objElwMgr.GetData(0,i))
 
        print("전체 종목 코드 #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicElwCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicElwCodes)
 
        print("ELW 전종목", len(self.dicElwCodes))
        for key in self.dicElwCodes :
            self.dicElwCodes[key].debugPrint(0)
 
    def ReqeustFutreMst(self):
        allcodelist = []
        for i in range(g_objFutureMgr.GetCount()) :
            allcodelist.append(g_objFutureMgr.GetData(0,i))
 
        print("전체 종목 코드 #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicFutreCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicFutreCodes)
 
        print("선물 전종목 ", len(self.dicFutreCodes))
        for key in self.dicFutreCodes :
            self.dicFutreCodes[key].debugPrint(1)
 
    def ReqeustOptionMst(self):
        allcodelist = []
        for i in range(g_objOptionMgr.GetCount()) :
            allcodelist.append(g_objOptionMgr.GetData(0,i))
 
        print("전체 종목 코드 #", len(allcodelist))
 
 
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                self.obj.Request(rqCodeList, self.dicOptionCodes)
                rqCodeList = []
                continue
 
        if len(rqCodeList) > 0 :
            self.obj.Request(rqCodeList, self.dicOptionCodes)
 
        print("옵션 전종목 ", len(self.dicOptionCodes))
        for key in self.dicOptionCodes :
            self.dicOptionCodes[key].debugPrint(1)
 
    def ReqeustUpjongMst(self):
        codeList = g_objCodeMgr.GetIndustryList() # 증권 산업 업종 리스트
 
        allcodelist = codeList
        print("전체 종목 코드 #", len(allcodelist))
 
 
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
 
        print("증권산업업종 리스트", len(self.dicUpjongCodes))
        for key in self.dicUpjongCodes :
            self.dicUpjongCodes[key].debugPrint(1)
 
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.main = testMain()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 500,  300, 500)
 
        nHeight = 20
        btnStock = QPushButton("주식 전 종목", self)
        btnStock.move(20, nHeight)
        btnStock.clicked.connect(self.btnStock_clicked)
 
        nHeight += 50
        btnElw = QPushButton("ELW 전 종목", self)
        btnElw.move(20, nHeight)
        btnElw.clicked.connect(self.btnElw_clicked)
 
        nHeight += 50
        btnFuture = QPushButton("선물 전 종목", self)
        btnFuture.move(20, nHeight)
        btnFuture.clicked.connect(self.btnFuture_clicked)
 
        nHeight += 50
        btnOption= QPushButton("옵션 전 종목", self)
        btnOption.move(20, nHeight)
        btnOption.clicked.connect(self.btnOption_clicked)
 
        nHeight += 50
        btnUpjong= QPushButton("업종 전 종목", self)
        btnUpjong.move(20, nHeight)
        btnUpjong.clicked.connect(self.btnUpjong_clicked)
 
        nHeight += 50
        btnExit = QPushButton("종료", self)
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

