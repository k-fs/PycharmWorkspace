import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

import pyodbc
from datetime import date

################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
 
 
################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False
 
    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False
 
    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False
 
    return True
 
 
class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0
 
 
    def Request(self, codes, dataInfo):
        # 0:종목코드 4:현재가 10:거래량 11:거래대금 20:상장주식수, 23:전일가격
        rqField = [0, 4, 10, 11, 20, 23]  # 요청 필드
 
        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()
 
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        if rqStatus != 0:
            print("통신상태", rqStatus, self.objRq.GetDibMsg1())
            return False
 
        cnt = self.objRq.GetHeaderValue(2)
 
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            cur = self.objRq.GetDataValue(1, i)  # 종가
            lvolume = self.objRq.GetDataValue(2, i)  # 거래량
            lvalue = self.objRq.GetDataValue(3, i)  # 거래대금
            listedStock = self.objRq.GetDataValue(4, i)  # 상장주식수
            junga = self.objRq.GetDataValue(5, i)  # 전일가격
 
            maketAmt = listedStock * cur
            if g_objCodeMgr.IsBigListingStock(code) :
                maketAmt *= 1000
#            print(code, maketAmt)
 
            # key(종목코드) = tuple(거래량, 거래대금, 상장주식수, 현재가, 시가총액, 전일가격)
            dataInfo[code] = (lvolume, lvalue, listedStock, cur, maketAmt, junga)
 
        return True
 
class CMarketTotal():
    def __init__(self):
        self.dataInfo = {}
 
 
    def GetAllMarketTotal(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
#        codeList3 = g_objCodeMgr.GetStockListByMarket(3)  # K-OTC
        codeList4 = g_objCodeMgr.GetStockListByMarket(4)  # KRX
        codeList5 = g_objCodeMgr.GetStockListByMarket(5)  # KONEX
#        allcodelist = codeList + codeList2 + codeList3 + codeList4 + codeList5
        allcodelist = codeList + codeList2 + codeList4 + codeList5
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d, K-OTC %d, KRX %d, KONEX %d' % (len(allcodelist), len(codeList), len(codeList2), len(codeList3), len(codeList4), len(codeList5)))
 
        objMarket = CpMarketEye()
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo)
                rqCodeList = []
                continue
        # end of for
 
        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo)
 
    def PrintMarketTotal(self):

        today = date.today()
        # 시가총액 순으로 소팅
        data2 = sorted(self.dataInfo.items(), key=lambda x: x[1][1], reverse=True)
 
        # Some other example server values are
        server = 'tcp:172.20.200.252,8629' 
        cnxn = pyodbc.connect('DSN=TDB;UID=icam;PWD=kfam1801')
        curs = cnxn.cursor()

        try:
#        print('전종목 시가총액 순 조회 (%d 종목)' % (len(data2)))
            for item in data2:
                code = item[0][1:]
                name = g_objCodeMgr.CodeToName(item[0])
                lvolume = item[1][0]
                lvalue = item[1][1]
                listed = item[1][2]
                cur = item[1][3]
                markettot = item[1][4]
                junga = item[1][5]
                print('[%s] %s [%s] 현재가: %s, 전일가격: %s, 상장주식수: %s, 시가총액 %s' %(today, code, name, format(cur, ','), format(junga, ','), format(listed, ','), format(markettot, ',')))

            curs.execute("insert into oms9jj (eToDate, eStockCode, eStockname, eNujucksu, eNujuckaek, eSangJangSu, eHyunJaeGa, eJunilGa) values (%s, %s, trim(%s), %s, %s, %s, %s, %s)'" % (today, record.code, record.name, record.vol, record.tradeAek, record.listedStock, record.cur, record.junga))

#           curs.execute("update oms9jj set eStockname = trim('%s'), eNujucksu = %s, eNujuckaek = %s, eSangJangSu = %s, eHyunJaeGa = %s , eJunilGa = %s where eToDate = '%s' and eStockCode = '%s'" % (name, lvolume, lvalue, listed, cur, junga, today, code))
#           curs.execute("update sjt1tg set close = %s, preclose = %s, volume = %s, value = %s, sangj_jusu = %s, cybos_dt = sysdate where ymd = '%s' and koscom_cd = '%s'" % (cur, junga, lvolume, lvalue, listed, today, code))

            cnxn.commit()

        except:
#	    cnxn.rollback():
            print ('pyodbc Error [%d]: %s' % (e.args[0], e.args[1]))

        finally:
            curs.close()
            cnxn.close()
            
#            row = cursor.fetchone()    #all()
#            if row:                     #for row in rows:
#                print(row)
            #    print(row.eToDate, eStockCode, row.eSangJangSu) 
 
if __name__ == "__main__":
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()
 
