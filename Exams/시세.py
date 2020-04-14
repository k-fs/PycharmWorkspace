import win32com.client
import ctypes

################################################
# PLUS 공통 OBJECT
#g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
#g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
#g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

#explore = win32com.client.Dispatch("InternetExplorer.Application")
#explore.Visible = True
################################################

objRq=win32com.client.Dispatch("Dscbo1.StockMst2")
codeList = 'A003540,A000660'
objRq.SetInputValue(0, codeList)
objRq.BlockRequest()
cnt = objRq.GetHeaderValue(0)

for i in range(cnt) :
    cur = objRq.GetDataValue(3, i)
    stocks = objRq.GetDataValue(17,i)
    print(cur, stocks)
