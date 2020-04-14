# find valuable items
import win32com.client
 
# declare Cybosplus API
cpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
cpStockMst = win32com.client.Dispatch("dscbo1.StockMst") # not CpDib.StockMst
 
# total num of stock items
itNum = cpStockCode.GetCount()
print(itNum)
 
for it in range(0, itNum):
 
   # get item info
   itCode = cpStockCode.GetData(0, it)
   itName = cpStockCode.GetData(1, it)
 
   # request info
   cpStockMst.SetInputValue(0, itCode)
   cpStockMst.BlockRequest()
 
   # get header values
   managedStatus = chr(cpStockMst.GetHeaderValue(66)) # 관리구분
   alarmInvestStatus = chr(cpStockMst.GetHeaderValue(67)) # 투자 경고 구분
   tradeStopStatus = chr(cpStockMst.GetHeaderValue(68)) # 거래 정지 구분
   unfaithfulNoticeStatus = chr(cpStockMst.GetHeaderValue(69)) # 불성실 공시 구분
   exchangeType = chr(cpStockMst.GetHeaderValue(45)) # 소속구분 '1' : 거래소, '5': 코스닥
 
   # filtering items
   if managedStatus == 'N' and alarmInvestStatus == '1' and \
         tradeStopStatus == 'N' and unfaithfulNoticeStatus == '0':
 
      # only to consider KOSPI or KOSDAQ
      if exchangeType == '1' or exchangeType == '5':
 
         currentPrice = int(cpStockMst.GetHeaderValue(11)) # current price (현재가)
         stockNum = int(cpStockMst.GetHeaderValue(31)) # 발행주식수
         marketCap = currentPrice * stockNum
         EPS = float(cpStockMst.GetHeaderValue(20)) # EPS
         BPS = float(cpStockMst.GetHeaderValue(70)) # BPS
         PER = float(cpStockMst.GetHeaderValue(28)) # PER
 
         if BPS != 0:
            PBR = float(currentPrice/BPS) # PBR calculation
         else:
            PBR = float('inf')
 
         # calc essencial value (본질 가치)
         assetValue = BPS
         profitValue = EPS / 0.01 # 10% 할인률 적용
         eValue = (assetValue * 2 + profitValue * 3) / 5
 
         if marketCap > 200000000000 and PER != 0 and 1 / PER > 0.12 and PBR < 1.0:
            if eValue > currentPrice:
               print(itName + "<" + itCode + ">")
