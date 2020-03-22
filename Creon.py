import win32com.client
import pandas as pd


class Creon():
    def __init__(self):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.Isconnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음.")
            exit()

    def requestCode(self, market):
        # 코드 데이터 요청
        instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        codelist = instCpCodeMgr.GetStockListByMarket(market)  # 1 코스피, 2 코스닥
        return codelist

    def requestData(self, code, startdate, enddate):
        # Create object
        chart = win32com.client.Dispatch("CpSysDib.StockChart")

        # SetInputValue
        chart.SetInputValue(0, code) #종목코드
        chart.SetInputValue(1, ord('1')) # 요청구분(개수)
        chart.SetInputValue(2, enddate) # 종료일시
        chart.SetInputValue(3, startdate) #시작일시
        chart.SetInputValue(5, (0,2,3,4,5,8)) #날짜,시가,고가,저가,종가,거래량
        chart.SetInputValue(6, ord('D')) #일 단위

        # BlockRequest
        chart.BlockRequest()

        # 통신 결과 확인
        rqStatus = chart.GetDibStatus()
        rqRet = chart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # GetHeaderValue
        numData = chart.GetHeaderValue(3)
        resultDf = pd.DataFrame(columns=("Open", "High", "Low", "Close", "Volume"))
        for i in range(numData):
            date = chart.GetDataValue(0, i)
            openPR = chart.GetDataValue(1, i)
            highPR = chart.GetDataValue(2, i)
            lowPR = chart.GetDataValue(3, i)
            closePR = chart.GetDataValue(4, i)
            volume = chart.GetDataValue(5, i)

            # 데이터 프레임으로 변환
            resultDf = resultDf.append(pd.DataFrame({'Open': openPR, 'High': highPR, 'Low': lowPR, 'Close': closePR, 'Volume': volume}, index=[date]))
        return resultDf

    def requestData2(self, code, numHist):
        # Create object
        chart = win32com.client.Dispatch("CpSysDib.StockChart")

        # SetInputValue
        chart.SetInputValue(0, code)  #종목코드
        chart.SetInputValue(1, ord('2'))  # 요청구분(개수)
        chart.SetInputValue(4, numHist)  #요청개수
        chart.SetInputValue(5, 13)  # 시가총액
        chart.SetInputValue(6, ord('D'))  #일 단위

        # BlockRequest
        chart.BlockRequest()

        # 통신 결과 확인
        rqStatus = chart.GetDibStatus()
        rqRet = chart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # GetHeaderValue
        numData = chart.GetHeaderValue(3)
        for i in range(numData):
            a=chart.GetDataValue(0, i)
            print( a / 1e8)


    def getOpen(self, code):
        # Create object
        chart = win32com.client.Dispatch("CpSysDib.StockChart")

        # SetInputValue
        chart.SetInputValue(0, code) #종목코드
        chart.SetInputValue(1, ord('2')) # 요청구분(개수)
        chart.SetInputValue(4, 1) #요청개수
        chart.SetInputValue(5, 2) #시가
        chart.SetInputValue(6, ord('D')) #일 단위

        # BlockRequest
        chart.BlockRequest()

        # 통신 결과 확인
        rqStatus = chart.GetDibStatus()
        rqRet = chart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # GetHeaderValue
        numData = chart.GetHeaderValue(3)
        for i in range(numData):
            openPR = chart.GetDataValue(0, i)

        return openPR

    def getName(self, code):
        instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        name = instCpCodeMgr.CodeToName(code)

        return name