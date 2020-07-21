import sys
import win32com.client
import ctypes
import time
from dt_alimi import *


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

    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False

    return True


################################################
# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self, acc):
        accFlag = g_objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }

    # 실제적인 6033 통신 처리
    def requestJango(self):
        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(7)
            print(cnt)

            jangoData = {}

            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목코드'] = code
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['현금신용'] = self.dicflag1[self.objRq.GetDataValue(1, i)]  # 신용구분
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                # item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                # item['평가손익'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
                item['현재가'] = 0
                item['대비'] = 0
                item['거래량'] = 0

                # 잔고 추가
                #                key = (code, item['현금신용'],item['대출일'] )
                key = code
                jangoData[key] = item

                if len(jangoData) >= 200:  # 최대 200 종목만,
                    break

            if len(jangoData) >= 200:
                break
            if (self.objRq.Continue == False):
                break
        return jangoData


################################################
# 주식 주문 처리
class CpRPOrder:
    def __init__(self, acc):
        self.acc = acc
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        #print(self.acc, self.accFlag[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 매수

    def buyOrder(self, code, price, amount):
        # 주식 매수 주문
        #print("신규 매수", code, price, amount)
        #self.caller.textBrowser.append("매수 %s %d %d" % (code, price, amount))

        self.objOrder.SetInputValue(0, "2")  # 2: 매수
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 03: 시장가

        # 매수 주문 요청
        ret = self.objOrder.BlockRequest()
        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('주의: 주문 연속 통신 제한에 걸렸음. 대기해서 주문할 지 여부 판단이 필요 남은 시간', remainTime)
            return False

        rqStatus = self.objOrder.GetDibStatus()
        rqRet = self.objOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        return True

    def sellOrder(self, code, amount):
        # 주식 매도 주문
        #print("신규 매도", code, price, amount)
        #self.caller.textBrowser.append("매도 %s %d %d" % (code, price, amount))

        self.objOrder.SetInputValue(0, "1")  # 1: 매도
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        #self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "03")  # 주문호가 구분코드 - 01: 보통 03: 시장가

        # 매도 주문 요청
        ret = self.objOrder.BlockRequest()
        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('주의: 주문 연속 통신 제한에 걸렸음. 대기해서 주문할 지 여부 판단이 필요 남은 시간', remainTime)
            return False

        rqStatus = self.objOrder.GetDibStatus()
        rqRet = self.objOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        return True


if __name__ == "__main__":
    bot.sendMessage(myId, "매도 주문 시작")

    # plus 상태 체크
    if InitPlusCheck() == False:
        exit()

    # 계좌 정의
    account = g_objCpTrade.AccountNumber[0]  # 계좌번호지정 (복수계좌 사인온 실패했을 경우를 위해 기본0번계좌 지정)

    # 6033 잔고 object
    obj6033 = Cp6033(account)
    jangoData = {}

    # 잔고 요청
    jango = obj6033.requestJango(account)

    order = CpRPOrder()

    msg = list()
    msg.append('계좌번호: %s\n' % account)

    for cd, item in jango.items():
        temp = "%s %s %i" % (cd, item['종목명'], item['매도가능'])
        print(temp)
        if order.sellOrder(cd, item['매도가능']):
            print("매도 주문 완료")
            msg.append(temp + '\n')
        time.sleep(1)

    if msg:
        bot.sendMessage(myId, "".join(msg))
