import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot, Qt, QAbstractTableModel
import win32com.client
import ctypes

# cp object
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
g_objFutureMgr = win32com.client.Dispatch("CpUtil.CpFutureCode")

def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if g_objCpStatus.IsConnect == 0:
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # 주문 관련 초기화
    ret = g_objCpTrade.TradeInit(0)
    if ret != 0:
        print("주문 초기화 실패, 오류번호 ", ret)
        return False
 
    return True

# 현재가 정보 저장 구조체
class stockPricedData:
    def __init__(self):
        self.dicEx = {10: "시가단일가", 11: "시가단일가연장", 20: "장중단일가", 21:"장중단일가연장", 30:"종가단일가", 40:"장중"}
        self.code = ""
        self.name = ""
        self.cur = 0        # 현재가
        self.diff = 0       # 대비
        self.diffp = 0      # 대비율
        self.offer = [0 for _ in range(5)]     # 매도호가
        self.bid = [0 for _ in range(5)]       # 매수호가
        self.offervol = [0 for _ in range(5)]     # 매도호가 잔량
        self.bidvol = [0 for _ in range(5)]       # 매수호가 잔량
        self.totOffer = 0       # 총매도잔량
        self.totBid = 0         # 총매수 잔량
        self.vol = 0            # 거래량
        self.baseprice = 0      # 기준가

        # 예상체결가 정보
        self.exFlag= 40
        self.expcur = 0         # 예상체결가
        self.expdiff = 0        # 예상 대비
        self.expdiffp = 0       # 예상 대비율
        self.expvol = 0         # 예상 거래량
        self.objCur = CpPBStockCur()
        self.objOfferbid = CpPBStockBid()

    def __del__(self):
        self.objCur.Unsubscribe()
        self.objOfferbid.Unsubscribe()


    # 전일 대비 계산
    def makediffp(self):
        lastday = 0
        if (self.exFlag != 40):  # 동시호가 시간 (예상체결)
            if self.baseprice > 0  :
                lastday = self.baseprice
            else:
                lastday = self.expcur - self.expdiff
            if lastday:
                self.expdiffp = (self.expdiff / lastday) * 100
            else:
                self.expdiffp = 0
        else:
            if self.baseprice > 0  :
                lastday = self.baseprice
            else:
                lastday = self.cur - self.diff
            if lastday:
                self.diffp = (self.diff / lastday) * 100
            else:
                self.diffp = 0

    def getCurColor(self):
        diff = self.diff
        if (self.exFlag != 40):  # 동시호가 시간 (예상체결)
            diff = self.expdiff
        if (diff > 0):
            return 'color: red'
        elif (diff == 0):
            return  'color: black'
        elif (diff < 0):
            return 'color: blue'


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    '''
    CpEvent Coding
    '''
    def set_params(self, client, name, rpMst, parent):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.parent = parent  # callback 을 위해 보관
        self.rpMst = rpMst

    def OnReceived(self):
        '''
        PLUS로 부터 실제로 시세를 수신받는 이벤트 핸들러
        '''
        if self.name == "futurecur":
            # 현재가 체결 데이터 실시간 업데이트
            self.rpMst.exFlag = self.client.GetHeaderValue(28)  # 예상체결 플래그
            code = self.client.GetHeaderValue(0)
            diff = self.client.GetHeaderValue(2)
            cur= self.client.GetHeaderValue(1)  # 현재가
            vol = self.client.GetHeaderValue(13)  # 거래량

            # 예제는 장중만 처리 함.
            if (self.rpMst.exFlag != 40):  # 동시호가 시간 (예상체결)
                # 예상체결가 정보
                self.rpMst.expcur = cur
                self.rpMst.expdiff = diff
                self.rpMst.expvol = vol
            else:
                self.rpMst.cur = cur
                self.rpMst.diff = diff
                self.rpMst.makediffp()
                self.vol = vol

            self.rpMst.makediffp()
            # 현재가 업데이트
            self.parent.monitorPriceChange()

            return

        elif self.name == "futurebid":
            # 현재가 10차 호가 데이터 실시간 업데이c
            code = self.client.GetHeaderValue(0)
            dataindex = [2, 7, 11, 15, 19, 27, 31, 35, 39, 43]
            obi = 0
            for i in range(5):
                self.rpMst.offer[i] = self.client.GetHeaderValue(i+2)
                self.rpMst.bid[i] = self.client.GetHeaderValue(i+19)
                self.rpMst.offervol[i] = self.client.GetHeaderValue(i+7)
                self.rpMst.bidvol[i] = self.client.GetHeaderValue(i+24)

            self.rpMst.totOffer = self.client.GetHeaderValue(12)
            self.rpMst.totBid = self.client.GetHeaderValue(29)
            # 10차 호가 변경 call back 함수 호출
            self.parent.monitorOfferbidChange()
            return

# SB/PB 요청 ROOT 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, rpMst, parent):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, rpMst, parent)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False

class CpPBStockCur(CpPublish):
    '''
    CpPBStockCur: 실시간 현재가 요청 클래스
    '''
    def __init__(self):
        super().__init__("futurecur", "DsCbo1.FutureCurOnly")


class CpPBStockBid(CpPublish):
    '''
    CpPBStockBid: 실시간 10차 호가 요청 클래스
    '''
    def __init__(self):
        super().__init__("futurebid", "CpSysDib.FutureJpBid")



class CpPBConnection:
    '''
    SB/PB 요청 ROOT 클래스
    '''
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCybos")
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, "connection", None)


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objStockMst = win32com.client.Dispatch("Dscbo1.FutureMst")
        return


    def Request(self, code, rtMst, callbackobj):
        # 현재가 통신
        rtMst.objCur.Unsubscribe()
        rtMst.objOfferbid.Unsubscribe()

        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False


        # 수신 받은 현재가 정보를 rtMst 에 저장
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur = self.objStockMst.GetHeaderValue(71)  # 현재가 11
        rtMst.diff = self.objStockMst.GetHeaderValue(77)  # 전일대비 12
        rtMst.baseprice = self.objStockMst.GetHeaderValue(13)  # 기준가 27
        rtMst.vol = self.objStockMst.GetHeaderValue(75)  # 거래량 18
        rtMst.exFlag = self.objStockMst.GetHeaderValue(115)  # 예상플래그 58
        rtMst.expcur = self.objStockMst.GetHeaderValue(113)  # 예상체결가 55
        rtMst.expdiff = self.objStockMst.GetHeaderValue(114)  # 예상체결대비 56
        rtMst.makediffp()

        rtMst.totOffer = self.objStockMst.GetHeaderValue(47)  # 총매도잔량 71
        rtMst.totBid = self.objStockMst.GetHeaderValue(64)  # 총매수잔량 73


        # 10차호가
        for i in range(5):
            rtMst.offer[i] = (self.objStockMst.GetHeaderValue(i+37))  # 매도호가
            rtMst.bid[i] = (self.objStockMst.GetHeaderValue(i+54) ) # 매수호가
            rtMst.offervol[i] = (self.objStockMst.GetHeaderValue(i+42))  # 매도호가 잔량
            rtMst.bidvol[i] = (self.objStockMst.GetHeaderValue(i+59) ) # 매수호가 잔량


        rtMst.objCur.Subscribe(code,rtMst, callbackobj)
        rtMst.objOfferbid.Subscribe(code,rtMst, callbackobj)

# CpFutureBalance: 선물 잔고
class CpFutureBalance:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd0723")
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 2)  # 선물/옵션 계좌구분
        print(self.acc, self.accFlag[0])
 
    def request(self,  retList):
        self.objRq.SetInputValue(0, self.acc)
        self.objRq.SetInputValue(1, self.accFlag[0])
        self.objRq.SetInputValue(4, 50)
 
 
        while True:
            self.objRq.BlockRequest()
 
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
 
            if rqStatus != 0:
                print("통신상태", rqStatus, rqRet)
                return False
 
            cnt = self.objRq.GetHeaderValue(2)
 
            for i in range(cnt):
                item = []
                item.append(self.objRq.GetDataValue(0, i))
                item.append(self.objRq.GetDataValue(1, i))
                flag = self.objRq.GetDataValue(2, i)
                if flag == '1':
                    item.append('매도')
                elif flag == '2':
                    item.append('매수')
 
                item.append(self.objRq.GetDataValue(3, i))
                item.append(self.objRq.GetDataValue(5, i))
                item.append(self.objRq.GetDataValue(9, i))
 
                retList.append(item)
            # end of for
 
            if self.objRq.Continue == False :
                break
        # end of while
 
        '''
        for item in  retList:
            data = ''
            for key, value in item.items():
                if (type(value) == float):
                    data += '%s:%.2f' % (key, value)
                elif (type(value) == str):
                    data += '%s:%s' % (key, value)
                elif (type(value) == int):
                    data += '%s:%d' % (key, value)
 
                data += ' '
            print(data)
        return True
        '''
 
# CpFutureNContract: 선물 미체결 조회
class CpFutureNContract:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd5371")
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 2)  # 선물/옵션 계좌구분
        print(self.acc, self.accFlag[0])
 
    def request(self,  retList):
        self.objRq.SetInputValue(0, self.acc)
        self.objRq.SetInputValue(1, self.accFlag[0])
        self.objRq.SetInputValue(6, '3') # '3' : 미체결
 
 
        while True:
            self.objRq.BlockRequest()
 
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            if rqStatus != 0:
                print("통신상태", rqStatus, rqRet)
                return False
 
            cnt = self.objRq.GetHeaderValue(6)
 
            for i in range(cnt):
                item = {}
                item['주문번호'] = self.objRq.GetDataValue(2, i)
                item['코드'] = self.objRq.GetDataValue(4, i)
                item['종목명'] = self.objRq.GetDataValue(5, i)
                item['주문가격'] = self.objRq.GetDataValue(8, i)
                item['잔량'] = self.objRq.GetDataValue(9, i)
                item['거래구분']= self.objRq.GetDataValue(6, i)
 
                retList.append(item)
            # end of for
 
            if self.objRq.Continue == False :
                break
        # end of while
 
 
        for item in  retList:
            data = ''
            for key, value in item.items():
                if (type(value) == float):
                    data += '%s:%.2f' % (key, value)
                elif (type(value) == str):
                    data += '%s:%s' % (key, value)
                elif (type(value) == int):
                    data += '%s:%d' % (key, value)
 
                data += ' '
            print(data)
        return True



#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
#form_class = uic.loadUiType("pytrainer.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow):
    def __init__(self) :
        super().__init__()
        self.ui = uic.loadUi("pytrainer.ui", self)

        self.fcodelist = []
 
        for i in range(g_objFutureMgr.GetCount()):
            code = g_objFutureMgr.GetData(0, i)
            name = g_objFutureMgr.GetData(1, i)
            if (code[0] == '4') :   # spread skip
                continue
            if (code[0] == '10100') : #연결선물 skip
                continue
            self.fcodelist.append((code, name))
        self.fcodelist.append(('165Q6', 'KTBF'))
        self.fcodelist.append(('167Q6', 'LKTBF'))

        # self.comboCodeList = QComboBox(self)
        for code, name in self.fcodelist :
            self.comboCodeList.addItem(code)
 
        self.comboCodeList.currentIndexChanged.connect(self.OnComboChanged)

        self.btnQuote.clicked.connect(self.btnQuote_Clicked)
        self.btnRefresh.clicked.connect(self.btnRefresh_Clicked)

        self.btnBuy0.clicked.connect(self.btnBuy0_Clicked)
        self.btnBuy1.clicked.connect(self.btnBuy1_Clicked)
        self.btnBuy2.clicked.connect(self.btnBuy2_Clicked)
        self.btnBuy3.clicked.connect(self.btnBuy3_Clicked)

        self.btnSell0.clicked.connect(self.btnSell0_Clicked)
        self.btnSell1.clicked.connect(self.btnSell1_Clicked)
        self.btnSell2.clicked.connect(self.btnSell2_Clicked)
        self.btnSell3.clicked.connect(self.btnSell3_Clicked)

        self.objMst = CpRPCurrentPrice()
        self.objBal = CpFutureBalance()

        self.item = stockPricedData()
        self.setCode("101Q6")

    def btnQuote_Clicked(self):
        code = self.comboCodeList.currentText()
        self.setCode(code)

    def OnComboChanged(self):
        code = self.comboCodeList.currentText()
        self.setCode(code)
    
    def btnRefresh_Clicked(self):
        retList = []
        self.objBal.request(retList)

        item_count = len(retList)
        self.tableBalance.setRowCount(item_count)
        for i in range(item_count):
            row = retList[i]
            for j in range(len(row)):
                wgtItem = QTableWidgetItem(str(row[j]))
                wgtItem.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableBalance.setItem(i, j, wgtItem)
        self.tableBalance.resizeColumnsToContents()

    def btnBuy0_Clicked(self):
        # 현재가 조회 > 매수 1주 주문
        code = self.comboCodeList.currentText()
        objFutureMst = CpFutureMst()
        retItem = {}
        objFutureMst.request(code, retItem)
 
        objOrder = CpFutureOrder()
        price = retItem['현재가']
        retOrder = {}
        objOrder.buyOrder(code, price, 1, retOrder)
 
        print(retOrder)

    def btnBuy1_Clicked(self):
        pass

    def btnBuy2_Clicked(self):
        pass

    def btnBuy3_Clicked(self):
        pass

    def btnSell0_Clicked(self):
        pass

    def btnSell1_Clicked(self):
        pass

    def btnSell2_Clicked(self):
        pass

    def btnSell3_Clicked(self):
        pass

    def monitorPriceChange(self):
        self.displyHoga()

    def monitorOfferbidChange(self):
        self.displyHoga()

    def setCode(self, code):
        if len(code) < 5 :
            return

        print(code)
        #if not (code[0] == "A"):
        #    code = "A" + code

        name = g_objCodeMgr.CodeToName(code)
        if len(name) == 0:
            print("종목코드 확인")
            return

        self.ui.label_name.setText(name)

        if (self.objMst.Request(code, self.item, self) == False):
            return
        self.displyHoga()

    def displyHoga(self):
        self.ui.label_offer5.setText(format(self.item.offer[4],'03.2f'))
        self.ui.label_offer4.setText(format(self.item.offer[3],'03.2f'))
        self.ui.label_offer3.setText(format(self.item.offer[2],'03.2f'))
        self.ui.label_offer2.setText(format(self.item.offer[1],'03.2f'))
        self.ui.label_offer1.setText(format(self.item.offer[0],'03.2f'))

        self.ui.label_offer_v5.setText(format(self.item.offervol[4],','))
        self.ui.label_offer_v4.setText(format(self.item.offervol[3],','))
        self.ui.label_offer_v3.setText(format(self.item.offervol[2],','))
        self.ui.label_offer_v2.setText(format(self.item.offervol[1],','))
        self.ui.label_offer_v1.setText(format(self.item.offervol[0],','))

        self.ui.label_bid5.setText(format(self.item.bid[4],'03.2f'))
        self.ui.label_bid4.setText(format(self.item.bid[3],'03.2f'))
        self.ui.label_bid3.setText(format(self.item.bid[2],'03.2f'))
        self.ui.label_bid2.setText(format(self.item.bid[1],'03.2f'))
        self.ui.label_bid1.setText(format(self.item.bid[0],'03.2f'))

        self.ui.label_bid_v5.setText(format(self.item.bidvol[4], ','))
        self.ui.label_bid_v4.setText(format(self.item.bidvol[3], ','))
        self.ui.label_bid_v3.setText(format(self.item.bidvol[2], ','))
        self.ui.label_bid_v2.setText(format(self.item.bidvol[1], ','))
        self.ui.label_bid_v1.setText(format(self.item.bidvol[0], ','))

        cur = self.item.cur
        diff = self.item.diff
        diffp = self.item.diffp
        if (self.item.exFlag != 40):  # 동시호가 시간 (예상체결)
            cur = self.item.expcur
            diff = self.item.expdiff
            diffp = self.item.expdiffp


        strcur = format(cur, '03.2f')
        if (self.item.exFlag != 40):  # 동시호가 시간 (예상체결)
            strcur = "*" + strcur

        curcolor = self.item.getCurColor()
        self.ui.label_cur.setStyleSheet(curcolor)
        self.ui.label_cur.setText(strcur)
        strdiff = format(diff, '03.2f') + "  " + format(diffp, '.2f')
        strdiff += "%"
        self.ui.label_diff.setText(strdiff)
        self.ui.label_diff.setStyleSheet(curcolor)

        self.ui.label_totoffer.setText(format(self.item.totOffer, ','))
        self.ui.label_totbid.setText(format(self.item.totBid, ','))

if __name__ == "__main__":
    if False == InitPlusCheck() :
        exit()
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()