import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
from pandas import DataFrame
import pandas as pd
import sqlite3

TR_REQ_TIME_INTERVAL = 0.2

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.orderNum = ""

    # openAPI 사용을 위해선 COM 오브젝트 생성이 필요
    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    # 서버로부터 발생한 이벤트(signal)와 처리할 메서드(slot)을 연결하는 메서드
    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self._receive_tr_data)
        #self.OnReceiveRealData.connect(self._receive_real_data)
        self.OnReceiveChejanData.connect(self._receive_chejan_data)

    # 로그인 함수
    def comm_connect(self):
        # CommConnect 함수를 dynamicCall 함수를 통해 호출
        self.dynamicCall("CommConnect()")
        # 이벤트 루프(이벤트가 발생할 때까지 프로그램 종료되지 않음) / GUI형태로 만들지 않을 시 필요
        self.login_event_loop = QEventLoop()
        # exec_메서드를 호출하여 생성
        self.login_event_loop.exec_()

    # 로그인 요청을 받으면 OnEventConnect가 발생
    # 연결 상태 확인 후 이벤트 루프 종료
    def _event_connect(self, err_code):
        if err_code == 0:
            print('connected')
        else:
            print('disconnected')

        self.login_event_loop.exit()

    # GetCodeListByMarket 메서드 호출
    # 종목 코드를 파이썬 리스트로 변환
    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')
        return code_list[:-1]

    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_name

    def get_connect_state(self):
        ret = self.dynamicCall("GetConnectState()")
        return ret

    def get_login_info(self, tag):
        ret = self.dynamicCall("GetLoginInfo(QString)", tag)
        return ret

    def set_input_value(self, id, value):
        self.dynamicCall("setInputValue(QString, QString)", id, value)

    # TR을 요청하면 데이터가 바로 반환되지 않음. 이벤트를 통해 알려주기 때문에
    # 이벤트가 올 때까지 대기 상태여야 해서 루프 코드가 있어야 함
    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    # TR데이터 가져오기, 반환값에 문자열 양쪽에 공백이 있기 때문에 strip으로 공백 제거
    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)", code,
                               real_type, field_name, index, item_name)
        return ret.strip()

    # 총 몇 개의 데이터가 왔는지 알기 위한 것
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def send_order(self, rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no):
        self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                         [rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no])

    def get_chejan_data(self, fid):
        ret = self.dynamicCall("GetChejanData(int)", fid)
        return ret

    def get_server_gubun(self):
        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret

    def _receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(gubun)
        # 주문번호
        print(self.get_chejan_data(9203))
        # 종목명
        print(self.get_chejan_data(302))
        # 주문수량
        print(self.get_chejan_data(900))
        # 주문가격
        print(self.get_chejan_data(901))

    # 이벤트 발생 시 처리하는 메서드
    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        self.orderNum = self._comm_get_data(trcode, "", rqname, 0, "주문번호")
        # 연속조회가 필요한 경우 2로 반환함
        if next == '2':
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == "opt10081_req":
            self._opt10081(rqname, trcode)

        elif rqname == "opw00001_req":
            self._opw00001(rqname, trcode)

        elif rqname == "opw00018_req":
            self._opw00018(rqname, trcode)

        # 필요하지 않은 이벤트 루프 종료
        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    """def _receive_real_data(self, scode, realtype, realdata):
        if realtype == "주식체결":
            self.item[scode] = {}
            self.scode = scode
            self.time =int(self.GetCommGetReal(scode, 20))
            self.price_up_down = self.GetCommRealData(scode, 12)
            self.volume = int(self.GetCommRealData(scode, 15))
            self.hour = self.time / 10000
            self.min = self.time / 100 % 100
            self.sec = self.time % 100
            self.realtime = self.hour * 3600 + self.min * 60 + self.sec
            self.item[scode].append([self.time, self.price_up_down, self.volume])"""


    @staticmethod
    def change_format(data):
        strip_data = data.lstrip('-0')
        if strip_data == '' or strip_data == '.00':
            strip_data = '0'

        format_data = format(int(strip_data), ',d')
        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    @staticmethod
    def change_format2(data):
        strip_data = data.lstrip('-0')

        if strip_data == '':
            strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    # OnReceiveTrData 이벤트 발생 시 수신 데이터를 가져옴
    def _opw00001(self, rqname, trcode):
        d2_deposit = self._comm_get_data(trcode, "", rqname, 0, "d+2추정예수금")
        self.d2_deposit = Kiwoom.change_format(d2_deposit)

    def _opt10081(self, rqname, trcode):
        # 데이터를 얻어오기 전에 데이터의 개수를 get
        data_cnt = self._get_repeat_cnt(trcode, rqname)
        for i in range(data_cnt):
            date = self._comm_get_data(trcode, "", rqname, i, "일자")
            open = self._comm_get_data(trcode, "", rqname, i, "시가")
            high = self._comm_get_data(trcode, "", rqname, i, "고가")
            low = self._comm_get_data(trcode, "", rqname, i, "저가")
            close = self._comm_get_data(trcode, "", rqname, i, "현재가")
            volume = self._comm_get_data(trcode, "", rqname, i, "거래량")
            #time.sleep(0.2)

            self.ohlcv['date'].append(date)
            #self.ohlcv['open'].append(int(open))
            #self.ohlcv['high'].append(int(high))
            #self.ohlcv['low'].append(int(low))
            self.ohlcv['close'].append(int(close))
            #self.ohlcv['volume'].append(int(volume))
        print("날짜: ", self.ohlcv['date'][1])
        print("종가: ", self.ohlcv['close'][1])
        self.final['close'].append(self.ohlcv['close'][1])
        self.current['current'].append(self.ohlcv['close'][0])

        #self.final_close.append(self.ohlcv['close'][1])
            #print("date   open   high   low    close    volume")
            #print(date, open, high, low, close, volume)

    def _opw00018(self, rqname, trcode):
        total_purchase_price = self._comm_get_data(trcode, "", rqname, 0, "총매입금액")
        total_eval_price = self._comm_get_data(trcode, "", rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, 0, "총평가손익금액")
        total_earning_rate = self._comm_get_data(trcode, "", rqname, 0, "총수익률(%)")
        estimated_deposit = self._comm_get_data(trcode, "", rqname, 0, "추정예탁자산")

        self.opw00018_output['single'].append(Kiwoom.change_format(total_purchase_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_profit_loss_price))

        total_earning_rate = Kiwoom.change_format2(total_earning_rate)

        if self.get_server_gubun():
            total_earning_rate = float(total_earning_rate) / 100
            total_earning_rate = str(total_earning_rate)

        self.opw00018_output['single'].append(total_earning_rate)
        self.opw00018_output['single'].append(Kiwoom.change_format(estimated_deposit))


        #multi data
        rows = self._get_repeat_cnt(trcode, rqname)
        for i in range(rows):
            name = self._comm_get_data(trcode, "", rqname, i, "종목명")
            quantity = self._comm_get_data(trcode, "", rqname, i, "보유수량")
            purchase_price = self._comm_get_data(trcode, "", rqname, i, "매입가")
            current_price = self._comm_get_data(trcode, "", rqname, i, "현재가")
            eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, i, "평가손익")
            earning_rate = self._comm_get_data(trcode, "", rqname, i, "수익률(%)")

            quantity = Kiwoom.change_format(quantity)
            purchase_price = Kiwoom.change_format(purchase_price)
            current_price = Kiwoom.change_format(current_price)
            eval_profit_loss_price = Kiwoom.change_format(eval_profit_loss_price)
            earning_rate = Kiwoom.change_format2(earning_rate)

            self.opw00018_output['multi'].append([name, quantity, purchase_price, current_price,
                                                  eval_profit_loss_price, earning_rate])

    def reset_opw00018_output(self):
        self.opw00018_output = {'single': [], 'multi': []}

    def store_fianl_close(self):
        self.final_close = []
        self.current_close = []


if __name__ == "__main__":
    # Kiwoom클래스는 QAxWidget클래스를 상속 받아서 Kiwoom 클래스에 대한 인스턴스를 생성하려면
    # QApplication 클래스의 인스턴스를 생성해야 함
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    # 로그인 수행
    kiwoom.comm_connect()

    kiwoom.reset_opw00018_output()
    account_number = kiwoom.get_login_info("ACCNO")
    account_number = account_number.split(';')[0]

    kiwoom.set_input_value("계좌번호", account_number)
    kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")
    print(kiwoom.opw00018_output['single'])
    print(kiwoom.opw00018_output['multi'])
    print(kiwoom.opw00018_output['multi'][0][3])