from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errCode import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom class")

        #### event loop 모음
        self.login_event_loop = None
        self.account_info_event_loop = QEventLoop()

        #### 변수
        self.account_num = None
        self.account_password = "0000"
        self.use_money = 0
        self.use_money_pct = 0.1
        self.account_stock_dict = {}
        self.screen_my_info = "2000"
        self.not_book_dict = {}

        ###
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_comm_connect()
        self.get_account_info()
        self.detail_account_info()
        self.account_eval_history()
        self.query_not_book()


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def login_slot(self, errCode):
        print(errors(errCode))
        self.login_event_loop.exit()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr 요청 받는 구역
        :param sScrNo: 화면번호
        :param sRQNamem: 사용자 구분명
        :param sTrCode: TR이름
        :param sRecordName: 레코드 이름
        :param sPrevNext: 연속조회 유무를 판단하는 값 0: 연속(추가조회)데이터 없음, 2:연속(추가조회) 데이터 있음
        :return:
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, 0, "예수금")
            print("예수금 : %s" % deposit)
            self.use_money = int(deposit) * self.use_money_pct
            self.account_info_event_loop.exit()
            
        elif sRQName == "계좌평가잔고내역요청":
            total_pulchase = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, 0, "총수익률(%)")
            print("총수익률 : %s" % total_pulchase)
            
            # 보유 종목 조회. 20개가 넘으면 한번더 요청해야함. sPrevNext = 2면 다음페이지 존재하는거. 0이나 ""이면 다음페이지 없음
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRecordName)
            cnt = 0

            for i in range(rows):
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "종목번호")
                stock_code = stock_code.strip()[1:]

                stock_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "종목명")
                purchase_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "매입가")
                yield_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "수익률(%)")
                present_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "현재가")
                purchase_amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "매입금액")
                available_sale_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "매매가능수량")

                if stock_code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({stock_code: {}})

                stock_name = stock_name.strip()
                purchase_price = int(purchase_price.strip())
                yield_rate = float(yield_rate.strip())
                present_price = int(present_price.strip())
                purchase_amount = int(purchase_amount.strip())
                available_sale_quantity = int(available_sale_quantity.strip())

                query_data = self.account_stock_dict[stock_code]

                query_data.update({"종목명": stock_name})
                query_data.update({"매입가": purchase_price})
                query_data.update({"수익률(%)": yield_rate})
                query_data.update({"현재가": present_price})
                query_data.update({"매입금액": purchase_amount})
                query_data.update({"매매가능수량": available_sale_quantity})

                cnt += 1

            print("가진 종목 : %s" % self.account_stock_dict)

            if sPrevNext == "2":  # 다음 페이지가 존재하는 경우 한번 더 요청
                self.detail_account_info(sPrevNext="2")
            else:
                self.account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            total_pulchase = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                              sTrCode, sRecordName, 0, "총수익률(%)")

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRecordName)
            cnt = 0

            for i in range(rows):
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "종목번호")
                stock_code = stock_code.strip()

                stock_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i,"종목명")
                stock_name = stock_name.strip()

                order_num = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "주문번호")
                order_num = int(order_num.strip())

                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "주문상태")
                order_status = order_status.strip()

                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "주문수량")
                order_quantity = int(order_quantity.strip())

                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "주문가격")
                order_price = int(order_price.strip())

                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "주문구분")
                order_gubun = order_gubun.strip().lstrip(' +').lstrip(' -')

                not_book_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "미체결수량")
                not_book_quantity = int(not_book_quantity.strip())

                book_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRecordName, i, "체결량")
                book_quantity = int(book_quantity.strip())

                if order_num in self.not_book_dict:
                    pass
                else:
                    self.not_book_dict[order_num] = {}

                not_book_data = self.not_book_dict[order_num]

                not_book_data.update({"종목번호": stock_code})
                not_book_data.update({"종목명": stock_name})
                not_book_data.update({"주문번호": order_num})
                not_book_data.update({"주문상태": order_status})
                not_book_data.update({"주문수량": order_quantity})
                not_book_data.update({"주문가격": order_price})
                not_book_data.update({"주문구분": order_gubun})
                not_book_data.update({"미체결수량": not_book_quantity})
                not_book_data.update({"체결량": book_quantity})

            print("미체결 종목 : %s" % self.not_book_dict)
            self.account_info_event_loop.exit()

    def signal_login_comm_connect(self):  # 로그인
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):  # 계좌 번호 조회
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print("my account num %s " % self.account_num)

    def detail_account_info(self):  # 예수금 가져오기
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", self.account_password)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info)

        self.account_info_event_loop.exec_()

    def account_eval_history(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", self.account_password)
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "1")
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", "0", self.screen_my_info)

        self.account_info_event_loop.exec_()

    def query_not_book(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(String, String)", "전체종목구분", "0")
        self.dynamicCall("SetInputValue(String, String)", "매매구분", "0")
        self.dynamicCall("SetInputValue(String, String)", "종목코드", "00")
        self.dynamicCall("SetInputValue(String, String)", "체결구분", "1")
        self.dynamicCall("CommRqData(String, String, int, String)", "실시간미체결요청", "opt10075", "0", self.screen_my_info)

        self.account_info_event_loop.exec_()


#
