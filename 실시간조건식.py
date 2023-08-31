import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 300, 200)
        self.setWindowTitle("Kiwoom 실시간 조건식 테스트")

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveConditionVer.connect(self._handler_condition_load)
        self.ocx.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.CommConnect()

        btn1 = QPushButton("condition down")
        btn2 = QPushButton("condition list")
        btn3 = QPushButton("condition send")

        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(btn1)
        layout.addWidget(btn2)
        layout.addWidget(btn3)
        self.setCentralWidget(widget)

        # event
        btn1.clicked.connect(self.GetConditionLoad)
        btn2.clicked.connect(self.GetConditionNameList)
        btn3.clicked.connect(self.send_condition)

        self.stock_codes = []  # 조건을 만족하는 주식 코드들을 저장할 리스트

    def CommConnect(self):
        self.ocx.dynamicCall("CommConnect()")

    def _handler_login(self, err_code):
        print("handler login", err_code)

    def _handler_condition_load(self, ret, msg):
        print("handler condition load", ret, msg)

    def _handler_real_condition(self, code, type, cond_name, cond_index):
        print(cond_name, code, type)

        # 새로운 주식 코드 추가
        if type == "I" and code not in self.stock_codes:
            self.stock_codes.append(code)
            self.print_stock_codes()

        # 주식 코드 삭제
        # elif type == "D" and code in self.stock_codes:
        #     self.stock_codes.remove(code)
        #     self.print_stock_codes()

    def print_stock_codes(self):
        print("========= 현재 조건을 만족하는 주식 코드 =========")
        for code in self.stock_codes:
            print(f"종목코드: {code}")
        print("============================================")

    def GetConditionLoad(self):
        self.ocx.dynamicCall("GetConditionLoad()")

    def GetConditionNameList(self):
        data = self.ocx.dynamicCall("GetConditionNameList()")
        conditions = data.split(";")[:-1]
        for condition in conditions:
            index, name = condition.split('^')
            print(index, name)

    def SendCondition(self, screen, cond_name, cond_index, search):
        ret = self.ocx.dynamicCall("SendCondition(QString, QString, int, int)", screen, cond_name, cond_index, search)

    def SendConditionStop(self, screen, cond_name, cond_index):
        ret = self.ocx.dynamicCall("SendConditionStop(QString, QString, int)", screen, cond_name, cond_index)

    def send_condition(self):
        self.SendCondition("100", "시가갭검색식_돌파", "006", 1)
        print(""""100", "시가갭검색식_돌파", "006", 1""")

    def print_all_stock_codes(self):
        if not self.stock_codes:  # 딕셔너리가 비어있는 경우
            print("조건에 만족하는 종목이 존재하지 않습니다")
        else:
            print("========= 모든 주식 코드 =========")
            for code in self.stock_codes:
                print(f"종목코드: {code}")
            print("=================================")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.aboutToQuit.connect(window.print_all_stock_codes)  # 프로그램 종료 직전에 호출되는 메소드 연결
    app.exec_()

