import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from PyQt5.QAxContainer import QAxWidget
import pythoncom


class Kiwoom:
    def __init__(self):
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self.on_login)
        self.ocx.OnReceiveRealData.connect(self.on_receive_real_data)
        self.connected = False

    def comm_connect(self):
        self.ocx.dynamicCall("CommConnect()")

    def on_login(self, err_code):
        if err_code == 0:
            print("Login Successful")
            self.connected = True

    def set_real_reg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_no, code_list, fid_list, real_type)

    def on_receive_real_data(self, code, real_type, real_data):
        if real_type == "주식체결":
            time = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, 20)
            price = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, 10)
            price=int(price)
            print(f"Time: {time}, Stock Price ({code}): {abs(price)}")

    def event_loop(self):
        while not self.connected:
            pythoncom.PumpWaitingMessages()

    def run(self):
        self.comm_connect()
        self.event_loop()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()

    kiwoom.run()

    # Register to receive real-time data for Samsung Electronics (삼성전자) with code '005930'
    # and Duksoong Machinery (덕성기업) with code '004830'
    kiwoom.set_real_reg("1001", "005930;004830", "20;10", "0")  # fid 20: 체결시간, fid 10: 현재가

    # Run the application event loop to continuously fetch and display the real-time stock price.
    sys.exit(app.exec_())
