import datetime
import sys

import requests
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QFont, QPalette
from PyQt5.QtWidgets import QApplication, QWidget, QToolTip, QPushButton, QMessageBox, QMainWindow, QLineEdit, \
    QInputDialog, QLabel


def send_post(touser, message):
    send_url = 'http://10.10.100.126:5001/testsendmsg/'+touser
    respone = requests.post(send_url, data={'msg': message})

def show_w():
    '显示窗口'

    app = QApplication(sys.argv)  # 所有的PyQt5应用必须创建一个应用（Application）对象。
    # sys.argv参数是一个来自命令行的参数列表。

    w = QWidget()  # Qwidget组件是PyQt5中所有用户界面类的基础类。我们给QWidget提供了默认的构造方法。
    # 默认构造方法没有父类。没有父类的widget组件将被作为窗口使用。

    w.resize(500, 500)  # resize()方法调整了widget组件的大小。它现在是500px宽，500px高。
    w.move(500, 100)  # move()方法移动widget组件到一个位置，这个位置是屏幕上x=500,y=200的坐标。
    w.setWindowTitle('Simple')  # 设置了窗口的标题。这个标题显示在标题栏中。
    w.show()  # show()方法在屏幕上显示出widget。一个widget对象在这里第一次被在内存中创建，并且之后在屏幕上显示。

    sys.exit(app.exec_())  # 应用进入主循环。在这个地方，事件处理开始执行。主循环用于接收来自窗口触发的事件，
    # 并且转发他们到widget应用上处理。如果我们调用exit()方法或主widget组件被销毁，主循环将退出。
    # sys.exit()方法确保一个不留垃圾的退出。系统环境将会被通知应用是怎样被结束的。


# ##***提示文本***## #
class PromptText(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        QToolTip.setFont(QFont('SansSerif', 10))  # 这个静态方法设置了用于提示框的字体。
        # 这里使用10px大小的SansSerif字体。
        self.setToolTip('This is a <b>QWidget</b> widget')  # 调用setTooltip()方法创建提示框。
        # 可以在提示框中使用富文本格式。
        btn = QPushButton('Button', self)  # 创建按钮
        btn.setToolTip('This is a <b>QPushButton</b> widget')  # 设置按钮提示框
        btn.resize(btn.sizeHint())  # 改变按钮大小
        btn.move(300, 100)  # 移动按钮位置
        self.setGeometry(300, 100, 600, 600)
        self.setWindowTitle('Tooltips')
        self.show()


class CloseW(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        qbtn = QPushButton('Quit', self)  # 创建了一个按钮。按钮是一个QPushButton类的实例。
        # 构造方法的第一个参数是显示在button上的标签文本。第二个参数是父组件。
        # 父组件是Example组件，它继承了QWiget类。
        qbtn.clicked.connect(QCoreApplication.instance().quit)
        qbtn.resize(qbtn.sizeHint())
        qbtn.move(500, 50)
        self.setGeometry(300, 100, 600, 600)
        self.setWindowTitle('excise')
        self.show()


class MessageBox(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        qbtn = QPushButton('Quit', self)  # 创建了一个按钮。按钮是一个QPushButton类的实例。
        # 构造方法的第一个参数是显示在button上的标签文本。第二个参数是父组件。
        # 父组件是Example组件，它继承了QWiget类。
        qbtn.clicked.connect(QCoreApplication.instance().quit)
        qbtn.resize(qbtn.sizeHint())
        qbtn.move(500, 50)
        self.setGeometry(300, 100, 600, 600)
        self.setWindowTitle('excise')
        self.show()

    def closeEvent(self, event):

        reply = QMessageBox.question(self, 'Message', "Are you sure to quit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


# ##***状态栏***## #
class StatusBar(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()



    def initUI(self):
        self.statusBar().showMessage('Ready')

        self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Statusbar')
        self.show()


class EventSender(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

        self.name = ''

    def initUI(self):
        btn1 = QPushButton('1', self)
        btn1.move(30, 50)

        btn2 = QPushButton("Button 2", self)
        btn2.move(150, 50)

        btn1.clicked.connect(self.buttonClicked)
        btn2.clicked.connect(self.buttonClicked)

        print(self.statusBar())

        # self.setGeometry(300, 300, 290, 150)
        # self.setWindowTitle('Event sender')
        #
        # self.show()

        self.btn = QPushButton('分享', self)

        self.btn.move(20, 20)
        self.btn.clicked.connect(self.showDialog)

        self.le = QLineEdit(self)
        self.le.move(130, 22)



    def buttonClicked(self):
        sender = self.sender()
        self.statusBar().showMessage(sender.text() + ' was pressed')
        self.name = sender.text()
        print(self.name)

        # if self.name == '1':
        #     self.btn = QPushButton('分享', self)
        #     self.btn.move(20, 20)
        #     self.btn.clicked.connect(self.showDialog)
        #
        #     self.le = QLineEdit(self)
        #     self.le.move(130, 22)


    def showDialog(self):
        # 这一行会显示一个输入对话框。第一个字符串参数是对话框的标题，第二个字符串参数是对话框内的消息文本。
        # 对话框返回输入的文本内容和一个布尔值。如果我们点击了Ok按钮，布尔值就是true，反之布尔值是false
        # （也只有按下Ok按钮时，返回的文本内容才会有值）。
        text, ok = QInputDialog.getText(self, '请输入分享连接',
                                        '请输入分享连接:')

        if ok:
            self.le.setText(str(text))  # 把我们从对话框接收到的文本设置到单行编辑框组件上显示。
            self.msg = str(text)
            self.sent_msg()



class InputDialog(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.msg = ''
        self.userid = ''
        self.username = ''

    def initUI(self):

        # 创建一个标签
        label = QLabel(self)
        # 设置标签文本
        label.setText('毛坯一进一出消息发送')
        # 设置标签位置
        label.move(200, 50)
        # label.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, Qt.red)


        self.btn = QPushButton('分享给沈陈芳', self)
        self.btn.move(100, 130)
        self.btn.clicked.connect(self.showDialog)

        self.btn1 = QPushButton('分享给黄艳春', self)
        self.btn1.move(100, 230)
        self.btn1.clicked.connect(self.showDialog1)


        self.le = QLineEdit(self)
        self.le.move(200, 130)

        self.le1 = QLineEdit(self)
        self.le1.move(200, 230)

        self.setGeometry(300, 300, 600, 400)
        self.setWindowTitle('消息推送')

        self.show()

    def showDialog(self):
        # 这一行会显示一个输入对话框。第一个字符串参数是对话框的标题，第二个字符串参数是对话框内的消息文本。
        # 对话框返回输入的文本内容和一个布尔值。如果我们点击了Ok按钮，布尔值就是true，反之布尔值是false
        # （也只有按下Ok按钮时，返回的文本内容才会有值）。
        text, ok = QInputDialog.getText(self, '分享给沈陈芳', '请输入分享连接:')
        if ok:
            self.le.setText(str(text))  # 把我们从对话框接收到的文本设置到单行编辑框组件上显示。
            self.msg = str(text)
            self.userid = '17971'
            self.username = '沈陈芳'
            self.sent_msg()

    def showDialog1(self):

        text, ok = QInputDialog.getText(self, '分享给黄艳春', '请输入分享连接:')
        if ok:
            self.le1.setText(str(text))  # 把我们从对话框接收到的文本设置到单行编辑框组件上显示。
            self.msg = str(text)
            self.userid = '17959'
            self.username = '黄艳春'
            self.sent_msg()


    def sent_msg(self):
        now_times = datetime.datetime.now()
        now_data = now_times.strftime('%Y.%m.%d')
        # msg = f"{now_data}一进一出单据号汇总，<a href=\"https://drive.weixin.qq.com/s?k=AKQAZwcoAAc1UBRuR4\">点击可查看</a>"
        msgs = f"{now_data}-{self.username}一进一出单据号汇总，<a href=\"{self.msg}\">点击可查看</a>"
        print(msgs)

        # send_post('20743|17971', msg)    ## 沈
        # send_post('20743|17959', msg)  ## 黄艳春
        send_post(f'20743|{self.userid}', msgs)
        # send_post('20743', msgs)


if __name__ == '__main__':
    print("消息推送")
    app = QApplication(sys.argv)

    ex = InputDialog()

    sys.exit(app.exec_())


