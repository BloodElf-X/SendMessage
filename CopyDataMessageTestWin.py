# -*- coding: utf-8 -*-
import win32con
import win32gui
import ctypes
import ctypes.wintypes
from PyQt4 import QtGui, QtCore
from CopyDataMessageTestUI import Ui_MainWindow


class COPYDATASTRUCT(ctypes.Structure):
    _fields_ = [
        ('dwData', ctypes.wintypes.LPARAM),
        ('cbData', ctypes.wintypes.DWORD),
        ('lpData', ctypes.c_void_p)
    ]


PCOPYDATASTRUCT = ctypes.POINTER(COPYDATASTRUCT)


class SendMessageTest(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(SendMessageTest, self).__init__()
        self.setupUi(self)

        self.button_send_message.clicked.connect(self.send_message)
        self.button_refresh_win_list.clicked.connect(self.refresh_win_list)
        self.tabMsg.currentChanged.connect(self.tab_change)

        self.win_name = None
        self.message_content = None
        self.is_find_win = False
        self.tool_hwnd = None

        self.init_win_name_list()

        self.editWndHandle.setText(str(self.tool_hwnd))

    def init_win_name_list(self):
        win32gui.EnumWindows(self.add_win_name, None)
        self.com_win_name_list.setCurrentIndex(-1)

    def add_win_name(self, hwnd, extra):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            win_name = win32gui.GetWindowText(hwnd).decode('GB2312')
            if len(win_name) == 0:
                return

            self.com_win_name_list.addItem(win_name)

    def refresh_win_list(self):
        self.com_win_name_list.clear()
        self.init_win_name_list()

    def send_message(self):
        try:
            self.win_name = unicode(self.com_win_name_list.currentText())
            if len(self.win_name) == 0:
                QtGui.QMessageBox.critical(None, u'错误信息',
                                           u'接受消息的进程窗口名称不能为空。',
                                           QtGui.QMessageBox.Close)
                return

            self.message_content = unicode(self.pedit_message_content.toPlainText())
            if len(self.message_content) == 0:
                QtGui.QMessageBox.critical(None, u'错误信息',
                                           u'发送的消息内容不能为空。',
                                           QtGui.QMessageBox.Close)
                return

            self.is_find_win = False
            win32gui.EnumWindows(self.handle_window, None)

            if not self.is_find_win:
                QtGui.QMessageBox.critical(None, u'错误信息',
                                           u'未找到窗口对应的进程。',
                                           QtGui.QMessageBox.Close)
        except Exception, error:
            QtGui.QMessageBox.critical(None, u'错误信息',
                                       error,
                                       QtGui.QMessageBox.Close)

    def handle_window(self, hwnd, extra):
        try:
            if win32gui.IsWindowVisible(hwnd):
                a = win32gui.GetWindowText(hwnd).decode('gb2312')

                if self.win_name == a:
                    self.is_find_win = True
                    sender_hwnd = 0
                    buf = ctypes.create_unicode_buffer(self.message_content)

                    copydata = COPYDATASTRUCT()
                    copydata.dwData = 0
                    copydata.cbData = len(buf)
                    copydata.lpData = ctypes.cast(buf, ctypes.c_void_p)

                    ctypes.windll.user32.SendMessageA(hwnd,
                                                      win32con.WM_COPYDATA,
                                                      sender_hwnd,
                                                      ctypes.byref(copydata))
        except Exception, error:
            QtGui.QMessageBox.critical(None, u'错误信息',
                                       error,
                                       QtGui.QMessageBox.Close)

    def tab_change(self, index):
        if index == 1 and not self.tool_hwnd:
            self.tool_hwnd = win32gui.FindWindow(None, u'CopyData消息测试工具');
            self.editWndHandle.setText(str(self.tool_hwnd))

    def winEvent(self, *args, **kwargs):
        for msg in args:
            if msg.message == win32con.WM_COPYDATA:
                pCDS = ctypes.cast(msg.lParam, PCOPYDATASTRUCT)
                msg_content = ctypes.wstring_at(pCDS.contents.lpData)
                self.pTextMsgContent.setPlainText(msg_content)

        return super(SendMessageTest, self).winEvent(*args, **kwargs)


if __name__ == '__main__':
    import sys

    app = QtGui.QApplication(sys.argv)
    wnd = SendMessageTest()
    wnd.show()

    sys.exit(app.exec_())
