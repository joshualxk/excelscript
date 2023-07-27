import logging
import pathlib
import re
import sys
import traceback

from PyQt6.QtCore import Qt, pyqtSignal, QObject, QTimer, QThread
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtWidgets import QMainWindow, QApplication, QLabel, QMessageBox, QWidget, QStackedWidget, \
    QHBoxLayout, QVBoxLayout, QListWidget, QListWidgetItem, QLineEdit, QPushButton, QDialog, QCheckBox
from openpyxl.utils import exceptions

from .data import DataHolder

logger = logging.getLogger(__name__)


class WorkerSignals(QObject):
    started = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(str)


class WorkerThread(QThread):

    def __init__(self, fn, *args, **kwargs) -> None:
        super().__init__()

        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.kwargs['progress_callback'] = self.signals.progress.emit

    def run(self):

        try:
            self.signals.started.emit()
            result = self.fn(*self.args, **self.kwargs)

        except:
            logging.exception("<exception>")
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value))
        else:
            self.signals.result.emit(result)


class WaitingDialog(QDialog):
    closeSignal = pyqtSignal()

    def __init__(self, parent=None):
        super(WaitingDialog, self).__init__(parent)

        self.setFixedWidth(200)
        self._want_to_close = False
        self.label = QLabel(self)
        layout = QHBoxLayout()
        self.setLayout(layout)
        layout.addWidget(self.label)

        self.setWindowTitle("等待中")

        self.timer = QTimer()
        self.timer.setInterval(600)
        self.timer.timeout.connect(self.tick)
        self.timer.start()
        self.count = 0

    def tick(self):
        self.count += 1
        if self.count >= 4:
            self.count = 0
        self.setWindowTitle(f"等待中{'.' * self.count}")

    def showEvent(self, a0) -> None:
        super().showEvent(a0)
        self.showMessage('')

    def closeEvent(self, event) -> None:
        self.closeSignal.emit()
        super(WaitingDialog, self).closeEvent(event)

    def showMessage(self, msg):
        self.label.setText(msg)


def messageBox(parent, content):
    dlg = QMessageBox(parent)
    dlg.setWindowTitle(" ")
    dlg.setText(content)
    logger.debug(f'messageBox:{content}')
    dlg.exec()


def messageDialog(parent, title, ls):
    dlg = QDialog(parent)
    dlg.setWindowTitle(" ")
    layout = QVBoxLayout()
    dlg.setLayout(layout)
    label = QLabel(title)
    layout.addWidget(label)
    listWidget = QListWidget()
    layout.addWidget(listWidget)
    for v in ls:
        listWidget.addItem(str(v))
    dlg.exec()


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setWindowTitle("excel拆表 1.0")
        icon_file = pathlib.Path(__file__).parent / 'Shin-chan.png'
        icon = QIcon(str(icon_file))
        self.setWindowIcon(icon)
        self.workerThread = None
        self.central_widget = QStackedWidget()
        self.setCentralWidget(self.central_widget)
        self.waitDialog = WaitingDialog(self)
        self.waitDialog.closeSignal.connect(self.thread_terminate_fn)

        self.mainWidget = MainWidget(self)
        self.central_widget.addWidget(self.mainWidget)
        self.loader_widget = LoaderWidget(self)
        self.central_widget.addWidget(self.loader_widget)

        self.showLoaderWidget()

    def showLoaderWidget(self):
        self.central_widget.setCurrentWidget(self.loader_widget)
        # self.showMainWidget('test.xlsx')

    def showMainWidget(self, file):
        def fn(progress_callback):
            return DataHolder.create(file)

        self.long_time_task(fn)

    def outputExcel(self):
        def fn(progress_callback):
            return self.mainWidget.dataHolder.gen(progress_callback)

        self.long_time_task(fn)

    def long_time_task(self, fn, *args):
        logger.debug('正在执行:long_time_task')
        if self.workerThread and self.workerThread.isRunning():
            messageBox(self, '另一个任务运行中..')
            return
        self.workerThread = WorkerThread(fn, *args)
        self.workerThread.signals.started.connect(self.thread_started)
        self.workerThread.signals.result.connect(self.success_fn)
        self.workerThread.signals.progress.connect(self.progress_fn)
        self.workerThread.signals.error.connect(self.error_fn)

        # Execute
        self.workerThread.start()

    def success_fn(self, x):
        logger.debug('正在执行:success_fn')
        self.waitDialog.hide()
        if isinstance(x, DataHolder):
            self.mainWidget.setDataHolder(x)
            self.central_widget.setCurrentWidget(self.mainWidget)
        else:
            logger.info(f'导出excel结果:{x}')
            if len(x) > 0:
                ls = []
                for s in x:
                    ls.append(s)
                messageDialog(self, '导出excel成功，以下分组未归类', ls)
            else:
                messageBox(self, '导出excel成功')

    def error_fn(self, arg):
        logger.debug('正在执行:error_fn')
        e, v = arg
        if e == exceptions.InvalidFileException:
            v = '不支持的文件类型'
        else:
            v = str(v)
        self.waitDialog.hide()

        messageBox(self, v)

    def thread_started(self):
        logger.debug('正在执行:thread_started')
        self.waitDialog.exec()

    def thread_terminate_fn(self):
        logger.debug('正在执行:thread_terminate_fn')
        if self.workerThread and self.workerThread.isRunning():
            logger.info('>>>>>>>>>>中止线程<<<<<<<<<')
            self.workerThread.terminate()

    def progress_fn(self, n):
        self.waitDialog.showMessage(n)


class LoaderWidget(QWidget):
    def __init__(self, parent):
        super(LoaderWidget, self).__init__()

        self.setAcceptDrops(True)
        self.parent = parent
        layout = QHBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label = QLabel('将excel拖至此处', self)
        label.setFont(QFont('Arial', 20))
        layout.addWidget(label)
        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if len(event.mimeData().urls()) != 1:
            messageBox(self, '导入文件错误')
            return
        file = event.mimeData().urls()[0].toLocalFile()
        self.parent.showMainWidget(file)


class MainWidget(QWidget):
    def __init__(self, parent):
        super(MainWidget, self).__init__()
        self.parent = parent
        self.dataHolder = None
        self.selectedSheet = None
        self.sheetWidget = None
        self.right_widget4 = None
        self.right_widget3 = None
        self.right_widget2 = None
        self.right_widget1 = None
        self.right_msg_label = None

        pageLayout = QVBoxLayout()
        topLayout = QHBoxLayout()
        topLayout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        bottomLayout = QHBoxLayout()
        bottomLayout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        leftWidget = QWidget()
        leftWidget.setFixedWidth(400)
        rightWidget = QWidget()
        bottomLayout.addWidget(leftWidget)
        bottomLayout.addWidget(rightWidget)

        pageLayout.addLayout(topLayout)
        pageLayout.addLayout(bottomLayout)
        self.setLayout(pageLayout)

        btn = QPushButton('导出')
        btn.clicked.connect(self.on_top_btn_click)
        topLayout.addWidget(btn)
        btn2 = QPushButton('重新导入excel')
        btn2.clicked.connect(self.on_top_btn2_click)
        topLayout.addWidget(btn2)

        leftLayout = QVBoxLayout()
        leftTitle = QLabel()
        leftTitle.setText('sheet（勾选需要导出的表格）')
        leftLayout.addWidget(leftTitle)

        self.sheetWidget = QListWidget()
        self.sheetWidget.currentItemChanged.connect(self.index_changed)
        self.sheetWidget.itemChanged.connect(self.item_change_fn)
        leftLayout.addWidget(self.sheetWidget)

        self.allCheck = QCheckBox()
        self.allCheck.setText('全选')
        self.allCheck.stateChanged.connect(self.all_check_fn)
        leftLayout.addWidget(self.allCheck)
        leftWidget.setLayout(leftLayout)

        rightLayout = QVBoxLayout()
        rightLayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout1 = QVBoxLayout()
        label = QLabel('表格详情')
        layout1.addWidget(label)
        rightLayout.addLayout(layout1)
        layout2 = QHBoxLayout()
        label = QLabel('<空>')
        self.right_widget1 = label
        layout2.addWidget(label)
        rightLayout.addLayout(layout2)

        layout3 = QHBoxLayout()
        layout3.setAlignment(Qt.AlignmentFlag.AlignLeft)
        label = QLabel('表头:  第')
        label.setFixedWidth(50)
        layout3.addWidget(label)
        lineEdit = QLineEdit()
        lineEdit.setText('1')
        self.right_widget2 = lineEdit
        lineEdit.setFixedWidth(40)
        layout3.addWidget(lineEdit)
        label = QLabel('行 - 第')
        label.setFixedWidth(45)
        layout3.addWidget(label)
        lineEdit2 = QLineEdit()
        lineEdit2.setFixedWidth(40)
        self.right_widget3 = lineEdit2
        lineEdit2.setText('1')
        layout3.addWidget(lineEdit2)
        label = QLabel('行')
        layout3.addWidget(label)
        rightLayout.addLayout(layout3)

        layout4 = QHBoxLayout()
        layout4.setAlignment(Qt.AlignmentFlag.AlignLeft)
        label = QLabel('根据单元格')
        label.setFixedWidth(60)
        layout4.addWidget(label)
        lineEdit3 = QLineEdit()
        lineEdit3.setText('A1')
        lineEdit3.setFixedWidth(60)
        self.right_widget4 = lineEdit3
        layout4.addWidget(lineEdit3)
        label = QLabel('分组')
        layout4.addWidget(label)
        rightLayout.addLayout(layout4)

        layout5 = QHBoxLayout()
        layout5.setAlignment(Qt.AlignmentFlag.AlignLeft)
        btn = QPushButton('保存')
        btn.setFixedWidth(100)
        btn.clicked.connect(self.on_btn_click)
        btn2 = QPushButton('重置')
        btn2.setFixedWidth(100)
        btn2.clicked.connect(self.on_btn2_click)
        layout5.addWidget(btn)
        layout5.addWidget(btn2)
        rightLayout.addLayout(layout5)

        layout6 = QHBoxLayout()
        layout6.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.right_msg_label = QLabel()
        layout6.addWidget(self.right_msg_label)
        rightLayout.addLayout(layout6)

        rightWidget.setLayout(rightLayout)

        self.check_state_lock = False

    def all_check_fn(self, s):
        if self.check_state_lock:
            return
        self.check_state_lock = True
        state = Qt.CheckState.Checked if s == Qt.CheckState.Checked.value else Qt.CheckState.Unchecked
        lw = self.sheetWidget
        for i in range(lw.count()):
            item = lw.item(i)
            item.setCheckState(state)
        self.check_state_lock = False

    def item_change_fn(self, item):
        if self.check_state_lock:
            return
        self.check_state_lock = True
        if item.checkState() == Qt.CheckState.Checked:
            lw = self.sheetWidget
            for i in range(lw.count()):
                item = lw.item(i)
                if item.checkState() != Qt.CheckState.Checked:
                    break
            else:
                self.allCheck.setCheckState(Qt.CheckState.Checked)
        else:
            self.allCheck.setCheckState(Qt.CheckState.Unchecked)
        logger.debug(f'选中状态:{item.checkState()}')
        self.check_state_lock = False

    def hideEvent(self, ev):
        self.dataHolder = None
        self.sheetWidget.clear()

    def setDataHolder(self, dataHolder):
        self.dataHolder = dataHolder
        wb = dataHolder.wb
        sheets = [ws.title for ws in wb]
        logger.info(f'设置DataHolder:{sheets}')
        for sheet in sheets:
            item = QListWidgetItem()
            item.setText(sheet)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.sheetWidget.addItem(item)
        self.allCheck.setCheckState(Qt.CheckState.Unchecked)

    def updateRightWidgets(self):
        if not self.selectedSheet:
            return False
        ws_title = self.selectedSheet
        self.right_widget1.setText(ws_title)
        self.right_widget2.setText(
            str(self.dataHolder.sheet_detail[ws_title]['title_row1']))
        self.right_widget3.setText(
            str(self.dataHolder.sheet_detail[ws_title]['title_row2']))
        self.right_widget4.setText(
            str(self.dataHolder.sheet_detail[ws_title]['key_cell']))
        self.hide_msg()
        return True

    def index_changed(self, i):
        logger.debug(f'index_changed:{i}')
        if i is None:
            return
        self.selectedSheet = i.text()
        self.updateRightWidgets()

    def on_btn_click(self):
        if not self.selectedSheet:
            return False
        txt2 = self.right_widget2.text()
        if not re.match(r'^\d+$', txt2):
            self.show_msg(False, '表头行数填写错误')
            return True
        txt3 = self.right_widget3.text()
        if not re.match(r'^\d+$', txt3):
            self.show_msg(False, '表头行数填写错误')
            return True
        txt4 = self.right_widget4.text()
        if not re.match(r'^[a-zA-Z]+[1-9]\d*$', txt4):
            self.show_msg(False, '分组单元格填写错误')
            return True
        data2 = int(txt2)
        data3 = int(txt3)

        if data2 > data3 or data3 < 1:
            self.show_msg(False, '表头行数填写错误')
            return True

        self.dataHolder.sheet_detail[self.selectedSheet]['title_row1'] = data2
        self.dataHolder.sheet_detail[self.selectedSheet]['title_row2'] = data3
        self.dataHolder.sheet_detail[self.selectedSheet]['key_cell'] = txt4

        self.show_msg(True, '成功!')
        return True

    def show_msg(self, isSuccess, msg):
        self.right_msg_label.setText(msg)
        self.setObjectName('nom_plan_label')
        if isSuccess:
            self.right_msg_label.setStyleSheet('color: green')
        else:
            self.right_msg_label.setStyleSheet('color: red')

    def hide_msg(self):
        self.right_msg_label.setText('')

    def on_btn2_click(self):
        self.updateRightWidgets()
        return True

    def on_top_btn_click(self):
        lw = self.sheetWidget
        for i in range(lw.count()):
            item = lw.item(i)
            txt = item.text()
            self.dataHolder.sheet_detail[txt]['output'] = item.checkState(
            ) == Qt.CheckState.Checked
        logger.info(f'表格详细信息:{self.dataHolder.sheet_detail}')
        self.parent.outputExcel()

    def on_top_btn2_click(self):
        self.parent.showLoaderWidget()
        return True


if __name__ == '__main__':
    try:
        logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', encoding='utf-8',
                            level=logging.DEBUG, filename='log.txt')
    except:
        logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', encoding='utf-8',
                            level=logging.DEBUG)
    app = QApplication(sys.argv)
    ui = MainWindow()
    ui.show()
    sys.exit(app.exec())
