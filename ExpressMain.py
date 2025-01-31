import os
import sys
import subprocess
import darkdetect
import ExpressRes
from config import cfg
from ctypes import CDLL, c_int
from winotify import Notification, audio
from win32api import GetVolumeInformation
from PySide6.QtGui import QIcon, QColor
from PySide6.QtCore import Qt, QThread, Signal, QEvent
from PySide6.QtWidgets import QApplication, QHBoxLayout, QVBoxLayout, QLabel, QWidget, QGridLayout, QFrame, QPushButton
from qfluentwidgets import setTheme, Theme, BodyLabel, isDarkTheme, PushButton, SubtitleLabel, ProgressBar, \
    InfoBar, InfoBarIcon, InfoBarPosition, IndeterminateProgressBar, setThemeColor, PrimaryPushButton, TextWrap
from qframelesswindow.titlebar import MinimizeButton, CloseButton, MaximizeButton
from qframelesswindow.utils import startSystemMove
from qframelesswindow import FramelessDialog
from qfluentwidgets.common.style_sheet import FluentStyleSheet
from qfluentwidgets import FluentIcon as FIF


def isWin11():
    return sys.platform == 'win32' and sys.getwindowsversion().build >= 22000


if isWin11():
    from qframelesswindow import AcrylicWindow as Window, TitleBarButton
else:
    from qframelesswindow import FramelessWindow as Window

DLL_PATH = "./taskbar-progress.dll"

TBPF_NOPROGRESS = 0x0
TBPF_INDETERMINATE = 0x1
TBPF_NORMAL = 0x2
TBPF_ERROR = 0x4
TBPF_PAUSED = 0x8


class TaskbarProgress:
    def __init__(self, dll_path: str = DLL_PATH) -> None:
        """Windows progress bar."""
        if sys.platform == "win32":
            self._dll = CDLL(dll_path)
        else:
            raise NotImplementedError("Only Windows is supported")
        self._is_init = False

    def init(self) -> int:
        """Initialize the progress bar."""
        ret = self._dll.init()
        self._is_init = True
        return ret

    def set_mode(self, mode: int) -> int:
        """Set the progress bar mode."""
        if not self._is_init:
            raise RuntimeError("ProgressBar is not initialized")
        return self._dll.set_mode(c_int(mode))

    def set_progress(self, progress: int, total: int) -> int:
        """Set the progress bar progress and total."""
        if not self._is_init:
            raise RuntimeError("ProgressBar is not initialized")
        return self._dll.set_value(c_int(progress), c_int(total))

    def end(self) -> int:
        """End the progress bar."""
        return self._dll.end()


class TitleBarBase(QWidget):
    """ Title bar base class """

    def __init__(self, parent):
        super().__init__(parent)
        self.minBtn = MinimizeButton(parent=self)
        self.closeBtn = CloseButton(parent=self)
        self.maxBtn = MaximizeButton(parent=self)

        self._isDoubleClickEnabled = True

        self.resize(200, 32)
        self.setFixedHeight(32)

        self.window().installEventFilter(self)

    def eventFilter(self, obj, e):
        if obj is self.window():
            if e.type() == QEvent.WindowStateChange:
                self.maxBtn.setMaxState(self.window().isMaximized())
                return False

        return super().eventFilter(obj, e)

    def mouseDoubleClickEvent(self, event):
        """ Toggles the maximization state of the window """
        if event.button() != Qt.LeftButton or not self._isDoubleClickEnabled:
            return

        self.__toggleMaxState()

    def mouseMoveEvent(self, e):
        if sys.platform != "win32" or not self.canDrag(e.pos()):
            return

        startSystemMove(self.window(), e.globalPos())

    def mousePressEvent(self, e):
        if sys.platform == "win32" or not self.canDrag(e.pos()):
            return

        startSystemMove(self.window(), e.globalPos())

    def __toggleMaxState(self):
        """ Toggles the maximization state of the window and change icon """
        if self.window().isMaximized():
            self.window().showNormal()
        else:
            self.window().showMaximized()

        if sys.platform == "win32":
            from qframelesswindow.utils.win32_utils import releaseMouseLeftButton
            releaseMouseLeftButton(self.window().winId())

    def _isDragRegion(self, pos):
        """ Check whether the position belongs to the area where dragging is allowed """
        width = 0
        for button in self.findChildren(TitleBarButton):
            if button.isVisible():
                width += button.width()

        return 0 < pos.x() < self.width() - width

    def _hasButtonPressed(self):
        """ whether any button is pressed """
        return any(btn.isPressed() for btn in self.findChildren(TitleBarButton))

    def canDrag(self, pos):
        """ whether the position is draggable """
        return self._isDragRegion(pos) and not self._hasButtonPressed()

    def setDoubleClickEnabled(self, isEnabled):
        """ whether to switch window maximization status when double clicked
        Parameters
        ----------
        isEnabled: bool
            whether to enable double click
        """
        self._isDoubleClickEnabled = isEnabled


class TitleBar(TitleBarBase):
    def __init__(self, parent):
        super().__init__(parent)
        self.hBoxLayout = QHBoxLayout(self)

        # add buttons to layout
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.minBtn, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.maxBtn, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.closeBtn, 0, Qt.AlignRight)


class FluentTitleBar(TitleBar):
    def __init__(self, parent):
        super().__init__(parent)
        self.setFixedHeight(45)
        self.hBoxLayout.removeWidget(self.minBtn)
        self.hBoxLayout.removeWidget(self.maxBtn)
        self.minBtn.setVisible(False)
        self.maxBtn.setVisible(False)
        self.hBoxLayout.removeWidget(self.closeBtn)
        self.titleLabel = QLabel(self)
        self.titleLabel.setText(self.GetDriveName() + ' (' + drive + ')' + ' - Express')
        self.titleLabel.setObjectName('titleLabel')
        self.vBoxLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.setSpacing(0)
        self.buttonLayout.setContentsMargins(0, 0, 0, 0)
        self.buttonLayout.setAlignment(Qt.AlignTop)
        self.buttonLayout.addWidget(self.closeBtn)
        self.titleLayout = QHBoxLayout()
        self.titleLayout.setContentsMargins(0, 8, 0, 0)
        self.titleLayout.addWidget(self.titleLabel)
        self.titleLayout.setAlignment(Qt.AlignTop)
        self.vBoxLayout.addLayout(self.buttonLayout)
        self.vBoxLayout.addStretch(1)
        self.hBoxLayout.addLayout(self.titleLayout)
        self.hBoxLayout.addStretch(50)
        self.hBoxLayout.addLayout(self.vBoxLayout)
        FluentStyleSheet.FLUENT_WINDOW.apply(self)

    def GetDriveName(self):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "U盘"
        except:
            sys.exit()


class MicaWindow(Window):

    def __init__(self):
        super().__init__()
        self.setTitleBar(FluentTitleBar(self))
        if isWin11():
            self.windowEffect.setMicaEffect(self.winId(), isDarkTheme())


class Ui_MessageBox:
    """ Ui of message box """

    yesSignal = Signal()

    def __init__(self, *args, **kwargs):
        pass

    def _setUpUi(self, title, content, parent):
        self.content = content
        self.titleLabel = QLabel(title, parent)
        self.contentLabel = BodyLabel(content, parent)

        self.buttonGroup = QFrame(parent)
        self.yesButton = PrimaryPushButton(self.tr('确定'), self.buttonGroup)

        self.vBoxLayout = QVBoxLayout(parent)
        self.textLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout(self.buttonGroup)

        self.__initWidget()

    def __initWidget(self):
        self.__setQss()
        self.__initLayout()

        self.yesButton.setAttribute(Qt.WA_LayoutUsesWidgetRect)
        self.yesButton.setFixedWidth(150)
        self.yesButton.setFocus()
        self.buttonGroup.setFixedHeight(81)

        self.contentLabel.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._adjustText()

        self.yesButton.clicked.connect(self.__onYesButtonClicked)

    def _adjustText(self):
        if self.isWindow():
            if self.parent():
                w = max(self.titleLabel.width(), self.parent().width())
                chars = max(min(w / 9, 140), 30)
            else:
                chars = 100
        else:
            w = max(self.titleLabel.width(), self.window().width())
            chars = max(min(w / 9, 100), 30)

        self.contentLabel.setText(TextWrap.wrap(self.content, chars, False)[0])

    def __initLayout(self):
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.addLayout(self.textLayout, 1)
        self.vBoxLayout.addWidget(self.buttonGroup, 0, Qt.AlignBottom)
        self.vBoxLayout.setSizeConstraint(QVBoxLayout.SetMinimumSize)

        self.textLayout.setSpacing(12)
        self.textLayout.setContentsMargins(24, 24, 24, 24)
        self.textLayout.addWidget(self.titleLabel, 0, Qt.AlignTop)
        self.textLayout.addWidget(self.contentLabel, 0, Qt.AlignTop)

        self.buttonLayout.setContentsMargins(24, 24, 24, 24)
        self.buttonLayout.addWidget(self.yesButton, 1, Qt.AlignRight)

    def __onYesButtonClicked(self):
        self.accept()
        self.yesSignal.emit()

    def __setQss(self):
        self.titleLabel.setObjectName("titleLabel")
        self.contentLabel.setObjectName("contentLabel")
        self.buttonGroup.setObjectName('buttonGroup')

        FluentStyleSheet.DIALOG.apply(self)
        FluentStyleSheet.DIALOG.apply(self.contentLabel)

        self.yesButton.adjustSize()

    def setContentCopyable(self, isCopyable: bool):
        """ set whether the content is copyable """
        if isCopyable:
            self.contentLabel.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse)
        else:
            self.contentLabel.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)


class Dialog(FramelessDialog, Ui_MessageBox):
    """ Dialog box """

    yesSignal = Signal()
    cancelSignal = Signal()

    def __init__(self, title: str, content: str, parent=None):
        super().__init__(parent=parent)
        self._setUpUi(title, content, self)

        self.windowTitleLabel = QLabel(title, self)

        self.setResizeEnabled(False)
        self.resize(240, 192)
        self.titleBar.hide()

        self.vBoxLayout.insertWidget(0, self.windowTitleLabel, 0, Qt.AlignTop)
        self.windowTitleLabel.setObjectName('windowTitleLabel')
        FluentStyleSheet.DIALOG.apply(self)
        self.setFixedSize(self.size())

    def setTitleBarVisible(self, isVisible: bool):
        self.windowTitleLabel.setVisible(isVisible)


class DeleteThread(QThread):
    deleteFinished = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent=parent)

    def run(self):
        args = "fcp.exe /cmd=delete " + f"/bufsize={buf} /log=FALSE " + f'/force_start={concurrentProcess} "{destFolder}"'
        subprocess.call(args, shell=True)
        self.deleteFinished.emit(True)
        return


class SyncThread(QThread):
    valueChange = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.is_paused = bool(0)
        self.progress_value = int(0)
    #     self.initDeltaSize = self.getDeltaSize()
    #
    #     self.updateProgressTimer = QTimer()
    #     self.updateProgressTimer.start(3100)
    #     self.updateProgressTimer.timeout.connect(self.updateProgress)
    #
    # def getDeltaSize(self):
    #     destSize = 0
    #     sourceSize = 0
    #     try:
    #         for root, dirs, files in os.walk(destFolder):
    #             destSize += sum([os.path.getsize(os.path.join(root, name)) for name in files])
    #     except:
    #         pass
    #     try:
    #         for root, dirs, files in os.walk(sourceFolder):
    #             sourceSize += sum([os.path.getsize(os.path.join(root, name)) for name in files])
    #     except:
    #         pass
    #     return abs(destSize - sourceSize)
    #
    # def runFunction(self, f, max_wait, default_value):
    #     try:
    #         return func_timeout.func_timeout(max_wait, self.getDeltaSize)
    #     except func_timeout.FunctionTimedOut:
    #         pass
    #     return default_value
    #
    # def updateProgress(self):
    #     tempDeltaSize = self.runFunction(self.getDeltaSize, 3, -1)
    #     if tempDeltaSize/1024/1024 <= 100:
    #        self.valueChange.emit(-2)
    #     elif tempDeltaSize == -1:
    #         return
    #     else:
    #         self.progress_value = int(abs(self.initDeltaSize - tempDeltaSize) * 100 / self.initDeltaSize)
    #         if self.progress_value != 0:
    #             self.valueChange.emit(self.progress_value)

    def run(self):
        while True:
            if taskList:
                if taskList[0] == 1:
                    currentFolder = cfg.yuwenFolder.value
                elif taskList[0] == 2:
                    currentFolder = cfg.shuxueFolder.value
                elif taskList[0] == 3:
                    currentFolder = cfg.yingyuFolder.value
                elif taskList[0] == 4:
                    currentFolder = cfg.wuliFolder.value
                elif taskList[0] == 5:
                    currentFolder = cfg.huaxueFolder.value
                elif taskList[0] == 6:
                    currentFolder = cfg.shengwuFolder.value
                elif taskList[0] == 7:
                    currentFolder = cfg.zhengzhiFolder.value
                elif taskList[0] == 8:
                    currentFolder = cfg.lishiFolder.value
                elif taskList[0] == 9:
                    currentFolder = cfg.diliFolder.value
                elif taskList[0] == 10:
                    currentFolder = cfg.jishuFolder.value
                elif taskList[0] == 11:
                    currentFolder = cfg.ziliaoFolder.value

                taskList.pop(0)
                args = "fcp.exe /cmd=sync " + f"/bufsize={buf} /log=FALSE " + f'/force_start={concurrentProcess} {commandOption} "' + currentFolder + f'" /to="{destFolder}"'
                subprocess.call(args, shell=True)
                self.progress_value = int((taskNum - len(taskList)) / taskNum * 100)
                self.valueChange.emit(self.progress_value)
            else:
                self.progress_value = -1
                self.valueChange.emit(self.progress_value)
                return


class MainWindow(MicaWindow):

    def __init__(self):
        super().__init__()
        setThemeColor(QColor(113, 89, 249))
        self.resize(500, 130)
        self.setWindowTitle('Express - ' + self.GetDriveName() + ' (' + drive + ')')
        self.setWindowIcon(QIcon(':/icon.png'))
        self.setFixedHeight(150)
        self.setWindowOpacity(0.98)
        self.displayText = {1:"同步 (默认)", 2:"同步 (低占用)", 3:"复制 (最近文件)", 4:"复制 (从时间戳)"}[mode]
        if mode == 3 or mode == 4:
            self.displayText += ' - '
            self.displayText += "删除原有文件" if isDelete else "保留原有文件"
        self.subject = ""
        for i in taskList:
            self.subject += {1: '语文', 2: '数学', 3: '英语', 4: '物理', 5: '化学', 6: '生物', 7: '政治', 8: '历史', 9: '地理', 10: '技术', 11: '资料'}[i] + '; '

        self.titleBar.closeBtn.clicked.connect(self.onCancelBtn)

        self.mainLayout = QVBoxLayout(self)
        self.topLayout = QGridLayout(self)
        self.bottomLayout = QHBoxLayout(self)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.topLayout.setContentsMargins(20, 0, 20, 0)
        self.bottomLayout.setContentsMargins(20, 5, 20, 20)

        self.statusLabel = SubtitleLabel(self)
        self.detailLabel = BodyLabel(self)
        self.detailLabel.setText(self.displayText)
        self.detailLabel.setTextColor(QColor(114, 114, 114))
        self.isPrepare = True
        self.showDetailBtn = PushButton(FIF.CHECKBOX, '选项', self)
        self.cancelBtn = PushButton(FIF.CLOSE, '取消', self)
        self.showDetailBtn.clicked.connect(self.onShowDetailBtn)
        self.cancelBtn.clicked.connect(self.onCancelBtn)
        self.topLayout.addWidget(self.statusLabel, 0, 0)
        self.topLayout.addWidget(self.detailLabel, 1, 0)
        self.topLayout.setColumnStretch(0, 1)
        self.topLayout.addWidget(self.showDetailBtn, 0, 1, 2, 1)
        self.topLayout.addWidget(self.cancelBtn, 0, 2, 2, 1)

        self.inProgressBar = IndeterminateProgressBar(self)
        self.progressBar = ProgressBar(self)
        self.spaceLabel = QLabel(self)
        self.spaceLabel.setFixedWidth(5)
        self.progressLabel = BodyLabel(self)
        self.progressBar.setVisible(False)
        self.spaceLabel.setVisible(False)
        self.progressLabel.setVisible(False)
        self.bottomLayout.addWidget(self.inProgressBar)

        self.taskbarProgress = TaskbarProgress()
        self.taskbarProgress.init()
        self.taskbarProgress.set_mode(1)

        self.mainLayout.addWidget(self.titleBar)
        self.mainLayout.addStretch(1)
        self.mainLayout.addLayout(self.topLayout)
        self.mainLayout.addLayout(self.bottomLayout)

        self.deleteThread = DeleteThread()
        self.syncThread = SyncThread()
        self.syncThreadRunning = False
        self.deleteThreadRunning = False
        if isDelete:
            self.statusLabel.setText("正在删除原有文件")
            self.setupDeleteThread()
            self.startDeleteThread()
        else:
            self.statusLabel.setText("准备中")
            self.setupSyncThread()
            self.startSyncThread()

    def setupDeleteThread(self):
        self.deleteThread.deleteFinished.connect(self.deleteThreadFinished)
        self.deleteThreadRunning = True

    def startDeleteThread(self):
        if self.deleteThreadRunning:
            self.deleteThread.start()
        else:
            self.setupDeleteThread()
            self.deleteThread.start()

    def deleteThreadFinished(self):
        self.deleteThread.quit()
        self.deleteThreadRunning = False
        self.setupSyncThread()
        self.startSyncThread()
        self.statusLabel.setText("准备中")

    def setupSyncThread(self):
        self.syncThread.valueChange.connect(self.setSyncValue)
        self.syncThreadRunning = True

    def startSyncThread(self):
        if self.syncThreadRunning:
            self.syncThread.start()
        else:
            self.setupSyncThread()
            self.syncThread.start()

    def setSyncValue(self):
        if self.syncThread.progress_value == -1:
            self.syncThread.quit()
            self.syncThreadRunning = False
            self.taskbarProgress.set_mode(0)
            if cfg.Notify.value:
                toast = Notification(app_id="Express", title="同步完成", msg=self.GetDriveName() + ' (' + drive + ')', duration="short")
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            sys.exit()
        # elif self.syncThread.progress_value == -2:
        #     self.statusLabel.setText("即将完成")
        #     self.bottomLayout.removeWidget(self.progressBar)
        #     self.inProgressBar.setVisible(True)
        #     self.progressBar.setVisible(False)
        #     self.spaceLabel.setVisible(False)
        #     self.progressLabel.setVisible(False)
        #     self.taskbarProgress.set_mode(1)
        else:
            if self.isPrepare:
                self.statusLabel.setText("正在同步")
                self.bottomLayout.removeWidget(self.inProgressBar)
                self.inProgressBar.setVisible(False)
                self.progressBar.setVisible(True)
                self.spaceLabel.setVisible(True)
                self.progressLabel.setVisible(True)
                self.isPrepare = False
                self.taskbarProgress.set_mode(2)
                self.taskbarProgress.init()

            self.taskbarProgress.set_progress(self.syncThread.progress_value, 100)
            self.progressBar.setValue(self.syncThread.progress_value)
            self.progressLabel.setText(str(self.syncThread.progress_value) + '%')
            self.bottomLayout.addWidget(self.progressBar)
            self.bottomLayout.addWidget(self.spaceLabel)
            self.bottomLayout.addWidget(self.progressLabel)

    def stopThread(self):
        self.progressBar.pause()
        self.inProgressBar.pause()
        self.taskbarProgress.set_mode(4)

        self.deleteThread.quit()
        self.deleteThreadRunning = False
        self.syncThread.quit()
        self.syncThreadRunning = False
        subprocess.call(["taskkill", "-f", "-im", "fcp.exe"], shell=True)
        sys.exit()

    def GetDriveName(self):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "U盘"
        except:
            sys.exit()

    def onCancelBtn(self):
        yesBtn = PushButton('确定')
        yesBtn.clicked.connect(self.stopThread)
        w = InfoBar(icon=InfoBarIcon.WARNING, title='取消同步？', content='', orient=Qt.Horizontal, isClosable=True,
                    position=InfoBarPosition.BOTTOM, duration=-1, parent=self)
        w.addWidget(yesBtn)
        w.show()

    def onShowDetailBtn(self):
        title = 'Express 选项'
        content = f"目标驱动器: {drive}\\\n模式: {self.displayText}\n学科: {self.subject}\n命令行选项: {commandOption}\n缓冲区大小: {buf} MB\n并行进程数: {concurrentProcess}"
        w = Dialog(title, content, self)
        w.setTitleBarVisible(False)
        w.setContentCopyable(True)
        if w.exec():
            pass


if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    if darkdetect.isDark():
        setTheme(Theme.DARK)
    app = QApplication(sys.argv)

    """
    args
    1           drive
    2 - 12      subject
    13          mode{1:"sync(default)", 2:"sync(low)", 3:"copy(lately)", 4:"copy(from_date)"}
    14          isDelete
    15          commandOption
    """

    drive = sys.argv[1]
    taskList = []
    for i in range(2, 13):
        if sys.argv[i] == '1':
            taskList.append(i - 1)
    taskNum = len(taskList)
    buf = str(cfg.BufSize.value)[9:]
    concurrentProcess = cfg.ConcurrentProcess.value
    sourceFolder = os.path.normpath(cfg.sourceFolder.value)
    destFolder = drive + '\\' + os.path.basename(sourceFolder) + '\\'
    mode = int(sys.argv[13])
    isDelete = False if sys.argv[14] == 'False' else True
    commandOption = sys.argv[15]

    w = MainWindow()
    w.show()
    app.exec()
