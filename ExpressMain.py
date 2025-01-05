import os
import sys
import subprocess
import darkdetect
from config import cfg
from ctypes import CDLL, c_int
from winotify import Notification, audio
from win32api import GetVolumeInformation
from PySide6.QtGui import QIcon, QColor
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QTimer
from PySide6.QtWidgets import QApplication, QHBoxLayout, QVBoxLayout, QLabel, QWidget
from qfluentwidgets import setTheme, Theme, BodyLabel, isDarkTheme, PushButton, SubtitleLabel, ProgressBar, \
    InfoBar, InfoBarIcon, InfoBarPosition, IndeterminateProgressBar, setThemeColor
from qfluentwidgets.common.style_sheet import FluentStyleSheet
from qframelesswindow import TitleBar
from qfluentwidgets import FluentIcon as FIF


def isWin11():
    return sys.platform == 'win32' and sys.getwindowsversion().build >= 22000


if isWin11():
    from qframelesswindow import AcrylicWindow as Window
else:
    from qframelesswindow import FramelessWindow as Window

DLL_PATH = "./_internal/taskbar-progress.dll"

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


class FluentTitleBar(TitleBar):
    def __init__(self, parent):
        super().__init__(parent)
        self.setFixedHeight(45)
        self.hBoxLayout.removeWidget(self.minBtn)
        self.hBoxLayout.removeWidget(self.maxBtn)
        self.minBtn.deleteLater()
        self.maxBtn.deleteLater()
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
        self.hBoxLayout.addStretch(-1)
        self.hBoxLayout.addLayout(self.vBoxLayout)
        FluentStyleSheet.FLUENT_WINDOW.apply(self)

    def GetDriveName(self):
        if GetVolumeInformation(drive)[0] != '':
            return GetVolumeInformation(drive)[0]
        else:
            return "U盘"


class MicaWindow(Window):

    def __init__(self):
        super().__init__()
        self.setTitleBar(FluentTitleBar(self))
        if isWin11():
            self.windowEffect.setMicaEffect(self.winId(), isDarkTheme())


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
                    currentFolder = folder + "/语文"
                elif taskList[0] == 2:
                    currentFolder = folder + "/数学"
                elif taskList[0] == 3:
                    currentFolder = folder + "/英语"
                elif taskList[0] == 4:
                    currentFolder = folder + "/物理"
                elif taskList[0] == 5:
                    currentFolder = folder + "/化学"
                elif taskList[0] == 6:
                    currentFolder = folder + "/生物"
                elif taskList[0] == 7:
                    currentFolder = folder + "/政治"
                elif taskList[0] == 8:
                    currentFolder = folder + "/历史"
                elif taskList[0] == 9:
                    currentFolder = folder + "/地理"
                elif taskList[0] == 10:
                    currentFolder = folder + "/技术"
                elif taskList[0] == 11:
                    currentFolder = folder + "/资料"
                taskList.pop(0)
                args = "fcp.exe /cmd=sync " + f"/bufsize={buf} " + '/log=FALSE "' + currentFolder.replace('/', '\\') + f'" /to="{destFolder}"'
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
        self.setWindowIcon(QIcon("icon.png"))
        self.resize(500, 130)
        self.setFixedHeight(130)
        self.setWindowOpacity(0.98)

        self.thread_running = False
        self.setup_thread()
        self.start_thread()

        self.titleBar.closeBtn.clicked.connect(self.cancelBtnOn)

        self.mainLayout = QVBoxLayout(self)
        self.topLayout = QHBoxLayout(self)
        self.bottomLayout = QHBoxLayout(self)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.topLayout.setContentsMargins(20, 0, 20, 5)
        self.bottomLayout.setContentsMargins(20, 5, 20, 20)

        self.statusLabel = SubtitleLabel(self)
        self.statusLabel.setText("准备中")
        self.isPrepare = True
        self.cancelBtn = PushButton(FIF.CLOSE, '取消同步', self)
        self.cancelBtn.clicked.connect(self.cancelBtnOn)
        self.topLayout.addWidget(self.statusLabel)
        self.topLayout.addStretch(1)
        self.topLayout.addWidget(self.cancelBtn)

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

    def setup_thread(self):
        self.thread = SyncThread()
        self.thread.valueChange.connect(self.set_value)
        self.thread_running = True

    def set_value(self):
        if self.thread.progress_value == -1:
            self.thread.quit()
            self.thread_running = False
            self.taskbarProgress.set_mode(0)
            if cfg.Notify.value:
                toast = Notification(app_id="Express", title="同步完成", msg=self.GetDriveName() + ' (' + drive + ')', icon=os.getcwd() + "\\SyncFinished.png", duration="short")
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            sys.exit()
        # elif self.thread.progress_value == -2:
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

            self.taskbarProgress.set_progress(self.thread.progress_value, 100)
            self.progressBar.setValue(self.thread.progress_value)
            self.progressLabel.setText(str(self.thread.progress_value) + '%')
            self.bottomLayout.addWidget(self.progressBar)
            self.bottomLayout.addWidget(self.spaceLabel)
            self.bottomLayout.addWidget(self.progressLabel)

    def start_thread(self):
        if self.thread_running:
            self.thread.start()
        if not self.thread_running:
            self.setup_thread()
            self.thread.start()

    def stop(self):
        self.progressBar.pause()
        self.inProgressBar.pause()
        self.taskbarProgress.set_mode(4)

        self.thread.quit()
        self.thread_running = False
        subprocess.call(["taskkill", "-f", "-im", "fcp.exe"], shell=True)
        sys.exit()

    def GetDriveName(self):
        if GetVolumeInformation(drive)[0] != '':
            return GetVolumeInformation(drive)[0]
        else:
            return "U盘"

    def cancelBtnOn(self):
        yesBtn = PushButton('确定')
        yesBtn.clicked.connect(self.stop)
        w = InfoBar(icon=InfoBarIcon.WARNING, title='取消同步？', content='', orient=Qt.Horizontal, isClosable=True,
                    position=InfoBarPosition.BOTTOM, duration=-1, parent=self)
        w.addWidget(yesBtn)
        w.show()


if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    if darkdetect.isDark():
        setTheme(Theme.DARK)
    app = QApplication(sys.argv)
    drive = sys.argv[1]
    taskList = []
    for i in range(2, 13):
        if sys.argv[i] == '1':
            taskList.append(i - 1)
    taskNum = len(taskList)
    buf = str(cfg.BufSize.value)[9:]
    folder = cfg.sourceFolder.value
    sourceFolder = folder.replace('/', '\\')
    destFolder = drive + '\\' + os.path.basename(folder) + '\\'
    w = MainWindow()
    w.show()
    app.exec()
