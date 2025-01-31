import sys
import os
import msvcrt
import ExpressRes
import subprocess
from typing import Union
from webbrowser import open as webopen
from pygetwindow import getWindowsWithTitle as GetWindow
from config import cfg, BufSize, VERSION, YEAR
from PySide6.QtCore import Qt, Signal, QTimer, QThread, QRectF, QEasingCurve
from PySide6.QtGui import QColor, QIcon, QPainter, QTextCursor, QAction, QPainterPath
from PySide6.QtWidgets import QFrame, QApplication, QWidget, QHBoxLayout, QFileDialog, QLabel, QVBoxLayout, \
    QPushButton, QButtonGroup, QTextBrowser, QTextEdit, QSizePolicy, QLineEdit, QSpinBox, QScrollArea, QScroller
from qfluentwidgets import MSFluentWindow, NavigationItemPosition, SubtitleLabel, MessageBox, ExpandLayout, \
    SettingCardGroup, PrimaryPushSettingCard, SmoothScrollArea, RadioButton, ExpandSettingCard, \
    ComboBox, SwitchButton, IndicatorPosition, qconfig, isDarkTheme, ConfigItem, OptionsConfigItem, \
    FluentStyleSheet, HyperlinkButton, Slider, IconWidget, drawIcon, setThemeColor, ImageLabel, MessageBoxBase, \
    SmoothScrollDelegate, setFont, themeColor
from qfluentwidgets.components.widgets.line_edit import EditLayer
from qfluentwidgets.components.widgets.menu import MenuAnimationType, RoundMenu
from qfluentwidgets.components.widgets.spin_box import SpinButton, SpinIcon
from qfluentwidgets import FluentIcon as FIF


class Mutex:
    def __init__(self):
        self.lockfile = None

    def __enter__(self):
        self.lockfile = open('ExpressSetting.lockfile', 'w')
        try:
            msvcrt.locking(self.lockfile.fileno(), msvcrt.LK_NBLCK, 1)
        except IOError:
            try:
                window = GetWindow("Express 设置")[0]
                if window.isMinimized:
                    window.restore()
                window.activate()
            except:
                pass
            sys.exit()

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.lockfile:
            msvcrt.locking(self.lockfile.fileno(), msvcrt.LK_UNLCK, 1)
            self.lockfile.close()
            os.remove('ExpressSetting.lockfile')


class SmoothScrollArea(QScrollArea):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.delegate = SmoothScrollDelegate(self, True)
        QScroller.grabGesture(self.viewport(), QScroller.TouchGesture)

    def setScrollAnimation(self, orient, duration, easing=QEasingCurve.OutCubic):
        """ set scroll animation

        Parameters
        ----------
        orient: Orient
            scroll orientation

        duration: int
            scroll duration

        easing: QEasingCurve
            animation type
        """
        bar = self.delegate.hScrollBar if orient == Qt.Horizontal else self.delegate.vScrollBar
        bar.setScrollAnimation(duration, easing)

    def enableTransparentBackground(self):
        self.setStyleSheet("QScrollArea{border: none; background: transparent}")

        if self.widget():
            self.widget().setStyleSheet("QWidget{background: transparent}")


class EditMenu(RoundMenu):
    """ Edit menu """

    def createActions(self):
        self.cutAct = QAction(FIF.CUT.icon(), self.tr("剪切"), self, shortcut="Ctrl+X", triggered=self.parent().cut,)
        self.copyAct = QAction(FIF.COPY.icon(), self.tr("复制"), self, shortcut="Ctrl+C", triggered=self.parent().copy,)
        self.pasteAct = QAction(FIF.PASTE.icon(), self.tr("粘贴"), self, shortcut="Ctrl+V", triggered=self.parent().paste,)
        self.cancelAct = QAction(FIF.CANCEL.icon(), self.tr("撤销"), self, shortcut="Ctrl+Z", triggered=self.parent().undo,)
        self.selectAllAct = QAction(self.tr("全选"), self, shortcut="Ctrl+A", triggered=self.parent().selectAll)
        self.action_list = [self.cutAct, self.copyAct, self.pasteAct, self.cancelAct, self.selectAllAct]

    def _parentText(self):
        raise NotImplementedError

    def _parentSelectedText(self):
        raise NotImplementedError

    def exec(self, pos, ani=True, aniType=MenuAnimationType.DROP_DOWN):
        self.clear()
        self.createActions()

        if QApplication.clipboard().mimeData().hasText():
            if self._parentText():
                if self._parentSelectedText():
                    if self.parent().isReadOnly():
                        self.addActions([self.copyAct, self.selectAllAct])
                    else:
                        self.addActions(self.action_list)
                else:
                    if self.parent().isReadOnly():
                        self.addAction(self.selectAllAct)
                    else:
                        self.addActions(self.action_list[2:])
            elif not self.parent().isReadOnly():
                self.addAction(self.pasteAct)
            else:
                return
        else:
            if not self._parentText():
                return

            if self._parentSelectedText():
                if self.parent().isReadOnly():
                    self.addActions([self.copyAct, self.selectAllAct])
                else:
                    self.addActions(
                        self.action_list[:2] + self.action_list[3:])
            else:
                if self.parent().isReadOnly():
                    self.addAction(self.selectAllAct)
                else:
                    self.addActions(self.action_list[3:])

        super().exec(pos, ani, aniType)


class LineEditMenu(EditMenu):
    """ Line edit menu """

    def __init__(self, parent: QLineEdit):
        super().__init__("", parent)
        self.selectionStart = parent.selectionStart()
        self.selectionLength = parent.selectionLength()

    def _onItemClicked(self, item):
        if self.selectionStart >= 0:
            self.parent().setSelection(self.selectionStart, self.selectionLength)

        super()._onItemClicked(item)

    def _parentText(self):
        return self.parent().text()

    def _parentSelectedText(self):
        return self.parent().selectedText()

    def exec(self, pos, ani=True, aniType=MenuAnimationType.DROP_DOWN):
        return super().exec(pos, ani, aniType)


class SpinBoxBase:
    """ Spin box ui """

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.hBoxLayout = QHBoxLayout(self)

        self.setProperty('transparent', True)
        FluentStyleSheet.SPIN_BOX.apply(self)
        self.setButtonSymbols(QSpinBox.NoButtons)
        self.setFixedHeight(33)
        setFont(self)

        self.setAttribute(Qt.WA_MacShowFocusRect, False)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._showContextMenu)

    def setReadOnly(self, isReadOnly: bool):
        super().setReadOnly(isReadOnly)
        self.setSymbolVisible(not isReadOnly)

    def setSymbolVisible(self, isVisible: bool):
        """ set whether the spin symbol is visible """
        self.setProperty("symbolVisible", isVisible)
        self.setStyle(QApplication.style())

    def _showContextMenu(self, pos):
        menu = LineEditMenu(self.lineEdit())
        menu.exec_(self.mapToGlobal(pos))

    def _drawBorderBottom(self):
        if not self.hasFocus():
            return

        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)

        path = QPainterPath()
        w, h = self.width(), self.height()
        path.addRoundedRect(QRectF(0, h-10, w, 10), 5, 5)

        rectPath = QPainterPath()
        rectPath.addRect(0, h-10, w, 8)
        path = path.subtracted(rectPath)

        painter.fillPath(path, themeColor())

    def paintEvent(self, e):
        super().paintEvent(e)
        self._drawBorderBottom()


class InlineSpinBoxBase(SpinBoxBase):
    """ Inline spin box base """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.upButton = SpinButton(SpinIcon.UP, self)
        self.downButton = SpinButton(SpinIcon.DOWN, self)

        self.hBoxLayout.setContentsMargins(0, 4, 4, 4)
        self.hBoxLayout.setSpacing(5)
        self.hBoxLayout.addWidget(self.upButton, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.downButton, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.upButton.clicked.connect(self.stepUp)
        self.downButton.clicked.connect(self.stepDown)

    def setSymbolVisible(self, isVisible: bool):
        super().setSymbolVisible(isVisible)
        self.upButton.setVisible(isVisible)
        self.downButton.setVisible(isVisible)

    def setAccelerated(self, on: bool):
        super().setAccelerated(on)
        self.upButton.setAutoRepeat(on)
        self.downButton.setAutoRepeat(on)


class SpinBox(InlineSpinBoxBase, QSpinBox):
    """ Spin box """


class TextEditMenu(EditMenu):
    def __init__(self, parent: QTextEdit):
        super().__init__("", parent)
        cursor = parent.textCursor()
        self.selectionStart = cursor.selectionStart()
        self.selectionLength = cursor.selectionEnd() - self.selectionStart + 1

    def _parentText(self):
        return self.parent().toPlainText()

    def _parentSelectedText(self):
        return self.parent().textCursor().selectedText()

    def _onItemClicked(self, item):
        if self.selectionStart >= 0:
            cursor = self.parent().textCursor()
            cursor.setPosition(self.selectionStart)
            cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, self.selectionLength)

        super()._onItemClicked(item)

    def exec(self, pos, ani=True, aniType=MenuAnimationType.DROP_DOWN):
        return super().exec(pos, ani, aniType)


class TextBrowser(QTextBrowser):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layer = EditLayer(self)
        self.scrollDelegate = SmoothScrollDelegate(self)
        FluentStyleSheet.LINE_EDIT.apply(self)
        setFont(self)

    def contextMenuEvent(self, e):
        menu = TextEditMenu(self)
        menu.exec(e.globalPos())


class SettingIconWidget(IconWidget):

    def paintEvent(self, e):
        painter = QPainter(self)

        if not self.isEnabled():
            painter.setOpacity(0.36)

        painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        drawIcon(self._icon, painter, self.rect())


class SettingCard(QFrame):
    def __init__(self, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
        super().__init__(parent=parent)
        self.iconLabel = SettingIconWidget(icon, self)
        self.titleLabel = QLabel(title, self)
        self.contentLabel = QLabel(content or '', self)
        self.hBoxLayout = QHBoxLayout(self)
        self.vBoxLayout = QVBoxLayout()

        if not content:
            self.contentLabel.hide()

        self.setFixedHeight(70 if content else 50)
        self.iconLabel.setFixedSize(16, 16)

        # initialize layout
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(16, 0, 0, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.setAlignment(Qt.AlignVCenter)

        self.hBoxLayout.addWidget(self.iconLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)

        self.hBoxLayout.addLayout(self.vBoxLayout)
        self.vBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.vBoxLayout.addWidget(self.contentLabel, 0, Qt.AlignLeft)

        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)

        self.contentLabel.setObjectName('contentLabel')
        FluentStyleSheet.SETTING_CARD.apply(self)

    def setTitle(self, title: str):
        self.titleLabel.setText(title)

    def setContent(self, content: str):
        self.contentLabel.setText(content)
        self.contentLabel.setVisible(bool(content))

    def setValue(self, value):
        pass

    def setIconSize(self, width: int, height: int):
        self.iconLabel.setFixedSize(width, height)

    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)

        if isDarkTheme():
            painter.setBrush(QColor(255, 255, 255, 13))
            painter.setPen(QColor(0, 0, 0, 50))
        else:
            painter.setBrush(QColor(255, 255, 255, 170))
            painter.setPen(QColor(0, 0, 0, 19))

        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 6, 6)


class SwitchSettingCard(SettingCard):
    checkedChanged = Signal(bool)

    def __init__(self, icon: Union[str, QIcon, FIF], title, content=None,
                 configItem: ConfigItem = None, parent=None):
        """
        Parameters
        ----------
        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        configItem: ConfigItem
            configuration item operated by the card

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.switchButton = SwitchButton(
            self.tr('关'), self, IndicatorPosition.RIGHT)

        if configItem:
            self.setValue(qconfig.get(configItem))
            configItem.valueChanged.connect(self.setValue)

        # add switch button to layout
        self.hBoxLayout.addWidget(self.switchButton, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        self.switchButton.checkedChanged.connect(self.__onCheckedChanged)

    def __onCheckedChanged(self, isChecked: bool):
        self.setValue(isChecked)
        self.checkedChanged.emit(isChecked)

    def setValue(self, isChecked: bool):
        if self.configItem:
            qconfig.set(self.configItem, isChecked)

        self.switchButton.setChecked(isChecked)
        self.switchButton.setText(
            self.tr('开') if isChecked else self.tr('关'))

    def setChecked(self, isChecked: bool):
        self.setValue(isChecked)

    def isChecked(self):
        return self.switchButton.isChecked()


class RangeSettingCard(SettingCard):
    valueChanged = Signal(int)

    def __init__(self, configItem, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        configItem: RangeConfigItem
            configuration item operated by the card

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.slider = Slider(Qt.Horizontal, self)
        self.valueLabel = QLabel(self)
        self.slider.setMinimumWidth(268)

        self.slider.setSingleStep(1)
        self.slider.setRange(*configItem.range)
        self.slider.setValue(configItem.value)
        self.valueLabel.setText(str(configItem.value / 10) + ' s')

        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.valueLabel, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(6)
        self.hBoxLayout.addWidget(self.slider, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        self.valueLabel.setObjectName('valueLabel')
        configItem.valueChanged.connect(self.setValue)
        self.slider.valueChanged.connect(self.__onValueChanged)

    def __onValueChanged(self, value: int):
        self.setValue(value)
        self.valueChanged.emit(value)

    def setValue(self, value):
        qconfig.set(self.configItem, value)
        self.valueLabel.setText(str(value / 10) + ' s')
        self.valueLabel.adjustSize()
        self.slider.setValue(value)


class PushSettingCard(SettingCard):
    clicked = Signal()

    def __init__(self, text, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        text: str
            the text of push button

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.button = QPushButton(text, self)
        self.hBoxLayout.addWidget(self.button, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)
        self.button.clicked.connect(self.clicked)


class PrimaryPushSettingCard(PushSettingCard):
    def __init__(self, text, icon, title, content=None, parent=None):
        super().__init__(text, icon, title, content, parent)
        self.button.setObjectName('primaryButton')


class HyperlinkCard(SettingCard):
    def __init__(self, url, text, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        url: str
            the url to be opened

        text: str
            text of url

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        text: str
            the text of push button

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.linkButton = HyperlinkButton(url, text, self)
        self.hBoxLayout.addWidget(self.linkButton, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)


class SpinBoxSettingCard(SettingCard):
    valueChanged = Signal(int)

    def __init__(self, configItem: ConfigItem, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        configItem: ConfigItem
            configuration item operated by the card

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.spinBox = SpinBox(self)
        self.spinBox.setFixedWidth(130)
        self.spinBox.setAccelerated(True)
        self.spinBox.setMaximum(5)
        self.spinBox.setMinimum(1)
        self.spinBox.setValue(configItem.value)

        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.spinBox, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        configItem.valueChanged.connect(self.setValue)
        self.spinBox.valueChanged.connect(self.__onValueChanged)

    def __onValueChanged(self, value: int):
        self.setValue(value)
        self.valueChanged.emit(value)

    def setValue(self, value):
        qconfig.set(self.configItem, value)
        self.spinBox.setValue(value)


class ComboBoxSettingCard(SettingCard):
    def __init__(self, configItem: OptionsConfigItem, icon: Union[str, QIcon, FIF], title, content=None, texts=None, parent=None):
        """
        Parameters
        ----------
        configItem: OptionsConfigItem
            configuration item operated by the card

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        texts: List[str]
            the text of items

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.comboBox = ComboBox(self)
        self.hBoxLayout.addWidget(self.comboBox, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        self.optionToText = {o: t for o, t in zip(configItem.options, texts)}
        for text, option in zip(texts, configItem.options):
            self.comboBox.addItem(text, userData=option)

        self.comboBox.setCurrentText(self.optionToText[qconfig.get(configItem)])
        self.comboBox.currentIndexChanged.connect(self._onCurrentIndexChanged)
        configItem.valueChanged.connect(self.setValue)

    def _onCurrentIndexChanged(self, index: int):

        qconfig.set(self.configItem, self.comboBox.itemData(index))

    def setValue(self, value):
        if value not in self.optionToText:
            return

        self.comboBox.setCurrentText(self.optionToText[value])
        qconfig.set(self.configItem, value)


class OptionsSettingCard(ExpandSettingCard):
    optionChanged = Signal(OptionsConfigItem)

    def __init__(self, configItem, icon: Union[str, QIcon, FIF], title, content=None, texts=None, parent=None):
        """
        Parameters
        ----------
        configItem: OptionsConfigItem
            options config item

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of setting card

        content: str
            the content of setting card

        texts: List[str]
            the texts of radio buttons

        parent: QWidget
            parent window
        """
        super().__init__(icon, title, content, parent)
        self.texts = texts or []
        self.configItem = configItem
        self.configName = configItem.name
        self.choiceLabel = QLabel(self)
        self.buttonGroup = QButtonGroup(self)

        self.addWidget(self.choiceLabel)

        self.viewLayout.setSpacing(19)
        self.viewLayout.setContentsMargins(48, 18, 0, 18)
        for text, option in zip(texts, configItem.options):
            button = RadioButton(text, self.view)
            self.buttonGroup.addButton(button)
            self.viewLayout.addWidget(button)
            button.setProperty(self.configName, option)

        self._adjustViewSize()
        self.setValue(qconfig.get(self.configItem))
        configItem.valueChanged.connect(self.setValue)
        self.buttonGroup.buttonClicked.connect(self.__onButtonClicked)

    def __onButtonClicked(self, button: RadioButton):
        if button.text() == self.choiceLabel.text():
            return

        value = button.property(self.configName)
        qconfig.set(self.configItem, value)

        self.choiceLabel.setText(button.text())
        self.choiceLabel.adjustSize()
        self.optionChanged.emit(self.configItem)

    def setValue(self, value):
        qconfig.set(self.configItem, value)

        for button in self.buttonGroup.buttons():
            isChecked = button.property(self.configName) == value
            button.setChecked(isChecked)

            if isChecked:
                self.choiceLabel.setText(button.text())
                # self.choiceLabel.adjustSize()


class FolderItem(QWidget):
    def __init__(self, folder: str, parent=None):
        super().__init__(parent=parent)
        self.folder = folder
        self.hBoxLayout = QHBoxLayout(self)
        self.folderLabel = QLabel(folder, self)
        self.changeButton = HyperlinkButton(self)
        self.changeButton.setText("更改")

        self.setFixedHeight(53)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.hBoxLayout.setContentsMargins(48, 0, 60, 0)
        self.hBoxLayout.addWidget(self.folderLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.changeButton, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)

    def setFolder(self, text):
        self.folderLabel.setText(text)


class CustomFolderListSettingCard(ExpandSettingCard):
    folderChanged = Signal(list)

    def __init__(self, title: str, content: str = None, directory=cfg.sourceFolder.value, parent=None):
        """
        Parameters
        ----------
        title: str
            the title of card

        content: str
            the content of card

        directory: str
            working directory of file dialog

        parent: QWidget
            parent widget
        """
        super().__init__(FIF.FOLDER_ADD, title, content, parent)
        self._dialogDirectory = directory
        self.__initWidget()

    def __initWidget(self):
        self.viewLayout.setSpacing(0)
        self.viewLayout.setAlignment(Qt.AlignTop)
        self.viewLayout.setContentsMargins(0, 0, 0, 0)

        self.yuwenItem = FolderItem("语文: " + cfg.yuwenFolder.value, self.view)
        self.shuxueItem = FolderItem("数学: " + cfg.shuxueFolder.value, self.view)
        self.yingyuItem = FolderItem("英语: " + cfg.yingyuFolder.value, self.view)
        self.wuliItem = FolderItem("物理: " + cfg.wuliFolder.value, self.view)
        self.huaxueItem = FolderItem("化学: " + cfg.huaxueFolder.value, self.view)
        self.shengwuItem = FolderItem("生物: " + cfg.shengwuFolder.value, self.view)
        self.zhengzhiItem = FolderItem("政治: " + cfg.zhengzhiFolder.value, self.view)
        self.lishiItem = FolderItem("历史: " + cfg.lishiFolder.value, self.view)
        self.diliItem = FolderItem("地理: " + cfg.diliFolder.value, self.view)
        self.jishuItem = FolderItem("技术: " + cfg.jishuFolder.value, self.view)
        self.ziliaoItem = FolderItem("资料: " + cfg.ziliaoFolder.value, self.view)
        self.yuwenItem.changeButton.clicked.connect(lambda: self.showFolderDialog(1))
        self.shuxueItem.changeButton.clicked.connect(lambda: self.showFolderDialog(2))
        self.yingyuItem.changeButton.clicked.connect(lambda: self.showFolderDialog(3))
        self.wuliItem.changeButton.clicked.connect(lambda: self.showFolderDialog(4))
        self.huaxueItem.changeButton.clicked.connect(lambda: self.showFolderDialog(5))
        self.shengwuItem.changeButton.clicked.connect(lambda: self.showFolderDialog(6))
        self.zhengzhiItem.changeButton.clicked.connect(lambda: self.showFolderDialog(7))
        self.lishiItem.changeButton.clicked.connect(lambda: self.showFolderDialog(8))
        self.diliItem.changeButton.clicked.connect(lambda: self.showFolderDialog(9))
        self.jishuItem.changeButton.clicked.connect(lambda: self.showFolderDialog(10))
        self.ziliaoItem.changeButton.clicked.connect(lambda: self.showFolderDialog(11))

        self.viewLayout.addWidget(self.yuwenItem)
        self.viewLayout.addWidget(self.shuxueItem)
        self.viewLayout.addWidget(self.yingyuItem)
        self.viewLayout.addWidget(self.wuliItem)
        self.viewLayout.addWidget(self.huaxueItem)
        self.viewLayout.addWidget(self.shengwuItem)
        self.viewLayout.addWidget(self.zhengzhiItem)
        self.viewLayout.addWidget(self.lishiItem)
        self.viewLayout.addWidget(self.diliItem)
        self.viewLayout.addWidget(self.jishuItem)
        self.viewLayout.addWidget(self.ziliaoItem)

        self._adjustViewSize()

    def showFolderDialog(self, index):
        folder = QFileDialog.getExistingDirectory(self, self.tr("选择文件夹"), self._dialogDirectory)
        if not folder:
            return

        if index==1:
            cfg.set(cfg.yuwenFolder, folder)
            self.yuwenItem.setFolder("语文: " + cfg.yuwenFolder.value)
        elif index==2:
            cfg.set(cfg.shuxueFolder, folder)
            self.shuxueItem.setFolder("数学: " + cfg.shuxueFolder.value)
        elif index==3:
            cfg.set(cfg.yingyuFolder, folder)
            self.yingyuItem.setFolder("英语: " + cfg.yingyuFolder.value)
        elif index==4:
            cfg.set(cfg.wuliFolder, folder)
            self.wuliItem.setFolder("物理: " + cfg.wuliFolder.value)
        elif index==5:
            cfg.set(cfg.huaxueFolder, folder)
            self.huaxueItem.setFolder("化学: " + cfg.huaxueFolder.value)
        elif index==6:
            cfg.set(cfg.shengwuFolder, folder)
            self.shengwuItem.setFolder("生物: " + cfg.shengwuFolder.value)
        elif index==7:
            cfg.set(cfg.zhengzhiFolder, folder)
            self.zhengzhiItem.setFolder("政治: " + cfg.zhengzhiFolder.value)
        elif index==8:
            cfg.set(cfg.lishiFolder, folder)
            self.lishiItem.setFolder("历史: " + cfg.lishiFolder.value)
        elif index==9:
            cfg.set(cfg.lishiFolder, folder)
            self.diliItem.setFolder("地理: " + cfg.diliFolder.value)
        elif index==10:
            cfg.set(cfg.jishuFolder, folder)
            self.jishuItem.setFolder("技术: " + cfg.jishuFolder.value)
        else:
            cfg.set(cfg.ziliaoFolder, folder)
            self.ziliaoItem.setFolder("资料: " + cfg.zilaioFolder.value)

    def updateContent(self):
        self.yuwenItem.setFolder("语文: " + cfg.yuwenFolder.value)
        self.shuxueItem.setFolder("数学: " + cfg.shuxueFolder.value)
        self.yingyuItem.setFolder("英语: " + cfg.yingyuFolder.value)
        self.wuliItem.setFolder("物理: " + cfg.wuliFolder.value)
        self.huaxueItem.setFolder("化学: " + cfg.huaxueFolder.value)
        self.shengwuItem.setFolder("生物: " + cfg.shengwuFolder.value)
        self.zhengzhiItem.setFolder("政治: " + cfg.zhengzhiFolder.value)
        self.lishiItem.setFolder("历史: " + cfg.lishiFolder.value)
        self.diliItem.setFolder("地理: " + cfg.diliFolder.value)
        self.jishuItem.setFolder("技术: " + cfg.jishuFolder.value)
        self.ziliaoItem.setFolder("资料: " + cfg.ziliaoFolder.value)


class ClearCache(QThread):
    isFinished = Signal(bool)
    def __init__(self):
        super(ClearCache, self).__init__()

    def run(self):
        if os.path.exists('./Log'):
            subprocess.call("del /s /q Log", shell=True)
        if os.path.exists('FastCopy2.ini'):
            subprocess.call("del /q FastCopy2.ini", shell=True)
        self.isFinished.emit(True)


class HomeInterface(SmoothScrollArea):
    sourceFolderChanged = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.scrollWidget = QWidget()
        self.stateTooltip = None
        self.expandLayout = ExpandLayout(self.scrollWidget)
        self.sourceGroup = SettingCardGroup(self.tr('源'), self.scrollWidget)
        self.actGroup = SettingCardGroup(self.tr('行为'), self.scrollWidget)
        self.performanceGroup = SettingCardGroup(self.tr('性能'), self.scrollWidget)
        self.storageGroup = SettingCardGroup(self.tr('存储'), self.scrollWidget)
        self.advanceGroup = SettingCardGroup(self.tr('高级'), self.scrollWidget)
        self.optionSourceCard = ComboBoxSettingCard(
            cfg.IsSourceCloud,
            FIF.FOLDER,
            self.tr('源文件夹'),
            self.tr('选择源文件夹来源'),
            texts=['云上春晖', '自定义'],
            parent=self.sourceGroup)
        self.autoRunCard = SwitchSettingCard(
            FIF.POWER_BUTTON,
            self.tr("开机时启动"),
            self.tr(""),
            configItem=cfg.AutoRun,
            parent=self.actGroup)
        self.notifyCard = SwitchSettingCard(
            FIF.RINGER,
            self.tr("完成后通知"),
            self.tr(""),
            configItem=cfg.Notify,
            parent=self.actGroup)
        self.cloudCard = PushSettingCard(
            self.tr('选择文件夹'),
            FIF.CLOUD,
            self.tr("云上春晖"),
            cfg.get(cfg.sourceFolder),
            self.sourceGroup)
        self.customFolderCard = CustomFolderListSettingCard(
            self.tr("自定义"),
            self.tr("展开选项卡以设置"),
            directory=cfg.sourceFolder.value,
            parent=self.sourceGroup)
        self.scanCycleCard = RangeSettingCard(
            cfg.ScanCycle,
            FIF.STOP_WATCH,
            self.tr('扫描周期'),
            parent=self.performanceGroup)
        self.concurrentProcessCard = SpinBoxSettingCard(
            cfg.ConcurrentProcess,
            FIF.ALIGNMENT,
            self.tr('并行进程数'),
            parent=self.performanceGroup)
        self.bufSizeCard = OptionsSettingCard(
            cfg.BufSize,
            FIF.PIE_SINGLE,
            self.tr('缓冲区大小'),
            texts=[
                self.tr('32 MB'), self.tr('64 MB'),
                self.tr('128 MB'), self.tr('256 MB'),
                self.tr('512 MB'), self.tr('1 GB')],
            parent=self.performanceGroup)
        self.clearCard = PushSettingCard(
            self.tr('清除'),
            FIF.BROOM,
            self.tr('清除缓存'),
            self.tr(self.getSize()),
            self.storageGroup)
        self.recoverCard = PushSettingCard(
            self.tr('恢复'),
            FIF.CLEAR_SELECTION,
            self.tr('恢复默认设置'),
            self.tr('重置所有参数为初始值'),
            self.advanceGroup)
        self.devCard = PushSettingCard(
            self.tr('打开'),
            FIF.DEVELOPER_TOOLS,
            self.tr('开发者选项'),
            self.tr('打开配置文件'),
            self.advanceGroup)
        self.helpCard = PrimaryPushSettingCard(
            self.tr('转到帮助'),
            FIF.HELP,
            self.tr('帮助'),
            self.tr('提示与常见问题'),
            self.advanceGroup)
        self.__initWidget()

    def __initWidget(self):
        self.resize(1000, 800)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, 0, 0, 0)
        self.setWidget(self.scrollWidget)
        self.setWidgetResizable(True)
        self.__initLayout()
        self.__connectSignalToSlot()
        self.onOptionSourceCard()

    def __initLayout(self):
        self.sourceGroup.addSettingCard(self.optionSourceCard)
        self.sourceGroup.addSettingCard(self.cloudCard)
        self.sourceGroup.addSettingCard(self.customFolderCard)
        self.actGroup.addSettingCard(self.autoRunCard)
        self.actGroup.addSettingCard(self.notifyCard)
        self.performanceGroup.addSettingCard(self.scanCycleCard)
        self.performanceGroup.addSettingCard(self.concurrentProcessCard)
        self.performanceGroup.addSettingCard(self.bufSizeCard)
        self.storageGroup.addSettingCard(self.clearCard)
        self.advanceGroup.addSettingCard(self.recoverCard)
        self.advanceGroup.addSettingCard(self.devCard)
        self.advanceGroup.addSettingCard(self.helpCard)
        self.expandLayout.setSpacing(28)
        self.expandLayout.setContentsMargins(60, 10, 60, 0)
        self.expandLayout.addWidget(self.sourceGroup)
        self.expandLayout.addWidget(self.actGroup)
        self.expandLayout.addWidget(self.performanceGroup)
        self.expandLayout.addWidget(self.storageGroup)
        self.expandLayout.addWidget(self.advanceGroup)

    def noSourceFolderDialog(self):
        w = MessageBox(
            '警告',
            '未设置班级文件夹',
            self)
        w.yesButton.setText('设置')
        w.cancelButton.hide()
        if w.exec():
            self.__onCloudCardClicked()

    def getSize(self):
        try:
            size = os.path.getsize('FastCopy2.ini')
        except:
            size = 0
        if os.path.exists('./Log'):
            for root, dirs, files in os.walk('./Log'):
                try:
                    size += sum([os.path.getsize(os.path.join(root, name)) for name in files])
                except:
                    pass
        kbSize = float(size / 1024)
        if kbSize >= 1024*1024:
            return str(round(kbSize/1024/1024, 1)) + ' GB'
        elif kbSize >= 1024:
            return str(round(kbSize/1024, 1)) + ' MB'
        else:
            return str(int(kbSize)) + ' KB'

    def onOptionSourceCard(self):
        if self.optionSourceCard.comboBox.text() == "云上春晖":
            self.cloudCard.setDisabled(False)
            self.customFolderCard.setDisabled(True)

            folder = cfg.sourceFolder.value
            cfg.set(cfg.yuwenFolder, os.path.join(folder, '语文'))
            cfg.set(cfg.shuxueFolder, os.path.join(folder, '数学'))
            cfg.set(cfg.yingyuFolder, os.path.join(folder, '英语'))
            cfg.set(cfg.wuliFolder, os.path.join(folder, '物理'))
            cfg.set(cfg.huaxueFolder, os.path.join(folder, '化学'))
            cfg.set(cfg.shengwuFolder, os.path.join(folder, '生物'))
            cfg.set(cfg.zhengzhiFolder, os.path.join(folder, '政治'))
            cfg.set(cfg.lishiFolder, os.path.join(folder, '历史'))
            cfg.set(cfg.diliFolder, os.path.join(folder, '地理'))
            cfg.set(cfg.jishuFolder, os.path.join(folder, '技术'))
            cfg.set(cfg.ziliaoFolder, os.path.join(folder, '资料'))
        else:
            self.cloudCard.setDisabled(True)
            self.customFolderCard.setDisabled(False)

    def __onCloudCardClicked(self):
        folder = QFileDialog.getExistingDirectory(self, self.tr("选择文件夹"), "./")
        if not folder or cfg.get(cfg.sourceFolder) == folder:
            return

        self.cloudCard.setContent(folder)
        cfg.set(cfg.sourceFolder, folder)
        cfg.set(cfg.yuwenFolder, os.path.join(folder, '语文'))
        cfg.set(cfg.shuxueFolder, os.path.join(folder, '数学'))
        cfg.set(cfg.yingyuFolder, os.path.join(folder, '英语'))
        cfg.set(cfg.wuliFolder, os.path.join(folder, '物理'))
        cfg.set(cfg.huaxueFolder, os.path.join(folder, '化学'))
        cfg.set(cfg.shengwuFolder, os.path.join(folder, '生物'))
        cfg.set(cfg.zhengzhiFolder, os.path.join(folder, '政治'))
        cfg.set(cfg.lishiFolder, os.path.join(folder, '历史'))
        cfg.set(cfg.diliFolder, os.path.join(folder, '地理'))
        cfg.set(cfg.jishuFolder, os.path.join(folder, '技术'))
        cfg.set(cfg.ziliaoFolder, os.path.join(folder, '资料'))
        self.customFolderCard.updateContent()

    def clearFinished(self):
        self.clearCacheThread.exit(0)
        self.clearCard.contentLabel.setText(self.getSize())
        self.clearCard.button.setText('已清除')
        QTimer.singleShot(2000, lambda: self.clearCard.button.setText('清除'))
        self.clearCard.button.setDisabled(False)

    def clearCache(self):
        w = MessageBox(
            '清除缓存',
            '缓存包含日志文件，确定清除吗？',
            self)
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            self.clearCard.button.setText('清除中')
            self.clearCard.button.setDisabled(True)
            self.clearCacheThread = ClearCache()
            self.clearCacheThread.start()
            self.clearCacheThread.isFinished.connect(self.clearFinished)

    def recoverConfig(self):
        w = MessageBox(
            '恢复默认设置',
            '确定要重置所有设置吗？',
            self)
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            self.autoRunCard.setChecked(True)
            self.notifyCard.setChecked(True)
            self.scanCycleCard.setValue(10)
            self.concurrentProcessCard.setValue(3)
            self.bufSizeCard.setValue(BufSize._256)

    def openConfig(self):
        w = MessageBox(
            '打开配置文件',
            '随意修改参数可能导致 Express 无法正常运行，确定继续吗？',
            self)
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            os.startfile("config\\config.json")

    def __connectSignalToSlot(self):
        self.optionSourceCard.comboBox.currentTextChanged.connect(self.onOptionSourceCard)
        self.cloudCard.clicked.connect(self.__onCloudCardClicked)
        self.clearCard.clicked.connect(self.clearCache)
        self.recoverCard.clicked.connect(self.recoverConfig)
        self.devCard.clicked.connect(self.openConfig)


class DetailMessageBox(MessageBoxBase):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleLabel = SubtitleLabel('关于 Express', self)
        self.textBox = TextBrowser(self)
        self.textBox.setText(
            f'Express\n版本 {VERSION}\nCopyright © {YEAR} BUG STUDIO\n\n' +
            'MIT License\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n' +
            'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n' +
            'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\n'
        )

        self.githubBtn = HyperlinkButton(self)
        self.websiteBtn = HyperlinkButton(self)
        self.onlineDocBtn = HyperlinkButton(self)
        self.githubBtn.setText('源代码')
        self.websiteBtn.setText('网站主页')
        self.onlineDocBtn.setText('在线文档')
        self.githubBtn.setIcon(FIF.GITHUB)
        self.websiteBtn.setIcon(FIF.GLOBE)
        self.onlineDocBtn.setIcon(FIF.DOCUMENT)
        self.githubBtn.clicked.connect(lambda: webopen("https://github.com/sudo0015/Express"))
        self.websiteBtn.clicked.connect(lambda: webopen("https://sudo0015.github.io/"))
        self.onlineDocBtn.clicked.connect(lambda: webopen("https://github.com/sudo0015/Express"))

        self.btnLayout = QHBoxLayout(self)
        self.btnLayout.addWidget(self.githubBtn)
        self.btnLayout.addWidget(self.websiteBtn)
        self.btnLayout.addWidget(self.onlineDocBtn)

        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addWidget(self.textBox)
        self.viewLayout.addLayout(self.btnLayout)

        self.yesButton.setText('确定')
        self.hideCancelButton()

        self.widget.setMinimumWidth(350)


class AboutInterface(SmoothScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.scrollWidget = QWidget()
        self.stateTooltip = None
        self.expandLayout = ExpandLayout(self.scrollWidget)
        self.aboutGroup = SettingCardGroup(self.tr(''), self.scrollWidget)
        self.aboutESCard = PushSettingCard(
            self.tr('详细信息'),
            FIF.INFO,
            self.tr('关于 Express'),
            self.tr(f'版本 {VERSION}'),
            self.aboutGroup)
        self.aboutBSCard = PushSettingCard(
            self.tr('了解更多'),
            FIF.PEOPLE,
            self.tr('关于作者'),
            self.tr('BUG STUDIO'),
            self.aboutGroup)
        self.helpCard = PrimaryPushSettingCard(
            self.tr('转到帮助'),
            FIF.HELP,
            self.tr('帮助'),
            self.tr('提示与常见问题'),
            self.aboutGroup)
        self.feedbackCard = PrimaryPushSettingCard(
            self.tr('提供反馈'),
            FIF.FEEDBACK,
            self.tr('反馈'),
            self.tr('报告问题或提出建议'),
            self.aboutGroup)
        self.__initWidget()

    def __initWidget(self):
        self.resize(1000, 800)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, 0, 0, 0)
        self.setWidget(self.scrollWidget)
        self.setWidgetResizable(True)

        self.imgLabel = ImageLabel(self)
        if isDarkTheme():
            self.imgLabel.setImage(':/BannerDark.png')
        else:
            self.imgLabel.setImage(':/BannerLight.png')
        self.imgLabel.setFixedSize(401, 150)

        self.__initLayout()
        self.__connectSignalToSlot()

    def __initLayout(self):
        self.aboutGroup.addSettingCard(self.aboutESCard)
        self.aboutGroup.addSettingCard(self.aboutBSCard)
        self.aboutGroup.addSettingCard(self.helpCard)
        self.aboutGroup.addSettingCard(self.feedbackCard)
        self.expandLayout.setContentsMargins(60, 10, 60, 0)
        self.expandLayout.addWidget(self.imgLabel)
        self.expandLayout.addWidget(self.aboutGroup)

    def onAboutESCardClicked(self):
        w = DetailMessageBox(self)
        if w.exec():
            pass

    def onAboutBSCardClicked(self):
        print("111")

    def onFeedbackCardClicked(self):
        webopen("https://github.com/sudo0015/Express/issues")

    def onHelpCardClicked(self):
        print("111")

    def __connectSignalToSlot(self):
        self.aboutESCard.clicked.connect(self.onAboutESCardClicked)
        self.aboutBSCard.clicked.connect(self.onAboutBSCardClicked)
        self.helpCard.clicked.connect(self.onHelpCardClicked)
        self.feedbackCard.clicked.connect(self.onFeedbackCardClicked)


class Main(MSFluentWindow):

    def __init__(self, parent=None):
        super().__init__(parent)
        setThemeColor(QColor(113, 89, 249))
        self.homeInterface = HomeInterface(self)
        self.aboutInterface = AboutInterface(self)
        self.homeInterface.setObjectName('homeInterface')
        self.aboutInterface.setObjectName('aboutInterface')
        self.addSubInterface(self.homeInterface, FIF.HOME, '设置', FIF.HOME_FILL)
        self.navigationInterface.addItem(
            routeKey='Log',
            icon=FIF.DOCUMENT,
            text='日志',
            onClick=self.onLogBtn,
            selectable=False,
            position=NavigationItemPosition.TOP, )
        self.navigationInterface.addItem(
            routeKey='Help',
            icon=FIF.HELP,
            text='帮助',
            onClick=self.onHelpBtn,
            selectable=False,
            position=NavigationItemPosition.BOTTOM, )
        self.addSubInterface(self.aboutInterface, FIF.INFO, '关于', FIF.INFO, NavigationItemPosition.BOTTOM)
        self.navigationInterface.setCurrentItem(self.homeInterface.objectName())
        self.homeInterface.helpCard.clicked.connect(self.onHelpBtn)
        self.resize(800, 600)
        self.setWindowTitle('Express 设置')
        self.setWindowIcon(QIcon(':/icon.png'))
        self.titleBar.raise_()
        desktop = QApplication.screens()[0].size()
        self.move(desktop.width() // 2 - self.width() // 2, desktop.height() // 2 - self.height() // 2)

    def onHelpBtn(self):
        print("clicked")

    def onLogBtn(self):
        try:
            os.startfile('Log')
        except:
            pass


if __name__ == '__main__':
    with Mutex():
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
        app = QApplication(sys.argv)
        w = Main()
        w.show()
        app.exec()
