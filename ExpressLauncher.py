import os
import sys
import msvcrt
import darkdetect
import subprocess
import ExpressRes
from config import cfg
from typing import Union, Iterable
from pygetwindow import getWindowsWithTitle as GetWindow
from PySide6.QtCore import Qt, QPoint, Signal, QEvent, Property, QPropertyAnimation, QTimer, QRectF
from PySide6.QtGui import QColor, QPainter, QIcon, QAction, QPainterPath, QCursor
from PySide6.QtWidgets import QApplication, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QCompleter, QFileDialog, \
    QLineEdit
from qfluentwidgets import setTheme, Theme, isDarkTheme, setThemeColor, FluentStyleSheet, \
    PillPushButton, TransparentToolButton, RoundMenu, EditableComboBox, HyperlinkButton, PrimaryPushButton, \
    IndicatorPosition, SwitchButton, ThemeColor, ToolButton, themeColor, PushButton, InfoBar, InfoBarIcon, \
    InfoBarPosition, setFont, LineEditButton, MenuAnimationType
from qfluentwidgets.common.overload import singledispatchmethod
from qfluentwidgets.components.widgets.combo_box import ComboItem, ComboBoxMenu
from qfluentwidgets.components.widgets.line_edit import CompleterMenu
from qframelesswindow import TitleBar
from qfluentwidgets import FluentIcon as FIF


def isWin11():
    return sys.platform == 'win32' and sys.getwindowsversion().build >= 22000


if isWin11():
    from qframelesswindow import AcrylicWindow as Window
else:
    from qframelesswindow import FramelessWindow as Window


class Mutex:
    def __init__(self):
        self.lockfile = None

    def __enter__(self):
        self.lockfile = open('ExpressLauncher.lockfile', 'w')
        try:
            msvcrt.locking(self.lockfile.fileno(), msvcrt.LK_NBLCK, 1)
        except IOError:
            try:
                window = GetWindow("ExpressLauncher")[0]
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
            os.remove('ExpressLauncher.lockfile')


class EditMenu(RoundMenu):
    """ Edit menu """

    def createActions(self):
        self.cutAct = QAction(
            FIF.CUT.icon(),
            self.tr("剪切"),
            self,
            shortcut="Ctrl+X",
            triggered=self.parent().cut,
        )
        self.copyAct = QAction(
            FIF.COPY.icon(),
            self.tr("复制"),
            self,
            shortcut="Ctrl+C",
            triggered=self.parent().copy,
        )
        self.pasteAct = QAction(
            FIF.PASTE.icon(),
            self.tr("粘贴"),
            self,
            shortcut="Ctrl+V",
            triggered=self.parent().paste,
        )
        self.cancelAct = QAction(
            FIF.CANCEL.icon(),
            self.tr("Cancel"),
            self,
            shortcut="Ctrl+Z",
            triggered=self.parent().undo,
        )
        self.selectAllAct = QAction(
            self.tr("全选"),
            self,
            shortcut="Ctrl+A",
            triggered=self.parent().selectAll
        )
        self.action_list = [
            self.cutAct, self.copyAct,
            self.pasteAct, self.cancelAct, self.selectAllAct
        ]

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


class FluentTitleBar(TitleBar):
    def __init__(self, parent):
        super().__init__(parent)
        self.setFixedHeight(45)
        self.hBoxLayout.removeWidget(self.minBtn)
        self.hBoxLayout.removeWidget(self.maxBtn)
        self.hBoxLayout.removeWidget(self.closeBtn)
        self.maxBtn.setVisible(False)
        self.titleLabel = QLabel(self)
        self.titleLabel.setText('Express')
        self.titleLabel.setObjectName('titleLabel')
        self.vBoxLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.setSpacing(0)
        self.buttonLayout.setContentsMargins(0, 0, 0, 0)
        self.buttonLayout.setAlignment(Qt.AlignTop)
        self.buttonLayout.addWidget(self.minBtn)
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


class MicaWindow(Window):

    def __init__(self):
        super().__init__()
        self.setTitleBar(FluentTitleBar(self))
        if isWin11():
            self.windowEffect.setMicaEffect(self.winId(), isDarkTheme())


class LineEdit(QLineEdit):
    """ Line edit """

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self._isClearButtonEnabled = False
        self._completer = None  # type: QCompleter
        self._completerMenu = None  # type: CompleterMenu
        self._isError = False

        self.leftButtons = []   # type: List[LineEditButton]
        self.rightButtons = []  # type: List[LineEditButton]

        self.setProperty("transparent", True)
        FluentStyleSheet.LINE_EDIT.apply(self)
        self.setFixedHeight(33)
        self.setAttribute(Qt.WA_MacShowFocusRect, False)
        setFont(self)

        self.hBoxLayout = QHBoxLayout(self)
        self.clearButton = LineEditButton(FIF.CLOSE, self)

        self.clearButton.setFixedSize(29, 25)
        self.clearButton.hide()

        self.hBoxLayout.setSpacing(3)
        self.hBoxLayout.setContentsMargins(4, 4, 4, 4)
        self.hBoxLayout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.hBoxLayout.addWidget(self.clearButton, 0, Qt.AlignRight)

        self.clearButton.clicked.connect(self.clear)
        self.textChanged.connect(self.__onTextChanged)
        self.textEdited.connect(self.__onTextEdited)

    def isError(self):
        return self._isError

    def setError(self, isError: bool):
        if isError == self.isError():
            return

        self._isError = isError
        self.update()

    def focusedBorderColor(self):
        if not self.isError():
            return themeColor()

        return QColor("#ff99a4") if isDarkTheme() else QColor("#c42b1c")

    def setClearButtonEnabled(self, enable: bool):
        self._isClearButtonEnabled = enable
        self._adjustTextMargins()

    def isClearButtonEnabled(self) -> bool:
        return self._isClearButtonEnabled

    def setCompleter(self, completer: QCompleter):
        self._completer = completer

    def completer(self):
        return self._completer

    def addAction(self, action: QAction, position=QLineEdit.ActionPosition.TrailingPosition):
        QWidget.addAction(self, action)

        button = LineEditButton(action.icon())
        button.setAction(action)
        button.setFixedWidth(29)

        if position == QLineEdit.ActionPosition.LeadingPosition:
            self.hBoxLayout.insertWidget(len(self.leftButtons), button, 0, Qt.AlignLeading)
            if not self.leftButtons:
                self.hBoxLayout.insertStretch(1, 1)

            self.leftButtons.append(button)
        else:
            self.rightButtons.append(button)
            self.hBoxLayout.addWidget(button, 0, Qt.AlignRight)

        self._adjustTextMargins()

    def addActions(self, actions, position=QLineEdit.ActionPosition.TrailingPosition):
        for action in actions:
            self.addAction(action, position)

    def _adjustTextMargins(self):
        left = len(self.leftButtons) * 30
        right = len(self.rightButtons) * 30 + 28 * self.isClearButtonEnabled()
        m = self.textMargins()
        self.setTextMargins(left, m.top(), right, m.bottom())

    def focusOutEvent(self, e):
        super().focusOutEvent(e)
        self.clearButton.hide()

    def focusInEvent(self, e):
        super().focusInEvent(e)
        if self.isClearButtonEnabled():
            self.clearButton.setVisible(bool(self.text()))

    def __onTextChanged(self, text):
        """ text changed slot """
        if self.isClearButtonEnabled():
            self.clearButton.setVisible(bool(text) and self.hasFocus())

    def __onTextEdited(self, text):
        if not self.completer():
            return

        if self.text():
            QTimer.singleShot(50, self._showCompleterMenu)
        elif self._completerMenu:
            self._completerMenu.close()

    def setCompleterMenu(self, menu):
        """ set completer menu

        Parameters
        ----------
        menu: CompleterMenu
            completer menu
        """
        menu.activated.connect(self._completer.activated)
        self._completerMenu = menu

    def _showCompleterMenu(self):
        if not self.completer() or not self.text():
            return

        # create menu
        if not self._completerMenu:
            self.setCompleterMenu(CompleterMenu(self))

        # add menu items
        self.completer().setCompletionPrefix(self.text())
        changed = self._completerMenu.setCompletion(self.completer().completionModel())
        self._completerMenu.setMaxVisibleItems(self.completer().maxVisibleItems())

        # show menu
        if changed:
            self._completerMenu.popup()

    def contextMenuEvent(self, e):
        menu = LineEditMenu(self)
        menu.exec(e.globalPos(), ani=True)

    def paintEvent(self, e):
        super().paintEvent(e)
        if not self.hasFocus():
            return

        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)

        m = self.contentsMargins()
        path = QPainterPath()
        w, h = self.width()-m.left()-m.right(), self.height()
        path.addRoundedRect(QRectF(m.left(), h-10, w, 10), 5, 5)

        rectPath = QPainterPath()
        rectPath.addRect(m.left(), h-10, w, 8)
        path = path.subtracted(rectPath)

        painter.fillPath(path, self.focusedBorderColor())


class ComboBoxBase:
    """ Combo box base """
    activated = Signal(int)
    textActivated = Signal(str)

    def __init__(self, parent=None, **kwargs):
        pass

    def _setUpUi(self):
        self.isHover = False
        self.isPressed = False
        self.items = []     # type: List[ComboItem]
        self._currentIndex = -1
        self._maxVisibleItems = -1
        self.dropMenu = None
        self._placeholderText = ""

        FluentStyleSheet.COMBO_BOX.apply(self)
        self.installEventFilter(self)

    def addItem(self, text, icon: Union[str, QIcon, FIF] = None, userData=None):
        """ add item

        Parameters
        ----------
        text: str
            the text of item

        icon: str | QIcon | FluentIconBase
        """
        item = ComboItem(text, icon, userData)
        self.items.append(item)
        if len(self.items) == 1:
            self.setCurrentIndex(0)

    def addItems(self, texts: Iterable[str]):
        """ add items

        Parameters
        ----------
        text: Iterable[str]
            the text of item
        """
        for text in texts:
            self.addItem(text)

    def removeItem(self, index: int):
        """ Removes the item at the given index from the combobox.
        This will update the current index if the index is removed.
        """
        if not 0 <= index < len(self.items):
            return

        self.items.pop(index)

        if index < self.currentIndex():
            self.setCurrentIndex(self._currentIndex - 1)
        elif index == self.currentIndex():
            if index > 0:
                self.setCurrentIndex(self._currentIndex - 1)
            else:
                self.setText(self.itemText(0))
                self.currentTextChanged.emit(self.currentText())
                self.currentIndexChanged.emit(0)

        if self.count() == 0:
            self.clear()

    def currentIndex(self):
        return self._currentIndex

    def setCurrentIndex(self, index: int):
        """ set current index

        Parameters
        ----------
        index: int
            current index
        """
        if not 0 <= index < len(self.items) or index == self.currentIndex():
            return

        oldText = self.currentText()

        self._currentIndex = index
        self.setText(self.items[index].text)

        if oldText != self.currentText():
            self.currentTextChanged.emit(self.currentText())

        self.currentIndexChanged.emit(index)

    def setText(self, text: str):
        super().setText(text)
        self.adjustSize()

    def currentText(self):
        if not 0 <= self.currentIndex() < len(self.items):
            return ''

        return self.items[self.currentIndex()].text

    def currentData(self):
        if not 0 <= self.currentIndex() < len(self.items):
            return None

        return self.items[self.currentIndex()].userData

    def setCurrentText(self, text):
        """ set the current text displayed in combo box,
        text should be in the item list

        Parameters
        ----------
        text: str
            text displayed in combo box
        """
        if text == self.currentText():
            return

        index = self.findText(text)
        if index >= 0:
            self.setCurrentIndex(index)

    def setItemText(self, index: int, text: str):
        """ set the text of item

        Parameters
        ----------
        index: int
            the index of item

        text: str
            new text of item
        """
        if not 0 <= index < len(self.items):
            return

        self.items[index].text = text
        if self.currentIndex() == index:
            self.setText(text)

    def itemData(self, index: int):
        """ Returns the data in the given index """
        if not 0 <= index < len(self.items):
            return None

        return self.items[index].userData

    def itemText(self, index: int):
        """ Returns the text in the given index """
        if not 0 <= index < len(self.items):
            return ''

        return self.items[index].text

    def itemIcon(self, index: int):
        """ Returns the icon in the given index """
        if not 0 <= index < len(self.items):
            return QIcon()

        return self.items[index].icon

    def setItemData(self, index: int, value):
        """ Sets the data role for the item on the given index """
        if 0 <= index < len(self.items):
            self.items[index].userData = value

    def setItemIcon(self, index: int, icon: Union[str, QIcon, FIF]):
        """ Sets the data role for the item on the given index """
        if 0 <= index < len(self.items):
            self.items[index].icon = icon

    def setItemEnabled(self, index: int, isEnabled: bool):
        """ Sets the enabled status of the item on the given index """
        if 0 <= index < len(self.items):
            self.items[index].isEnabled = isEnabled

    def findData(self, data):
        """ Returns the index of the item containing the given data, otherwise returns -1 """
        for i, item in enumerate(self.items):
            if item.userData == data:
                return i

        return -1

    def findText(self, text: str):
        """ Returns the index of the item containing the given text; otherwise returns -1. """
        for i, item in enumerate(self.items):
            if item.text == text:
                return i

        return -1

    def clear(self):
        """ Clears the combobox, removing all items. """
        if self.currentIndex() >= 0:
            self.setText('')

        self.items.clear()
        self._currentIndex = -1

    def count(self):
        """ Returns the number of items in the combobox """
        return len(self.items)

    def insertItem(self, index: int, text: str, icon: Union[str, QIcon, FIF] = None, userData=None):
        """ Inserts item into the combobox at the given index. """
        item = ComboItem(text, icon, userData)
        self.items.insert(index, item)

        if index <= self.currentIndex():
            self.setCurrentIndex(self.currentIndex() + 1)

    def insertItems(self, index: int, texts: Iterable[str]):
        """ Inserts items into the combobox, starting at the index specified. """
        pos = index
        for text in texts:
            item = ComboItem(text)
            self.items.insert(pos, item)
            pos += 1

        if index <= self.currentIndex():
            self.setCurrentIndex(self.currentIndex() + pos - index)

    def setMaxVisibleItems(self, num: int):
        self._maxVisibleItems = num

    def maxVisibleItems(self):
        return self._maxVisibleItems

    def _closeComboMenu(self):
        if not self.dropMenu:
            return

        # drop menu could be deleted before this method
        try:
            self.dropMenu.close()
        except:
            pass

        self.dropMenu = None

    def _onDropMenuClosed(self):
        if sys.platform != "win32":
            self.dropMenu = None
        else:
            pos = self.mapFromGlobal(QCursor.pos())
            if not self.rect().contains(pos):
                self.dropMenu = None

    def _createComboMenu(self):
        return ComboBoxMenu(self)

    def _showComboMenu(self):
        if not self.items:
            return

        menu = self._createComboMenu()
        for item in self.items:
            action = QAction(item.icon, item.text)
            action.setEnabled(item.isEnabled)
            menu.addAction(action)

        # fixes issue #468
        menu.view.itemClicked.connect(lambda i: self._onItemClicked(self.findText(i.text().lstrip())))

        if menu.view.width() < self.width():
            menu.view.setMinimumWidth(self.width())
            menu.adjustSize()

        menu.setMaxVisibleItems(self.maxVisibleItems())
        menu.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        menu.closedSignal.connect(self._onDropMenuClosed)
        self.dropMenu = menu

        # set the selected item
        if self.currentIndex() >= 0 and self.items:
            menu.setDefaultAction(menu.actions()[self.currentIndex()])

        # determine the animation type by choosing the maximum height of view
        x = -menu.width()//2 + menu.layout().contentsMargins().left() + self.width()//2
        pd = self.mapToGlobal(QPoint(x, self.height()))
        hd = menu.view.heightForAnimation(pd, MenuAnimationType.DROP_DOWN)

        pu = self.mapToGlobal(QPoint(x, 0))
        hu = menu.view.heightForAnimation(pu, MenuAnimationType.PULL_UP)

        if hd >= hu:
            menu.view.adjustSize(pd, MenuAnimationType.DROP_DOWN)
            menu.exec(pd, aniType=MenuAnimationType.DROP_DOWN)
        else:
            menu.view.adjustSize(pu, MenuAnimationType.PULL_UP)
            menu.exec(pu, aniType=MenuAnimationType.PULL_UP)

    def _toggleComboMenu(self):
        if self.dropMenu:
            self._closeComboMenu()
        else:
            self._showComboMenu()

    def _onItemClicked(self, index):
        if not self.items[index].isEnabled:
            return

        if index != self.currentIndex():
            self.setCurrentIndex(index)

        self.activated.emit(index)
        self.textActivated.emit(self.currentText())


class EditableComboBox(LineEdit, ComboBoxBase):
    """ Editable combo box """

    currentIndexChanged = Signal(int)
    currentTextChanged = Signal(str)
    activated = Signal(int)
    textActivated = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self._setUpUi()

        self.dropButton = LineEditButton(FIF.ARROW_DOWN, self)

        self.setTextMargins(0, 0, 29, 0)
        self.dropButton.setFixedSize(30, 25)
        self.hBoxLayout.addWidget(self.dropButton, 0, Qt.AlignRight)

        self.dropButton.clicked.connect(self._toggleComboMenu)
        self.textChanged.connect(self._onComboTextChanged)
        self.returnPressed.connect(self._onReturnPressed)

        FluentStyleSheet.LINE_EDIT.apply(self)

        self.clearButton.clicked.disconnect()
        self.clearButton.clicked.connect(self._onClearButtonClicked)

    def setCompleterMenu(self, menu):
        super().setCompleterMenu(menu)
        menu.activated.connect(self.__onActivated)

    def __onActivated(self, text):
        index = self.findText(text)
        if index >= 0:
            self.setCurrentIndex(index)

    def currentText(self):
        return self.text()

    def setCurrentIndex(self, index: int):
        if index >= self.count() or index == self.currentIndex():
            return

        if index < 0:
            self._currentIndex = -1
            self.setText("")
            self.setPlaceholderText(self._placeholderText)
        else:
            self._currentIndex = index
            self.setText(self.items[index].text)

    def clear(self):
        ComboBoxBase.clear(self)

    def setPlaceholderText(self, text: str):
        self._placeholderText = text
        super().setPlaceholderText(text)

    def _onReturnPressed(self):
        if not self.text():
            return

        index = self.findText(self.text())
        if index >= 0 and index != self.currentIndex():
            self._currentIndex = index
            self.currentIndexChanged.emit(index)
        elif index == -1:
            self.addItem(self.text())
            self.setCurrentIndex(self.count() - 1)

    def eventFilter(self, obj, e: QEvent):
        if obj is self:
            if e.type() == QEvent.MouseButtonPress:
                self.isPressed = True
            elif e.type() == QEvent.MouseButtonRelease:
                self.isPressed = False
            elif e.type() == QEvent.Enter:
                self.isHover = True
            elif e.type() == QEvent.Leave:
                self.isHover = False

        return super().eventFilter(obj, e)

    def _onComboTextChanged(self, text: str):
        self._currentIndex = -1
        self.currentTextChanged.emit(text)

        for i, item in enumerate(self.items):
            if item.text == text:
                self._currentIndex = i
                self.currentIndexChanged.emit(i)
                return

    def _onDropMenuClosed(self):
        self.dropMenu = None

    def _onClearButtonClicked(self):
        LineEdit.clear(self)
        self._currentIndex = -1


class Indicator(ToolButton):
    """ Indicator of switch button """

    checkedChanged = Signal(bool)

    def __init__(self, parent):
        super().__init__(parent=parent)
        self.setCheckable(True)
        self.setFixedSize(42, 22)

        self._sliderX = 5
        self.slideAni = QPropertyAnimation(self, b'sliderX', self)
        self.slideAni.setDuration(120)

        self.toggled.connect(self._toggleSlider)

    def mouseReleaseEvent(self, e):
        """ toggle checked state when mouse release"""
        super().mouseReleaseEvent(e)
        self.checkedChanged.emit(self.isChecked())

    def _toggleSlider(self):
        self.slideAni.setEndValue(25 if self.isChecked() else 5)
        self.slideAni.start()

    def toggle(self):
        self.setChecked(not self.isChecked())

    def setDown(self, isDown: bool):
        self.isPressed = isDown
        super().setDown(isDown)

    def setHover(self, isHover: bool):
        self.isHover = isHover
        self.update()

    def paintEvent(self, e):
        """ paint indicator """
        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)
        self._drawBackground(painter)
        self._drawCircle(painter)

    def _drawBackground(self, painter: QPainter):
        r = self.height() / 2
        painter.setPen(self._borderColor())
        painter.setBrush(self._backgroundColor())
        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), r, r)

    def _drawCircle(self, painter: QPainter):
        painter.setPen(Qt.NoPen)
        painter.setBrush(self._sliderColor())
        painter.drawEllipse(int(self.sliderX), 5, 12, 12)

    def _backgroundColor(self):
        isDark = isDarkTheme()

        if self.isChecked():
            if not self.isEnabled():
                return QColor(255, 255, 255, 41) if isDark else QColor(0, 0, 0, 56)
            if self.isPressed:
                return ThemeColor.LIGHT_2.color()
            elif self.isHover:
                return ThemeColor.LIGHT_1.color()

            return themeColor()
        else:
            if not self.isEnabled():
                return QColor(0, 0, 0, 0)
            if self.isPressed:
                return QColor(255, 255, 255, 18) if isDark else QColor(0, 0, 0, 23)
            elif self.isHover:
                return QColor(255, 255, 255, 10) if isDark else QColor(0, 0, 0, 15)

            return QColor(0, 0, 0, 0)

    def _borderColor(self):
        isDark = isDarkTheme()

        if self.isChecked():
            return self._backgroundColor() if self.isEnabled() else QColor(0, 0, 0, 0)
        else:
            if self.isEnabled():
                return QColor(255, 255, 255, 153) if isDark else QColor(0, 0, 0, 133)

            return QColor(255, 255, 255, 41) if isDark else QColor(0, 0, 0, 56)

    def _sliderColor(self):
        isDark = isDarkTheme()

        if self.isChecked():
            if self.isEnabled():
                return QColor(Qt.black if isDark else Qt.white)

            return QColor(255, 255, 255, 77) if isDark else QColor(255, 255, 255)
        else:
            if self.isEnabled():
                return QColor(255, 255, 255, 201) if isDark else QColor(0, 0, 0, 156)

            return QColor(255, 255, 255, 96) if isDark else QColor(0, 0, 0, 91)

    def getSliderX(self):
        return self._sliderX

    def setSliderX(self, x):
        self._sliderX = max(x, 5)
        self.update()

    sliderX = Property(float, getSliderX, setSliderX)


class SwitchButton(QWidget):
    """ Switch button class

    Constructors
    ------------
    * SwitchButton(`parent`: QWidget = None)
    * SwitchButton(`text`: str = "Off", `parent`: QWidget = None, `indicatorPos`=IndicatorPosition.LEFT)
    """

    checkedChanged = Signal(bool)

    @singledispatchmethod
    def __init__(self, parent: QWidget = None, indicatorPos=IndicatorPosition.LEFT):
        """
        Parameters
        ----------
        parent: QWidget
            parent widget

        indicatorPosition: IndicatorPosition
            the position of indicator
        """
        super().__init__(parent=parent)
        self._text = self.tr('关')
        self._offText =  self.tr('关')
        self._onText =  self.tr('开')
        self.__spacing = 12

        self.indicatorPos = indicatorPos
        self.hBox = QHBoxLayout(self)
        self.indicator = Indicator(self)
        self.label = QLabel(self._text, self)

        self.__initWidget()

    @__init__.register
    def _(self, text: str = 'Off', parent: QWidget = None, indicatorPos=IndicatorPosition.LEFT):
        """
        Parameters
        ----------
        text: str
            the text of switch button

        parent: QWidget
            parent widget

        indicatorPosition: IndicatorPosition
            the position of indicator
        """
        self.__init__(parent, indicatorPos)
        self._offText = text
        self.setText(text)

    def __initWidget(self):
        """ initialize widgets """
        self.setAttribute(Qt.WA_StyledBackground)
        self.installEventFilter(self)
        self.setFixedHeight(22)

        # set layout
        self.hBox.setSpacing(self.__spacing)
        self.hBox.setContentsMargins(2, 0, 0, 0)

        if self.indicatorPos == IndicatorPosition.LEFT:
            self.hBox.addWidget(self.indicator)
            self.hBox.addWidget(self.label)
            self.hBox.setAlignment(Qt.AlignLeft)
        else:
            self.hBox.addWidget(self.label, 0, Qt.AlignRight)
            self.hBox.addWidget(self.indicator, 0, Qt.AlignRight)
            self.hBox.setAlignment(Qt.AlignRight)

        # set default style sheet
        FluentStyleSheet.SWITCH_BUTTON.apply(self)

        # connect signal to slot
        self.indicator.toggled.connect(self._updateText)
        self.indicator.toggled.connect(self.checkedChanged)

    def eventFilter(self, obj, e: QEvent):
        if obj is self and self.isEnabled():
            if e.type() == QEvent.MouseButtonPress:
                self.indicator.setDown(True)
            elif e.type() == QEvent.MouseButtonRelease:
                self.indicator.setDown(False)
                self.indicator.toggle()
            elif e.type() == QEvent.Enter:
                self.indicator.setHover(True)
            elif e.type() == QEvent.Leave:
                self.indicator.setHover(False)

        return super().eventFilter(obj, e)

    def isChecked(self):
        return self.indicator.isChecked()

    def setChecked(self, isChecked):
        """ set checked state """
        self._updateText()
        self.indicator.setChecked(isChecked)

    def toggleChecked(self):
        """ toggle checked state """
        self.indicator.setChecked(not self.indicator.isChecked())

    def _updateText(self):
        self.setText(self.onText if self.isChecked() else self.offText)
        self.adjustSize()

    def getText(self):
        return self._text

    def setText(self, text):
        self._text = text
        self.label.setText(text)
        self.adjustSize()

    def getSpacing(self):
        return self.__spacing

    def setSpacing(self, spacing):
        self.__spacing = spacing
        self.hBox.setSpacing(spacing)
        self.update()

    def getOnText(self):
        return self._onText

    def setOnText(self, text):
        self._onText = text
        self._updateText()

    def getOffText(self):
        return self._offText

    def setOffText(self, text):
        self._offText = text
        self._updateText()

    spacing = Property(int, getSpacing, setSpacing)
    checked = Property(bool, isChecked, setChecked)
    text = Property(str, getText, setText)
    onText = Property(str, getOnText, setOnText)
    offText = Property(str, getOffText, setOffText)


class Window(MicaWindow):
    def __init__(self):
        super().__init__()
        setThemeColor(QColor(113, 89, 249))
        self.titleBar.raise_()
        self.setWindowTitle('ExpressLauncher')
        self.setWindowIcon(QIcon(':/icon.png'))
        self.resize(360, 155)
        self.move(QApplication.screens()[0].size().width() // 2 - self.width() // 2, QApplication.screens()[0].size().height() // 2 - self.height() // 2)

        self.statusBtn = PillPushButton(self)
        self.statusBtn.setCheckable(False)
        self.openSettingBtn = TransparentToolButton(self)
        self.openHelpBtn = TransparentToolButton(self)
        self.openSettingBtn.setIcon(FIF.SETTING)
        self.openHelpBtn.setIcon(FIF.HELP)
        self.openSettingBtn.clicked.connect(lambda: subprocess.Popen("ExpressSetting.exe", shell=True))
        self.openHelpBtn.clicked.connect(lambda: print("clicked"))
        self.upperLayout = QHBoxLayout()
        self.upperLayout.addWidget(self.statusBtn)
        self.upperLayout.addStretch(1)
        self.upperLayout.addWidget(self.openSettingBtn)
        self.upperLayout.addWidget(self.openHelpBtn)
        self.statusBtn.clicked.connect(self.onStatusBtn)
        self.updateStatus()

        items = ["D:", "E:", "F:", "G:", "H:", "I:", "J:"]
        self.comboBox = EditableComboBox(self)
        self.comboBox.setPlaceholderText("输入或选择盘符")
        self.comboBox.addItems(items)
        self.comboBox.setCurrentIndex(-1)
        self.completer = QCompleter(items, self)
        self.comboBox.setCompleter(self.completer)
        self.chooseBtn = HyperlinkButton(self)
        self.chooseBtn.setText("浏览")
        self.chooseBtn.setIcon(FIF.FOLDER)
        self.chooseBtn.clicked.connect(self.onChooseBtn)
        self.middleLayout = QHBoxLayout()
        self.middleLayout.addWidget(self.comboBox)
        self.middleLayout.addWidget(self.chooseBtn)

        self.bottomLayout = QHBoxLayout()
        self.exeBtn = PrimaryPushButton(self)
        self.exeBtn.setText("唤出U盘弹窗")
        self.exeBtn.clicked.connect(self.onYesBtn)
        self.cancelBtn = PushButton(self)
        self.cancelBtn.setText("取消")
        self.cancelBtn.clicked.connect(lambda: sys.exit())
        self.bottomLayout.addWidget(self.exeBtn)
        self.bottomLayout.addWidget(self.cancelBtn)

        self.upperLayout.setContentsMargins(15, 0, 15, 10)
        self.middleLayout.setContentsMargins(15, 0, 15, 10)
        self.bottomLayout.setContentsMargins(15, 0, 15, 15)
        self.mainLayout = QVBoxLayout(self)
        self.mainLayout.addWidget(self.titleBar)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.mainLayout.addLayout(self.upperLayout)
        self.mainLayout.addLayout(self.middleLayout)
        self.mainLayout.addStretch(1)
        self.mainLayout.addLayout(self.bottomLayout)

    def getStatus(self):
        try:
            window = GetWindow("ExpressSacn")[0]
            if window:
                return True
            else:
                return False
        except:
            return False

    def updateStatus(self):
        if self.getStatus():
            self.statusBtn.setText("U盘扫描已开启")
            self.statusBtn.setIcon(FIF.ACCEPT)
        else:
            self.statusBtn.setText("U盘扫描未开启")
            self.statusBtn.setIcon(FIF.CLOSE)

    def onStatusBtn(self):
        menu = RoundMenu(parent=self)
        switchBtn = SwitchButton(self)
        menu.addWidget(switchBtn, selectable=False)
        menu.exec(QPoint(self.x() + 33, self.y() + 80))

    def onChooseBtn(self):
        folder = QFileDialog.getExistingDirectory(self, self.tr("选择驱动器"), "./")
        if not folder:
            return
        else:
            self.comboBox.setText(os.path.splitdrive(folder)[0])

    def onYesBtn(self):
        if self.comboBox.text():
            subprocess.Popen(["ExpressUsbService.exe", self.comboBox.text().replace('：', ':')], shell=True)
            sys.exit()
        else:
            yesBtn = PushButton('确定')

            w = InfoBar(icon=InfoBarIcon.WARNING, title='未选择盘符', content='', orient=Qt.Horizontal, isClosable=True,
                        position=InfoBarPosition.BOTTOM, duration=-1, parent=self)
            yesBtn.clicked.connect(w.close)
            w.addWidget(yesBtn)
            w.show()


if __name__ == '__main__':
    if len(sys.argv) == 2 and sys.argv[1] == '-a' and cfg.AutoRun.value:
        subprocess.Popen("ExpressScan.exe", shell=True)
        sys.exit()
    with Mutex():
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
        if darkdetect.isDark():
            setTheme(Theme.DARK)
        app = QApplication(sys.argv)
        w = Window()
        w.show()
        app.exec()
