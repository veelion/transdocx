#!/usr/bin/env python3
# coding:utf-8
# Author: veelion



from PyQt5.QtCore import QThread, QSize
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QIcon

import googletrans

import docx_translator as docxtr
from pdf2text import resource_path


class LogHandler(QtCore.QObject):
    # sig = QtCore.pyqtSignal(str)
    show = QtCore.pyqtSignal(str)


class TranslateTask(QThread):
    done = QtCore.pyqtSignal(str)

    def set_attr(self, src, dst, fn):
        self.src = src
        self.dst = dst
        self.fn = fn

    def run(self):
        to_save = docxtr.translate(self.fn, self.src, self.dst)
        msg = '翻译成功，保存为：<b>{}</b>'.format(to_save)
        self.done.emit(msg)


class Window(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(Window, self).__init__(parent)

        languages = googletrans.LANGUAGES.copy()
        languages['zh-cn'] = '中文(简体)'
        languages['zh-tw'] = '中文(繁体)'
        self.langcodes = dict(map(reversed, languages.items()))
        self.langlist = [k.capitalize() for k in self.langcodes.keys()]

        self.browseButton = self.createButton("&浏览...", self.browse)
        self.transButton = self.createButton("&翻译", self.translate)

        self.lang_srcComboBox = self.createComboBox()
        srcIndex = self.langlist.index('English')
        self.lang_srcComboBox.setCurrentIndex(srcIndex)
        self.lang_dstComboBox = self.createComboBox()
        dstIndex = self.langlist.index('中文(简体)')
        self.lang_dstComboBox.setCurrentIndex(dstIndex)
        self.fileComboBox = self.createComboBox('file')

        srcLabel = QtWidgets.QLabel("文档语言:")
        dstLabel = QtWidgets.QLabel("目标语言:")
        docLabel = QtWidgets.QLabel("选择文档:")
        self.filesFoundLabel = QtWidgets.QLabel()

        self.logPlainText = QtWidgets.QPlainTextEdit()
        self.logPlainText.setReadOnly(True)


        buttonsLayout = QtWidgets.QHBoxLayout()
        buttonsLayout.addStretch()
        buttonsLayout.addWidget(self.transButton)

        mainLayout = QtWidgets.QGridLayout()
        mainLayout.addWidget(srcLabel, 0, 0)
        mainLayout.addWidget(self.lang_srcComboBox, 0, 1, 1, 2)
        mainLayout.addWidget(dstLabel, 1, 0)
        mainLayout.addWidget(self.lang_dstComboBox, 1, 1, 1, 2)
        mainLayout.addWidget(docLabel, 2, 0)
        mainLayout.addWidget(self.fileComboBox, 2, 1)
        mainLayout.addWidget(self.browseButton, 2, 2)
        # mainLayout.addWidget(self.filesTable, 3, 0, 1, 3)
        mainLayout.addWidget(self.logPlainText, 3, 0, 1, 3)
        mainLayout.addWidget(self.filesFoundLabel, 4, 0)
        mainLayout.addLayout(buttonsLayout, 5, 0, 1, 3)
        self.setLayout(mainLayout)

        app_icon = QIcon()#icon = https://imgur.com/NV7Ugfd
        icon_path = resource_path('icon.png')
        app_icon.addFile(icon_path, QSize(16, 16))
        app_icon.addFile(icon_path, QSize(24, 24))
        app_icon.addFile(icon_path, QSize(32, 32))
        app_icon.addFile(icon_path, QSize(48, 48))
        app_icon.addFile(icon_path, QSize(256, 256))
        self.setWindowIcon(app_icon)

        self.setWindowTitle("文档翻译")
        self.resize(800, 600)

        # translate
        self.logger = LogHandler()
        self.logger.show.connect(self.onLog)
        docxtr.g_log = self.logger
        self.task = TranslateTask()
        self.task.done.connect(self.onLog)

    def init_lang(self):
        self.lang_srcComboBox.currentIndexChanged('english')
        self.lang_dstComboBox.currentIndexChanged('中文(简体)')

    def browse(self):
        sfile, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择文件",
            QtCore.QDir.currentPath(),
            "Docx, PDF(*.docx *.pdf *.PDF)",
        )
        print(sfile)

        if sfile:
            self.logPlainText.clear()
            msg = '选择了文件: <b>{}</b>'.format(sfile)
            self.logger.show.emit(msg)
            if self.fileComboBox.findText(sfile) == -1:
                self.fileComboBox.addItem(sfile)

            self.fileComboBox.setCurrentIndex(self.fileComboBox.findText(sfile))

    def onLog(self, msg):
        self.logPlainText.appendHtml(msg)

    def translate(self):
        lang_src_select = self.lang_srcComboBox.currentText()
        lang_dst_select = self.lang_dstComboBox.currentText()
        fileName = self.fileComboBox.currentText()
        if not fileName:
            self.logger.show.emit('请先选择要翻译的Word文档')
            return
        lang_src = self.langcodes.get(lang_src_select.lower())
        lang_dst = self.langcodes.get(lang_dst_select.lower())
        if not lang_src:
            self.logger.show.emit('无效的源语言:{}'.format(lang_src_select))
            return
        if not lang_dst:
            self.logger.show.emit('无效的目标语言:{}'.format(lang_dst_select))
            return
        print(lang_src, lang_dst, fileName)
        self.logger.show.emit('开始翻译文档：{}'.format(fileName))
        msg = '从 <b>{}</b> 到 <b>{}</b>'.format(lang_src_select, lang_dst_select)
        self.logger.show.emit(msg)
        # self.logPlainText.textChanged()
        self.task.set_attr(lang_src, lang_dst, fileName)
        self.task.start()

    def createButton(self, text, member):
        button = QtWidgets.QPushButton(text)
        button.clicked.connect(member)
        return button

    def createComboBox(self, btype=''):
        comboBox = QtWidgets.QComboBox(self)
        if btype != 'file':
            comboBox.setEditable(True)
            comboBox.addItems(self.langlist)
        comboBox.setSizePolicy(QtWidgets.QSizePolicy.Expanding,
                QtWidgets.QSizePolicy.Preferred)
        return comboBox


if __name__ == '__main__':

    import sys

    app = QtWidgets.QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
