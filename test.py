# -*- coding: utf-8 -*-

from PyQt5 import QtWidgets, QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtWidgets.QApplication.UnicodeUTF8


    def _translate(context, text, disambig):
        return QtWidgets.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtWidgets.QApplication.translate(context, text, disambig)


class Ui_Form(QtWidgets.QDialog):
    def setupUi(self, Form):
        Form.setObjectName(_fromUtf8("Form"))
        Form.resize(800, 600)
        self.tableView = QtWidgets.QTableView(Form)
        self.tableView.setGeometry(QtCore.QRect(5, 30, 790, 560))
        self.tableView.setObjectName(_fromUtf8("tableView"))
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(10, 10, 300, 12))
        self.label.setObjectName(_fromUtf8("label"))

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        Form.setWindowTitle(_translate("Form", "定值设置", None))
        self.label.setText(_translate("Form", "请在下面的表格中设置定值参数：", None))

        # 初始化表格
        self.tableView_set()

    def tableView_set(self):

        # 添加表头：
        self.model = QtGui.QStandardItemModel(self.tableView)

        # 设置表格属性：
        self.model.setRowCount(17)
        self.model.setColumnCount(8)

        # 设置表头
        self.model.setHeaderData(0, QtCore.Qt.Horizontal, _fromUtf8(u"类型"))
        self.model.setHeaderData(1, QtCore.Qt.Horizontal, _fromUtf8(u"值"))
        self.model.setHeaderData(2, QtCore.Qt.Horizontal, _fromUtf8(u""))
        self.model.setHeaderData(3, QtCore.Qt.Horizontal, _fromUtf8(u"类型"))
        self.model.setHeaderData(4, QtCore.Qt.Horizontal, _fromUtf8(u"值"))
        self.model.setHeaderData(5, QtCore.Qt.Horizontal, _fromUtf8(u""))
        self.model.setHeaderData(6, QtCore.Qt.Horizontal, _fromUtf8(u"类型"))
        self.model.setHeaderData(7, QtCore.Qt.Horizontal, _fromUtf8(u"值"))

        self.tableView.setModel(self.model)

        # 设置列宽
        self.tableView.setColumnWidth(0, 100)
        self.tableView.setColumnWidth(1, 80)
        self.tableView.setColumnWidth(2, 80)
        self.tableView.setColumnWidth(3, 100)
        self.tableView.setColumnWidth(4, 80)
        self.tableView.setColumnWidth(5, 80)
        self.tableView.setColumnWidth(6, 100)
        self.tableView.setColumnWidth(7, 80)

        # 合并单元格的效果
        # 第一个参数：要改变的单元格行数
        # 第二个参数：要改变的单元格列数
        # 第三个参数：需要合并的行数
        # 第四个参数：需要合并的列数
        self.tableView.setSpan(0, 2, 17, 1)
        self.tableView.setSpan(0, 5, 17, 1)

        # 设置单元格禁止更改
        # self.tableView.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)

        # 表头信息显示居左
        # self.tableView.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignLeft)

        # 表头信息显示居中
        self.tableView.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignCenter)

        '''''
        #添加表项
        for i in range(0,3):
            self.model.setItem(i,0,QtGui.QStandardItem("2009441676"))
            #设置字符颜色
            self.model.item(i,0).setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
            #设置字符位置
            self.model.item(i,0).setTextAlignment(QtCore.Qt.AlignCenter)

            self.model.setItem(i,1,QtGui.QStandardItem(_fromUtf8("哈哈")))
        '''
        # 添加表项
        set_value_type_list = [u'过流I段定值', u'过流II段定值', u'过流III段定值', u'零序过流定值', u'后加速定值', u'PT有压定值', u'PT过压定值', u'PT无压定值',
                               u'过流I段延时', u'过流II段延时', u'过流III段延时', u'零序过流延时', u'后加速延时', u'重合闸I段延时', u'重合闸II段延时',
                               u'重合闸III段延时', u'重合复归延时', u'反时限曲线', u'反时限启动', u'反时限倍数', u'环境温度', u'控制字', u'手机号码']
        index = 0
        for i in range(0, 30):
            self.model.setItem(i, 0, QtGui.QStandardItem(_fromUtf8(set_value_type_list[index])))
            index += 1
            # 设置字符颜色
            self.model.item(i, 0).setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
            # 设置字符位置
            self.model.item(i, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            if index >= len(set_value_type_list):
                break

            self.model.setItem(i, 3, QtGui.QStandardItem(_fromUtf8(set_value_type_list[index])))
            index += 1
            # 设置字符颜色
            self.model.item(i, 3).setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
            # 设置字符位置
            self.model.item(i, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            if index >= len(set_value_type_list):
                break

            self.model.setItem(i, 6, QtGui.QStandardItem(_fromUtf8(set_value_type_list[index])))
            index += 1
            # 设置字符颜色
            self.model.item(i, 6).setForeground(QtGui.QBrush(QtGui.QColor(255, 0, 0)))
            # 设置字符位置
            self.model.item(i, 6).setTextAlignment(QtCore.Qt.AlignCenter)
            if index >= len(set_value_type_list):
                break

        self.tableView.setModel(self.model)


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())