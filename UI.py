# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'UI.ui'
#
# Created: Sat Mar 04 18:11:02 2017
#      by: PyQt5 UI code generator 5.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import time
from multiprocessing import Process
import shutil
from Main import mainProcess
import os
import cPickle as pickle


class WorkerProcess(QtCore.QObject):
    '''
    自定义进程，进程由Python模块multiprocess创建
    并由pyqt的信号处理机制创建与主调进程的通信，通知子进程是否结束
    '''
    trigger = QtCore.pyqtSignal()

    def __init__(self):
        super(WorkerProcess, self).__init__()
        self.timer = QtCore.QTimer(self)  #创建计时器
        self.timer.timeout.connect(self.timeOut)  #连接计时器周期结束事件

    def beginRun(self, in_parameters):
        self.p = Process(target= mainProcess, kwargs=in_parameters) #设置子进程调用程序和参数
        self.p.start()         #启动子进程
        self.timer.start(1000) #启动计时器，计时器周期为1秒（1000毫秒）

    def timeOut(self):  #计时器一个周期结束即触发一个timeout事件，然后调用此函数，主要用于检测子进程是否结束
        if not self.p.is_alive():  # 每隔一秒钟检查进程是否结束，当前子进程完成后
            self.trigger.emit()  # 发出处理完成信息
            self.timer.stop()  # 并停止计时器，否则在同一进程下再次运行会导致过早判断子进程的状态

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.last_year = int(time.strftime('%Y', time.localtime(time.time()))) - 1
        self.in_parameters = {u'datetime': str(self.last_year) + u'年',
                              u'province':u'河南',
                              u'target_area': u'新乡市',
                              u'density_cell': u'10',
                              u'density_class': 10,
                              u'day_cell': u'15',
                              u'day_class': 10,
                              u'out_type': u'png'}
        self.setupUi()

    def setupUi(self):
        self.setObjectName("MainWindow")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("resource/weather-thunder.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(self)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.centralwidget.sizePolicy().hasHeightForWidth())
        self.centralwidget.setSizePolicy(sizePolicy)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.horizontalLayout_1 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_1.setObjectName("horizontalLayout_1")
        self.year_label = QtWidgets.QLabel(self.centralwidget)
        self.year_label.setObjectName("year_label")
        self.horizontalLayout_1.addWidget(self.year_label)
        self.year_DateEdit = QtWidgets.QDateEdit(self.centralwidget)
        self.year_DateEdit.setDateTime(QtCore.QDateTime(QtCore.QDate(self.last_year, 1, 1), QtCore.QTime(0, 0, 0)))
        self.year_DateEdit.setObjectName("year_DateEdit")
        self.horizontalLayout_1.addWidget(self.year_DateEdit)
        spacerItem = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_1.addItem(spacerItem)
        self.province_label = QtWidgets.QLabel(self.centralwidget)
        self.province_label.setObjectName("province_label")
        self.horizontalLayout_1.addWidget(self.province_label)
        self.province_comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.province_comboBox.setObjectName("province_comboBox")
        self.province_comboBox.addItem("")
        self.province_comboBox.addItem("")
        self.horizontalLayout_1.addWidget(self.province_comboBox)
        spacerItem1 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_1.addItem(spacerItem1)
        self.target_area_label = QtWidgets.QLabel(self.centralwidget)
        self.target_area_label.setObjectName("target_area_label")
        self.horizontalLayout_1.addWidget(self.target_area_label)
        self.target_area_comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.target_area_comboBox.setObjectName("target_area_comboBox")
        self.target_area_comboBox.addItem("")
        self.target_area_comboBox.addItem("")
        self.target_area_comboBox.addItem("")
        self.target_area_comboBox.addItem("")
        self.target_area_comboBox.addItem("")
        self.target_area_comboBox.addItem("")
        self.target_area_comboBox.addItem("")
        self.horizontalLayout_1.addWidget(self.target_area_comboBox)
        spacerItem2 = QtWidgets.QSpacerItem(80, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_1.addItem(spacerItem2)
        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_1.addItem(spacerItem3)
        self.gridLayout.addLayout(self.horizontalLayout_1, 0, 0, 1, 1)
        self.verticalLayout_1 = QtWidgets.QVBoxLayout()
        self.verticalLayout_1.setObjectName("verticalLayout_1")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setText("")
        self.label.setObjectName("label")
        self.verticalLayout_1.addWidget(self.label)
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.density_tab = QtWidgets.QWidget()
        self.density_tab.setObjectName("density_tab")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.density_tab)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.density_cell_label = QtWidgets.QLabel(self.density_tab)
        self.density_cell_label.setObjectName("density_cell_label")
        self.horizontalLayout_2.addWidget(self.density_cell_label)
        self.density_cell_spinBox = QtWidgets.QSpinBox(self.density_tab)
        self.density_cell_spinBox.setMaximum(30)
        self.density_cell_spinBox.setProperty("value", 10)
        self.density_cell_spinBox.setObjectName("density_cell_spinBox")
        self.horizontalLayout_2.addWidget(self.density_cell_spinBox)
        spacerItem4 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem4)
        self.density_class_label = QtWidgets.QLabel(self.density_tab)
        self.density_class_label.setObjectName("density_class_label")
        self.horizontalLayout_2.addWidget(self.density_class_label)
        self.density_class_spinBox = QtWidgets.QSpinBox(self.density_tab)
        self.density_class_spinBox.setMaximum(15)
        self.density_class_spinBox.setProperty("value", 10)
        self.density_class_spinBox.setObjectName("density_class_spinBox")
        self.horizontalLayout_2.addWidget(self.density_class_spinBox)
        spacerItem5 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem5)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.density_view = QtWidgets.QGraphicsView(self.density_tab)
        self.density_view.setObjectName("density_view")
        self.verticalLayout_2.addWidget(self.density_view)
        self.tabWidget.addTab(self.density_tab, "")
        self.day_tab = QtWidgets.QWidget()
        self.day_tab.setObjectName("day_tab")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.day_tab)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.day_cell_label = QtWidgets.QLabel(self.day_tab)
        self.day_cell_label.setObjectName("day_cell_label")
        self.horizontalLayout_3.addWidget(self.day_cell_label)
        self.day_cell_spinBox = QtWidgets.QSpinBox(self.day_tab)
        self.day_cell_spinBox.setMinimum(5)
        self.day_cell_spinBox.setMaximum(30)
        self.day_cell_spinBox.setProperty("value", 15)
        self.day_cell_spinBox.setObjectName("day_cell_spinBox")
        self.horizontalLayout_3.addWidget(self.day_cell_spinBox)
        spacerItem6 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem6)
        self.day_class_label = QtWidgets.QLabel(self.day_tab)
        self.day_class_label.setObjectName("day_class_label")
        self.horizontalLayout_3.addWidget(self.day_class_label)
        self.day_class_spinBox = QtWidgets.QSpinBox(self.day_tab)
        self.day_class_spinBox.setMinimum(5)
        self.day_class_spinBox.setMaximum(15)
        self.day_class_spinBox.setProperty("value", 10)
        self.day_class_spinBox.setObjectName("day_class_spinBox")
        self.horizontalLayout_3.addWidget(self.day_class_spinBox)
        spacerItem7 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem7)
        self.verticalLayout_3.addLayout(self.horizontalLayout_3)
        self.day_view = QtWidgets.QGraphicsView(self.day_tab)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.day_view.sizePolicy().hasHeightForWidth())
        self.day_view.setSizePolicy(sizePolicy)
        self.day_view.setObjectName("day_view")
        self.verticalLayout_3.addWidget(self.day_view)
        self.tabWidget.addTab(self.day_tab, "")
        self.regions_stats_tab = QtWidgets.QWidget()
        self.regions_stats_tab.setObjectName("regions_stats_tab")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.regions_stats_tab)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.province_stats_table = QtWidgets.QTableView(self.regions_stats_tab)
        self.province_stats_table.setObjectName("province_stats_table")
        self.verticalLayout_4.addWidget(self.province_stats_table)
        self.region_stats_table = QtWidgets.QTableView(self.regions_stats_tab)
        self.region_stats_table.setObjectName("region_stats_table")
        self.verticalLayout_4.addWidget(self.region_stats_table)
        self.verticalLayout_5.addLayout(self.verticalLayout_4)
        self.tabWidget.addTab(self.regions_stats_tab, "")
        self.month_stats_tab = QtWidgets.QWidget()
        self.month_stats_tab.setObjectName("month_stats_tab")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.month_stats_tab)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.month_stats_table = QtWidgets.QTableView(self.month_stats_tab)
        self.month_stats_table.setObjectName("month_stats_table")
        self.verticalLayout_6.addWidget(self.month_stats_table)
        self.month_stats_view = QtWidgets.QGraphicsView(self.month_stats_tab)
        self.month_stats_view.setObjectName("month_stats_view")
        self.verticalLayout_6.addWidget(self.month_stats_view)
        self.tabWidget.addTab(self.month_stats_tab, "")
        self.hour_stats_tab = QtWidgets.QWidget()
        self.hour_stats_tab.setObjectName("hour_stats_tab")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout(self.hour_stats_tab)
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.hours_stats_table = QtWidgets.QTableView(self.hour_stats_tab)
        self.hours_stats_table.setObjectName("hours_stats_table")
        self.verticalLayout_9.addWidget(self.hours_stats_table)
        self.hour_stats_view = QtWidgets.QGraphicsView(self.hour_stats_tab)
        self.hour_stats_view.setObjectName("hour_stats_view")
        self.verticalLayout_9.addWidget(self.hour_stats_view)
        self.tabWidget.addTab(self.hour_stats_tab, "")
        self.intensity_stats_tab = QtWidgets.QWidget()
        self.intensity_stats_tab.setObjectName("intensity_stats_tab")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.intensity_stats_tab)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.intensity_stats_table = QtWidgets.QTableView(self.intensity_stats_tab)
        self.intensity_stats_table.setObjectName("intensity_stats_table")
        self.horizontalLayout_4.addWidget(self.intensity_stats_table)
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.negative_view = QtWidgets.QGraphicsView(self.intensity_stats_tab)
        self.negative_view.setObjectName("negative_view")
        self.verticalLayout_7.addWidget(self.negative_view)
        self.positive_view = QtWidgets.QGraphicsView(self.intensity_stats_tab)
        self.positive_view.setObjectName("positive_view")
        self.verticalLayout_7.addWidget(self.positive_view)
        self.horizontalLayout_4.addLayout(self.verticalLayout_7)
        self.gridLayout_2.addLayout(self.horizontalLayout_4, 0, 0, 1, 1)
        self.tabWidget.addTab(self.intensity_stats_tab, "")
        self.verticalLayout_1.addWidget(self.tabWidget)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_1.addWidget(self.progressBar)
        self.gridLayout.addLayout(self.verticalLayout_1, 1, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 955, 23))
        self.menubar.setObjectName("menubar")
        self.file_menu = QtWidgets.QMenu(self.menubar)
        self.file_menu.setObjectName("file_menu")
        self.help_menu = QtWidgets.QMenu(self.menubar)
        self.help_menu.setObjectName("help_menu")
        self.action_menu = QtWidgets.QMenu(self.menubar)
        self.action_menu.setObjectName("action_menu")
        self.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        self.toolBar = QtWidgets.QToolBar(self)
        self.toolBar.setStatusTip("")
        self.toolBar.setObjectName("toolBar")
        self.addToolBar(QtCore.Qt.TopToolBarArea, self.toolBar)
        self.help_action = QtWidgets.QAction(self)
        self.help_action.setObjectName("help_action")
        self.about_action = QtWidgets.QAction(self)
        self.about_action.setObjectName("about_action")
        self.export_charts_action = QtWidgets.QAction(self)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("resource/export_charts.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.export_charts_action.setIcon(icon1)
        self.export_charts_action.setObjectName("export_charts_action")
        self.export_doc_action = QtWidgets.QAction(self)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("resource/export_doc.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.export_doc_action.setIcon(icon2)
        self.export_doc_action.setObjectName("export_doc_action")
        self.load_data_action = QtWidgets.QAction(self)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("resource/load_data.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.load_data_action.setIcon(icon3)
        self.load_data_action.setObjectName("load_data_action")
        self.execute_action = QtWidgets.QAction(self)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("resource/excute.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.execute_action.setIcon(icon4)
        self.execute_action.setObjectName("execute_action")
        self.open_doc_action = QtWidgets.QAction(self)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap("resource/open_doc.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.open_doc_action.setIcon(icon5)
        self.open_doc_action.setObjectName("open_doc_action")
        self.open_charts_action = QtWidgets.QAction(self)
        icon6 = QtGui.QIcon()
        icon6.addPixmap(QtGui.QPixmap("resource/open_charts.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.open_charts_action.setIcon(icon6)
        self.open_charts_action.setObjectName("open_charts_action")
        self.exit_action = QtWidgets.QAction(self)
        self.exit_action.setObjectName("exit_action")
        self.file_menu.addAction(self.open_doc_action)
        self.file_menu.addAction(self.open_charts_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.export_doc_action)
        self.file_menu.addAction(self.export_charts_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.exit_action)
        self.help_menu.addAction(self.help_action)
        self.help_menu.addAction(self.about_action)
        self.action_menu.addAction(self.load_data_action)
        self.action_menu.addAction(self.execute_action)
        self.menubar.addAction(self.file_menu.menuAction())
        self.menubar.addAction(self.action_menu.menuAction())
        self.menubar.addAction(self.help_menu.menuAction())
        self.toolBar.addAction(self.load_data_action)
        self.toolBar.addAction(self.execute_action)
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.open_doc_action)
        self.toolBar.addAction(self.open_charts_action)
        self.toolBar.addSeparator()
        self.toolBar.addAction(self.export_doc_action)
        self.toolBar.addAction(self.export_charts_action)
        self.year_label.setBuddy(self.year_DateEdit)
        self.province_label.setBuddy(self.province_comboBox)
        self.target_area_label.setBuddy(self.target_area_comboBox)
        self.tabWidget.setCurrentIndex(0)

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)

        self.initSizePosition()

        self.show()

        self.handleEvents()

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "雷电公报"))
        self.year_label.setText(_translate("MainWindow", "年份"))
        self.year_DateEdit.setDisplayFormat(_translate("MainWindow", "yyyy"))
        self.province_label.setText(_translate("MainWindow", "省份"))
        self.province_comboBox.setItemText(0, _translate("MainWindow", "河南"))
        self.province_comboBox.setItemText(1, _translate("MainWindow", "浙江"))
        self.target_area_label.setText(_translate("MainWindow", "目标区域"))
        self.target_area_comboBox.setItemText(0, _translate("MainWindow", "新乡市"))
        self.target_area_comboBox.setItemText(1, _translate("MainWindow", "延津县"))
        self.target_area_comboBox.setItemText(2, _translate("MainWindow", "新乡县"))
        self.target_area_comboBox.setItemText(3, _translate("MainWindow", "辉县市"))
        self.target_area_comboBox.setItemText(4, _translate("MainWindow", "卫辉市"))
        self.target_area_comboBox.setItemText(5, _translate("MainWindow", "获嘉县"))
        self.target_area_comboBox.setItemText(6, _translate("MainWindow", "封丘县"))
        self.density_cell_label.setText(_translate("MainWindow", "插值网格大小"))
        self.density_class_label.setText(_translate("MainWindow", "制图分类数目"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.density_tab), _translate("MainWindow", "闪电密度分布图"))
        self.day_cell_label.setText(_translate("MainWindow", "插值半径大小"))
        self.day_class_label.setText(_translate("MainWindow", "制图分类数目"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.day_tab), _translate("MainWindow", "雷暴日分布图"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.regions_stats_tab), _translate("MainWindow", "分区统计"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.month_stats_tab), _translate("MainWindow", "分月统计"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.hour_stats_tab), _translate("MainWindow", "分时段统计"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.intensity_stats_tab), _translate("MainWindow", "强度统计"))
        self.file_menu.setTitle(_translate("MainWindow", "文件"))
        self.help_menu.setTitle(_translate("MainWindow", "帮助"))
        self.action_menu.setTitle(_translate("MainWindow", "操作"))
        self.toolBar.setWindowTitle(_translate("MainWindow", "toolBar"))
        self.help_action.setText(_translate("MainWindow", "使用说明"))
        self.help_action.setStatusTip(_translate("MainWindow", "使用说明"))
        self.help_action.setShortcut(_translate("MainWindow", "F1"))
        self.about_action.setText(_translate("MainWindow", "关于"))
        self.about_action.setStatusTip(_translate("MainWindow", "关于"))
        self.about_action.setShortcut(_translate("MainWindow", "F2"))
        self.export_charts_action.setText(_translate("MainWindow", "导出图表"))
        self.export_charts_action.setToolTip(_translate("MainWindow", "导出所有GIS图和统计图表"))
        self.export_charts_action.setStatusTip(_translate("MainWindow", "导出所有GIS图和统计图表"))
        self.export_charts_action.setShortcut(_translate("MainWindow", "Alt+C"))
        self.export_doc_action.setText(_translate("MainWindow", "导出文档"))
        self.export_doc_action.setToolTip(_translate("MainWindow", "导出公报word文档"))
        self.export_doc_action.setStatusTip(_translate("MainWindow", "导出公报word文档"))
        self.export_doc_action.setShortcut(_translate("MainWindow", "Alt+D"))
        self.load_data_action.setText(_translate("MainWindow", "加载数据"))
        self.load_data_action.setToolTip(_translate("MainWindow", "制作雷电公报图表前，必须加载电闪数据"))
        self.load_data_action.setStatusTip(_translate("MainWindow", "制作雷电公报图表前，必须加载电闪数据"))
        self.load_data_action.setShortcut(_translate("MainWindow", "Ctrl+N"))
        self.execute_action.setText(_translate("MainWindow", "执行"))
        self.execute_action.setToolTip(_translate("MainWindow", "点击开始执行"))
        self.execute_action.setStatusTip(_translate("MainWindow", "点击开始执行"))
        self.execute_action.setShortcut(_translate("MainWindow", "F5"))
        self.open_doc_action.setText(_translate("MainWindow", "打开文档"))
        self.open_doc_action.setToolTip(_translate("MainWindow", "打开公报word文档"))
        self.open_doc_action.setStatusTip(_translate("MainWindow", "打开公报word文档"))
        self.open_doc_action.setShortcut(_translate("MainWindow", "Ctrl+D"))
        self.open_charts_action.setText(_translate("MainWindow", "打开图表"))
        self.open_charts_action.setToolTip(_translate("MainWindow", "打开公报统计图表"))
        self.open_charts_action.setStatusTip(_translate("MainWindow", "打开公报统计图表"))
        self.open_charts_action.setShortcut(_translate("MainWindow", "Ctrl+C"))
        self.exit_action.setText(_translate("MainWindow", "退出"))
        self.exit_action.setShortcut(_translate("MainWindow", "F4"))

    def initSizePosition(self):
        # 当前屏幕的大小和位置，QtCore.QRect(左上角x坐标,左上角y坐标,x方向宽度，y方向高度)
        rect_monitor = QtWidgets.QDesktopWidget().availableGeometry()
        #根据屏幕大小改变窗口大小
        height_monitor = rect_monitor.height()
        height_main_win = int(height_monitor*0.9)
        width_main_win = int(height_main_win*1.2)
        self.resize(width_main_win,height_main_win)
        # 当前主窗口的大小和位置，也是QtCore.QRect类型
        qr = self.frameGeometry()
        #屏幕中心点坐标
        center_point = rect_monitor.center()#center函数返回当前矩形的中心点位置坐标QtCore.QPoint类型
        qr.moveCenter(center_point)
        self.move(qr.topLeft())

    def handleEvents(self):
        self.about_action.triggered.connect(self.showAbout)
        self.load_data_action.triggered.connect(self.loadData)
        self.execute_action.triggered.connect(self.execute)
        self.exit_action.triggered.connect(self.close)
        self.export_doc_action.triggered.connect(self.exportDoc)
        self.export_charts_action.triggered.connect(self.exportCharts)
        self.province_comboBox .activated[str].connect(self.updateProvince)
        self.target_area_comboBox.activated[str].connect(self.updateTargetArea)
        self.year_DateEdit .dateChanged.connect(self.updateDatetime)
        self.density_cell_spinBox.valueChanged.connect(self.updateDensityCell)
        self.density_class_spinBox .valueChanged.connect(self.updateDensityClass)
        self.day_cell_spinBox .valueChanged.connect(self.updateDayCell)
        self.day_class_spinBox.valueChanged.connect(self.updateDayClass)

    def showAbout(self):
        self.about = AboutDialog()

    def exportDoc(self):
        self.setStatsTables()
        self.setStatsPics()

    def exportCharts(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self,u'请选择图片保存位置',
                                                     u'E:/Documents/工作/雷电公报',
                                    QtWidgets.QFileDialog.ShowDirsOnly|QtWidgets.QFileDialog.DontResolveSymlinks)
        if not directory:
            return

        cwd = os.getcwd()
        densityPic = ''.join([cwd,u'/temp/',self.in_parameters[u'province'],'/', self.in_parameters[u'datetime'],u'/',
            self.in_parameters[u'target_area'],'.gdb',u'/',self.in_parameters[u'datetime'],
            self.in_parameters[u'target_area'],u'闪电密度空间分布.', self.in_parameters[u'out_type']])

        dayPic = ''.join([cwd,u'/temp/',self.in_parameters[u'province'],'/',self.in_parameters[u'datetime'],u'/',
            self.in_parameters[u'target_area'],'.gdb',u'/',self.in_parameters[u'datetime'],
            self.in_parameters[u'target_area'],u'地闪雷暴日空间分布.', self.in_parameters[u'out_type']])

        dest_density = os.path.join(directory,os.path.basename(densityPic))
        dest_day = os.path.join(directory,os.path.basename(dayPic))

        if os.path.isfile(dest_day) or os.path.isfile(dest_density):
            message = u"文件已经存在！"
            msgBox = QtWidgets.QMessageBox()
            msgBox.setText(message)
            msgBox.setIcon(QtWidgets.QMessageBox.Information)
            icon = QtGui.QIcon()
            icon.addPixmap(QtGui.QPixmap("./resource/weather-thunder.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msgBox.setWindowIcon(icon)
            msgBox.setWindowTitle(" ")
            msgBox.exec_()
            return

        shutil.copy2(dayPic,directory)
        shutil.copy2(densityPic,directory)

    def loadData(self):
        fnames = QtWidgets.QFileDialog.getOpenFileNames(self, u'请选择原始的电闪数据',
                                              u'Z:/ZHUF/数据/闪电数据',
                                              'Text files (*.txt);;All(*.*)')

        self.in_parameters[u'origin_data_path'] = fnames[0]

    def updateProvince(self,area):
        self.in_parameters[u'province'] = area
        self.target_area_comboBox.clear()
        if area == u'浙江':
            self.target_area_comboBox.addItems(['绍兴市','柯桥区','上虞区','诸暨市','嵊州市','新昌县'])
            self.in_parameters[u'target_area'] = u'绍兴市'
        elif area == u'河南':
            self.target_area_comboBox.addItems(['新乡市', '延津县', '新乡县', '辉县市', '卫辉市', '获嘉县','封丘县'])
            self.in_parameters[u'target_area'] = u'新乡市'

        if self.in_parameters.has_key(u'origin_data_path'):
            self.in_parameters.__delitem__(u'origin_data_path')

    def updateTargetArea(self,area):
        self.in_parameters[u'target_area'] = area

    def updateDatetime(self, date):
        self.in_parameters[u'datetime'] = str(date.year()) + u'年'
        if self.in_parameters.has_key(u'origin_data_path'):
            self.in_parameters.__delitem__(u'origin_data_path')

    def updateDensityCell(self, cell):
        self.in_parameters[u'density_cell'] = str(cell)

    def updateDensityClass(self, nclass):
        self.in_parameters[u'density_class'] = nclass

    def updateDayCell(self, cell):
        self.in_parameters[u'day_cell'] = str(cell)

    def updateDayClass(self, nclass):
        self.in_parameters[u'day_class'] = nclass

    def execute(self):
        reply = QtWidgets.QMessageBox.question(self, u'确定', u'确定是制作%s的%s公报吗？'%(
                                                self.in_parameters[u'datetime'],self.in_parameters[u'target_area']),
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                               QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.No:
            return

        dir = u"E:/Documents/工作/雷电公报/闪电定位原始文本数据/" + self.in_parameters[u'datetime']

        if os.path.exists(dir):
            datafiles = os.listdir(dir)
            datafiles = map(lambda x:os.path.join(dir,x),datafiles)
            self.in_parameters[u'origin_data_path'] = datafiles

        if not self.in_parameters.has_key(u'origin_data_path'):
            message = u"请加载%s的数据" % self.in_parameters[u'datetime']
            msgBox = QtWidgets.QMessageBox()
            msgBox.setText(message)
            msgBox.setIcon(QtWidgets.QMessageBox.Information)
            icon = QtGui.QIcon()
            icon.addPixmap(QtGui.QPixmap('./resource/weather-thunder.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msgBox.setWindowIcon(icon)
            msgBox.setWindowTitle(" ")
            msgBox.exec_()
            return

        self.execute_action.setDisabled(True)
        self.statusbar.showMessage(u'正在执行中……')
        self.progressBar.setMaximum(0)
        self.progressBar.setMinimum(0)
        self.progressBar.setFormat(u'执行中……')

        self.load_data_action.setDisabled(True)
        self.execute_action.setDisabled(True)
        self.export_charts_action.setDisabled(True)
        self.export_doc_action.setDisabled(True)
        self.open_charts_action.setDisabled(True)
        self.open_doc_action.setDisabled(True)
        self.province_comboBox.setDisabled(True)
        self.target_area_comboBox.setDisabled(True)
        self.year_DateEdit.setDisabled(True)
        self.density_cell_spinBox.setDisabled(True)
        self.density_class_spinBox.setDisabled(True)
        self.day_cell_spinBox.setDisabled(True)
        self.day_class_spinBox.setDisabled(True)

        # for outfile in self.in_parameters[u'origin_data_path']:
        #     infile =
        #     try:
        #         with open(infile, 'w+') as in_f:
        #             for line in in_f:
        #                 line = line.replace(u"：",":")
        #                 in_f.write(line)
        #     except Exception,inst:
        #         print infile

        self.worker_process = WorkerProcess()
        self.worker_process.trigger.connect(self.finished)
        self.worker_process.beginRun(self.in_parameters)

    def finished(self):
        #设置图片显示
        self.setStatsPics()

        #设置分区统计的表格
        self.setStatsTables()

        #处理进度条和执行按钮状态
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.progressBar.setValue(100)
        self.progressBar.setFormat(u'完成!')

        #改变一些控件的状态
        self.load_data_action.setDisabled(False)
        self.execute_action.setDisabled(False)
        self.export_charts_action.setDisabled(False)
        self.export_doc_action.setDisabled(False)
        self.open_charts_action.setDisabled(False)
        self.open_doc_action.setDisabled(False)
        self.province_comboBox.setDisabled(False)
        self.target_area_comboBox.setDisabled(False)
        self.year_DateEdit.setDisabled(False)
        self.density_cell_spinBox.setDisabled(False)
        self.density_class_spinBox.setDisabled(False)
        self.day_cell_spinBox.setDisabled(False)
        self.day_class_spinBox.setDisabled(False)
        self.statusbar.showMessage(u'完成！')
        #self.action_save_pic.setDisabled(False)

    def closeEvent(self, event):

        reply = QtWidgets.QMessageBox.question(self, u'退出',u'真的要退出程序吗？',
                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                               QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def setStatsTables(self):
        cwd = os.getcwd()
        workspath = ''.join([cwd, u"/temp/", self.in_parameters['province'],
                             '/',self.in_parameters[u'datetime'], '/', self.in_parameters[u'target_area'] , '.gdb'])
        # 读取查询参数
        f = open(os.path.join(workspath, 'query_results.pkl'), 'rb')
        query_results = pickle.load(f)
        f.close()

        #******省分区统计******
        stats_of_province = query_results[u'Stats_of_Province']
        if self.in_parameters[u'province'] == u'浙江':
            regions = [u'杭州',u'宁波',u'湖州', u'嘉兴', u'绍兴', u'金华', u'台州', u'温州', u'衢州', u'丽水', u'舟山',u'总计']
        elif self.in_parameters[u'province'] == u'河南':
            regions = [u'郑州',u'开封',u'洛阳',u'平顶山',u'安阳',u'鹤壁',u'新乡',u'焦作',u'濮阳',
                       u'许昌',u'漯河',u'三门峡',u'商丘',u'周口',u'驻马店',u'南阳',u'信阳',u'济源', u'总计']

        n_regions = len(regions)
        self.model_province_stats = QtGui.QStandardItemModel(2, n_regions, self.province_stats_table)
        self.model_province_stats.setHorizontalHeaderLabels(regions)
        self.model_province_stats.setVerticalHeaderLabels([u'地闪次数', u'平均密度\n(次/km²)'])

        for i in range(n_regions):
            for j in range(2):
                self.model_province_stats.setItem(j,i, QtGui.QStandardItem(str(stats_of_province[regions[i]][j])))
                self.model_province_stats.item(j,i).setTextAlignment(QtCore.Qt.AlignCenter)

        self.province_stats_table.setModel(self.model_province_stats)
        #表格占满窗口，并可以活动
        self.province_stats_table.resizeRowsToContents()
        self.province_stats_table.horizontalHeader().setStretchLastSection(True)
        self.province_stats_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        #self.province_stats_table.resizeColumnsToContents()
        #设置表格内容不可编辑
        self.province_stats_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        #*********分县统计************
        stats_of_region = query_results[u'Stats_of_Region']
        if self.in_parameters[u'province'] == u'浙江':
            counties = [u'越城区', u'柯桥区', u'上虞区', u'诸暨市', u'嵊州市', u'新昌县', u'总计']
        elif self.in_parameters[u'province'] == u'河南':
            counties = [u'新乡市', u'新乡县', u'辉县市', u'卫辉市', u'获嘉县', u'原阳县',u'延津县', u'封丘县', u'长垣县']

        n_counties = len(counties)
        self.model_region_stats = QtGui.QStandardItemModel(2, n_counties, self.region_stats_table)
        self.model_region_stats.setHorizontalHeaderLabels(counties)
        self.model_region_stats.setVerticalHeaderLabels([u'地闪次数', u'平均密度\n(次/km²)',u'最大密度\n(次/km²)',
                                            u'最小密度\n(次/km²)',u'平均雷暴日(天)',u'最多雷暴日(天)',u'最少雷暴日(天)'])

        for i in range(n_counties):
            for j in range(7):
                self.model_region_stats.setItem(j,i, QtGui.QStandardItem(str(stats_of_region[counties[i]][j])))
                self.model_region_stats.item(j,i).setTextAlignment(QtCore.Qt.AlignCenter)

        self.region_stats_table.setModel(self.model_region_stats)
        #表格占满窗口，并可以活动
        self.region_stats_table.resizeRowsToContents()
        self.region_stats_table.horizontalHeader().setStretchLastSection(True)
        self.region_stats_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        #self.province_stats_table.resizeColumnsToContents()
        #设置表格内容不可编辑
        self.region_stats_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        #********分月统计***********
        stats_of_month = query_results[u'Stats_of_Month']#0负闪次数、1负闪强度、2正闪次数、3正闪强度、4总次数
        self.model_month_stats = QtGui.QStandardItemModel(12, 5, self.month_stats_table)
        self.model_month_stats.setHorizontalHeaderLabels([u'负闪次数',u'正闪次数',u'地闪总数',u'负闪平均强度',u'正闪平均强度'])
        self.model_month_stats.setVerticalHeaderLabels([u'1月', u'2月',u'3月',u'4月',u'5月',u'6月',u'7月',u'8月',u'9月',u'10月',u'11月',u'12月'])
        for i in range(12):
            self.model_month_stats.setItem(i,0,QtGui.QStandardItem(str(stats_of_month[i+1][0])))#0负闪次数
            self.model_month_stats.setItem(i,1,QtGui.QStandardItem(str(stats_of_month[i+1][2])))#2正闪次数
            self.model_month_stats.setItem(i,2,QtGui.QStandardItem(str(stats_of_month[i+1][4])))#4总次数
            self.model_month_stats.setItem(i,3,QtGui.QStandardItem('%.2f'%stats_of_month[i+1][1]))#1负闪强度
            self.model_month_stats.setItem(i,4,QtGui.QStandardItem('%.2f'%stats_of_month[i+1][3]))#3正闪强度
        for i in range(12):
            for j in range(5):
                self.model_month_stats.item(i,j).setTextAlignment(QtCore.Qt.AlignCenter)
        self.month_stats_table.setModel(self.model_month_stats)

        #表格占满窗口，并可以活动
        self.month_stats_table.resizeRowsToContents()
        self.month_stats_table.horizontalHeader().setStretchLastSection(True)
        self.month_stats_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        #self.province_stats_table.resizeColumnsToContents()
        #设置表格内容不可编辑
        self.month_stats_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        #*****分时段统计******
        stats_of_hour = query_results[u'Stats_of_Hour']#0负闪次数、1负闪强度、2正闪次数、3正闪强度、4总次数
        self.model_hour_stats = QtGui.QStandardItemModel(12, 5, self.hours_stats_table)
        self.model_hour_stats.setHorizontalHeaderLabels([u'负闪次数',u'正闪次数',u'地闪总数',u'负闪平均强度',u'正闪平均强度'])
        self.model_hour_stats.setVerticalHeaderLabels([u'0-1', u'1-2',u'2-3',u'3-4',u'4-5',u'5-6',u'6-7',u'7-8',u'8-9',u'9-10',u'10-11',u'11-12',
                                                        u'12-13', u'13-14',u'14-15',u'15-16',u'16-17',u'17-18',u'18-19',u'19-20',u'20-21',u'21-22',
                                                        u'22-23',u'23-24'])
        for i in range(24):
            self.model_hour_stats.setItem(i,0,QtGui.QStandardItem(str(stats_of_hour[i][0])))#0负闪次数
            self.model_hour_stats.setItem(i,1,QtGui.QStandardItem(str(stats_of_hour[i][2])))#2正闪次数
            self.model_hour_stats.setItem(i,2,QtGui.QStandardItem(str(stats_of_hour[i][4])))#4总次数
            self.model_hour_stats.setItem(i,3,QtGui.QStandardItem('%.2f'%stats_of_hour[i][1]))#1负闪强度
            self.model_hour_stats.setItem(i,4,QtGui.QStandardItem('%.2f'%stats_of_hour[i][3]))#3正闪强度
        for i in range(24):
            for j in range(5):
                self.model_hour_stats.item(i,j).setTextAlignment(QtCore.Qt.AlignCenter)
        self.hours_stats_table.setModel(self.model_hour_stats)

        #表格占满窗口，并可以活动
        self.hours_stats_table.resizeRowsToContents()
        self.hours_stats_table.horizontalHeader().setStretchLastSection(True)
        self.hours_stats_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        #self.province_stats_table.resizeColumnsToContents()
        #设置表格内容不可编辑
        self.hours_stats_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        #*****分强度统计*******
        stats_of_intensity = query_results[u'Stats_of_Intensity']
        self.model_intensity_stats = QtGui.QStandardItemModel(12, 5, self.hours_stats_table)
        self.model_intensity_stats.setHorizontalHeaderLabels([u'负闪次数',u'正闪次数',u'地闪总数',u'负闪累计比例',u'正闪累计比例'])
        self.model_intensity_stats.setVerticalHeaderLabels([u'0-5', u'5-10', u'10-15', u'15-20',u'20-25',u'25-30',u'30-35',u'35-40',u'40-45',u'45-50',u'50-55',
                                                            u'55-60', u'60-65', u'65-70', u'70-75', u'75-80', u'80-85', u'85-90',u'90-95',u'95-100',
                                                            u'100-150',u'150-200',u'200-250',u'250-300',u'300-1000'])

        for i in range(25):
            self.model_intensity_stats.setItem(i,0,QtGui.QStandardItem(str(stats_of_intensity[i][0])))#0负闪次数
            self.model_intensity_stats.setItem(i,1,QtGui.QStandardItem(str(stats_of_intensity[i][1])))#1正闪次数
            self.model_intensity_stats.setItem(i,2,QtGui.QStandardItem(str(stats_of_intensity[i][0]+stats_of_intensity[i][1])))#4总次数
            self.model_intensity_stats.setItem(i,3,QtGui.QStandardItem('%.2f%%'%(100*stats_of_intensity[i][2])))#负闪累计比例
            self.model_intensity_stats.setItem(i,4,QtGui.QStandardItem('%.2f%%'%(100*stats_of_intensity[i][3])))#正闪累计比例

        for i in range(25):
            for j in range(5):
                self.model_intensity_stats.item(i,j).setTextAlignment(QtCore.Qt.AlignCenter)
                self.intensity_stats_table.setModel(self.model_intensity_stats)

        #表格占满窗口，并可以活动
        self.intensity_stats_table.resizeRowsToContents()
        self.intensity_stats_table.horizontalHeader().setStretchLastSection(True)
        self.intensity_stats_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        #self.province_stats_table.resizeColumnsToContents()
        #设置表格内容不可编辑
        self.intensity_stats_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

    def setStatsPics(self):
        cwd = os.getcwd()
        # ***********绘制闪电密度图*************
        #清除上一次的QGraphicsView对象，防止其记录上一次图片结果，影响显示效果
        self.density_view.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.verticalLayout_2.removeWidget(self.density_view)
        size = self.density_view.size()
        self.density_view.close()

        self.density_view = QtWidgets.QGraphicsView(self.density_tab)
        self.density_view.setObjectName("density_view")
        self.density_view.resize(size)
        self.verticalLayout_2.addWidget(self.density_view)

        densityPic = ''.join([cwd, u'/temp/', self.in_parameters[u'province'], u'/',
                              self.in_parameters[u'datetime'], u'/', self.in_parameters[u'target_area'], u'.gdb/',
                              self.in_parameters[u'datetime'], self.in_parameters[u'target_area'], u'闪电密度空间分布.',self.in_parameters[u'out_type']])

        scene = QtWidgets.QGraphicsScene()
        pixmap_density = QtGui.QPixmap(densityPic)
        scene.addPixmap(pixmap_density)
        self.density_view.setScene(scene)
        scale = float(self.density_view.width()) / pixmap_density.width()
        self.density_view.scale(scale, scale)

        # ***********绘制雷暴日图*************
        self.day_view.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.verticalLayout_3.removeWidget(self.day_view)
        size = self.day_view.size()
        self.day_view.close()

        self.day_view = QtWidgets.QGraphicsView(self.day_tab)
        self.day_view.setObjectName("day_view")
        self.day_view.resize(size)
        self.verticalLayout_3.addWidget(self.day_view)

        dayPic = ''.join([cwd, u'/temp/', self.in_parameters[u'province'], u'/',
                          self.in_parameters[u'datetime'], u'/', self.in_parameters[u'target_area'], u'.gdb/',
                          self.in_parameters[u'datetime'], self.in_parameters[u'target_area'], u'地闪雷暴日空间分布.',self.in_parameters[u'out_type']])

        pixmap_day = QtGui.QPixmap(dayPic)
        scene = QtWidgets.QGraphicsScene()
        scene.addPixmap(pixmap_day)
        self.day_view.resize(self.density_view.width(), self.density_view.height())
        self.day_view.setScene(scene)
        scale = float(self.day_view.width()) / pixmap_day.width()
        self.day_view.scale(scale, scale)

        #*****绘制分月统计图******
        self.month_stats_view.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.verticalLayout_6.removeWidget(self.month_stats_view)
        size = self.month_stats_view.size()
        self.month_stats_view.close()

        self.month_stats_view = QtWidgets.QGraphicsView(self.month_stats_tab)
        self.month_stats_view.setObjectName("month_stats_view")
        self.month_stats_view.resize(size)
        self.verticalLayout_6.addWidget(self.month_stats_view)

        month_stats_pic = ''.join([cwd, u'/temp/', self.in_parameters[u'province'], u'/',
                                   self.in_parameters[u'datetime'], u'/', self.in_parameters[u'target_area'], u'.gdb/',
                                   u'month_stats_pic.png'])

        pixmap_month_stats = QtGui.QPixmap(month_stats_pic)
        scene = QtWidgets.QGraphicsScene()
        scene.addPixmap(pixmap_month_stats)
        self.month_stats_view.resize(self.month_stats_view.width(), self.month_stats_view.height())
        self.month_stats_view.setScene(scene)
        # scale = float(self.month_stats_view.width()) / pixmap_month_stats.width()
        # self.month_stats_view.scale(scale, scale)

        #******绘制分时段统计图*******
        self.hour_stats_view.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.verticalLayout_9.removeWidget(self.hour_stats_view)
        size = self.hour_stats_view.size()
        self.hour_stats_view.close()

        self.hour_stats_view = QtWidgets.QGraphicsView(self.hour_stats_tab)
        self.hour_stats_view.setObjectName("hour_stats_view")
        self.hour_stats_view.resize(size)
        self.verticalLayout_9.addWidget(self.hour_stats_view)

        hour_stats_pic = ''.join([cwd, u'/temp/', self.in_parameters[u'province'], u'/',
                                  self.in_parameters[u'datetime'], u'/', self.in_parameters[u'target_area'], u'.gdb/',
                                  u'hour_stats_pic.png'])

        pixmap_hour_stats = QtGui.QPixmap(hour_stats_pic)
        scene = QtWidgets.QGraphicsScene()
        scene.addPixmap(pixmap_hour_stats)
        self.hour_stats_view.resize(self.hour_stats_view.width(), self.hour_stats_view.height())
        self.hour_stats_view.setScene(scene)

        #****绘制负闪强度分布图****
        self.negative_view.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.verticalLayout_7.removeWidget(self.negative_view)
        size = self.negative_view.size()
        self.negative_view.close()

        self.negative_view = QtWidgets.QGraphicsView(self.intensity_stats_tab)
        self.negative_view.setObjectName("negative_view")
        self.negative_view.resize(size)
        self.verticalLayout_7.addWidget(self.negative_view)

        negative_pic = ''.join([cwd, u'/temp/', self.in_parameters[u'province'], u'/',
                                self.in_parameters[u'datetime'], u'/', self.in_parameters[u'target_area'], u'.gdb/',
                                u'negative_stats_pic.png'])

        pixmap_negative = QtGui.QPixmap(negative_pic)
        scene = QtWidgets.QGraphicsScene()
        scene.addPixmap(pixmap_negative)
        self.negative_view.resize(self.negative_view.width(), self.negative_view.height())
        self.negative_view.setScene(scene)

        #*****绘制正闪强度分布图*****
        self.positive_view.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.verticalLayout_7.removeWidget(self.positive_view)
        size = self.positive_view.size()
        self.positive_view.close()

        self.positive_view = QtWidgets.QGraphicsView(self.intensity_stats_tab)
        self.positive_view.setObjectName("positive_view")
        self.positive_view.resize(size)
        self.verticalLayout_7.addWidget(self.positive_view)

        positive_pic = ''.join([cwd, u'/temp/', self.in_parameters[u'province'], u'/',
                                self.in_parameters[u'datetime'], u'/', self.in_parameters[u'target_area'], u'.gdb/',
                                u'positive_stats_pic.png'])

        pixmap_positive = QtGui.QPixmap(positive_pic)
        scene = QtWidgets.QGraphicsScene()
        scene.addPixmap(pixmap_positive)
        self.positive_view.resize(self.positive_view.width(), self.positive_view.height())
        self.positive_view.setScene(scene)


class AboutDialog(QtWidgets.QDialog):
    def __init__(self):
        super(AboutDialog, self).__init__()
        self.setupUi()

    def setupUi(self):
        self.setObjectName("Dialog")
        self.resize(491, 261)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        self.setMinimumSize(QtCore.QSize(491, 261))
        self.setMaximumSize(QtCore.QSize(491, 261))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("resource/weather-thunder.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(0, 0, 491, 141))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("resource/about.jpg"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.textBrowser = QtWidgets.QTextBrowser(self)
        self.textBrowser.setGeometry(QtCore.QRect(0, 140, 491, 141))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.textBrowser.sizePolicy().hasHeightForWidth())
        self.textBrowser.setSizePolicy(sizePolicy)
        self.textBrowser.setObjectName("textBrowser")

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.show()

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("Dialog", "关于"))
        self.textBrowser.setHtml(_translate("Dialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt; font-weight:600;\"> 雷电公报-v1.0</span></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-weight:600;\"><br /></p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:11pt;\"> 联系：976309391@qq.com; zhufeng314@163.com </span></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:11pt;\"><br /></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:11pt;\"><br /></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:11pt;\">©</span> 新乡市气象局 延津县气象局  All rights reserved.</p></body></html>"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = MainWindow()
    sys.exit(app.exec_())

