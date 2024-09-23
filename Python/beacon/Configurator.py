import os
import sys
import xlrd3
import json
from datetime import datetime
from time import localtime, mktime

import serial
import serial.tools.list_ports
from PyQt5.QtCore import QTimer, QRegExp, Qt, QThread, pyqtSignal
from PyQt5.QtGui import QRegExpValidator, QStandardItemModel, QStandardItem, QBrush, QColor, QFont, QPixmap
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem, QMainWindow, qApp, QApplication, QFileDialog, QTableWidget

import CS
import re
import time
from Ui_window import Ui_MainWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.setWindowTitle("Configurator - Beacon(Plus) | Spark | Beam V2.32")
        self.ser = serial.Serial()
        self.init()
        self.portCheck()
        self.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)  # 仅显示最小化和关闭按钮
        self.pbNarrow.setVisible(False)  # 启动时隐藏pbNarrwow按钮

        # 接收数据和发送数据数目置零
        self.dataNumReceive = 0
        self.leNumReceive.setText(str(self.dataNumReceive))
        self.dataNumSend = 0
        self.leNumSend.setText(str(self.dataNumSend))
        self.leNumReceive.textChanged.connect(self.checkC4)

        # 限制输入范围
        self.cbBeaAddr.setValidator(
            QRegExpValidator(QRegExp("^\\d{10}1234|A4A4A4A4A41234|AAAAAAAAAA1234")))  # 限制Beacon地址为14位数字
        self.cbSpkAddr.setValidator(
            QRegExpValidator(QRegExp("^\\d{10}1235|A4A4A4A4A41235|AAAAAAAAAA1235")))  # 限制Spark地址为14位数字
        self.leBeaAddr.setValidator(
            QRegExpValidator(QRegExp("^\\d{10}1234|A4A4A4A4A41234|AAAAAAAAAA1234")))  # 限制Beacon地址为14位数字
        self.leSpkAddr.setValidator(
            QRegExpValidator(QRegExp("^\\d{10}1235|A4A4A4A4A41235|AAAAAAAAAA1235")))  # 限制Spark地址为14位数字
        self.cbBeamAddr.setValidator(QRegExpValidator(QRegExp("^\\d{10}1236|A4A4A4A4A41236|AAAAAAAAAA1236")))  # 限制Beam地址为14位数字
        self.leBeamAddr.setValidator(QRegExpValidator(QRegExp("^\\d{10}1236")))  # 限制Beam地址为14位数字
        self.leBatchAddr.setValidator(QRegExpValidator(QRegExp("^\\d{10}1236")))  # 限制Beam地址为14位数字
        '''
        self.leSetPos.setValidator(QRegExpValidator(QRegExp("[0-9]\.[0-9][0-9]|[1-9][0-9]\.[0-9][0-9]|100"), self))   # 限制阀位设定值0.00~100.00
        self.leSetTemp.setValidator(QRegExpValidator(QRegExp("[0-9]\.[0-9][0-9]|[1-9][0-9]\.[0-9][0-9]|100"), self))   # 限制温度设定值0.00~100.00
        self.leSetPosMin.setValidator(QRegExpValidator(QRegExp("[0-9]|[1-2][0-9]|30"), self))   # 限制阀位下限0~30
        self.leSetPosMax.setValidator(QRegExpValidator(QRegExp("[7-9][0-9]|100"), self))  # 限制阀位上限70~100
        self.leSetDead.setValidator(QRegExpValidator(QRegExp("0\.[1-9]|[1-9]|10"), self))  # 限制死区0.1~10
        self.leSetP.setValidator(QRegExpValidator(QRegExp("[1-9]|1[0-9]|20"), self))  # 限制P值1~20
        self.leSetI.setValidator(QRegExpValidator(QRegExp("[1-9][0-9]|[1-9][0-9][0-9]|1[0-5][0-9][0-9]|1600"), self))  # 限制I值10~1600
        '''
        # 设置tableData
        self.tableData.setSpan(0, 0, 1, 3)  # 使单元格(0,0)合并成1行3列
        item = QTableWidgetItem("Beacon 数据 ⬇")
        item.setFont(QFont('Microsoft YaHei', 10, QFont.Bold))
        item.setBackground(QBrush(QColor(185, 185, 185)))
        self.tableData.setItem(0, 0, item)

        self.tableData.setSpan(21, 0, 1, 3)
        item = QTableWidgetItem("Spark 数据 ⬇")
        item.setFont(QFont('Microsoft YaHei', 10, QFont.Bold))
        item.setBackground(QBrush(QColor(185, 185, 185)))
        self.tableData.setItem(21, 0, item)

        self.tableData.setColumnWidth(2, 225)
        self.tableData.setItem(12, 2, QTableWidgetItem("01:防拆 02:校准后堵塞 03:磨损堵塞"))
        self.tableData.setItem(13, 2, QTableWidgetItem("00:正常 01:超过室温限制 02:低温报警"))
        self.tableData.setItem(15, 2, QTableWidgetItem("00:正常 01:未配对 02:连接超时"))
        self.tableData.setItem(16, 2, QTableWidgetItem("00:正常 非0:异常"))
        self.tableData.setItem(23, 2, QTableWidgetItem("00:正常 01:超过室温限制 02:低温报警"))
        self.tableData.setItem(27, 2, QTableWidgetItem("00:处于非供暖季 01:处于供暖季"))
        self.tableData.setItem(29, 2, QTableWidgetItem("00:正常 01:低电量"))
        self.tableData.setItem(30, 2, QTableWidgetItem("00:关机 01:开机"))
        self.tableData.setItem(31, 2, QTableWidgetItem("00:正常 01:未配对 02:连接超时"))
        self.tableData.setItem(33, 2, QTableWidgetItem("00:停用 01:启用"))
        self.tableData.setItem(34, 2, QTableWidgetItem("00:停用 01:启用"))
        self.tableData.setItem(35, 2, QTableWidgetItem("00:禁用 01:允许"))

        # 格式化批量导入列表
        self.tableBatchNum = 0
        self.cwd = os.getcwd()
        self.pbImpExcel.clicked.connect(self.importExcel)
        self.pbBatchPair1.clicked.connect(self.startBatchPair)
        self.pbBatchPair0.clicked.connect(self.stopBatchPair)
        self.pbBatchPair0.setEnabled(False)
        self.timerBatch = QTimer()
        self.timerBatch.timeout.connect(self.alsBatchPair)
        self.batchProcess = 0
        self.tableBatch.horizontalHeader().setDisabled(1)  # 禁止调整列宽
        self.tableBatch.verticalHeader().setDisabled(1)  # 禁止调整行高
        self.tableBatch.setColumnWidth(0, 80)
        self.tableBatch.setColumnWidth(1, 110)
        self.tableBatch.setColumnWidth(2, 110)
        self.addr1 = ''
        self.addr2 = ''
        self.addr3 = ''

        # Beam列表
        self.tableBeam.horizontalHeader().setDisabled(1)  # 禁止调整列宽
        self.tableBeam.verticalHeader().setDisabled(1)  # 禁止调整行高
        self.tableBeam.setColumnWidth(0, 120)
        self.pbReadBeam.clicked.connect(self.readBeam)
        self.timerBeam = QTimer()
        self.timerBeam.timeout.connect(self.alsBeam)
        self.pbBeamRTSet.clicked.connect(self.beamRTSet)
        self.pbBeamRTLimit.clicked.connect(self.beamRTLmit)
        self.pbBeamRTOffset.clicked.connect(self.beamRTOffset)
        self.pbBeamAlarm.clicked.connect(self.beamAlarm)
        self.pbBeamHint.clicked.connect(self.beamHint)
        self.pbBeamExact.clicked.connect(self.beamExact)
        self.pbBeamLockRT.clicked.connect(self.beamLockRT)
        self.pbBeamAway.clicked.connect(self.beamAway)
        self.tableBeam.setContextMenuPolicy(Qt.CustomContextMenu)  # 允许打开上下文菜单
        self.tableBeam.customContextMenuRequested.connect(self.rightClick)  # 绑定右键点击事件

        # 限制地址comboBox列表数量
        self.cbBeaAddr.setMaxCount(10)
        self.cbSpkAddr.setMaxCount(10)

        # 禁用套用格式
        self.tbReceive.setAcceptRichText(False)
        self.teSend.setAcceptRichText(False)

    def init(self):
        # 串口检测按钮
        self.pbCheckPort.clicked.connect(self.portCheck)
        # 串口信息显示
        self.cbPortList.currentTextChanged.connect(self.portInfo)
        # 打开串口按钮
        self.pbOpenPort.clicked.connect(self.portOpen)
        # 关闭串口按钮
        self.pbClosePort.clicked.connect(self.portClose)
        # 发送数据按钮
        self.pbSend.clicked.connect(self.dataSend)
        # 定时器接收数据
        self.timerReceive = QTimer(self)
        self.timerReceive.timeout.connect(self.dataReceive)
        # 定时器发送数据
        self.timerSend = QTimer(self)
        self.timerSend.timeout.connect(self.dataSend)
        self.cbTimerSend.stateChanged.connect(self.dataTimerSend)
        # 接收窗口可编辑
        self.tbReceive.setStyleSheet("background:rgb(240,240,240)")
        self.cbEdit.stateChanged.connect(self.receiveEditable)
        # 清除发送窗口
        self.pbClearReceive.clicked.connect(self.dataReceiveClear)
        # 清除接收窗口
        self.pbClearSend.clicked.connect(self.dataSendClear)
        # 计算CS
        self.pbCS.clicked.connect(self.calCS)

        # Beacon广播地址
        self.cbBeaBC.stateChanged.connect(self.beaBC)
        # Plus勾选框
        self.cbBeaType.stateChanged.connect(self.selectPlus)
        self.notPlus()
        # 读取Beacon汇总数据
        self.pbBeaRead.clicked.connect(self.beaRead)
        # 解析Beacon汇总数据
        self.pbBeaAnalysis.clicked.connect(self.beaAnalysis)
        # 延时解析Beacon汇总数据
        self.timerBeaAll = QTimer()
        self.timerBeaAll.timeout.connect(self.beaAnalysis)
        # 设定阀门类型
        self.pbSetType.clicked.connect(self.setType)
        # 设定流量特性
        self.pbSetCharacter.clicked.connect(self.setCharacter)
        # 设定控制方式
        self.pbSetMode.clicked.connect(self.setMode)
        # 远程校准
        self.pbCalibration.clicked.connect(self.setCalibration)
        # 冲洗
        self.pbFlush.clicked.connect(self.setFlush)
        # 排气
        self.pbDeAir.clicked.connect(self.setDeAir)
        # 恢复出厂设置
        self.pbFactory.clicked.connect(self.setFactory)
        # 设定阀位上下限和死区
        self.pbSetPosLimit.clicked.connect(self.setPosLimit)
        self.pbSetPosMin.clicked.connect(self.setPosMin)
        self.pbSetPosMax.clicked.connect(self.setPosMax)
        self.pbSetDead.clicked.connect(self.setDead)
        # 设定PI
        self.pbSetPI.clicked.connect(self.setPI)
        self.pbSetP.clicked.connect(self.setP)
        self.pbSetI.clicked.connect(self.setI)
        # 阀位设定
        self.pbSetPos.clicked.connect(self.setPos)
        # 温度设定
        self.pbSetReturnT.clicked.connect(self.setReturnT)
        # 设定Spark模式位置限制
        self.pbSetSpkLimit.clicked.connect(self.setSpkLimit)
        # Spark广播地址
        self.cbSpkBC.stateChanged.connect(self.spkBC)
        # 读取Spark汇总数据
        self.pbSpkRead.clicked.connect(self.spkRead)
        # 解析Spark汇总数据
        self.pbSpkAnalysis.clicked.connect(self.spkAnalysis)
        # 延时解析Spark汇总数据
        self.timerSpkAll = QTimer()
        self.timerSpkAll.timeout.connect(self.spkAnalysis)
        # 室温设定
        self.pbSetRoomT.clicked.connect(self.setRoomT)
        # 室温锁定
        self.pbLockRoomT.clicked.connect(self.lockRoomT)
        # 室温设定限制
        self.pbSetRTLimit.clicked.connect(self.setRTLimit)
        # 读取室温补偿
        self.pbReadRToffset.clicked.connect(self.readRToffset)
        # 延时计息室温补偿
        self.timerRToffset = QTimer()
        self.timerRToffset.timeout.connect(self.alsRToffset)
        # 设定室温补偿
        self.pbSetRToffset.clicked.connect(self.setRTOffset)
        # 读取报警阈值
        self.pbReadAlmT.clicked.connect(self.readAlmT)
        # 延时解析报警阈值
        self.timerAlmT = QTimer()
        self.timerAlmT.timeout.connect(self.alsAlmT)
        # 设定报警阈值
        self.pbSetAlmT.clicked.connect(self.setAlmT)
        # 读取供暖季状态
        self.pbReadSeason.clicked.connect(self.readSeason)
        # 延时解析供暖季状态
        self.timerSeason = QTimer()
        self.timerSeason.timeout.connect(self.alsSeason)
        # 供暖季设定
        self.pbSetSeason.clicked.connect(self.setSeason)
        # 强制供暖季
        self.pbSetSeason1.clicked.connect(self.setSeason1)
        # 强制非供暖季
        self.pbSetSeason0.clicked.connect(self.setSeason0)
        # 精确温控
        self.pbExact.clicked.connect(self.exact)
        # 提示符号
        self.pbHint.clicked.connect(self.hint)
        # 离开模式
        self.pbAway.clicked.connect(self.away)
        # 读取Spark时钟
        self.pbReadSpkClk.clicked.connect(self.readSpkClk)
        # 延时解析Spark时钟
        self.timerSpkClk = QTimer()
        self.timerSpkClk.timeout.connect(self.alsSpkClk)
        # 同步PC时间至Spark
        self.pbSyncSpkClk.clicked.connect(self.syncSpkClk)
        # 读取历史数据
        self.pbReadHis.clicked.connect(self.readHis)
        # 解析历史数据
        self.pbAlsHis.clicked.connect(self.analysisHis)
        # 延时解析历史数据
        self.timerSpkHis = QTimer()
        self.timerSpkHis.timeout.connect(self.analysisHis)
        # 读取日志总数
        self.pbReadHisNum.clicked.connect(self.readHisNum)
        # 解析日志总数
        self.pbAlsHisNum.clicked.connect(self.alsHisNum)
        # 延时解析日志总数
        self.timerHisNum = QTimer()
        self.timerHisNum.timeout.connect(self.alsHisNum)
        # 读取日期索引
        self.pbReadHisDate.clicked.connect(self.readHisDate)
        # 解析日期索引
        self.pbAlsHisDate.clicked.connect(self.alsHisDate)
        # 延时解析日期
        self.timerHisDate = QTimer()
        self.timerHisDate.timeout.connect(self.alsHisDate)
        # 清空日期索引列表
        self.pbClrHisDate.clicked.connect(lambda: self.lvHisDate.clear())
        # 根据索引定位日期
        self.lvHisDate.clicked.connect(self.setHisDate)
        # 同步Beacon Plus地址用于配对
        self.cbBeaSync.stateChanged.connect(self.beaSync)
        self.cbBeaAddr.currentTextChanged.connect(self.beaSync)
        # 同步Spark地址用于配对
        self.cbSpkSync.stateChanged.connect(self.spkSync)
        self.cbSpkAddr.currentTextChanged.connect(self.spkSync)
        # 同步Beam地址用于配对
        self.cbBeamSync.stateChanged.connect(self.beamSync)
        self.cbBeamAddr.currentTextChanged.connect(self.beamSync)
        # 写入Spark配对信息
        self.pbSetSpkInfo.clicked.connect(self.setSpkInfo)
        # 清除Spark配对信息
        self.pbClrSpkInfo.clicked.connect(self.clrSpkInfo)
        # 读取Spark配对信息
        self.pbReadSpkInfo.clicked.connect(self.readSpkInfo)
        # 延时解析Spark配对信息
        self.timerSpkInfo = QTimer()
        self.timerSpkInfo.timeout.connect(self.spkAlsInfo)
        # 复制Spark ID到剪贴板
        self.lvSpkID.doubleClicked.connect(self.copySpkID)
        # 宽列表
        self.pbWiden.clicked.connect(self.tableWiden)
        # 窄列表
        self.pbNarrow.clicked.connect(self.tableNarrow)
        # 按钮禁用
        if not self.ser.isOpen():
            self.pbDisable()
            self.pbOpenPort.setEnabled(True)
            self.pbClosePort.setEnabled(False)
        else:
            self.pbEnable()
            self.pbOpenPort.setEnabled(False)
            self.pbClosePort.setEnabled(True)
        # Beam和Beacon Plus选择（批量配对）
        self.rbSelBeam.toggled.connect(self.selBeamOrBeacon)
        self.rbSelBeacon.toggled.connect(self.selBeamOrBeacon)
        # Beam和Beacon Plus选择（Spark配对）
        self.rbSpkSelBeam.toggled.connect(self.spkSelBeamOrBeacon)
        self.rbSpkSelBeacon.toggled.connect(self.spkSelBeamOrBeacon)
        # 导入和导出地址json
        self.importCfg.triggered.connect(self.importJson)
        self.exportCfg.triggered.connect(self.exportJson)
        # 删除地址记录列表
        self.pbDelBeaAddr.clicked.connect(lambda: self.cbBeaAddr.removeItem(self.cbBeaAddr.currentIndex()))
        self.pbDelSpkAddr.clicked.connect(lambda: self.cbSpkAddr.removeItem(self.cbSpkAddr.currentIndex()))
        self.pbDelBeamAddr.clicked.connect(lambda: self.cbBeamAddr.removeItem(self.cbBeamAddr.currentIndex()))

    # 导入json
    def importJson(self):
        name = 'config.json'
        if os.path.exists(name) and os.path.isfile(name):   # 检测文件是否存在
            with open(name, 'r') as f:
                content = f.read()
                c1, c2, c3 = [], [], []
                a = json.loads(content)
                if 'cbBeaAddr' in a:
                    b = a['cbBeaAddr']
                    for i in b:
                        c1.append(i)
                    self.cbBeaAddr.addItems(c1)
                else:
                    pass
                if 'cbSpkAddr' in a:
                    b = a['cbSpkAddr']
                    for i in b:
                        c2.append(i)
                    self.cbSpkAddr.addItems(c2)
                else:
                    pass
                if 'cbBeamAddr' in a:
                    b = a['cbBeamAddr']
                    for i in b:
                        c3.append(i)
                    self.cbBeamAddr.addItems(c3)
                else:
                    pass
        else:
            pass

    # 导出json
    def exportJson(self):
        print('abc')
        c1, c2, c3 = [], [], []
        for i in range(self.cbBeaAddr.count()):
            print(self.cbBeaAddr.itemText(i))
            c1.append(self.cbBeaAddr.itemText(i))
        for i in range(self.cbSpkAddr.count()):
            print(self.cbSpkAddr.itemText(i))
            c2.append(self.cbSpkAddr.itemText(i))
        for i in range(self.cbBeamAddr.count()):
            print(self.cbBeamAddr.itemText(i))
            c3.append(self.cbBeamAddr.itemText(i))
        a = {"cbBeaAddr": c1,
             "cbSpkAddr": c2,
             "cbBeamAddr": c3
             }
        b = json.dumps(a, indent=2)
        f = open('../config.json', 'w')
        f.write(b)
        f.close()

    # 按钮禁用
    def pbDisable(self):
        self.pbSend.setEnabled(False)
        self.pbBeaRead.setEnabled(False)
        self.pbCalibration.setEnabled(False)
        self.pbFlush.setEnabled(False)
        self.pbDeAir.setEnabled(False)
        self.pbFactory.setEnabled(False)
        self.pbSetType.setEnabled(False)
        self.pbSetCharacter.setEnabled(False)
        self.pbSetMode.setEnabled(False)
        self.pbSetPos.setEnabled(False)
        self.pbSetReturnT.setEnabled(False)
        self.pbSetPosMax.setEnabled(False)
        self.pbSetPosMin.setEnabled(False)
        self.pbSetDead.setEnabled(False)
        self.pbSetPosLimit.setEnabled(False)
        self.pbSetP.setEnabled(False)
        self.pbSetI.setEnabled(False)
        self.pbSetPI.setEnabled(False)
        self.pbSetSpkLimit.setEnabled(False)
        self.pbSpkRead.setEnabled(False)
        self.pbSetRoomT.setEnabled(False)
        self.pbSetRTLimit.setEnabled(False)
        self.pbReadRToffset.setEnabled(False)
        self.pbSetRToffset.setEnabled(False)
        self.pbReadAlmT.setEnabled(False)
        self.pbSetAlmT.setEnabled(False)
        self.pbReadSeason.setEnabled(False)
        self.pbSetSeason.setEnabled(False)
        self.pbSetSeason1.setEnabled(False)
        self.pbSetSeason0.setEnabled(False)
        self.pbLockRoomT.setEnabled(False)
        self.pbExact.setEnabled(False)
        self.pbHint.setEnabled(False)
        self.pbAway.setEnabled(False)
        self.pbReadSpkClk.setEnabled(False)
        self.pbSyncSpkClk.setEnabled(False)
        self.pbReadHis.setEnabled(False)
        self.pbReadHisNum.setEnabled(False)
        self.pbReadHisDate.setEnabled(False)
        self.pbSetSpkInfo.setEnabled(False)
        self.pbClrSpkInfo.setEnabled(False)
        self.pbReadSpkInfo.setEnabled(False)
        self.pbBatchPair1.setEnabled(False)
        self.pbBatchPair0.setEnabled(False)
        self.pbReadBeam.setEnabled(False)
        self.pbBeamRTLimit.setEnabled(False)
        self.pbBeamRTSet.setEnabled(False)
        self.pbBeamRTOffset.setEnabled(False)
        self.pbBeamAlarm.setEnabled(False)
        self.pbBeamHint.setEnabled(False)
        self.pbBeamExact.setEnabled(False)
        self.pbBeamLockRT.setEnabled(False)
        self.pbBeamAway.setEnabled(False)

    # 按钮启用
    def pbEnable(self):
        self.pbSend.setEnabled(True)
        self.pbBeaRead.setEnabled(True)
        self.pbCalibration.setEnabled(True)
        self.pbFlush.setEnabled(True)
        self.pbDeAir.setEnabled(True)
        self.pbFactory.setEnabled(True)
        self.pbSetType.setEnabled(True)
        self.pbSetCharacter.setEnabled(True)
        self.pbSetMode.setEnabled(True)
        self.pbSetPos.setEnabled(True)
        self.pbSetReturnT.setEnabled(True)
        self.pbSetPosLimit.setEnabled(True)
        if self.cbBeaType.isChecked():
            self.pbSetPosMax.setEnabled(True)
            self.pbSetPosMin.setEnabled(True)
            self.pbSetDead.setEnabled(True)
            self.pbSetP.setEnabled(True)
            self.pbSetI.setEnabled(True)
        self.pbSetPI.setEnabled(True)
        self.pbSetSpkLimit.setEnabled(True)
        self.pbSpkRead.setEnabled(True)
        self.pbSetRoomT.setEnabled(True)
        self.pbSetRTLimit.setEnabled(True)
        self.pbReadRToffset.setEnabled(True)
        self.pbSetRToffset.setEnabled(True)
        self.pbReadAlmT.setEnabled(True)
        self.pbSetAlmT.setEnabled(True)
        self.pbReadSeason.setEnabled(True)
        self.pbSetSeason.setEnabled(True)
        self.pbSetSeason1.setEnabled(True)
        self.pbSetSeason0.setEnabled(True)
        self.pbLockRoomT.setEnabled(True)
        self.pbExact.setEnabled(True)
        self.pbHint.setEnabled(True)
        self.pbAway.setEnabled(True)
        self.pbReadSpkClk.setEnabled(True)
        self.pbSyncSpkClk.setEnabled(True)
        self.pbReadHis.setEnabled(True)
        self.pbReadHisNum.setEnabled(True)
        self.pbReadHisDate.setEnabled(True)
        self.pbSetSpkInfo.setEnabled(True)
        self.pbClrSpkInfo.setEnabled(True)
        self.pbReadSpkInfo.setEnabled(True)
        self.pbBatchPair1.setEnabled(True)
        self.pbReadBeam.setEnabled(True)
        self.pbBeamRTLimit.setEnabled(True)
        self.pbBeamRTSet.setEnabled(True)
        self.pbBeamRTOffset.setEnabled(True)
        self.pbBeamAlarm.setEnabled(True)
        self.pbBeamHint.setEnabled(True)
        self.pbBeamExact.setEnabled(True)
        self.pbBeamLockRT.setEnabled(True)
        self.pbBeamAway.setEnabled(True)

    # 是否勾选Beacon Plus
    def selectPlus(self):
        if self.cbBeaType.isChecked():
            self.isPlus()
        else:
            self.notPlus()

    def isPlus(self):
        self.gbSpkLimit.setEnabled(True)
        self.radMode3.setEnabled(True)
        if self.ser.isOpen():
            self.pbSetP.setEnabled(True)
            self.pbSetI.setEnabled(True)
            self.pbSetPosMax.setEnabled(True)
            self.pbSetPosMin.setEnabled(True)
            self.pbSetDead.setEnabled(True)

    def notPlus(self):
        self.gbSpkLimit.setEnabled(False)
        self.radMode3.setEnabled(False)
        self.pbSetP.setEnabled(False)
        self.pbSetI.setEnabled(False)
        self.pbSetPosMax.setEnabled(False)
        self.pbSetPosMin.setEnabled(False)
        self.pbSetDead.setEnabled(False)

    def keyPressEvent(self, event) -> None:
        """ Ctrl + C复制表格内容 """
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_C:
            # 获取表格的选中行
            selected_ranges = self.tableData.selectedRanges()[0]  # 只取第一个数据块,其他的如果需要要做遍历,简单功能就不写得那么复杂了
            text_str = ""  # 最后总的内容
            # 行（选中的行信息读取）
            for row in range(selected_ranges.topRow(), selected_ranges.bottomRow() + 1):
                row_str = ""
                # 列（选中的列信息读取）
                for col in range(selected_ranges.leftColumn(), selected_ranges.rightColumn() + 1):
                    item = self.tableData.item(row, col)
                    row_str += item.text() + '\t'  # 制表符间隔数据
                text_str += row_str + '\n'  # 换行
            clipboard = qApp.clipboard()  # 获取剪贴板
            clipboard.setText(text_str)  # 内容写入剪贴板

    # 串口检测
    def portCheck(self):
        # 检测串口，将信息保存在字典中
        self.comDict = {}
        port_list = list(serial.tools.list_ports.comports())
        self.cbPortList.clear()
        for port in port_list:
            self.comDict["%s" % port[0]] = "%s" % port[1]
            self.cbPortList.addItem(port[0])
        if len(self.comDict) == 0:
            self.lbPortName.setText(" 无串口")

    # 串口信息
    def portInfo(self):
        # 显示选定的串口的详细信息
        info = self.cbPortList.currentText()
        if info != "":
            self.lbPortName.setText(self.comDict[self.cbPortList.currentText()])

    # 打开串口
    def portOpen(self):
        self.ser.port = self.cbPortList.currentText()
        self.ser.baudrate = int(self.cbBaudrate.currentText())
        self.ser.bytesize = int(self.cbBytesize.currentText())
        self.ser.stopbits = int(self.cbStopbits.currentText())
        self.ser.parity = self.cbParity.currentText()

        try:
            self.ser.open()
        except:
            QMessageBox.warning(self, "Port Error", "此不能打开此串口")
            return None

        # 打开串口接收定时器，周期为2ms
        self.timerReceive.start(2)

        if self.ser.isOpen():
            self.pbOpenPort.setEnabled(False)
            self.pbClosePort.setEnabled(True)
            self.gbPortStatus.setTitle("串口状态（已开启）")
            self.cbPortList.setEnabled(0)
            self.cbBaudrate.setEnabled(0)
            self.cbBytesize.setEnabled(0)
            self.cbParity.setEnabled(0)
            self.cbStopbits.setEnabled(0)
            self.pbEnable()

    # 关闭串口
    def portClose(self):
        self.timerReceive.stop()
        self.timerSend.stop()
        try:
            self.ser.close()
        except:
            pass
        self.pbOpenPort.setEnabled(True)
        self.pbClosePort.setEnabled(False)
        self.cbTimerSend.setChecked(False)
        # 接收数据和发送数据数目置零
        self.dataNumReceive = 0
        self.leNumReceive.setText(str(self.dataNumReceive))
        self.dataNumSend = 0
        self.leNumSend.setText(str(self.dataNumSend))
        self.gbPortStatus.setTitle("串口状态（已关闭）")
        self.cbPortList.setEnabled(1)
        self.cbBaudrate.setEnabled(1)
        self.cbBytesize.setEnabled(1)
        self.cbParity.setEnabled(1)
        self.cbStopbits.setEnabled(1)
        self.pbDisable()

    # 发送数据
    def dataSend(self):
        if self.ser.isOpen():
            input_s = self.teSend.toPlainText()
            # 非空字符串
            if input_s != "":
                # hex发送
                input_s = input_s.strip()
                send_list = []
                while input_s != '':
                    try:
                        num = int(input_s[0:2], 16)
                    except ValueError:
                        QMessageBox.warning(self, 'Wrong Data', '请输入十六进制数据，以空格分开!')
                        return None
                    input_s = input_s[2:].strip()
                    send_list.append(num)
                input_s = bytes(send_list)

                num = self.ser.write(input_s)
                self.dataNumSend += num
                self.leNumSend.setText(str(self.dataNumSend))
        else:
            pass

    # 接收数据
    def dataReceive(self):
        try:
            num = self.ser.inWaiting()
        except:
            self.portClose()
            return None
        if num > 0:
            data = self.ser.read(num)
            num = len(data)
            # hex显示
            out_s = ''
            for i in range(0, len(data)):
                out_s = out_s + '{:02X}'.format(data[i]) + ' '
            self.tbReceive.insertPlainText(out_s)
            # 统计接收字符的数量
            self.dataNumReceive += num
            self.leNumReceive.setText(str(self.dataNumReceive))
            # 获取到text光标
            textCursor = self.tbReceive.textCursor()
            # 滚动到底部
            textCursor.movePosition(textCursor.End)
            # 设置光标到text中去
            self.tbReceive.setTextCursor(textCursor)
        else:
            pass

    # 定时发送数据
    def dataTimerSend(self):
        if self.cbTimerSend.isChecked():
            self.timerSend.start(self.sbTimerSend.value() * 1000)
            self.sbTimerSend.setEnabled(False)
        else:
            self.timerSend.stop()
            self.sbTimerSend.setEnabled(True)

    # 接收窗口可编辑
    def receiveEditable(self):
        if self.cbEdit.isChecked():
            self.tbReceive.setReadOnly(False)
            self.tbReceive.setStyleSheet("background:rgb(255,255,255)")
        else:
            self.tbReceive.setReadOnly(True)
            self.tbReceive.setStyleSheet("background:rgb(240,240,240)")

    # 清除显示
    def dataReceiveClear(self):
        self.tbReceive.setText("")

    def dataSendClear(self):
        self.teSend.setText("")

    # 计算CS
    def calCS(self):
        data = self.teSend.toPlainText().replace(' ', '')
        if (len(data) % 2) != 0:
            QMessageBox.warning(self, 'Wrong Data', '回复长度错误')
        else:
            result = CS.Result(data).upper()
            if self.cbAutoInsert.isChecked():  # 自动把"CS 16"附加到报文结尾
                # 获取到text光标
                textCursor = self.teSend.textCursor()
                # 滚动到底部
                textCursor.movePosition(textCursor.End)
                # 设置光标到text中去
                self.teSend.setTextCursor(textCursor)
                self.teSend.insertPlainText(' ' + result + ' 16')
            self.leCS.setText(result)

    # 保存地址
    def saveAddress(self):
        # 地址长度位14位且只有列表中不存在的地址才会保存
        if len(self.cbBeaAddr.currentText()) == 14 and self.cbBeaAddr.findText(self.cbBeaAddr.currentText()) < 0:
            self.cbBeaAddr.insertItem(0, self.cbBeaAddr.currentText())
        else:
            pass
        if len(self.cbSpkAddr.currentText()) == 14 and self.cbSpkAddr.findText(self.cbSpkAddr.currentText()) < 0:
            self.cbSpkAddr.insertItem(0, self.cbSpkAddr.currentText())
        else:
            pass
        if len(self.cbBeamAddr.currentText()) == 14 and self.cbBeamAddr.findText(self.cbBeamAddr.currentText()) < 0:
            self.cbBeamAddr.insertItem(0, self.cbBeamAddr.currentText())
        else:
            pass

    # 检测控制命令是否有C4
    def checkC4(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(' ', '')
        if len(data) >= 20:
            if data[18:20] in ['c4', 'C4']:
                QMessageBox.warning(self, 'Wrong Response', '控制命令失败：C4')

    # 格式化Send报文并发送
    def format(self, data):
        cs = CS.Result(data)
        data = data + cs + "16"
        data = re.findall('.{2}', data)
        data = ' '.join(data)
        data = data.upper()
        self.teSend.setText(data)
        # 获取到text光标
        textCursor = self.teSend.textCursor()
        # 滚动到底部
        textCursor.movePosition(textCursor.End)
        # 设置光标到text中去
        self.teSend.setTextCursor(textCursor)
        self.tbReceive.setText("")
        self.dataSend()
        self.saveAddress()
        # 清空计数
        self.dataNumReceive = 0
        self.dataNumSend = 0

    # 判断地址长度是否合法
    def ifAddr(self, data):
        if len(data) != 14:
            QMessageBox.warning(self, 'Wrong Address', '地址长度应为14位')
            return False
        else:
            return True

    # 2字节转num
    def bytes2Num(self, data):
        lowByte = str(data)[:2]
        highByte = str(data)[2:]
        return int(highByte, 16) + (int(lowByte, 16) / 10)

    # num转2字节
    def num2bytes(self, data):
        data = str(int(data)).zfill(4)
        highByte = hex(int(data[:-1])).replace("0x", "").zfill(2)
        lowByte = hex(int(data[-1:])).replace("0x", "").zfill(2)
        return lowByte + highByte

    # num转单字节
    def num2byte(self, data):
        return hex(int(data)).replace("0x", "").zfill(2)

    # Beacon广播地址
    def beaBC(self):
        if self.cbBeaBC.isChecked():
            self.cbBeaAddr.setCurrentText("99999999991234")
            self.cbBeaAddr.setEnabled(0)
        else:
            self.cbBeaAddr.setEnabled(1)
            self.cbBeaAddr.setCurrentIndex(0)

    # 读取Beacon汇总数据
    def beaRead(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            if self.cbBeaType.isChecked():
                request = "6844" + addr + "0103111000"
            else:
                request = "6844" + addr + "0103101000"
            self.format(request)
            self.timerBeaAll.start(self.sbTimeout.value() * 1000)

    # 解析Beacon汇总数据
    def beaAnalysis(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) == 74 or len(data) == 96:
            address = data[4:14]  # 地址
            self.tableData.setItem(1, 1, QTableWidgetItem(address))
            pos = self.bytes2Num(data[30:34])  # 阀位反馈
            print(data[30:34])
            self.lePos.setText(str(pos))
            self.tableData.setItem(2, 1, QTableWidgetItem(str(pos) + " %"))
            returnT = self.bytes2Num(data[34:38])  # 回水温度反馈
            self.leReturnT.setText(str(returnT))
            self.tableData.setItem(3, 1, QTableWidgetItem(str(returnT) + " ℃"))
            setPos = self.bytes2Num(data[38:42])  # 阀位设定值
            self.sbSetPos.setValue(setPos)
            self.tableData.setItem(4, 1, QTableWidgetItem(str(setPos) + " %"))
            setReturnT = self.bytes2Num(data[42:46])  # 回水温度设定值
            self.sbSetReturnT.setValue(setReturnT)
            self.tableData.setItem(5, 1, QTableWidgetItem(str(setReturnT) + " ℃"))
            flag = int(data[28:30], 16)  # 运行模式
            if flag == 0:
                self.radMode0.setChecked(1)
            elif flag == 1:
                self.radMode1.setChecked(1)
            elif flag == 3:
                self.radMode3.setChecked(1)
            else:
                pass
            flag = int(data[46:48], 16)  # 阀门类型
            if flag == 1:
                self.radType1.setChecked(1)
            elif flag == 2:
                self.radType2.setChecked(1)
            elif flag == 3:
                self.radType3.setChecked(1)
            elif flag == 4:
                self.radType4.setChecked(1)
            elif flag == 5:
                self.radType5.setChecked(1)
            else:
                pass
            setP = int(data[48:50], 16)  # P系数
            self.sbSetP.setValue(setP)
            self.tableData.setItem(6, 1, QTableWidgetItem(str(setP)))
            setI = int(data[50:52], 16)  # I系数
            self.sbSetI.setValue(setI * 10)
            self.tableData.setItem(7, 1, QTableWidgetItem(str(setI)))
            max = int(data[52:54], 16)  # 阀位上限
            self.sbSetPosMax.setValue(max)
            self.tableData.setItem(8, 1, QTableWidgetItem(str(max) + " %"))
            min = int(data[54:56], 16)  # 阀位下限
            self.sbSetPosMin.setValue(min)
            self.tableData.setItem(9, 1, QTableWidgetItem(str(min) + " %"))
            dead = float(int(data[56:58], 16)) / 10  # 死区
            self.sbSetDead.setValue(dead)
            self.tableData.setItem(10, 1, QTableWidgetItem(str(dead) + " ℃"))
            flag = int(data[58:60], 16)  # 流量特性
            if flag == 0:
                self.radCharacter0.setChecked(1)
            elif flag == 1:
                self.radCharacter1.setChecked(1)
            elif flag == 2:
                self.radCharacter2.setChecked(1)
            else:
                pass
            beaVer = data[60:68]  # Beacon版本信息
            self.tableData.setItem(11, 1, QTableWidgetItem(beaVer))
            beaAlm = data[68:70]  # Beacon报警信息
            self.leBeaAlm.setText(beaAlm)
            self.tableData.setItem(12, 1, QTableWidgetItem(beaAlm))
            # 当报文长度为96时，判断为Beacon Plus回复报文
            if len(data) == 96:
                overTemp = data[70:72]  # 超温报警
                self.tableData.setItem(13, 1, QTableWidgetItem(overTemp))
                limitRoomT = self.bytes2Num(data[72:76])  # 室温限制值
                self.tableData.setItem(14, 1, QTableWidgetItem(str(limitRoomT) + " ℃"))
                loraBeaStat = data[76:78]  # Lora连接状态
                self.tableData.setItem(15, 1, QTableWidgetItem(loraBeaStat))
                loraInitRst = data[78:80]  # Lora初始化结果
                self.tableData.setItem(16, 1, QTableWidgetItem(loraInitRst))
                loraTxPower = int(data[80:82], 16)  # Lora发射功率
                self.tableData.setItem(17, 1, QTableWidgetItem(str(loraTxPower) + " dB"))
                loraOnPos = self.bytes2Num(data[82:86])  # Lora开位置
                self.sbSetSpkOn.setValue(int(loraOnPos))
                self.tableData.setItem(18, 1, QTableWidgetItem(str(loraOnPos) + " %"))
                loraOffPos = self.bytes2Num(data[86:90])  # Lora关位置
                self.sbSetSpkOff.setValue(int(loraOffPos))
                self.tableData.setItem(19, 1, QTableWidgetItem(str(loraOffPos) + " %"))
                loraBeaCh = int(data[90:92], 16)  # Lora工作信道
                self.tableData.setItem(20, 1, QTableWidgetItem(str(loraBeaCh)))
            else:
                pass
        else:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        self.timerBeaAll.stop()

    # 设定阀门类型
    def setType(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = ""
            if self.radType1.isChecked():
                request = "6844" + addr + "040401000001"
            elif self.radType2.isChecked():
                request = "6844" + addr + "040401000002"
            elif self.radType3.isChecked():
                request = "6844" + addr + "040401000003"
            elif self.radType4.isChecked():
                request = "6844" + addr + "040401000004"
            elif self.radType5.isChecked():
                request = "6844" + addr + "040401000005"
            else:
                QMessageBox.warning(self, 'Wrong Type', '未选择阀门类型')
            self.format(request)

    # 设定阀门特性
    def setCharacter(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = ""
            if self.radCharacter0.isChecked():
                request = "6844" + addr + "040408000000"
            elif self.radCharacter1.isChecked():
                request = "6844" + addr + "040408000001"
            elif self.radCharacter2.isChecked():
                request = "6844" + addr + "040408000002"
            else:
                QMessageBox.warning(self, 'Wrong Type', '未选择阀门特性')
            self.format(request)

    # 设定控制方式
    def setMode(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = ""
            if self.radMode0.isChecked():
                request = "6844" + addr + "04040F000000"
            elif self.radMode1.isChecked():
                request = "6844" + addr + "04040F000001"
            elif self.radMode3.isChecked():
                request = "6844" + addr + "04040F000003"
            else:
                QMessageBox.warning(self, 'Wrong Type', '未选择运行模式')
            self.format(request)

    # 远程校准
    def setCalibration(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = "6844" + addr + "0403050000"
            self.format(request)

    # 冲洗48秒
    def setFlush(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = "6844" + addr + "04050D00003000"
            self.format(request)

    # 排气
    def setDeAir(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = "6844" + addr + "04030E0000"
            self.format(request)

    # 恢复出厂设置
    def setFactory(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            request = "6844" + addr + "04030B0000"
            self.format(request)

    # 设定阀位
    def setPos(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2bytes(self.sbSetPos.value() * 10)
            request = "6844" + addr + "0405090000" + setpoint
            self.format(request)

    # 设定温度
    def setReturnT(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2bytes(self.sbSetReturnT.value() * 10)
            request = "6844" + addr + "0405100000" + setpoint
            self.format(request)

    # 设定阀位上下限和死区
    def setPosLimit(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            if float(self.sbSetDead.text()) * 10 in range(1, 100):
                dead = self.num2byte(self.sbSetDead.value() * 10)
                min = self.num2byte(self.sbSetPosMin.text())
                max = self.num2byte(self.sbSetPosMax.text())
                request = "6844" + addr + "0406017000" + dead + min + max
                self.format(request)

    # 设定Dead
    def setDead(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2byte(self.sbSetDead.value() * 10)
            request = "6844" + addr + "0404057000" + setpoint
            self.format(request)

    # 设定阀位下限
    def setPosMin(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2byte(self.sbSetPosMin.value())
            request = "6844" + addr + "0404047000" + setpoint
            self.format(request)

    # 设定阀位上限
    def setPosMax(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2byte(self.sbSetPosMax.value())
            request = "6844" + addr + "0404037000" + setpoint
            self.format(request)

    # 设定PI
    def setPI(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            p = self.num2byte(self.sbSetP.text())
            I = self.num2byte(self.sbSetI.value() / 10)
            request = "6844" + addr + "0405027000" + p + I
            self.format(request)

    # 设定P
    def setP(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2byte(self.sbSetP.value())
            request = "6844" + addr + "0404067000" + setpoint
            self.format(request)

    # 设定I
    def setI(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2byte(self.sbSetI.value() / 10)
            request = "6844" + addr + "0404077000" + setpoint
            self.format(request)

    # Spark控制模式开关位置限制
    def setSpkLimit(self):
        addr = self.cbBeaAddr.currentText()
        if self.ifAddr(addr):
            onPos = self.sbSetSpkOn.value() * 10
            offPos = self.sbSetSpkOff.value() * 10
            request = "6844" + addr + "0407211000" + self.num2bytes(onPos) + self.num2bytes(offPos)
            self.format(request)

    # Spark广播地址
    def spkBC(self):
        if self.cbSpkBC.isChecked():
            self.cbSpkAddr.setCurrentText("99999999991234")
            self.cbSpkAddr.setEnabled(0)
        else:
            self.cbSpkAddr.setEnabled(1)
            self.cbSpkAddr.setCurrentIndex(0)

    # 读取Spark汇总数据
    def spkRead(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            request = "6845" + addr + "0103003501"
            self.format(request)
            self.timerSpkAll.start(self.sbTimeout.value() * 1000)

    # 解析Spark汇总数据
    def spkAnalysis(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) != 82:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            if data[30:42] == '00EC00EC00EC':
                QMessageBox.information(self, 'Wrong Response', 'Spark数据未同步')
            address = data[4:14]  # 地址
            self.tableData.setItem(22, 1, QTableWidgetItem(address))
            roomAlm = data[28:30]  # 室内温度报警
            self.tableData.setItem(23, 1, QTableWidgetItem(roomAlm))
            roomT = self.bytes2Num(data[30:34])  # 实际室内温度
            self.leRoomT.setText(str(roomT))
            self.tableData.setItem(24, 1, QTableWidgetItem(str(roomT) + " ℃"))
            setRoomT = self.bytes2Num(data[34:38])  # 室内温度设定值
            self.sbSetRoomT.setValue(setRoomT)
            self.tableData.setItem(25, 1, QTableWidgetItem(str(setRoomT) + " ℃"))
            setRTLimit = self.bytes2Num(data[38:42])  # 室内温度限制值
            self.sbSetRTLimit.setValue(setRTLimit)
            self.tableData.setItem(26, 1, QTableWidgetItem(str(setRTLimit) + " ℃"))
            seasonStat = data[42:44]  # 供暖季状态
            if seasonStat == "00":
                self.leSeasonStat.setText("非供暖季")
            elif seasonStat == "01":
                self.leSeasonStat.setText("供暖季")
            else:
                self.leSeasonStat.setText("")
            self.tableData.setItem(27, 1, QTableWidgetItem(seasonStat))
            spkVer = data[44:52]  # Spark版本信息
            self.tableData.setItem(28, 1, QTableWidgetItem(spkVer))
            spkAlm = data[52:54]  # Spark报警信息
            self.tableData.setItem(29, 1, QTableWidgetItem(spkAlm))
            status = data[54:56]  # Spark开关机状态
            if status == "00":
                self.lbSpkStatus.setPixmap(QPixmap('../off.png'))
            elif status == "01":
                self.lbSpkStatus.setPixmap(QPixmap('../on.png'))
            else:
                pass
            self.tableData.setItem(30, 1, QTableWidgetItem(status))
            loraStat = data[56:58]  # Spark Lora连接状态
            self.tableData.setItem(31, 1, QTableWidgetItem(loraStat))
            loraMinute = str(int(data[58:60], 16))  # Sparl Load离线时间
            loraHour = str(int(data[60:62], 16))
            loraDay = str(int(data[62:64], 16))
            loraMonth = str(int(data[64:66], 16))
            loraYear = str(int(data[66:68], 16))
            self.tableData.setItem(32, 1, QTableWidgetItem(
                loraYear + '-' + loraMonth + '-' + loraDay + ' ' + loraHour + ':' + loraMinute))
            spkExact = data[68:70]  # 精确温控
            if spkExact == "00":
                self.cbExact.setChecked(False)
            elif spkExact == "01":
                self.cbExact.setChecked(True)
            else:
                pass
            self.tableData.setItem(33, 1, QTableWidgetItem(spkExact))
            lockRoomT = data[70:72]  # Spark 温度锁
            if lockRoomT == "00":
                self.cbLockRoomT.setChecked(False)
            elif lockRoomT == "01":
                self.cbLockRoomT.setChecked(True)
            else:
                pass
            self.tableData.setItem(34, 1, QTableWidgetItem(lockRoomT))
            away = data[72:74]  # Spark 离开模式
            if away == "00":
                self.cbAway.setChecked(False)
            elif away == "01":
                self.cbAway.setChecked(True)
            else:
                pass
            self.tableData.setItem(35, 1, QTableWidgetItem(away))
            loraSpkCH = int(data[74:76], 16)  # Spark 工作信道
            self.tableData.setItem(36, 1, QTableWidgetItem(str(loraSpkCH)))
            moduleVer = data[76:78]  # Spark 模块型号
            self.tableData.setItem(37, 1, QTableWidgetItem(moduleVer))
        self.timerSpkAll.stop()

    # 室温设定
    def setRoomT(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2bytes(self.sbSetRoomT.value() * 10)
            request = "6845" + addr + "0405063500" + setpoint
            self.format(request)

    # 室温锁定
    def lockRoomT(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            if self.cbLockRoomT.isChecked():
                request = "6845" + addr + "04040B350001"
                self.format(request)
            else:
                request = "6845" + addr + "04040B350000"
                self.format(request)

    # 室温设定限制
    def setRTLimit(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2bytes(self.sbSetRTLimit.value() * 10)
            request = "6845" + addr + "0405013500" + setpoint
            self.format(request)

    # 读取室温补偿
    def readRToffset(self):
        addr = self.leSpkAddr.text()
        if self.ifAddr(addr):
            request = "6845" + addr + "0103033500"
            self.format(request)
            self.timerRToffset.start(self.sbTimeout.value() * 1000)

    # 延时解析室温补偿
    def alsRToffset(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) != 36:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            if data[30] in ['F', 'f']:
                small = float(int(data[28:30], 16)) / 10
                big = float(int(data[30:32], 16)) - 256
                self.sbSetRToffset.setValue(small + big)
            else:
                self.sbSetRToffset.setValue(self.bytes2Num(data[28:32]))
        self.timerRToffset.stop()

    # 设定室温补偿
    def setRTOffset(self):
        addr = self.cbSpkAddr.currentText()
        data = self.sbSetRToffset.value()
        if self.ifAddr(addr):
            if data < 0:
                data = abs(data)
                if (data / 0.5) % 2 == 1:
                    setpoint = '05' + self.num2byte(256 - (data + 0.5))
                else:
                    setpoint = '00' + self.num2byte(256 - data)
            else:
                setpoint = self.num2bytes(data * 10)
            request = "6845" + addr + "0405033500" + setpoint
            self.format(request)

    # 读取报警阈值
    def readAlmT(self):
        addr = self.leSpkAddr.text()
        if self.ifAddr(addr):
            request = "6845" + addr + "0103023500"
            self.format(request)
            self.timerAlmT.start(self.sbTimeout.value() * 1000)

    # 延时解析报警阈值
    def alsAlmT(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) != 36:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            self.sbSetAlmT.setValue(self.bytes2Num(data[28:32]))
        self.timerAlmT.stop()

    # 设定报警阈值
    def setAlmT(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            setpoint = self.num2bytes(self.sbSetAlmT.value() * 10)
            request = "6845" + addr + "0405023500" + setpoint
            self.format(request)

    # 读取供暖季状态
    def readSeason(self):
        addr = self.leSpkAddr.text()
        if self.ifAddr(addr):
            request = "6845" + addr + "0103043500"
            self.format(request)
            self.timerSeason.start(self.sbTimeout.value() * 1000)

    # 延时解析供暖季状态
    def alsSeason(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) != 40:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            if data[28:36] == 'FFFFFFFF':
                self.leSeasonMode.setText('强制供暖季')
            elif data[28:36] == '00000000':
                self.leSeasonMode.setText('强制非供暖季')
            else:
                self.leSeasonMode.setText('自动')
                self.sbSetSeason1.setValue(int(data[28:30], 16))
                self.sbSetSeason2.setValue(int(data[30:32], 16))
                self.sbSetSeason3.setValue(int(data[32:34], 16))
                self.sbSetSeason4.setValue(int(data[34:36], 16))
        self.timerSeason.stop()

    # 供暖季设定
    def setSeason(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            on_m = self.num2byte(self.sbSetSeason1.value())
            on_d = self.num2byte(self.sbSetSeason2.value())
            off_m = self.num2byte(self.sbSetSeason3.value())
            off_d = self.num2byte(self.sbSetSeason4.value())
            request = "6845" + addr + "0407043500" + on_m + on_d + off_m + off_d
            self.format(request)

    # 强制供暖季
    def setSeason1(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            request = "6845" + addr + "0407043500FFFFFFFF"
            self.format(request)

    # 强制非供暖季
    def setSeason0(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            request = "6845" + addr + "040704350000000000"
            self.format(request)

    # 精确温控
    def exact(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            if self.cbExact.isChecked():
                request = "6845" + addr + "04040A350001"
                self.format(request)
            else:
                request = "6845" + addr + "04040A350000"
                self.format(request)

    # 提示符号
    def hint(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            if self.cbHint.isChecked():
                request = "6845" + addr + "040405350002"
                self.format(request)
            else:
                request = "6845" + addr + "040405350000"
                self.format(request)

    # 离开模式
    def away(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            if self.cbAway.isChecked():
                request = "6845" + addr + "04040C350001"
                self.format(request)
            else:
                request = "6845" + addr + "04040C350000"
                self.format(request)

    # 读取Spark时钟
    def readSpkClk(self):
        addr = self.leSpkAddr.text()
        if self.ifAddr(addr):
            request = "6845" + addr + "0103121000"
            self.format(request)
            self.timerSpkClk.start(self.sbTimeout.value() * 1000)

    # 解析Spark时钟
    def alsSpkClk(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) == 40:
            time_bytes = data[34:36] + data[32:34] + data[30:32] + data[28:30]
            time_stamp = int(time_bytes, 16) - 28800
            self.leSpkClk.setText('Epoch: ' + str(datetime.fromtimestamp(time_stamp)))
        elif len(data) == 46:
            Y = str(int(data[28:30], 16) + 2000)
            M = str(int(data[30:32], 16)).zfill(2)
            D = str(int(data[32:34], 16)).zfill(2)
            h = str(int(data[36:38], 16)).zfill(2)
            m = str(int(data[38:40], 16)).zfill(2)
            s = str(int(data[40:42], 16)).zfill(2)
            self.leSpkClk.setText('Direct: ' + Y + '-' + M + '-' + D + ' ' + h + ':' + m + ':' + s)
        else:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        self.timerSpkClk.stop()

    # 同步PC时钟至Spark
    def syncSpkClk(self):
        addr = self.leSpkAddr.text()
        if self.ifAddr(addr):
            if self.rbEpoch.isChecked():
                timeNow = datetime.now()
                unix_time = mktime(timeNow.timetuple()) + 28800
                bytes = hex(int(unix_time)).zfill(10)
                request = "6845" + addr + "0407121000" + bytes[8:] + bytes[6:8] + bytes[4:6] + bytes[2:4]
            elif self.rbDirect.isChecked():
                Y = self.num2byte(localtime().tm_year - 2000)
                M = self.num2byte(localtime().tm_mon)
                D = self.num2byte(localtime().tm_mday)
                W = self.num2byte(localtime().tm_wday + 1)
                h = self.num2byte(localtime().tm_hour)
                m = self.num2byte(localtime().tm_min)
                s = self.num2byte(localtime().tm_sec)
                request = "6845" + addr + "040A121000" + Y + M + D + W + h + m + s
            else:
                request = ''
            self.format(request)

    # 读取历史查询
    def readHis(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            year = self.num2byte(self.sbHisYear.value() - 2000)
            month = self.num2byte(self.sbHisMonth.value())
            day = self.num2byte(self.sbHisDay.value())
            request = "6845" + addr + "0106073500" + year + month + day
            self.format(request)
            self.timerSpkHis.start(self.sbTimeout.value() * 1000)

    # 解析历史数据
    def analysisHis(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) != 230:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            self.model = QStandardItemModel(48, 3)
            self.model.setHorizontalHeaderLabels(['时间', '室内温度', '设定温度'])
            self.tableView.setModel(self.model)
            time = ['00:00', '00:30', '01:00', '01:30', '02:00', '02:30', '03:00', '03:30', '04:00', '04:30', '05:00',
                    '05:30',
                    '06:00', '06:30', '07:00', '07:30', '08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00',
                    '11:30',
                    '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00',
                    '17:30',
                    '18:00', '18:30', '19:00', '19:30', '20:00', '20:30', '21:00', '21:30', '22:00', '22:30', '23:00',
                    '23:30']
            n = 0
            for x in time:
                self.model.setItem(n, 0, QStandardItem(x))
                n += 1
            his_act = list()
            his_set = list()
            for i in range(48):
                a1 = i * 4 + 34
                a2 = i * 4 + 36
                s1 = i * 4 + 36
                s2 = i * 4 + 38
                if data[a1:a2] == "FF":
                    his_act.append("--")
                else:
                    his_act.append(str(int(data[a1:a2], 16) / 2 - 20))
                if data[s1:s2] == "FF":
                    his_set.append("--")
                else:
                    his_set.append(str(int(data[s1:s2], 16) / 2 - 20))
            n = 0
            for x in his_act:
                self.model.setItem(n, 1, QStandardItem(x))
                n += 1
            n = 0
            for x in his_set:
                self.model.setItem(n, 2, QStandardItem(x))
                n += 1
        self.timerSpkHis.stop()

    # 读取日志总数
    def readHisNum(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            request = "6845" + addr + "0103083500"
            self.format(request)
            self.timerHisNum.start(self.sbTimeout.value() * 1000)

    # 解析日志总数
    def alsHisNum(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) != 36:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            num = int(data[30:32], 16) * 256 + int(data[28:30], 16)
            self.leHisNum.setText(str(num))
        self.timerHisNum.stop()

    # 读取日期索引
    def readHisDate(self):
        addr = self.cbSpkAddr.currentText()
        if self.ifAddr(addr):
            request = "6845" + addr + "0104093500" + self.num2byte(int(self.sbPage.text()))
            self.format(request)
            self.timerHisDate.start(self.sbTimeout.value() * 1000)
        self.lvHisDate.clear()

    # 解析日期索引
    def alsHisDate(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if float(len(data) - 32) % 6 != 0:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        else:
            l = int((len(data) - 32) / 6)
            dateList = ['' for i in range(l)]
            n = 0
            for n in range(l):
                m = n * 6
                dateList[n] = str(int(data[28 + m:30 + m], 16) + 2000).zfill(4) + '/' + str(
                    int(data[30 + m:32 + m], 16)).zfill(2) + '/' + str(int(data[32 + m:34 + m], 16)).zfill(2)
            self.lvHisDate.addItems(dateList)
        self.timerHisDate.stop()

    # 根据索引定位日期
    def setHisDate(self):
        date = self.lvHisDate.currentItem().text()
        Y = date[0:4]
        M = date[5:7]
        D = date[8:]
        self.sbHisYear.setValue(int(Y))
        self.sbHisMonth.setValue(int(M))
        self.sbHisDay.setValue(int(D))

    # 同步Beacon Plus地址用于配对
    def beaSync(self):
        if self.cbBeaSync.isChecked():
            self.leBeaAddr.setEnabled(False)
            self.leBeaAddr.setText(self.cbBeaAddr.currentText())
        else:
            self.leBeaAddr.setEnabled(True)

    # 同步Spark地址用于配对
    def spkSync(self):
        if self.cbSpkSync.isChecked():
            self.leSpkAddr.setEnabled(False)
            self.leSpkAddr.setText(self.cbSpkAddr.currentText())
        else:
            self.leSpkAddr.setEnabled(True)

    # 同步Beam地址用于配对
    def beamSync(self):
        if self.cbBeamSync.isChecked():
            self.leBeamAddr.setEnabled(False)
            self.leBeamAddr.setText(self.cbBeamAddr.currentText())
        else:
            self.leBeamAddr.setEnabled(True)

    # 写入Spark配对信息
    def setSpkInfo(self):
        addrBea = self.leBeaAddr.text()
        addrSpk = self.leSpkAddr.text()
        addrBeam = self.leBeamAddr.text()
        bcAddr = ['99999999991234',
                  'A4A4A4A4A41234',
                  'AAAAAAAAAA1234',
                  '99999999991235',
                  'A4A4A4A4A41235',
                  'AAAAAAAAAA1235',
                  '99999999991236',
                  'A4A4A4A4A41236',
                  'AAAAAAAAAA1236']
        if self.rbSpkSelBeacon.isChecked():
            if self.ifAddr(addrBea) and self.ifAddr(addrSpk):
                if addrBea in bcAddr or addrSpk in bcAddr:
                    QMessageBox.warning(self, 'Wrong Address', '广播地址不可用')
                else:
                    request = "6844" + addrBea + "040A201000" + addrSpk
                    self.format(request)
        elif self.rbSpkSelBeam.isChecked():
            if self.ifAddr(addrBeam) and self.ifAddr(addrSpk):
                seq = hex(self.sbBeamSeq.value())[2:].zfill(2)
                if addrBeam in bcAddr or addrSpk in bcAddr:
                    QMessageBox.warning(self, 'Wrong Address', '广播地址不可用')
                else:
                    request = "6846" + addrBeam + "040A" + seq + "3600" + addrSpk
                    self.format(request)

    # 清除Spark配对信息
    def clrSpkInfo(self):
        addrBea = self.leBeaAddr.text()
        addrBeam = self.leBeamAddr.text()
        if self.rbSpkSelBeacon.isChecked():
            if self.ifAddr(addrBea):
                request = "6844" + addrBea + "040A201000FFFFFFFFFF1235"
                self.format(request)
        if self.rbSpkSelBeam.isChecked():
            if self.ifAddr(addrBeam):
                seq = hex(self.sbBeamSeq.value())[2:].zfill(2)
                request = "6846" + addrBeam + "040A" + seq + "3600FFFFFFFFFFFFFF"
                self.format(request)

    # 读取Spark配对信息
    def readSpkInfo(self):
        addrBea = self.leBeaAddr.text()
        addrBeam = self.leBeamAddr.text()
        if self.rbSpkSelBeacon.isChecked():
            if self.ifAddr(addrBea):
                request = "6844" + addrBea + "0103201000"
                self.format(request)
                self.timerSpkInfo.start(self.sbTimeout.value() * 1000)
        if self.rbSpkSelBeam.isChecked():
            if self.ifAddr(addrBeam):
                request = "6846" + addrBeam + "0103010100"
                self.format(request)
                self.timerSpkInfo.start(self.sbTimeout.value() * 1000)
        self.lvSpkID.clear()

    # 延时解析Spark配对信息
    def spkAlsInfo(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        self.lvSpkID.clear()
        if self.rbSpkSelBeacon.isChecked():
            if len(data) != 46:
                QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
            else:
                self.lvSpkID.addItem(data[28:42])
        if self.rbSpkSelBeam.isChecked():
            if len(data) != 200:
                QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
            else:
                for i in range(0, 12):
                    info = data[28 + i * 14:42 + i * 14]
                    if info != 'FFFFFFFFFFFFFF':
                        self.lvSpkID.addItem(info + '(' + str(i) + ')')
        self.timerSpkInfo.stop()

    # 复制Spark ID到剪贴板
    def copySpkID(self):
        text_str = self.lvSpkID.currentItem().text()
        clipboard = qApp.clipboard()  # 获取剪贴板
        clipboard.setText(text_str)  # 内容写入剪贴板

    # 宽列表
    def tableWiden(self):
        self.resize(1645, 646)
        self.pbWiden.setVisible(False)
        self.pbNarrow.setVisible(True)
        self.tableData.setGeometry(1170, 34, 444, 565)

    # 窄列表
    def tableNarrow(self):
        self.resize(1420, 646)
        self.pbWiden.setVisible(True)
        self.pbNarrow.setVisible(False)
        self.tableData.setGeometry(1170, 34, 219, 565)

    # 导入Excel
    def importExcel(self):
        filename, filetype = QFileDialog.getOpenFileNames(self, "导入Excel文件", self.cwd, "Excel 文件(*.xls;*.xlsx)")
        if not filename:  # 如果filename值为空
            return
        else:
            data = xlrd3.open_workbook(filename[0])
            sheet = data.sheet_by_index(0)
            self.tableBatch.setRowCount(sheet.nrows)
            self.tableBatchNum = sheet.nrows
            self.tableBatch.clearContents()
            for i in range(0, sheet.nrows):
                v1 = sheet.cell(i, 0).value
                if v1:
                    d1 = float(v1)
                    s1 = "{:.0f}".format(d1)
                    s2 = "{:0>14s}".format(s1)
                    self.tableBatch.setItem(i, 1, QTableWidgetItem(s2))
                else:
                    continue
            for i in range(0, sheet.nrows):
                if self.rbSelBeam.isChecked():
                    s1 = self.leBatchAddr.text()
                else:
                    v1 = sheet.cell(i, 1).value
                    if v1:
                        d1 = float(v1)
                        s1 = "{:.0f}".format(d1)
                    else:
                        continue
                s2 = "{:0>14s}".format(s1)
                self.tableBatch.setItem(i, 2, QTableWidgetItem(s2))
            # 设定ProgressBar的最大值
            self.pgBatchPair.setMaximum(self.tableBatchNum)
            # 清空ProgressBar值
            self.pgBatchPair.setValue(0)

    # 开始批量配对
    def startBatchPair(self):
        if self.rbSelBeam.isChecked() and self.leBatchAddr.text() == '':
            QMessageBox.warning(self, 'Wrong Address', '请输入14位Beam地址')
            return
        else:
            if self.tableBatchNum == 0:
                QMessageBox.warning(self, 'Wrong Data', '导入列表为空')
            else:
                for i in range(0, self.tableBatchNum):
                    self.tableBatch.setItem(i, 0, QTableWidgetItem(''))
                self.work = workThread(self.tableBatchNum, 4)
                if self.rbSelBeacon.isChecked():
                    self.work.trigger.connect(self.batchBeaconPair)  # 并调用触发器返回的i值
                elif self.rbSelBeam.isChecked():
                    self.work.trigger.connect(self.batchBeamPair)  # 并调用触发器返回的i值
                self.work.start()
                self.gbBatchType.setEnabled(False)
                self.pbImpExcel.setEnabled(False)
                self.pbBatchPair1.setEnabled(False)
                self.pbBatchPair0.setEnabled(True)
                self.pgBatchPair.setValue(0)

    # 批量配对进度条
    def batchPairBar(self):
        print(self.tableBatch.rowCount())

    # 停止批量配对
    def stopBatchPair(self):
        self.work.stop()
        self.gbBatchType.setEnabled(True)
        self.pbImpExcel.setEnabled(True)
        self.pbBatchPair1.setEnabled(True)
        self.pbBatchPair0.setEnabled(False)

    # 批量配对
    def batchBeaconPair(self, i):
        if self.tableBatch.item(i, 2):  # 如果item(i, 2)不为空
            self.addr1 = self.tableBatch.item(i, 2).text()
        else:
            return
        if self.tableBatch.item(i, 1):
            self.addr2 = self.tableBatch.item(i, 1).text()
        else:
            return
        if self.ifAddr(self.addr1) and self.ifAddr(self.addr2):
            bcAddr = ['99999999991234', 'A4A4A4A4A41234', 'AAAAAAAAAA1234', '99999999991235', 'A4A4A4A4A41235',
                      'AAAAAAAAAA1235']
            if self.addr1 in bcAddr or self.addr2 in bcAddr:
                QMessageBox.warning(self, 'Wrong Address', '广播地址不可用')
                self.work.stop()
            else:
                request = "6844" + self.addr1 + "040A201000" + self.addr2
                self.tbReceive.clear()
                self.format(request)
                self.batchProcess = i
                self.timerBatch.start(2000)

    # 批量配对
    def batchBeamPair(self, i):
        self.addr1 = self.leBatchAddr.text()
        if self.tableBatch.item(i, 1):
            self.addr2 = self.tableBatch.item(i, 1).text()
        else:
            return
        if self.ifAddr(self.addr1) and self.ifAddr(self.addr2):
            bcAddr = ['99999999991236', 'A4A4A4A4A41236', 'AAAAAAAAAA1236', '99999999991235', 'A4A4A4A4A41235',
                      'AAAAAAAAAA1235']
            if self.addr1 in bcAddr or self.addr2 in bcAddr:
                QMessageBox.warning(self, 'Wrong Address', '广播地址不可用')
                self.work.stop()
            else:
                seq = hex(i)[2:].zfill(2)
                request = "6846" + self.addr1 + "040A" + seq + "3600" + self.addr2
                self.tbReceive.clear()
                self.format(request)
                self.batchProcess = i
                self.timerBatch.start(2000)

    # 判断配对是否成功
    def checkPair(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(' ', '')
        if len(data) != 46:
            return "回复异常"
        else:
            if self.rbSelBeacon.isChecked():
                if data[4:42] == self.addr1 + "840A201000" + self.addr2:
                    return "成功"
                else:
                    return "失败"
            elif self.rbSelBeam.isChecked():
                seq = hex(self.batchProcess)[2:].zfill(2)
                if data[4:42] == self.addr1 + "840A" + seq + "3600" + self.addr2:
                    return "成功"
                else:
                    return "失败"
            else:
                QMessageBox.warning(self, 'Wrong Type', '未选择访问参数')

    # 解析批量配对结果
    def alsBatchPair(self):
        self.timerBatch.stop()
        self.tableBatch.setItem(self.batchProcess, 0, QTableWidgetItem(self.checkPair()))
        barValue = self.batchProcess + 1
        self.pgBatchPair.setValue(barValue)  # 更新进度条数值
        if barValue == self.tableBatchNum:  # 当进度条100%时启用禁用响应按钮
            self.gbBatchType.setEnabled(True)
            self.pbImpExcel.setEnabled(True)
            self.pbBatchPair1.setEnabled(True)
            self.pbBatchPair0.setEnabled(False)

    # Beam和Beacon Plus选择（批量配对）
    def selBeamOrBeacon(self):
        if self.rbSelBeacon.isChecked():
            self.tableBatch.setHorizontalHeaderLabels(['Result', 'Spark', 'Beacon Plus'])
            self.tableBatch.setColumnHidden(2, False)
            self.leBatchAddr.setEnabled(False)
        if self.rbSelBeam.isChecked():
            self.leBatchAddr.setText(self.leBeamAddr.text())
            self.tableBatch.setHorizontalHeaderLabels(['Result', 'Spark', 'Beam'])
            self.tableBatch.setColumnHidden(2, True)
            self.leBatchAddr.setEnabled(True)

    # Beam和Beacon Plus选择（Spark配对）
    def spkSelBeamOrBeacon(self):
        if self.rbSpkSelBeacon.isChecked():
            if self.cbBeaSync.isChecked():
                self.leBeaAddr.setEnabled(False)
            else:
                self.leBeaAddr.setEnabled(True)
            self.cbBeaSync.setEnabled(True)
            self.leBeamAddr.setEnabled(False)
            self.sbBeamSeq.setEnabled(False)
            self.cbBeamSync.setEnabled(False)
        if self.rbSpkSelBeam.isChecked():
            self.leBeaAddr.setEnabled(False)
            self.cbBeaSync.setEnabled(False)
            if self.cbBeamSync.isChecked():
                self.leBeamAddr.setEnabled(False)
            else:
                self.leBeamAddr.setEnabled(True)
            self.sbBeamSeq.setEnabled(True)
            self.cbBeamSync.setEnabled(True)

    # 右键点击tableBeam
    def rightClick(self):
        if self.ser.isOpen():
            addr = self.cbBeamAddr.currentText()
            if self.ifAddr(addr):
                column = self.tableBeam.currentColumn()
                if column == 0:
                    self.rbBeamID.setChecked(True)
                elif column == 1:
                    self.rbBeamTx.setChecked(True)
                elif column == 2:
                    self.rbBeamSignal.setChecked(True)
                elif column == 3:
                    self.rbBeamTimeout.setChecked(True)
                elif column == 4:
                    self.rbBeamAlarm.setChecked(True)
                elif column == 5:
                    self.rbBeamRT.setChecked(True)
                elif column == 6:
                    self.rbBeamRTLimit.setChecked(True)
                elif column == 7:
                    self.rbBeamRTSet.setChecked(True)
                elif column == 8:
                    self.rbBeamRTOffset.setChecked(True)
                elif column == 9:
                    self.rbBeamAlarmT.setChecked(True)
                elif column == 10:
                    self.rbBeamHint.setChecked(True)
                elif column == 11:
                    self.rbBeamExact.setChecked(True)
                elif column == 12:
                    self.rbBeamLockRT.setChecked(True)
                elif column == 13:
                    self.rbBeamAway.setChecked(True)
                else:
                    QMessageBox.warning(self, 'Wrong Type', '访问参数错误')
                index = hex(column + 1)[2:].zfill(2)
                request = "6846" + addr + "0103" + index + "0100"
                self.format(request)
                self.timerBeam.start(self.sbTimeout.value() * 1000)
                self.gbReadBeam.setEnabled(False)
            else:
                return
        else:
            return

    # 读取Beam数据
    def readBeam(self):
        request = ''
        addr = self.cbBeamAddr.currentText()
        if self.ifAddr(addr):
            if self.rbBeamID.isChecked():
                request = "6846" + addr + "0103010100"
            elif self.rbBeamTx.isChecked():
                request = "6846" + addr + "0103020100"
            elif self.rbBeamSignal.isChecked():
                request = "6846" + addr + "0103030100"
            elif self.rbBeamTimeout.isChecked():
                request = "6846" + addr + "0103040100"
            elif self.rbBeamAlarm.isChecked():
                request = "6846" + addr + "0103050100"
            elif self.rbBeamRT.isChecked():
                request = "6846" + addr + "0103060100"
            elif self.rbBeamRTLimit.isChecked():
                request = "6846" + addr + "0103070100"
            elif self.rbBeamRTSet.isChecked():
                request = "6846" + addr + "0103080100"
            elif self.rbBeamRTOffset.isChecked():
                request = "6846" + addr + "0103090100"
            elif self.rbBeamAlarmT.isChecked():
                request = "6846" + addr + "01030A0100"
            elif self.rbBeamHint.isChecked():
                request = "6846" + addr + "01030B0100"
            elif self.rbBeamExact.isChecked():
                request = "6846" + addr + "01030C0100"
            elif self.rbBeamLockRT.isChecked():
                request = "6846" + addr + "01030D0100"
            elif self.rbBeamAway.isChecked():
                request = "6846" + addr + "01030E0100"
            else:
                QMessageBox.warning(self, 'Wrong Type', '未选择访问参数')
            self.format(request)
            self.timerBeam.start(self.sbTimeout.value() * 1000)
            self.gbReadBeam.setEnabled(False)
        else:
            return

    # 解析Beam数据
    def alsBeam(self):
        data = self.tbReceive.toPlainText()
        data = data.replace(" ", "")
        if len(data) == 200:
            if self.rbBeamID.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 14:42 + i * 14]
                    self.tableBeam.setItem(i, 0, QTableWidgetItem(info))
                self.transferSpkID()
            else:
                QMessageBox.warning(self, 'Wrong Type', '访问参数错误')
        elif len(data) == 56:
            if self.rbBeamTx.isChecked():
                for i in range(0, 12):
                    info = str(int(data[28 + i * 2:30 + i * 2]) + 1) + ' dB'
                    self.tableBeam.setItem(i, 1, QTableWidgetItem(info))
            elif self.rbBeamSignal.isChecked():
                for i in range(0, 12):
                    info = str(int(data[28 + i * 2:30 + i * 2], 16) - 256) + ' dbm'
                    self.tableBeam.setItem(i, 2, QTableWidgetItem(info))
            elif self.rbBeamTimeout.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 2:30 + i * 2]
                    self.tableBeam.setItem(i, 3, QTableWidgetItem(info))
            elif self.rbBeamAlarm.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 2:30 + i * 2]
                    self.tableBeam.setItem(i, 4, QTableWidgetItem(info))
            elif self.rbBeamHint.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 2:30 + i * 2]
                    self.tableBeam.setItem(i, 10, QTableWidgetItem(info))
            elif self.rbBeamExact.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 2:30 + i * 2]
                    self.tableBeam.setItem(i, 11, QTableWidgetItem(info))
            elif self.rbBeamLockRT.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 2:30 + i * 2]
                    self.tableBeam.setItem(i, 12, QTableWidgetItem(info))
            elif self.rbBeamAway.isChecked():
                for i in range(0, 12):
                    info = data[28 + i * 2:30 + i * 2]
                    self.tableBeam.setItem(i, 13, QTableWidgetItem(info))
            else:
                QMessageBox.warning(self, 'Wrong Type', '访问参数错误')
        elif len(data) == 80:
            if self.rbBeamRT.isChecked():
                for i in range(0, 12):
                    info = str(self.bytes2Num(data[28 + i * 4:32 + i * 4]))
                    self.tableBeam.setItem(i, 5, QTableWidgetItem(info))
            elif self.rbBeamRTLimit.isChecked():
                for i in range(0, 12):
                    info = str(self.bytes2Num(data[28 + i * 4:32 + i * 4]))
                    self.tableBeam.setItem(i, 6, QTableWidgetItem(info))
            elif self.rbBeamRTSet.isChecked():
                for i in range(0, 12):
                    info = str(self.bytes2Num(data[28 + i * 4:32 + i * 4]))
                    self.tableBeam.setItem(i, 7, QTableWidgetItem(info))
            elif self.rbBeamRTOffset.isChecked():
                for i in range(0, 12):
                    if data[30 + i * 4] in ['F', 'f']:
                        small = float(int(data[28 + i * 4:30 + i * 4], 16)) / 10
                        big = float(int(data[30 + i * 4:32 + i * 4], 16)) - 256
                        info = str(small + big)
                    else:
                        info = str(self.bytes2Num(data[28 + i * 4:32 + i * 4]))
                    self.tableBeam.setItem(i, 8, QTableWidgetItem(info))
            elif self.rbBeamAlarmT.isChecked():
                for i in range(0, 12):
                    info = str(self.bytes2Num(data[28 + i * 4:32 + i * 4]))
                    self.tableBeam.setItem(i, 9, QTableWidgetItem(info))
            else:
                QMessageBox.warning(self, 'Wrong Type', '访问参数错误')
        else:
            QMessageBox.warning(self, 'Wrong Data', '回复报文异常')
        self.timerBeam.stop()
        self.gbReadBeam.setEnabled(True)

    def transferSpkID(self):
        for i in range(0, 12):
            spkID = self.tableBeam.item(i, 0).text()
            if spkID != 'FFFFFFFFFFFFFF':
                if self.cbSpkAddr.findText(spkID) < 0:
                    self.cbSpkAddr.insertItem(99, spkID)
                else:
                    pass

    def beamRTLmit(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 6)
            if not data:
                data = 0
            else:
                data = float(data.text()) * 10
            data = self.num2bytes(data)
            s1 = s1 + data
        request = '6846' + addr + '041B070100' + s1
        self.format(request)

    def beamRTSet(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 7)
            if not data:
                data = 0
            else:
                data = float(data.text()) * 10
            data = self.num2bytes(data)
            s1 = s1 + data
        request = '6846' + addr + '041B080100' + s1
        self.format(request)

    def beamRTOffset(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 8)
            if not data:
                data = 0
            else:
                data = float(data.text())
                if data < 0:
                    data = abs(data)
                    if (data / 0.5) % 2 == 1:
                        data = '05' + self.num2byte(256 - (data + 0.5))
                        print(data)
                    else:
                        data = '00' + self.num2byte(256 - data)
                        print(data)
                else:
                    data = self.num2bytes(data * 10)
            s1 = s1 + data
        request = '6846' + addr + '041B090100' + s1
        self.format(request)

    def beamAlarm(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 9)
            if not data:
                data = 0
            else:
                data = float(data.text()) * 10
            data = self.num2bytes(data)
            s1 = s1 + data
        request = '6846' + addr + '041B0A0100' + s1
        self.format(request)

    def beamHint(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 10)
            if not data:
                data = 0
            else:
                data = float(data.text())
                if data == 0:
                    pass
                else:
                    data = 1
            data = self.num2byte(data)
            s1 = s1 + data
        request = '6846' + addr + '040F0B0100' + s1
        self.format(request)

    def beamExact(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 11)
            if not data:
                data = 0
            else:
                data = float(data.text())
                if data == 0:
                    pass
                else:
                    data = 1
            data = self.num2byte(data)
            s1 = s1 + data
        request = '6846' + addr + '040F0C0100' + s1
        self.format(request)

    def beamLockRT(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 12)
            if not data:
                data = 0
            else:
                data = float(data.text())
                if data == 0:
                    pass
                else:
                    data = 1
            data = self.num2byte(data)
            s1 = s1 + data
        request = '6846' + addr + '040F0D0100' + s1
        self.format(request)

    def beamAway(self):
        s1 = ''
        addr = self.cbBeamAddr.currentText()
        for i in range(0, 12):
            data = self.tableBeam.item(i, 13)
            if not data:
                data = 0
            else:
                data = float(data.text())
                if data == 0:
                    pass
                else:
                    data = 1
            data = self.num2byte(data)
            s1 = s1 + data
        request = '6846' + addr + '040F0E0100' + s1
        self.format(request)


# 继承QThread类
class workThread(QThread):
    trigger = pyqtSignal(int)

    def __init__(self, num, timeout):
        super().__init__()
        self.num = num
        self.timeout = timeout

    def run(self):
        for i in range(0, self.num):
            self.trigger.emit(int(i))  # 触发器返回一个i值
            time.sleep(self.timeout)

    def stop(self):
        self.terminate()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
