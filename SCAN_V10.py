# -*- coding: utf-8 -*-
# V10: 根据经验全面优化代码和界面
#      增加判断和结束FirmwareUpdate进程，使Booting的进程可以交替进行，最大化工作效率
import datetime
import logging
import os
import socket
import sys
import time
import pandas as pd
import pyodbc
import serial
import serial.tools.list_ports
import psutil
from PySide6.QtCore import Signal, QThread, Slot
from PySide6.QtWidgets import QApplication, QMessageBox, QMainWindow
from SCAN_V10_UI import Ui_MainWindow


def detect_com(com_port1, com_port2):
    """检测串口是否连接"""
    ports_list = list(serial.tools.list_ports.comports())
    com_flag = False
    for comport in ports_list:
        if comport[0] == com_port1 or comport[0] == com_port2:
            com_flag = True
    return com_flag


def port_open():
    """打开串口"""
    ser.port = COM_number  # 设置端口号
    ser.baudrate = 9600  # 设置波特率
    ser.bytesize = 8  # 设置数据位
    ser.stopbits = 1  # 设置停止位
    if ser.is_open:
        print("port is trying to close")
        send('<09200000000>')
        time.sleep(0.1)
        port_close()
        time.sleep(0.1)
        if not ser.is_open:
            print("port is successfully closed!")
        else:
            print("port fails to close")
    ser.open()  # 打开串口,要找到对的串口号才会成功


def port_close():
    """关闭串口"""
    ser.close()


def send(send_data):
    """发送串口编码"""
    ser.write(send_data.encode('utf-8'))  # utf-8 编码发送


def set_voltage(voltage):
    """设定电压编码函数，直接调用填入输入电压大小即可"""
    # ----------------------串口编码规则--------------------- #
    if 0 <= voltage < 10:
        voltage_str = "00" + str(voltage)
    elif 10 <= voltage < 100:
        voltage_str = "0" + str(voltage)
    else:
        voltage_str = str(voltage)
    # ------------------------连接串口----------------------- #
    port_open()
    # ------------------------发送编码----------------------- #
    send('<09100000000>')  # 连接
    time.sleep(0.1)
    send('<07000000000>')  # 启动电源
    time.sleep(0.1)
    send('<01' + voltage_str + '000000>')  # 设置电压
    time.sleep(0.1)
    send('<09200000000>')  # 断开连接
    time.sleep(0.1)
    # -----------------------断开串口------------------------- #
    port_close()
    # -----------------------结束通信------------------------- #


def current_read():
    try:
        """读取电源电流"""
        port_open()
        send('<09100000000>')  # 连接
        time.sleep(0.05)
        send("<04003300000>")  # 发送检测电流编码
        data_read = ser.read(26)
        # print(data_read.decode('utf-8'))
        current_code = data_read.decode('utf-8')
        current_read_str = current_code[16:22]
        current = int(current_read_str)
        print(current)
        time.sleep(0.05)
        send('<09200000000>')  # 断开连接
        time.sleep(0.05)
        port_close()
        return current
    except Exception as a:
        print("port has been open, cannot read current!")
        logging.debug(a)


def get_pid_by_name(process_name):
    pid = []
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        # print(proc)
        if proc.info['name'] == process_name:
            pid.append(proc.info['pid'])
            print(pid)
    return pid


def detect_process_pid(pid):
    flag = False
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['pid'] == pid:
            flag = True
        else:
            pass
    return flag


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.set_warning1 = 0
        self.set_warning2 = 0
        self.cp_number2 = None
        self.cp_number1 = None
        self.firmware_monitor_thread2 = None
        self.firmware_monitor_thread1 = None
        self.comport1_pid = None
        self.comport2_pid = None
        self.firmware_thread2 = None
        self.firmware_thread1 = None
        self.thread2 = None
        self.thread1 = None
        self.current_detect = 1
        self.monitor_thread = None
        self.check_status_2 = None
        self.pck2_7SJ586 = None
        self.pck4_path = None
        self.eth_pck = None
        self.pck1_7SJ586 = None
        self.eth_pck_6MD685 = None
        self.pck3_path = None
        self.check_status_1 = None
        self.setupUi(self)
        self.bind()
        self.init()
        self.detect_comport()  # 初始化串口
        self.detect_workstation()  # 初始化工作站
        self.start_current_monitor()  # 开始电流监视

    def init(self):
        """窗口初始化"""
        self.com_edit.setText("COM5")
        self.voltage_edit.setText("0")
        self.voltage_edit.setEnabled(False)
        self.com_edit.setEnabled(False)
        self.checkBox.setChecked(True)
        self.checkBox_1.setEnabled(False)
        self.checkBox_2.setEnabled(False)
        self.com_input1.setFocus()

    def bind(self):
        """窗体信号绑定"""
        self.pushButton_1.clicked.connect(self.start_comport1_firmware)
        self.pushButton_2.clicked.connect(self.start_comport2_firmware)
        self.pushButton_3.clicked.connect(self.custom_voltage)
        self.pushButton_4.clicked.connect(self.custom_source_port)
        self.com_input1.textChanged.connect(self.get_comport1_input)
        self.com_input2.textChanged.connect(self.get_comport2_input)
        self.pushButton_5.clicked.connect(self.set_voltage_warning)
        self.comboBox.currentIndexChanged.connect(self.choose_voltage_warning)
        self.pushButton.clicked.connect(self.reopen)
        self.actionopen_3.triggered.connect(lambda: os.startfile(r"运行日志"))
        self.actionopen_4.triggered.connect(lambda: os.startfile(r"版本说明V10.0.txt"))
        self.actionopen_5.triggered.connect(lambda: print("V10版本"))

    def set_voltage_warning(self):
        choice = QMessageBox.question(self, "提示", "是否电压回零？",
                                      QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if choice == QMessageBox.StandardButton.Yes:
            self.custom_voltage_0()
        else:
            pass

    def choose_voltage_warning(self):
        choice = QMessageBox.question(self, "提示", "是否设置电压？",
                                      QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if choice == QMessageBox.StandardButton.Yes:
            self.voltage_combobox_edit()
        else:
            pass

    def detect_workstation(self):
        """检测当前工作站是否存在"""
        if not flag_workstation:
            QMessageBox.information(self, 'warning',
                                    '(Your computer is not in the work Station!)\n该电脑不在产线电脑名单中!')
            sys.exit(0)
        global host_number
        self.label_workstation.setText(host_number)

    def detect_comport(self):
        """检测串口是否存在"""
        # 测试默认的电源串口通信是否通畅
        try:
            port_open()
            time.sleep(0.1)
            port_close()
        except:
            QMessageBox.information(self, 'warning', '(Cannot connect to COM5)\n无法通过COM5串口连接至电源')

    def set_comport1(self):
        pass

    def set_comport2(self):
        pass

    def get_comport1_input(self):
        """得到并处理第一个CP号信息"""
        """读取数据库装置信息"""
        self.set_warning1 = 0
        self.cp_number1 = self.com_input1.text()
        print(self.cp_number1)
        global MLFB1
        if len(self.cp_number1) == 12:
            """开始数据库检索"""
            logging.info('CP号为12')
            try:
                with pyodbc.connect(
                        u'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + database_path) as conn:  # 链接数据库
                    with conn.cursor() as cursor:  # 创建游标
                        sql = f"""SELECT ID,CPNum,MLFB,ISO_Date,ISO_Time FROM {tb_name} WHERE CPNum=?;"""
                        cursor.execute(sql, self.cp_number1)
                        data = cursor.fetchone()
                        MLFB1 = str(data[2])
                        self.textEdit_1.setText("MLFB: " + str(data[2]))
                        self.textEdit_1.append("ISO Date: " + str(data[3]))
                        self.textEdit_1.append("ISO Time: " + str(data[4]))
            except Exception as a:
                logging.debug(a)
                print("Access数据库繁忙，无法连接")
                MLFB1 = ""
            """检测PCK文件"""
            self.check_comport1_pck()
        elif len(self.cp_number1) == 0:
            logging.info('CP号为0')
            self.check_status_1 = True
            self.pck3_path = ""
            self.eth_pck_6MD685 = ""
            self.pck1_7SJ586 = ""
            self.eth_pck = ""
            self.checkBox_1.setChecked(False)
            self.textEdit_1.clear()
        else:
            self.check_status_1 = False
            self.pck3_path = ""
            self.eth_pck_6MD685 = ""
            self.pck1_7SJ586 = ""
            self.eth_pck = ""
            self.checkBox_1.setChecked(False)
            self.textEdit_1.clear()

    def get_comport2_input(self):
        """得到并处理第二个CP号信息"""
        """读取数据库装置信息"""
        self.set_warning2 = 0
        self.cp_number2 = self.com_input2.text()
        global MLFB2
        if len(self.cp_number2) == 12:
            """开始数据库检索"""
            try:
                with pyodbc.connect(
                        u'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + database_path) as conn:  # 链接数据库
                    # print('连接成功')
                    with conn.cursor() as cursor:  # 创建游标
                        sql = f"""SELECT ID,CPNum,MLFB,ISO_Date,ISO_Time FROM {tb_name} WHERE CPNum=?;"""
                        cursor.execute(sql, self.cp_number2)
                        data = cursor.fetchone()
                        MLFB2 = str(data[2])
                        self.textEdit_2.setText("MLFB: " + str(data[2]))
                        self.textEdit_2.append("ISO Date: " + str(data[3]))
                        self.textEdit_2.append("ISO Time: " + str(data[4]))
            except Exception as a:
                logging.debug(a)
                print("Access数据库繁忙，无法连接")
                MLFB2 = ""
            """检测PCK文件"""
            self.check_comport2_pck()
        elif len(self.cp_number2) == 0:
            self.pck4_path = ""
            self.eth_pck_6MD685 = ""
            self.pck2_7SJ586 = ""
            self.eth_pck = ""
            self.check_status_2 = True
            self.checkBox_2.setChecked(False)
            self.textEdit_2.clear()
        else:
            self.pck4_path = ""
            self.eth_pck_6MD685 = ""
            self.pck2_7SJ586 = ""
            self.eth_pck = ""
            self.check_status_2 = False
            self.checkBox_2.setChecked(False)
            self.textEdit_2.clear()

    def check_comport1_pck(self):
        """检查并处理第一个CP号对应的PCK文件"""
        """检测COM3 PCK文件"""
        cp_number1 = self.com_input1.text()
        pck3 = pck_path + cp_number1.upper() + ".PCK"
        print(pck3)
        try:
            if os.path.exists(pck3):
                self.checkBox_1.setChecked(True)
                self.textEdit_1.append("—成功获取PCK文件位置！")
                self.check_status_1 = True
                if MLFB1 == "1":
                    self.eth_pck = pck3
                    print('ethpck的值是', self.eth_pck)
                    return
                if MLFB1 == "":
                    self.pck3_path = pck3  # 暂定为这个
                    return
                elif len(self.com_input1.text()) == 12 and MLFB1[8] == "2":
                    try:
                        set_voltage(24)
                        self.textEdit_1.append("—注意使用24V电压(已自动设定)")
                        self.voltage_edit.setText("24")
                    except:
                        self.textEdit_1.append("(警告)自动设定24V电压失败，请手动设定")
                        print("try to close port")
                        set_voltage(24)
                elif len(self.com_input1.text()) == 12 and (
                        MLFB1[3:6] == "686" or MLFB1[3:5] == "58" or MLFB1[3:6] == "685") and MLFB1[8] == "4":
                    try:
                        set_voltage(24)
                        self.textEdit_1.append("—注意使用24V电压(已自动设定)")
                        self.voltage_edit.setText("24")
                    except:
                        self.textEdit_1.append("(警告)自动设定24V电压失败，请手动设定")
                        print("try to close port")
                        set_voltage(24)
                else:
                    try:
                        set_voltage(110)
                        self.textEdit_1.append("—已设定电源至110V")
                        self.voltage_edit.setText("110")
                    except:
                        self.textEdit_1.append("(警告)自动设定110V电压失败，请手动设定")
                        print("try to close port")
                        set_voltage(110)

                if (MLFB1[3:6] == "686") and (MLFB1[-2:] == "EE" or MLFB1[-2:] == "FF"):
                    if self.eth_pck != "":
                        self.check_status_1 = False
                        self.checkBox_1.setChecked(False)
                        QMessageBox.information(self, 'warning',
                                                '(ERROR! 2 pck files are both for MCP3!)\n不能同时进行两台MCP3装置')
                        self.com_input1.clear()
                        self.com_input2.clear()
                    else:
                        self.eth_pck = pck3
                        print('pck3=', self.eth_pck)
                elif MLFB1[0:6] == "6MD685":
                    self.eth_pck_6MD685 = pck3
                elif MLFB1[3:5] == "58":
                    self.pck1_7SJ586 = pck3
                else:
                    self.pck3_path = pck3
            else:
                self.textEdit_1.setText("COM1/3 PCK file do not find!)")
                self.check_status_1 = False
                self.checkBox_1.setChecked(False)
            if self.check_status_1 and self.checkBox.isChecked():
                print("开始执行comport1写入")
                try:
                    self.indicator_1.setText("调用中...")
                    self.thread1 = firmware_worker_thread(self.pck3_path, "", self.eth_pck, self.eth_pck_6MD685,
                                                          self.pck1_7SJ586, "")
                    self.thread1.finished_signal.connect(self.complete_comport1_firmware)
                    self.thread1.start()
                except Exception as a:
                    QMessageBox.warning(self, "提示", "未正常启动，请手动点击Start")
                    logging.info(a)
            else:
                pass
        except Exception as e:
            logging.debug("checkpck部分出现报错", e)
        if len(self.com_input1.text()) == 12 and len(self.com_input2.text()) == 12:
            self.pushButton_1.setFocus()
        else:
            self.com_input2.setFocus()

    def check_comport2_pck(self):
        """检查并处理第二个CP号对应的PCK文件"""
        """检测COM4 PCK文件"""
        cp_number2 = self.com_input2.text()
        pck4 = pck_path + cp_number2.upper() + ".PCK"
        if os.path.exists(pck4):
            self.checkBox_2.setChecked(True)
            self.textEdit_2.append("—成功获取PCK文件位置！")
            self.check_status_2 = True
            if MLFB2 == "1":
                self.eth_pck = pck4
                print('ethpck的值是', self.eth_pck)
                return
            if MLFB2 == "":
                logging.info("无法连接access，跳过")
                self.pck4_path = pck4  # 暂定为这个
                return
            elif len(self.com_input2.text()) == 12 and MLFB2[8] == "2":
                try:
                    set_voltage(24)
                    self.textEdit_2.append("—注意使用24V电压(已自动设定)")
                    self.voltage_edit.setText("24")
                except:
                    self.textEdit_2.append("(警告)自动设定24V电压失败，请手动设定")
                    print("try to close port")
                    set_voltage(24)
            elif len(self.com_input2.text()) == 12 and (
                    MLFB2[3:6] == "686" or MLFB2[3:5] == "58" or MLFB2[3:6] == "685") and MLFB2[8] == "4":
                try:
                    set_voltage(24)
                    self.textEdit_2.append("—注意使用24V电压(已自动设定)")
                    self.voltage_edit.setText("24")
                except:
                    self.textEdit_2.append("(警告)自动设定24V电压失败，请手动设定")
                    print("try to close port")
                    set_voltage(24)
            else:
                try:
                    set_voltage(110)
                    self.textEdit_2.append("—已设定电源至110V")
                    self.voltage_edit.setText("110")
                except:
                    self.textEdit_2.append("(警告)自动设定110V电压失败，请手动设定")
                    print("try to close port")
                    set_voltage(110)

            """根据MLFB寻找PCK文件"""
            if (MLFB2[3:6] == "686") and (MLFB2[-2:] == "EE" or MLFB2[-2:] == "FF"):
                if self.eth_pck != "":
                    self.check_status_2 = False
                    self.checkBox_2.setChecked(False)
                    QMessageBox.information(self, 'warning',
                                            '(ERROR! 2 pck files are both for MCP3!)\n不能同时进行两台MCP3装置')
                    self.com_input1.clear()
                    self.com_input2.clear()
                else:
                    self.eth_pck = pck4
            elif MLFB2[0:6] == "6MD685":
                self.eth_pck_6MD685 = pck4
            elif MLFB2[3:5] == "58":
                self.pck2_7SJ586 = pck4
            else:
                self.pck4_path = pck4
            if self.check_status_2 and self.checkBox.isChecked():
                print("start comport2 write")
                try:
                    self.indicator_2.setText("调用中...")
                    self.thread2 = firmware_worker_thread("", self.pck4_path, self.eth_pck, self.eth_pck_6MD685, "",
                                                          self.pck2_7SJ586)
                    self.thread2.finished_signal.connect(self.complete_comport2_firmware)
                    print("start2")
                    self.thread2.start()
                except Exception as a:
                    QMessageBox.warning(self, "提示", "未正常启动，请手动点击Start")
                    logging.info(a)
            else:
                pass
        else:
            self.textEdit_2.setText("(COM2/4 PCK file do not find!)")
            self.check_status_2 = False
            self.checkBox_2.setChecked(False)
        if len(self.com_input1.text()) == 12 and len(self.com_input2.text()) == 12:
            self.pushButton_2.setFocus()
        else:
            self.com_input1.setFocus()

    def access_write(self):
        """信息数据库写入"""
        """booting信息写入数据库"""
        print("开始数据库写入")
        now_date = time.strftime('%m/%d/%Y', time.localtime())  # 当前日期
        now_time = time.strftime('%H:%M', time.localtime())  # 当前具体时间
        """开始数据库写入"""
        try:
            with pyodbc.connect(
                    u'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + database_path) as conn:  # 链接数据库
                print('连接成功')
                conn.autocommit = True
                with conn.cursor() as cursor:  # 创建游标
                    cursor = conn.cursor()  # 创建游标
                    print('尝试写入')
                    if self.cp_number1 != "":
                        self.textEdit_1.append("—写入数据库中...")
                        try:
                            cursor.execute(
                                f"""UPDATE {tb_name} SET Runup_Date=?, Runup_Time=?, Runup_Station=? WHERE CPNum=?;""",
                                (now_date, now_time, host_number, self.cp_number1))  # 执行SQL更新数据库
                            self.textEdit_1.append("—写入数据库完成!")
                            print('写入完成')
                            self.cp_number1 = ""
                        except Exception as a:
                            print(f"写入access过程中出现{a}")
                    if self.cp_number2 != "":
                        self.textEdit_2.append("—写入数据库中...")
                        try:
                            cursor.execute(
                                f"""UPDATE {tb_name} SET Runup_Date=?, Runup_Time=?, Runup_Station=? WHERE CPNum=?;""",
                                (now_date, now_time, host_number, self.cp_number2))  # 执行SQL更新数据库
                            self.textEdit_2.append("—写入数据库完成!")
                            print('写入完成')
                            self.cp_number2 = ""
                        except Exception as a:
                            print(f"写入access过程中出现{a}")
        except Exception as a:
            logging.debug(time.strftime(
                '%y-%m-%d %H:%M:%S') + '数据库写入Running up数据失败。' + '\nRunup Station:' + host_number + '\nRunup Date:' + now_date + '\nRunup Time' + now_time)
            if 'conn' in locals() and conn is not None:
                # conn 已经初始化并赋值
                cursor.close()  # 关闭游标
                conn.close()  # 关闭链接
            else:
                pass

    def start_current_monitor(self):
        """启动电流监视"""
        self.monitor_thread = current_monitor_thread()
        self.monitor_thread.current_value_signal.connect(self.update_current_display)
        self.monitor_thread.start()

    @Slot(int)
    def update_current_display(self, current):
        self.current_detect = current
        self.current_monitor.setText(str(current))

    @Slot(bool)
    def time_out(self):
        QMessageBox.warning(self, "提示", "调用firmware程序等待超时！")

    def start_comport1_firmware(self):
        """调用firmware开始第一个串口写入"""
        self.set_warning1 = 0
        if self.check_status_1:
            print("start comport1")
            self.indicator_1.setText("调用中...")
            self.thread1 = firmware_worker_thread(self.pck3_path, "", self.eth_pck, self.eth_pck_6MD685,
                                                  self.pck1_7SJ586, "")
            self.thread1.timeout_signal.connect(self.time_out)
            self.thread1.finished_signal.connect(self.complete_comport1_firmware)
            self.thread1.start()
        else:
            print("cannot start thread1")

    def start_comport2_firmware(self):
        """调用firmware开始第二个串口写入"""
        self.set_warning2 = 0
        if self.check_status_2:
            print("start comport2")
            self.indicator_2.setText("调用中...")
            self.thread2 = firmware_worker_thread("", self.pck4_path, self.eth_pck, self.eth_pck_6MD685, "",
                                                  self.pck2_7SJ586)
            self.thread2.timeout_signal.connect(self.time_out)
            self.thread2.finished_signal.connect(self.complete_comport2_firmware)
            print("start2")
            self.thread2.start()
        else:
            print("cannot start thread2")

    def complete_comport1_firmware(self):
        """comport1写入程序调用完成"""
        self.setFocus()
        self.indicator_1.setText("writing...")
        self.firmware_thread1 = get_firmware_thread()
        self.firmware_thread1.pid_detect_signal.connect(self.update_comport1_pid)
        self.firmware_thread1.start()
        self.access_write()

    def complete_comport2_firmware(self):
        """comport2写入程序调用完成"""
        self.setFocus()
        self.indicator_2.setText("writing...")
        self.firmware_thread2 = get_firmware_thread()
        self.firmware_thread2.pid_detect_signal.connect(self.update_comport2_pid)
        self.firmware_thread2.start()
        self.access_write()

    @Slot(list)
    def update_comport1_pid(self, pid_list):
        """获取第一个firmware程序的pid"""
        if len(pid_list) == 1:
            self.comport1_pid = pid_list[0]
        elif len(pid_list) == 2:
            for pid in pid_list:
                if pid == self.comport2_pid:
                    pass
                else:
                    self.comport1_pid = pid
                    print("pid1:", self.comport1_pid)
        print(self.comport1_pid)
        self.firmware_monitor_thread1 = firmware_monitor_thread(self.comport1_pid)
        self.firmware_monitor_thread1.firmware_end_signal.connect(self.end_comport1_write)
        self.firmware_monitor_thread1.start()
        self.com_input1.setEnabled(False)

    @Slot(list)
    def update_comport2_pid(self, pid_list):
        """获取第二个firmware程序的pid"""
        if len(pid_list) == 1:
            self.comport2_pid = pid_list[0]
        elif len(pid_list) == 2:
            for pid in pid_list:
                if pid == self.comport1_pid:
                    pass
                else:
                    self.comport2_pid = pid
                    print("pid2:", self.comport2_pid)
        print(self.comport2_pid)
        self.firmware_monitor_thread2 = firmware_monitor_thread(self.comport2_pid)
        self.firmware_monitor_thread2.firmware_end_signal.connect(self.end_comport2_write)
        self.firmware_monitor_thread2.start()
        self.com_input2.setEnabled(False)

    def end_comport1_write(self):
        self.com_input1.clear()
        self.com_input1.setEnabled(True)
        self.indicator_1.setText("complete!")
        self.com_input1.setFocus()
        self.set_warning1 = 1
        self.judge_end_warning()

    def end_comport2_write(self):
        self.com_input2.clear()
        self.com_input2.setEnabled(True)
        self.indicator_2.setText("complete!")
        self.com_input2.setFocus()
        self.set_warning2 = 1
        self.judge_end_warning()

    def custom_voltage(self):
        """自定义电压"""
        try:
            if self.voltage_edit.isEnabled():
                choice = QMessageBox.question(self, "提示", "是否设置电压？",
                                              QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if choice == QMessageBox.StandardButton.Yes:
                    self.voltage_edit.setEnabled(False)
                    set_voltage(int(self.voltage_edit.text()))
                else:
                    self.voltage_edit.setEnabled(False)
            else:
                self.voltage_edit.setEnabled(True)
        except:
            print("cannot set voltage, trying to set again")
            set_voltage(int(self.voltage_edit.text()))

    def custom_voltage_0(self):
        """电压回零"""
        try:
            set_voltage(0)
        except:
            print("cannot set 0, trying to set again")
            set_voltage(0)

    def custom_source_port(self):
        global COM_number
        """自定义电源串口"""
        if self.com_edit.isEnabled():
            self.com_edit.setEnabled(False)
            COM_number = self.com_edit.text()
            print(COM_number)
        else:
            self.com_edit.setEnabled(True)

    def voltage_combobox_edit(self):
        try:
            self.voltage_edit.setText(self.comboBox.currentText())
            set_voltage(int(self.voltage_edit.text()))
        except:
            print("cannot set voltage, trying to set again")
            set_voltage(int(self.voltage_edit.text()))

    def judge_end_warning(self):
        if self.set_warning1 == 1 and self.set_warning2 == 1:
            choice = QMessageBox.question(self, "提示", "是否电压回零？",
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            self.set_warning1 = 0
            self.set_warning2 = 0
            if choice == QMessageBox.StandardButton.Yes:
                try:
                    self.custom_voltage_0()
                except:
                    print("cannot set voltage, trying to set again")
                    self.custom_voltage_0()
            else:
                pass
        else:
            pass

    def reopen(self):
        choice = QMessageBox.question(self, "提示", "是否要重启程序？",
                                      QMessageBox.StandardButton.No | QMessageBox.StandardButton.Yes)
        if choice == QMessageBox.StandardButton.Yes:
            os.system("start SCAN_V10.exe")
            app.exit()
        else:
            pass


class current_monitor_thread(QThread):
    """电流监视线程"""
    current_value_signal = Signal(int)
    end_monitor_signal = Signal(bool)

    def __init__(self):
        super().__init__()

    def run(self):
        while 1:
            time.sleep(10)
            current = current_read()
            self.current_value_signal.emit(current)
            # if current >= 1:
            #     pass
            # else:
            #     break
        # self.end_monitor_signal.emit(True)


class get_firmware_thread(QThread):
    pid_detect_signal = Signal(list)

    def __init__(self):
        super().__init__()

    def run(self):
        firmware_pid_list = get_pid_by_name("FirmwareUpdate.exe")
        print(firmware_pid_list)
        while 1:
            if firmware_pid_list:
                self.pid_detect_signal.emit(firmware_pid_list)
                break


class firmware_monitor_thread(QThread):
    firmware_end_signal = Signal(bool)

    def __init__(self, comport_pid):
        super().__init__()
        self.comport_pid = comport_pid

    def run(self):
        while 1:
            if detect_process_pid(self.comport_pid):
                print("检测到firmware程序正在运行，继续监视", self.comport_pid)
                pass
            else:
                self.firmware_end_signal.emit(True)
                print("结束监视")
                break
            time.sleep(5)


class firmware_worker_thread(QThread):
    """主工作线程firmware update"""
    finished_signal = Signal(bool)
    timeout_signal = Signal(bool)

    def __init__(self, pck3_path, pck4_path, eth_pck, eth_pck_6MD685, pck1_7SJ586, pck2_7SJ586, parent=None):
        super().__init__(parent)
        self.pck3_path = pck3_path
        self.pck4_path = pck4_path
        self.eth_pck = eth_pck
        self.eth_pck_6MD685 = eth_pck_6MD685
        self.pck1_7SJ586 = pck1_7SJ586
        self.pck2_7SJ586 = pck2_7SJ586
        self.x = 0
        self.y = 0

    def run(self):
        print('self.pck3_path is ', self.pck3_path)
        print('self.pck4_path is ', self.pck4_path)
        print('self.eth_pck is ', self.eth_pck)
        print('eth_pck_6MD685 is ', self.eth_pck_6MD685)
        while 1:
            try:
                if self.pck3_path == "" and self.pck4_path == "" and self.eth_pck == "" and self.eth_pck_6MD685 == "" and self.pck1_7SJ586 == "" and self.pck2_7SJ586 == "":
                    print('stop testing')
                    break

                if self.pck1_7SJ586 != "" and serial.Serial("com1").is_open:
                    os.system(
                        "start D:\\Firmware_update\\FirmwareUpdate.exe " + self.pck1_7SJ586 + " -c1")
                    self.pck1_7SJ586 = ""
                    time.sleep(1)

                if self.pck2_7SJ586 != "" and serial.Serial("com2").is_open:
                    time.sleep(2)  # 等待2s，防止两个线程同时调用firmware时发生错误
                    os.system(
                        "start D:\\Firmware_update\\FirmwareUpdate.exe " + self.pck2_7SJ586 + " -c2")
                    self.pck2_7SJ586 = ""
                    time.sleep(1)

                if self.pck3_path != "" and serial.Serial("com3").is_open:
                    os.system(
                        "start D:\\Firmware_update\\FirmwareUpdate.exe " + self.pck3_path + " -c3")
                    print('test_pck3')
                    self.pck3_path = ""
                    time.sleep(1)

                if self.pck4_path != "" and serial.Serial("com4").is_open:
                    time.sleep(2)  # 等待2s，防止两个线程同时调用firmware时发生错误
                    os.system(
                        "start D:\\Firmware_update\\FirmwareUpdate.exe " + self.pck4_path + " -c4")
                    #os.system(
                    #   "start D:\\Firmware_update\\FirmwareUpdate.exe " + pck4_path + " -c4")
                    print('test_pck4')
                    self.pck4_path = ""
                    time.sleep(1)

                # if self.eth_pck != "" and os.popen("ping 192.168.253.253").read().count("timed out") < 2:
                if self.eth_pck != "":  # and os.popen("ping 192.168.253.253").read().count("timed out") < 2
                    print('wait for 30s')
                    time.sleep(30)  ## delay for 7SJ686 EE， TFTP transfer
                    #os.system(
                    #   "start D:\\Firmware_update\\FirmwareUpdate.exe " + eth_pck + " -c256")
                    os.system(
                        "start D:\\Firmware_update\\FirmwareUpdate.exe " + self.eth_pck + " -c256")
                    self.eth_pck = ""
                    self.finished_signal.emit("Update complete_eth")
                    break

                if self.eth_pck_6MD685 != "" and os.popen("ping 192.168.253.253").read().count("timed out") < 2:
                    # os.system(
                    #    "start D:\\Firmware_update\\FirmwareUpdate.exe " + eth_pck_6MD685 + " -c255")
                    print('test_eth_pck4')
                    os.system(
                        "start D:\\Firmware_update\\FirmwareUpdate.exe " + self.eth_pck_6MD685 + " -c255")
                    self.eth_pck_6MD685 = ""

                # 发送完成信号
                self.finished_signal.emit("Update complete")
                break
            except Exception as a:
                self.x += 1
                print(self.x)
                time.sleep(1)
                if self.x >= 300:
                    self.timeout_signal.emit("Time Out!")
                    logging.info("调用firmware等待超时")
                    break


if __name__ == "__main__":
    # 获取今日日期
    today = datetime.date.today()
    # 打开日志文件
    logging.basicConfig(level='DEBUG', filename=f'./运行日志/运行日志{today}.log', filemode='a+')  # 创建/打开日志文件，模式为追加写入

    # 路径变量初始化
    host_filepath = ""  # MF固资盘点表
    database_path = ""  # Production_Database数据库
    pck_path = ""  # PCK文件所在路径位置

    # 获取相关文件路径
    filepath = r"J:\Department\MF\RPA\Configuration Flies\PO_DT_config.xlsx"  # 配置文件路径
    file = pd.read_excel(filepath)
    read_name = file['Name'].tolist()
    read_path = file['Value'].tolist()
    m = 0
    while m < len(read_name):
        if read_name[m] == "MF workstation+human":
            host_filepath = read_path[m]
        if read_name[m] == "SPA_ProductionDatabase":
            database_path = read_path[m]
        if read_name[m] == "PCK_Folder":
            pck_path = read_path[m]
        m = m + 1

    # 获取当前工作站信息
    host_number = ""
    hostname = socket.gethostname()  # 获取当前电脑的名称
    print(hostname)
    host_file = pd.read_excel(host_filepath)
    win10_client_name = host_file['Win10 Client\nName'].tolist()
    Number = host_file['编号'].tolist()
    n = 0
    flag_workstation = False
    while n < len(win10_client_name) and flag_workstation is False:
        if win10_client_name[n] == hostname:
            flag_workstation = True
            host_number = Number[n]
            print("工作站:", host_number)
        else:
            n += 1
    flag_workstation = True
    print(flag_workstation)

    # 初始化串口
    ser = serial.Serial()
    COM_number = "COM5"

    # 初始化数据库
    tb_name = "ProductionData"

    # 执行GUI
    app = QApplication([])  # Qt应用程序
    window = MyWindow()  # 打开窗口
    window.show()  # 显示窗口
    app.exec()  # 执行应用程序
