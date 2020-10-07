import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import filedialog
from tkinter import messagebox
import threading
import time
import subprocess
import types
import os
import sys
import struct
import re
import time
import platform
import ctypes
import hashlib
import configparser
import platform
import pyqrcode
import png
# import pyzbar.pyzbar import decode
import qrtools
from scanner import Scanner
import requests
import json
from datetime import datetime
import time

from serial import Serial, SerialException
from serial import PARITY_EVEN, PARITY_MARK, PARITY_NAMES, PARITY_NONE, PARITY_ODD, PARITY_SPACE
from serial import STOPBITS_ONE, STOPBITS_ONE_POINT_FIVE , STOPBITS_TWO
from serial import FIVEBITS, SIXBITS, SEVENBITS, EIGHTBITS
import glob

__author__ = "Cosmas Eric. s"
__copyright__ = "Copyright 2020, Serial communication project"

__license__ = "GPL"
__version__ = "1.1.1"
__maintainer__ = "Cosmas Eric "
__email__ = "cosmas.eric.septian@polytron.co.id"
__status__ = "Internal Test Beta"


CMD_GET_SN = "getprop ro.serialno\r"
#CMD_GET_MAC = "ip addr show wlan0  | grep 'link/ether '| cut -d' ' -f6\r"
CMD_GET_MAC = "ip addr show | grep 'link/ether' | cut -d' ' -f6\r"

listScanMac = []
VENDOR_ID = 0x27DD
PRODUCT_ID = 0x0103
URL_LINK = "http://10.8.42.44/mola/scan/scanHKC"
USERNAME = "snreader"
PASSWORD = "HkcSn20"
LOG_PATH = "./log.txt"

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack(fill=BOTH, expand=1)
        main_frame = Frame(master)
        main_frame.pack(fill="y", expand=1)
        
        #scan insert clear mac
        #self.insertMac("abc")
        #self.insertMac("def")
        #print(listScanMac)
        #for x in listScanMac:
        #    print(x)
        #self.clearListMac()
        #print(listScanMac)
        #create qrcode
        #qr = pyqrcode.create("test1")
        #qr.png("D:\Python Project\pyserialcom", scale=6)
        
        # qr = qrtools.QR()
        # qr.decode("test1.png")
        # s = qr.data
        # print("The decoded QR code is: %s" % s)

        self.serial_property = {
            "port1" : "",
            "port2" : "",
            "port3" : "",
            "baud" : 9600,
            "data" : 8,
            "parity" : "None",
            "stop" : 1,
        }

        self.default_filename = "hkc_sn_mac"

        self.uart1_open = False
        self.uart2_open = False
        self.uart3_open = False

        self.createWidgets(main_frame)
        self.dl_thread = 0
        self.directory_path = None

        master.protocol("WM_DELETE_WINDOW", self.closeWin)

        # threading.Thread(target=self.check_files).start()

        # create serial
        self.ser1 = Serial()
        self.ser2 = Serial()
        self.ser3 = Serial()

        self.rx_count = 0
        self.output_file = None

        # data to be capture
        self.data_query = {
            "sn" : "",
            "mac" : "",
        }
        
        self.isBarcodeRunning = False
        
    def getBarcodeRunning(self) :
        return self.isBarcodeRunning
    
    def setBtnStartEnable(self, status) :
        self.bStart["state"] = status

    def setBarcodeRunning(self, status) :
        self.isBarcodeRunning = status

    def insertMac(self, mac):
        listScanMac.append(mac)

    def clearListMac(self):
        listScanMac.clear()

    def getListSize(self):
        return len(listScanMac)

    def scan_available_ports(self):
        """ Lists serial port names

            :raises EnvironmentError:
                On unsupported or unknown platforms
            :returns:
                A list of the serial ports available on the system
        """
        ports = None
        
        if sys.platform.startswith('win'):
            ports = ['COM%s' % (i + 1) for i in range(256)]
        elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
            # this excludes your current terminal "/dev/tty"
            ports = glob.glob('/dev/tty[A-Za-z]*')
        elif sys.platform.startswith('darwin'):
            ports = glob.glob('/dev/tty.*')
        else:
            raise EnvironmentError('Unsupported platform')

        result = []
        for port in ports:
            try:
                s = Serial(port)
                s.close()
                result.append(port)
            except (OSError, SerialException):
                pass

        return result

    def terminate_all_process(self):
        if ( hasattr(self, "current_process") and type(self.current_process) is subprocess.Popen):
            try:
                self.current_process.terminate()
                self.current_process.wait()
            except OSError as e:
                print(e)

    def closeWin(self):
        self.terminate_all_process()
        os._exit(1)

    def get_path(self):
        # determine if application is a script file or frozen exe
        if getattr(sys, "frozen", False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        return application_path

    def read_file(self, filepath):
        text = None
        with open(filepath, "rb") as f:
            text = bytes(f.read())
        return text

    def get_directory_path(self):

        self.directory_path = filedialog.askdirectory(initialdir=self.get_path())
        #self.OutputText.insert(tk.END,self.directory_path)

        #self.dir_path = filedialog.askdirectory(initialdir=self.get_path())
        print("Current dir : {}".format(self.directory_path))

        if self.directory_path:
            # (filepath, tempfilename) = os.path.split(self.directory_path)
            # (shotname, extension) = os.path.splitext(tempfilename)
            self.lblDirectoryPath["text"] = self.directory_path           
        else:
            pass
    
    def write_to_textbox(self, message, tag):
        self.OutputText.insert(tk.END,  message + "\n", tag)
    
    def clearBox(self):
        self.OutputText.delete("1.0", "end")

    def enable_uart_component(self, state):
        if state == True:
            self.comboPort["state"] = "readonly"
            self.comboBaud["state"] = "readonly"
            self.comboData["state"] = "readonly"
            self.comboParity["state"]  = "readonly"
            self.comboStop["state"] = "readonly"
        else:
            self.comboPort["state"] = "disable"
            self.comboBaud["state"] = "disable"
            self.comboData["state"] = "disable"
            self.comboParity["state"]  = "disable"
            self.comboStop["state"] = "disable"


    def open_com_event1(self):
        if self.uart1_open == False:
            try:
                # print("port: {}".format(self.serial_property["port"].get()))
                # print("baud: {}".format(self.serial_property["baud"].get()))
                # print("data: {}".format(self.serial_property["data"].get()))
                # print("parity: {}".format(self.serial_property["parity"].get()))
                # print("stop: {}".format(self.serial_property["stop"].get()))

                self.ser1.port = self.serial_property["port1"].get()
                self.ser1.baudrate = 115200
                self.ser1.bytesize = EIGHTBITS
                self.ser1.parity = PARITY_NONE
                self.ser1.stopbits = STOPBITS_ONE

                self.ser1.open()

            except Exception as err:
                message_string = "Error while open serial : {}".format(err)
                tk.messagebox.showerror(title="Error", message=message_string)

            if self.ser1.isOpen():
                self.btnOpenCom1["text"] = "Close"
                self.comboPort1["state"] = "disable"
                self.uart1_open = True
                self.btnOpenCom1["bg"] = "#00FF00"
                # disable component
                #self.enable_uart_component(False)
        else:
            try:
                self.ser1.close()
            except Exception as err:
                print("Failed to close serial port : {}".format(err))
            
            if self.ser1.isOpen() == False:
                self.btnOpenCom1["text"] = "Open"
                self.comboPort1["state"] = "readonly"
                self.uart1_open = False
                self.btnOpenCom1["bg"] = "#6495ED"
                #self.enable_uart_component(True)

    def open_com_event2(self):
        if self.uart2_open == False:
            try:
                # print("port: {}".format(self.serial_property["port"].get()))
                # print("baud: {}".format(self.serial_property["baud"].get()))
                # print("data: {}".format(self.serial_property["data"].get()))
                # print("parity: {}".format(self.serial_property["parity"].get()))
                # print("stop: {}".format(self.serial_property["stop"].get()))

                self.ser2.port = self.serial_property["port2"].get()
                self.ser2.baudrate = 115200
                self.ser2.bytesize = EIGHTBITS
                self.ser2.parity = PARITY_NONE
                self.ser2.stopbits = STOPBITS_ONE

                self.ser2.open()

            except Exception as err:
                message_string = "Error while open serial : {}".format(err)
                tk.messagebox.showerror(title="Error", message=message_string)

            if self.ser2.isOpen():
                self.btnOpenCom2["text"] = "Close"
                self.comboPort2["state"] = "disable"
                self.uart2_open = True
                self.btnOpenCom2["bg"] = "#00FF00"
                # disable component
                #self.enable_uart_component(False)
        else:
            try:
                self.ser2.close()
            except Exception as err:
                print("Failed to close serial port : {}".format(err))
            
            if self.ser2.isOpen() == False:
                self.btnOpenCom2["text"] = "Open"
                self.comboPort2["state"] = "readonly"
                self.uart2_open = False
                self.btnOpenCom2["bg"] = "#6495ED"
                #self.enable_uart_component(True)

    def open_com_event3(self):
        if self.uart3_open == False:
            try:
                # print("port: {}".format(self.serial_property["port"].get()))
                # print("baud: {}".format(self.serial_property["baud"].get()))
                # print("data: {}".format(self.serial_property["data"].get()))
                # print("parity: {}".format(self.serial_property["parity"].get()))
                # print("stop: {}".format(self.serial_property["stop"].get()))

                self.ser3.port = self.serial_property["port3"].get()
                self.ser3.baudrate = 115200
                self.ser3.bytesize = EIGHTBITS
                self.ser3.parity = PARITY_NONE
                self.ser3.stopbits = STOPBITS_ONE

                self.ser3.open()

            except Exception as err:
                message_string = "Error while open serial : {}".format(err)
                tk.messagebox.showerror(title="Error", message=message_string)

            if self.ser3.isOpen():
                self.btnOpenCom3["text"] = "Close"
                self.comboPort3["state"] = "disable"
                self.uart3_open = True
                self.btnOpenCom3["bg"] = "#00FF00"
                # disable component
                #self.enable_uart_component(False)
        else:
            try:
                self.ser3.close()
            except Exception as err:
                print("Failed to close serial port : {}".format(err))
            
            if self.ser3.isOpen() == False:
                self.btnOpenCom3["text"] = "Open"
                self.comboPort3["state"] = "readonly"
                self.uart3_open = False
                self.btnOpenCom3["bg"] = "#6495ED"
                # self.enable_uart_component(True)

    def event_start(self) :
        listSerialSN = []
        listSerialMAC = []
        readSN = None
        readMac = None 

        if not self.ser1.isOpen():
            message_string = "Open com Port 1 first !!"
            tk.messagebox.showerror(title="Error", message=message_string)
            return

        if not self.ser2.isOpen():
            message_string = "Open com Port 2 first !!"
            tk.messagebox.showerror(title="Error", message=message_string)
            return
        
        if not self.ser3.isOpen():
            message_string = "Open com Port 3 first !!"
            tk.messagebox.showerror(title="Error", message=message_string)
            return
        
        self.bStart["text"] = "Process"
        self.setBtnStartEnable("disable")

        # === SERIAL 1 ====
        # get serial number
        message = CMD_GET_SN.encode(encoding='ascii')
        self.ser1.write(message)
        time.sleep(1)
        read_data_sn = self.ser1.read_all().decode(encoding='ascii')
        print("SN 1 = " + read_data_sn)

        # get mac address
        message = CMD_GET_MAC.encode(encoding='ascii')
        self.ser1.write(message)
        time.sleep(1)
        read_data_mac = self.ser1.read_all().decode(encoding='ascii')

        string_split_sn = read_data_sn.splitlines()
        string_split_mac = read_data_mac.splitlines()

        idx = 0
        for i in range (len(string_split_sn)) :
            print ("test : " , i , " - ", string_split_sn[i])
            if (string_split_sn[i] == CMD_GET_SN.replace('\r','')) :
                print(" bener : ", string_split_sn[i])
                idx = i+2
                break

        try:
            # filter data serial number from garbage character
            if re.match("[A-Z0-9]+$", string_split_sn[idx]):
                #self.data_query['sn'] = string_split_sn[2]
                readSN = string_split_sn[idx]
            else:
                readSN = "None"
        except Exception as identifier:
            readSN = "None"
        
        print("SN 1 : " + readSN)
        listSerialSN.append(readSN)

        try:
            # filter data mac address from garbage character
            if re.match("[0-9a-f]{2}([-:]?)[0-9a-f]{2}(\\1[0-9a-f]{2}){4}$", string_split_mac[4].lower()):
                #self.data_query['mac'] = string_split_mac[4]
                readMac = string_split_mac[4].upper()
            else:
                readMac = "None"
        except Exception as identifier:
            readMac = "None"
        
        print("MAC 1 : " + readMac)
        listSerialMAC.append(readMac)
        
        # === SERIAL 2 ====
        # get serial number
        message = CMD_GET_SN.encode(encoding='ascii')
        self.ser2.write(message)
        time.sleep(1)
        read_data_sn = self.ser2.read_all().decode(encoding='ascii')
        print("SN 2 = " + read_data_sn)

        # get mac address
        message = CMD_GET_MAC.encode(encoding='ascii')
        self.ser2.write(message)
        time.sleep(1)
        read_data_mac = self.ser2.read_all().decode(encoding='ascii')

        string_split_sn = read_data_sn.splitlines()
        string_split_mac = read_data_mac.splitlines()

        idx = 0
        for i in range (len(string_split_sn)) :
            print ("test : " , i , " - ", string_split_sn[i])
            if (string_split_sn[i] == CMD_GET_SN.replace('\r','')) :
                print(" bener : ", string_split_sn[i])
                idx = i+2
                break

        try:
            # filter data serial number from garbage character
            if re.match("[A-Z0-9]+$", string_split_sn[idx]):
                #self.data_query['sn'] = string_split_sn[2]
                readSN = string_split_sn[idx]
            else:
                readSN = "None"
        except Exception as identifier:
            readSN = "None"
        
        print("SN 2 : " + readSN)
        listSerialSN.append(readSN)

        try:
            # filter data mac address from garbage character
            if re.match("[0-9a-f]{2}([-:]?)[0-9a-f]{2}(\\1[0-9a-f]{2}){4}$", string_split_mac[4].lower()):
                #self.data_query['mac'] = string_split_mac[4]
                readMac = string_split_mac[4].upper()
            else:
                readMac = "None"
        except Exception as identifier:
            readMac = "None"
        
        print("MAC 2 : " + readMac)
        listSerialMAC.append(readMac)

        # === SERIAL 3 ====
        # get serial number
        message = CMD_GET_SN.encode(encoding='ascii')
        self.ser3.write(message)
        time.sleep(1)
        read_data_sn = self.ser3.read_all().decode(encoding='ascii')
        print("SN 3 = " + read_data_sn)

        # get mac address
        message = CMD_GET_MAC.encode(encoding='ascii')
        self.ser3.write(message)
        time.sleep(1)
        read_data_mac = self.ser3.read_all().decode(encoding='ascii')

        string_split_sn = read_data_sn.splitlines()
        string_split_mac = read_data_mac.splitlines()

        idx = 0
        for i in range (len(string_split_sn)) :
            print ("test : " , i , " - ", string_split_sn[i])
            if (string_split_sn[i] == CMD_GET_SN.replace('\r','')) :
                print(" bener : ", string_split_sn[i])
                idx = i+2
                break

        try:
            # filter data serial number from garbage character
            if re.match("[A-Z0-9]+$", string_split_sn[idx]):
                #self.data_query['sn'] = string_split_sn[2]
                readSN = string_split_sn[idx]
            else:
                readSN = "None"
        except Exception as identifier:
            readSN = "None"
        
        print("SN 3 : " + readSN)
        listSerialSN.append(readSN)

        try:
            # filter data mac address from garbage character
            if re.match("[0-9a-f]{2}([-:]?)[0-9a-f]{2}(\\1[0-9a-f]{2}){4}$", string_split_mac[4].lower()):
                #self.data_query['mac'] = string_split_mac[4]
                readMac = string_split_mac[4].upper()
            else:
                readMac = "None"
        except Exception as identifier:
            readMac = "None"
        
        print("MAC 3 : " + readMac)
        listSerialMAC.append(readMac)

        #check read listScanMac & read Serial
        
        i=0
        for x in listScanMac:
            j=0
            isFound = False
            for y in listSerialMAC:
                x = x.replace('\n','')
                y = y.replace(':','')
                if (x == y) :
                    isFound = True
                    # print(listSerialSN[j])
                    if(listSerialSN[j] != "None") : 
                        # print(self.sendDataToServer(listSerialSN[j], listSerialMAC[j]))
                        respon = self.sendDataToServer(listSerialSN[j], listSerialMAC[j])
                        data = json.loads(respon)
                        status = data['status']
                        msg = data['msg']
                        self.write_to_textbox(self.insertString(listScanMac[i]) + " - " +msg , status)
                        self.saveLog(listSerialSN[j], listSerialMAC[j], msg, status)
                    else :
                        self.write_to_textbox(self.insertString(listScanMac[i]) , "error")  
                        self.saveLog(listSerialSN[j], listSerialMAC[j], "msg", "status")                
                j = j+1
            if(isFound == False) :
                self.write_to_textbox(self.insertString(listScanMac[i]) , "error")
                self.saveLog("sn", listScanMac[i], "msg", "status") 
            i = i+1
        
        self.setBarcodeRunning(False)
        self.clearListMac()
        self.bStart["text"] = "Start"

    def saveLog(self, SN, MAC, msg, status) :
        out_file = open(LOG_PATH, 'a')
        now = datetime.now()
        current = now.strftime("%d/%m/%Y %H:%M:%S")
        out_file.writelines(current + "\n")
    
    def sendDataToServer(self, SN, MAC):
        PARAMS={'username':USERNAME, 'pass':PASSWORD, 'sn':SN, 'mac':MAC}
        # sending post request and saving response as response object 
        x = requests.post(URL_LINK, data = PARAMS, timeout=30)
        return x.text

    def event_start1(self):
        message_string = ""

        if not self.directory_path:
            # message_string = "Load directory path first!"
            # tk.messagebox.showerror(title="Error", message=message_string)
            # use directory data
            if not os.path.isdir("data"):
                os.mkdir("data")
            self.output_file = "data" + "/" + self.default_filename + ".txt"
        else :
            print("self directory : {}".format(self.directory_path))
            self.output_file = self.directory_path + "/" + self.default_filename + ".txt"

        # if self.eFIlename.get() == "Masukkan nama file.." or self.eFIlename.get() == '':
        #     message_string = "Set filename first!"
        #     tk.messagebox.showerror(title="Error", message=message_string)
        #     return

        # check serial com
        if not self.ser.isOpen():
            message_string = "Open com first !!"
            tk.messagebox.showerror(title="Error", message=message_string)
            return

        # get serial number
        message = CMD_GET_SN.encode(encoding='ascii')
        self.ser.write(message)
        time.sleep(.2)
        read_data_sn = self.ser.read_all().decode(encoding='ascii')

        # get mac address
        message = CMD_GET_MAC.encode(encoding='ascii')
        self.ser.write(message)
        time.sleep(.2)
        read_data_mac = self.ser.read_all().decode(encoding='ascii')

        # split data
        string_split_sn = read_data_sn.splitlines()
        string_split_mac = read_data_mac.splitlines()
        print(string_split_sn)
        print(string_split_mac)

        print("Serial number : {}".format(string_split_sn[2]))

        # filter data serial number from garbage character
        if re.match("[A-Z0-9]+$", string_split_sn[2]):
            self.data_query['sn'] = string_split_sn[2]
            tk.messagebox.showinfo(
                title="Status",
                message="SN : {}".format(self.data_query['sn'])
            )
        else:
            print("serial not valid")
            tk.messagebox.showerror(
                title="Error",
                message="SN data not valid"
            )
            return

        # filter data mac address from garbage character
        if re.match("[0-9a-f]{2}([-:]?)[0-9a-f]{2}(\\1[0-9a-f]{2}){4}$", string_split_mac[4].lower()):
            self.data_query['mac'] = string_split_mac[4]
            tk.messagebox.showinfo(
                title="Status",
                message="MAC : {}".format(self.data_query['mac'])
            )
        else:
            print("MAC Address not valid")
            tk.messagebox.showerror(
                title="Error",
                message="MAC data not valid"
            )
            return
        
        self.write_to_textbox(self.data_query['sn'] + " " +  self.data_query['mac'] + "\n")
        
        # try to create file
        try:
            out_file = open(self.output_file, 'a')
            out_file.writelines(self.data_query['sn'] + " " + self.data_query['mac'] + "\n")
            out_file.close()
        except IOError as err:
            print("Err : {}".format(err))
        

    def createWidgets(self, main_frame):
        ftLabel = font.Font(family="Lucida Grande", size=12, weight=font.BOLD)
        ftText = font.Font(family="Lucida Grande", size=8, weight=font.NORMAL)
        ftButton = font.Font(family="Lucida Grande", size=10, weight=font.BOLD)
        smallLabel = font.Font(family="Lucida Grande", size=8)

        self.row_count = 0

        self.menubar = Menu(main_frame, tearoff=False)
        self.frame1 = Frame(main_frame)
        self.frame1.pack(fill=None, expand=False)
        self.frame1.grid(row=0, column=0, columnspan=2, sticky=W + E + N + S, padx=10, pady=10)

        self.bBrowseFile = Button(self.frame1, text="Browse dir", width=12, font=ftButton, bg="#6495ED")
        self.bBrowseFile.grid(row=0, column=0, sticky=W + E + N + S)
        self.bBrowseFile["command"] = self.get_directory_path
        self.bBrowseFile["state"] = "disable"

        # dir
        self.lblDirectoryPath = Label(self.frame1, text="data", width=30, font=smallLabel, anchor="w")
        self.lblDirectoryPath.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)

        self.row_count = 1
        
        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)
        self.row_count = self.row_count + 1

        # filename
        self.lblFileName = Label(self.frame1, text="Filename", width=30, font = smallLabel, anchor='w')
        self.lblFileName.grid(row=self.row_count, column=0, sticky=W+E, columnspan=4, pady=3)

        
        self.initial_string_value_filename = self.default_filename
        self.eFIlename = StringVar(value='')
        self.eFilenameText = Entry(self.frame1, width=18, font = smallLabel, textvariable = self.eFIlename)
        self.eFilenameText.grid(row=self.row_count, column=1, sticky=W+E+N+S, pady=3, columnspan=4)
        self.eFilenameText.insert(0, self.initial_string_value_filename)
        self.eFilenameText.config(fg = 'grey')
        self.eFilenameText["state"] = "disabled"
        
        def on_entry_click_filename(event):
            if self.eFilenameText.get() == self.initial_string_value_filename:
                self.eFilenameText.delete(0, "end") # delete all the text in the entry
                self.eFilenameText.insert(0, '') #Insert blank for user input
                self.eFilenameText.config(fg = 'black')

        def on_focusout_filename(event):
            if self.eFilenameText.get() == '':
                self.eFilenameText.insert(0, self.initial_string_value_filename)
                self.eFilenameText.config(fg = 'grey')

        self.eFilenameText.bind('<FocusIn>', on_entry_click_filename)
        self.eFilenameText.bind('<FocusOut>', on_focusout_filename)

        self.row_count = self.row_count + 1

        # Port/com 1
        self.lblPort1 = Label(self.frame1, text="Port 1", width=30, font=smallLabel, anchor="w")
        self.lblPort1.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        # scan available port
        ports1 =  self.scan_available_ports()
        #print("len : {}".format(len(ports)))
        initial_port = None
        if len(ports1) > 0:
            initial_port = ports1[0]
            
        self.serial_property["port1"] = tk.StringVar(value=initial_port)
        self.comboPort1 = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["port1"])
        self.comboPort1["values"] = self.scan_available_ports() #("com1", "com2", "com3", "com4", "com5")
        self.comboPort1["state"] = "readonly"
        self.comboPort1.grid(row=self.row_count, column=1, sticky=W + E, columnspan=1, pady=3)

        self.btnOpenCom1 = Button(self.frame1, text="Open", width=12, font=ftButton, bg="#6495ED")
        self.btnOpenCom1.grid(row=self.row_count, column=3, sticky=W + E + N + S, columnspan=1)
        self.btnOpenCom1["command"] = self.open_com_event1
        self.row_count = self.row_count + 1

        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)

        #self.row_count = self.row_count + 1

        # Port/com 2
        self.lblPort2 = Label(self.frame1, text="Port 2", width=30, font=smallLabel, anchor="w")
        self.lblPort2.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        # scan available port
        ports2 =  self.scan_available_ports()
        #print("len : {}".format(len(ports)))
        initial_port = None
        if len(ports2) > 0:
            initial_port = ports2[0]
            
        self.serial_property["port2"] = tk.StringVar(value=initial_port)
        self.comboPort2 = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["port2"])
        self.comboPort2["values"] = self.scan_available_ports() #("com1", "com2", "com3", "com4", "com5")
        self.comboPort2["state"] = "readonly"
        self.comboPort2.grid(row=self.row_count, column=1, sticky=W + E, columnspan=1, pady=3)

        self.btnOpenCom2 = Button(self.frame1, text="Open", width=12, font=ftButton, bg="#6495ED")
        self.btnOpenCom2.grid(row=self.row_count, column=3, sticky=W + E + N + S, columnspan=1)
        self.btnOpenCom2["command"] = self.open_com_event2
        self.row_count = self.row_count + 1

        # Port/com 3
        self.lblPort3 = Label(self.frame1, text="Port 3", width=30, font=smallLabel, anchor="w")
        self.lblPort3.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        # scan available port
        ports3 =  self.scan_available_ports()
        #print("len : {}".format(len(ports)))
        initial_port = None
        if len(ports3) > 0:
            initial_port = ports3[0]
            
        self.serial_property["port3"] = tk.StringVar(value=initial_port)
        self.comboPort3 = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["port3"])
        self.comboPort3["values"] = self.scan_available_ports() #("com1", "com2", "com3", "com4", "com5")
        self.comboPort3["state"] = "readonly"
        self.comboPort3.grid(row=self.row_count, column=1, sticky=W + E, columnspan=1, pady=3)

        self.btnOpenCom3 = Button(self.frame1, text="Open", width=12, font=ftButton, bg="#6495ED")
        self.btnOpenCom3.grid(row=self.row_count, column=3, sticky=W + E + N + S, columnspan=1)
        self.btnOpenCom3["command"] = self.open_com_event3
        self.row_count = self.row_count + 1

        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)

        self.row_count = self.row_count + 1

        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)
        self.row_count = self.row_count + 1

        # frame for received data
        frameRecv = tk.Frame(self.frame1)
        frameRecv.grid(row=self.row_count, column=0, columnspan=4)

        labelOutText = tk.Label(frameRecv,text="Received Data:")
        labelOutText.grid(row = 1, column = 1, padx = 3, pady = 2, sticky = tk.W)
        frameRecvSon = tk.Frame(frameRecv)
        frameRecvSon.grid(row = 2, column =1)
        scrollbarRecv = tk.Scrollbar(frameRecvSon)
        scrollbarRecv.pack(side = tk.RIGHT, fill = tk.Y)
        self.OutputText = tk.Text(frameRecvSon, wrap = tk.WORD, width = 42, height = 10, yscrollcommand = scrollbarRecv.set)
        self.OutputText.pack()
        self.OutputText.tag_config('error', background="yellow", foreground="red")
        self.OutputText.tag_config('success', foreground="green")

        self.row_count = self.row_count + 1

        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)
        self.row_count = self.row_count + 1

        self.bStart = Button(self.frame1, text="Start", width=12, font=ftButton, bg="#6495ED")
        self.bStart.grid(row=self.row_count, column=0, sticky=W + E + N + S, columnspan=4)
        self.bStart["command"] = self.event_start
        self.bStart["state"] = "disable"

    def insertString(self, string):
        t = iter(string)
        temp = ':'.join(a+b for a,b in zip(t, t))
        return temp
       
if __name__ == "__main__":

    print("platform : [{}]".format(platform.platform()))
    print("platform version : [{}]".format(platform.version()))
    print("platform architecture : [{}]".format(platform.architecture()))
    print("platform machine : [{}]".format(platform.machine()))
    print("platform node : [{}]".format(platform.node()))
    print("platform processor : [{}]".format(platform.processor()))
    print("platform system : [{}]".format(platform.system()))
    # print('uname : [{}]'.format(platform.uname()))

    # print current os
    cur_os = platform.system()
    print("Current os : {}".format(cur_os))

    root = Tk()
    root.title("Serial Number Query V%s" % (__version__))
    # root.geometry("320x200+0+0")
    try:
        root.iconbitmap("app.ico")
    except Exception as e:
        print(e)
    # lock the root size
    root.resizable(False, False)
    root.lift()

    app = Application(master=root)

    if cur_os == "Windows":
        # to minimize the console windows
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)

    #app.mainloop()
    scans = Scanner()
    try:
        scans.findDevice(VENDOR_ID, PRODUCT_ID)
    except Exception as identifier:
        tk.messagebox.showerror(title="Barcode Scanner Not Found", message=identifier)
        sys.exit(1)
        
    while True:
        app.update()
        if(app.getBarcodeRunning() == False) :
            value = scans.startScan()
            if (value != None) :
                if(app.getListSize() == 0) :
                    app.clearBox()
                if (app.getListSize() < 3) :
                    app.insertMac(value)
                    app.write_to_textbox("Scan : " + app.insertString(value), "")
                if (app.getListSize() == 3) :
                    app.setBtnStartEnable("normal")
                    app.setBarcodeRunning(True)
                

