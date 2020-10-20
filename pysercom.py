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
from serial import Serial, SerialException
from serial import PARITY_EVEN, PARITY_MARK, PARITY_NAMES, PARITY_NONE, PARITY_ODD, PARITY_SPACE
from serial import STOPBITS_ONE, STOPBITS_ONE_POINT_FIVE , STOPBITS_TWO
from serial import FIVEBITS, SIXBITS, SEVENBITS, EIGHTBITS
import glob
from datetime import datetime
import time
import requests
import json

__author__ = "Cosmas Eric. s"
__copyright__ = "Copyright 2020, Serial communication project"

__license__ = "GPL"
__version__ = "1.1.0"
__maintainer__ = "Cosmas Eric "
__email__ = "cosmas.eric.septian@polytron.co.id"
__status__ = "Internal Test Beta"


CMD_GET_SN = "getprop ro.serialno\r"
#CMD_GET_MAC = "ip addr show wlan0  | grep 'link/ether '| cut -d' ' -f6\r"
CMD_GET_MAC = "ip addr show | grep 'link/ether' | cut -d' ' -f6\r"
LOG_PATH = "./log.txt"

URL_LINK = "http://10.8.42.44/mola/scan/readHKC"

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack(fill=BOTH, expand=1)
        main_frame = Frame(master)
        main_frame.pack(fill="y", expand=1)

        self.serial_property = {
            "port" : "",
            "baud" : 9600,
            "data" : 8,
            "parity" : "None",
            "stop" : 1,
        }

        self.default_filename = "hkc_sn_mac"

        self.uart_open = False

        self.createWidgets(main_frame)
        self.dl_thread = 0
        self.directory_path = None

        master.protocol("WM_DELETE_WINDOW", self.closeWin)

        # threading.Thread(target=self.check_files).start()

        # create serial
        self.ser = Serial()

        self.rx_count = 0
        self.output_file = None

        # data to be capture
        self.data_query = {
            "sn" : "",
            "mac" : "",
        }
        

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
        self.OutputText.insert(tk.END, message, tag)

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


    def open_com_event(self):
        if self.uart_open == False:
            try:
                # print("port: {}".format(self.serial_property["port"].get()))
                # print("baud: {}".format(self.serial_property["baud"].get()))
                # print("data: {}".format(self.serial_property["data"].get()))
                # print("parity: {}".format(self.serial_property["parity"].get()))
                # print("stop: {}".format(self.serial_property["stop"].get()))

                self.ser.port = self.serial_property["port"].get()
                self.ser.baudrate = self.serial_property["baud"].get()
                
                # serial data size
                if self.serial_property["data"] == "5":
                    self.ser.bytesize = FIVEBITS
                elif self.serial_property["data"] == "6":
                    self.ser.bytesize = SIXBITS
                elif self.serial_property["data"] == "7":
                    self.ser.bytesize = SEVENBITS
                elif self.serial_property["data"] == "8":
                    self.ser.bytesize = EIGHTBITS

                # set parity
                if self.serial_property["parity"] == "None":
                    self.ser.parity = PARITY_NONE
                elif self.serial_property["parity"] == "Odd":
                    self.ser.parity = PARITY_ODD
                elif self.serial_property["parity"] == "Even":
                    self.ser.parity = PARITY_EVEN
                elif self.serial_property["parity"] == "Mark":
                    self.ser.parity = PARITY_MARK
                elif self.serial_property["parity"] == "Space":
                    self.ser.parity = PARITY_SPACE

                # set stop bit
                if self.serial_property["stop"] == "1":
                    self.ser.stopbits = STOPBITS_ONE
                elif self.serial_property["stop"] == "1.5":
                    self.ser.stopbits = STOPBITS_ONE_POINT_FIVE
                elif self.serial_property["stop"] == "2":
                    self.ser.stopbits = STOPBITS_TWO

                self.ser.open()

            except Exception as err:
                message_string = "Error while open serial : {}".format(err)
                tk.messagebox.showerror(title="Error", message=message_string)

            if self.ser.isOpen():
                self.btnOpenCom["text"] = "Close"
                self.uart_open = True
                self.lblComStatusVal["text"] = "Open : {} {} {} {} {}".format(
                    self.ser.port,
                    self.ser.baudrate,
                    self.ser.bytesize,
                    self.ser.parity,
                    self.ser.stopbits
                )
                # disable component
                self.enable_uart_component(False)
        else:
            try:
                self.ser.close()
            except Exception as err:
                print("Failed to close serial port : {}".format(err))
            
            if self.ser.isOpen() == False:
                self.btnOpenCom["text"] = "Open"
                self.uart_open = False
                self.lblComStatusVal["text"] = "Closed"
                self.enable_uart_component(True)

    def event_start(self):
        message_string = ""
        string_split_sn = ""
        string_split_mac = ""
        readSN = "None"
        readMac = "None"

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
        
        try:
            # get serial number
            message = CMD_GET_SN.encode(encoding='ascii')
            self.ser.write(message)
            time.sleep(1)
            read_data_sn = self.ser.read_all().decode(encoding='ascii')
            string_split_sn = read_data_sn.splitlines()
            print("SN 1 = " + string_split_sn)
        except Exception as identifier:
            readSN = "None"
        
        try:
            # get mac address
            message = CMD_GET_MAC.encode(encoding='ascii')
            self.ser.write(message)
            time.sleep(1)
            read_data_mac = self.ser.read_all().decode(encoding='ascii')
            string_split_mac = read_data_mac.splitlines()
        except Exception as identifier:
            readMac = "None"

        idx = 0
        try :
            for i in range (len(string_split_sn)) :
                print ("test : " , i , " - ", string_split_sn[i])
                if (string_split_sn[i] == CMD_GET_SN.replace('\r','')) :
                    print(" bener : ", string_split_sn[i])
                    idx = i+2
                    break
        except Exception as identifier:
            readSN = "None"    

        try:
            # filter data serial number from garbage character
            if re.match("[A-Z0-9]+$", string_split_sn[idx]):
                #self.data_query['sn'] = string_split_sn[2]
                readSN = string_split_sn[idx]
            else:
                readSN = "None"
        except Exception as identifier:
            readSN = "None"
        
        print("SN nya : " + readSN)

        try:
            # filter data mac address from garbage character
            if re.match("[0-9a-f]{2}([-:]?)[0-9a-f]{2}(\\1[0-9a-f]{2}){4}$", string_split_mac[4].lower()):
                #self.data_query['mac'] = string_split_mac[4]
                readMac = string_split_mac[4].upper()
            else:
                readMac = "None"
        except Exception as identifier:
            readMac = "None"
        
        print("MAC nya : " + readMac)
        
        if(readSN == "None" or readMac == "None") :
            self.event_start()
            return
        
        readSN = "56FFEDCD21"
        readMac = "DC:BD:7A:62:5F:25"
        respon = self.sendDataToServer(readSN, readMac)
        data = json.loads(respon)
        status = data['status']
        sn = data['sn']
        mac = data['mac']
        self.write_to_textbox(readSN + " - "+readMac + " = "+status, status)
        self.saveLog(readSN, readMac, status)
        

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

        # Port/com
        self.lblPort = Label(self.frame1, text="Port", width=30, font=smallLabel, anchor="w")
        self.lblPort.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        # scan available port
        ports =  self.scan_available_ports()
        #print("len : {}".format(len(ports)))
        initial_port = None
        if len(ports) > 0:
            initial_port = ports[0]
            
        self.serial_property["port"] = tk.StringVar(value=initial_port)
        self.comboPort = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["port"])
        self.comboPort["values"] = self.scan_available_ports() #("com1", "com2", "com3", "com4", "com5")
        self.comboPort["state"] = "readonly"
        self.comboPort.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)

        self.row_count = self.row_count + 1

        # Baudrate
        self.serial_property["baud"] = tk.StringVar(value="115200")
        self.lblBaud = Label(self.frame1, text="Baudrate", width=30, font=smallLabel, anchor="w")
        self.lblBaud.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        self.comboBaud = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["baud"])
        self.comboBaud["values"] = ("9600", "19200", "38400", "57600", "115200", "230400")
        self.comboBaud["state"] = "readonly"
        self.comboBaud.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)

        self.row_count = self.row_count + 1

        # Data bits
        self.serial_property["data"] = tk.StringVar(value="8")
        self.lblDataBits = Label(self.frame1, text="Data Bits", width=30, font=smallLabel, anchor="w")
        self.lblDataBits.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        self.comboData = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["data"])
        self.comboData["values"] = ("5", "6", "7", "8")
        self.comboData["state"] = "readonly"
        self.comboData.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)

        self.row_count = self.row_count + 1

        # stop
        self.serial_property["stop"] = tk.StringVar(value="1")
        self.lblStopBit = Label(self.frame1, text="Stop Bits", width=30, font=smallLabel, anchor="w")
        self.lblStopBit.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        self.comboStop = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["stop"])
        self.comboStop["values"] = ("1", "1.5", "2")
        self.comboStop["state"] = "readonly"
        self.comboStop.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)

        self.row_count = self.row_count + 1

        # Parity
        self.serial_property["parity"] = tk.StringVar(value="None")
        self.lblParity = Label(self.frame1, text="Parity", width=30, font=smallLabel, anchor="w")
        self.lblParity.grid(row=self.row_count, column=0, sticky=W + E, columnspan=4, pady=3)

        self.comboParity = ttk.Combobox(self.frame1, width=17, textvariable=self.serial_property["parity"])
        self.comboParity["values"] = ("None", "Even", "Odd", "Mark", "Space")
        self.comboParity["state"] = "readonly"
        self.comboParity.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)

        self.row_count = self.row_count + 1

        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)

        self.row_count = self.row_count + 1

        self.btnOpenCom = Button(self.frame1, text="Open", width=12, font=ftButton, bg="#6495ED")
        self.btnOpenCom.grid(row=self.row_count, column=0, sticky=W + E + N + S, columnspan=1)
        self.btnOpenCom["command"] = self.open_com_event

        self.lblComStatusVal = Label(self.frame1, text="Closed", width=30, font=smallLabel, anchor="w")
        self.lblComStatusVal.grid(row=self.row_count, column=1, sticky=W + E, columnspan=4, pady=3)
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
        self.OutputText = tk.Text(frameRecvSon, wrap = tk.WORD, width = 42, height = 10, yscrollcommand = scrollbarRecv.set, font=("Helvetica", 18))
        self.OutputText.pack()
        self.OutputText.tag_config('NOK', background="yellow", foreground="red")
        self.OutputText.tag_config('OK', foreground="green")

        self.row_count = self.row_count + 1

        # space -----
        self.space = Label(self.frame1, text="", width=30, font=smallLabel, anchor="w")
        self.space.grid(row=self.row_count, column=0, sticky=W + E, columnspan=2)
        self.row_count = self.row_count + 1

        self.bStart = Button(self.frame1, text="Start", width=12, font=ftButton, bg="#6495ED")
        self.bStart.grid(row=self.row_count, column=0, sticky=W + E + N + S, columnspan=4)
        self.bStart["command"] = self.event_start

    def sendDataToServer(self, SN, MAC):
        PARAMS={'sn':SN, 'mac':MAC.lower()}
        # sending post request and saving response as response object 
        x = requests.post(URL_LINK, data = PARAMS, timeout=30)
        return x.text

    def saveLog(self, SN, MAC, status) :
        out_file = open(LOG_PATH, 'a')
        now = datetime.now()
        current = now.strftime("%d/%m/%Y %H:%M:%S")
        current = current + " "+ SN + " "+MAC + " "+status
        out_file.writelines(current + "\n")
       
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

    app.mainloop()
