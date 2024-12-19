#!/usr/bin/python3
# -*- coding: utf-8 -*-

# автор: Титовский С.А.
# 18/12/2024

import wx
from wx.adv import TaskBarIcon as TaskBarIcon
from wx.adv import SplashScreen as SplashScreen
import sys
import requests
import re
import time
import string # for buttons
import datetime  # импорт библиотеки дат и времени
from datetime import datetime, timedelta
import os  # импорт библиотеки для работы с операционной системой
import socket
import sqlite3
import openpyxl # импорт библиотек для работы Excel
from openpyxl import Workbook # импорт библиотек для работы Excel
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils.cell import get_column_letter
import locale # для работы с локалью
import threading
from threading import Thread # для работы с потоками
from pubsub import pub # для работы с подписчиками
import ftplib # для работы с ftp
import queue

#=============================================
#=============================================
#=============================================
#=============================================
def except_foo_dec(foo):
    def wrapper(*args, **kwargs):
        #ToLog(f"{foo.__name__}, *args = {args}, **kwargs = {kwargs} started")
        try:
            return foo(*args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {foo.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            #raise Exception
        else:
            ToLog(f"{foo.__name__}, *args = {args}, **kwargs = {kwargs} finished successfully")
    return wrapper

def except_foo_brief(foo):
    def wrapper(*args, **kwargs):
        #ToLog(f"{foo.__name__} started")
        try:
            return foo(*args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {foo.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            #raise Exception
        else:
            ToLog(f"{foo.__name__} finished successfully")
    return wrapper

def info_foo_dec(foo):
    def wrapper(*args, **kwargs):
        #ToLog(f"{foo.__name__}, *args = {args}, **kwargs = {kwargs} started")
        return foo(*args, **kwargs)
        ToLog(f"{foo.__name__}, *args = {args}, **kwargs = {kwargs} finished")
    return wrapper

def except_method_dec(method):
    def wrapper(self, *args, **kwargs):
        #ToLog(f"{method.__name__}, *args = {args}, **kwargs = {kwargs} started")
        try:
            method(self, *args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {method.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            #raise Exception
        else:
            ToLog(f"{method.__name__}, *args = {args}, **kwargs = {kwargs} finished successfully")
    return wrapper

def except_method_brief(method):
    def wrapper(self, *args, **kwargs):
        #ToLog(f"{method.__name__} started")
        try:
            method(self, *args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {method.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            #raise Exception
        else:
            ToLog(f"{method.__name__} finished successfully")
    return wrapper

def except_method_briefer(method):
    def wrapper(self, *args, **kwargs):
        #ToLog(f"{method.__name__} started")
        try:
            method(self, *args, **kwargs)
        except Exception as Err:
            ToLog(f"Error in {method.__name__}, *args = {args}, **kwargs = {kwargs}, Error code = {Err}")
            r#aise Exception
    return wrapper

def info_method_dec(method):
    def wrapper(self, *args, **kwargs):
        #ToLog(f"{method.__name__}, *args = {args}, **kwargs = {kwargs} started")
        method(self, *args, **kwargs)
        ToLog(f"{method.__name__}, *args = {args}, **kwargs = {kwargs} finished")
    return wrapper

#=================================
# HelloFrame
class HelloFrame(SplashScreen):
    def __init__(self, parent = None):
        super(HelloFrame, self).__init__(
            bitmap = wx.Bitmap(name = os.getcwd() + "\\images\\WritingPNG.png", type = wx.BITMAP_TYPE_PNG),
            splashStyle = wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT,
            milliseconds = 1500,
            parent = None,
            id = -1,
            pos = wx.DefaultPosition,
            size = (1000,1000),
            #wx.DefaultSize,
            style = wx.STAY_ON_TOP | wx.BORDER_NONE)
        self.Show(True)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def OnClose(self, event):
        event.Skip()
        self.Hide()
        
#=================================
# HelloFrame        
class IconTray(TaskBarIcon):
    def __init__(self, frame, docdir = "somedir", stop = False, pub = "SomePub"):
        TaskBarIcon.__init__(self)
        self.frame = frame
        self.docdir = docdir
        self.stop = stop
        self.pub = pub
        #image
        self.SetIcon(wx.Icon(os.getcwd() + "\\images\\WritingIco.ico"), "MonitSpiskiSS")
        self.imageidx = 1
        #popup menu
        self.START_STOP = wx.NewIdRef(count = 1)
        self.CONVERT = wx.NewIdRef(count = 1)
        self.CH_FRAME = wx.NewIdRef(count = 1)
        self.SHOW_STNGS = wx.NewIdRef(count = 1)
        self.SHOW_LIC = wx.NewIdRef(count = 1)
        self.EXIT_SCR = wx.NewIdRef(count = 1)
        #events
        #self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self.OnDClick)
        self.Bind(wx.EVT_MENU, self.OnStartStop, id = self.START_STOP)
        self.Bind(wx.EVT_MENU, self.OnConvert, id = self.CONVERT)
        self.Bind(wx.EVT_MENU, self.OnCheckFrame, id = self.CH_FRAME)
        self.Bind(wx.EVT_MENU, self.OnSettings, id = self.SHOW_STNGS)
        self.Bind(wx.EVT_MENU, self.OnLicense, id = self.SHOW_LIC)
        self.Bind(wx.EVT_MENU, self.OnExit, id = self.EXIT_SCR)

    #@except_method_dec
    def CreatePopupMenu(self):
        menu = wx.Menu()
        if self.stop == True:
            menu.Append(self.START_STOP, "Start...")
        else:
            menu.Append(self.START_STOP, "Stop...")
        menu.Append(self.CONVERT, "Convert to Excel...")
        menu.Append(self.CH_FRAME, "Show status...")
        menu.Append(self.SHOW_STNGS, "Settings...")
        menu.Append(self.SHOW_LIC, "License")
        menu.Append(self.EXIT_SCR, "Exit")
        return menu

    @except_method_dec
    def OnStartStop(self, evt):
        wx.CallAfter(
            pub.sendMessage, self.pub,
            mess = "StartStop")

    @except_method_brief
    def OnConvert(self, evt):
        if self.docdir + "\\temp\\temp.db" in os.listdir(self.docdir + "\\temp"):
            os.remove(self.DocDir + "\\temp\\temp.db")
            ToLog("Previous temp file removed")
                                                              
        DialogLoad = wx.FileDialog(
            None,
            "Load file",
            #defaultDir = self.DocDir,
            wildcard = "db files (*.db)|*db",
            style = wx.FD_OPEN)

        if DialogLoad.ShowModal() == wx.ID_CANCEL:
            ToLog("Cancel converting")
            return

        else:
            
            print("GetDir = ", DialogLoad.GetDirectory())
            print("GetFilename = ", DialogLoad.GetFilename())
            selected_file = DialogLoad.GetDirectory() + "\\" + DialogLoad.GetFilename()
            selected_dir = DialogLoad.GetDirectory()
            
            tempFile = self.docdir + "\\temp\\temp.db"
            CopyFile(selected_file, tempFile)
        
            ConvThread = ConvertThread(path_to_file = tempFile, path_to_dir = selected_dir)
            ConvThread.setDaemon(True)
            ConvThread.start()

    @except_method_dec
    def OnCheckFrame(self, evt):
        wx.CallAfter(
            pub.sendMessage, self.pub,
            mess = "CheckFrame")
    
    @except_method_dec
    def OnSettings(self, evt):
        wx.CallAfter(
            pub.sendMessage, self.pub,
            mess = "Settings")

    @except_method_brief
    def OnLicense(self, evt):
        LICENSE = (
            "Данная программа является свободным программным обеспечением\n"+
            "Вы вправе распространять её и/или модифицировать в соответствии\n"+
            "с условиями версии 2 либо по Вашему выбору с условиями более\n"+
            "поздней версии Стандартной общественной лицензии GNU, \n"+
            "опубликованной Free Software Foundation.\n\n\n"+
            "Эта программа создана в надежде, что будет Вам полезной, однако\n"+
            "на неё нет НИКАКИХ гарантий, в том числе гарантии товарного\n"+
            "состояния при продаже и пригодности для использования в\n"+
            "конкретных целях.\n"+
            "Для получения более подробной информации ознакомьтесь со \n"+
            "Стандартной Общественной Лицензией GNU.\n\n"+
            "Данная программа написана на Python\n"
            "Автор: Титовский С.А.")
        
        wx.MessageBox(LICENSE, "Лицензия", wx.OK)

    @except_method_dec
    def OnExit(self, evt):
        wx.CallAfter(
            pub.sendMessage, self.pub,
            mess = "ExitCmd")

#settings dialog
class dlg(wx.Dialog):
    def __init__(
        self, settings, label = "Settings"):
        self.def_dict = {
            "num_days": ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'], # numdays
            "time_renew": ['30', '60', '90', '120', '150', '180', '300', '600'], # time renew16
            "time_ftp": ['300', '600', '900', '1200', '1500'],
            "ftp_addr": "None", "ftp_login":"None", "ftp_password":"None"}

        choices = {}
        self.Data = []
        
        for key in self.def_dict.keys():
            if settings[key] in self.def_dict[key]:
                for i in range (0, len(self.def_dict[key])):
                    if self.def_dict[key][i] == settings[key]:
                        choices.update({key: i})
            else:
                choices.update({key: 0})
        
        wx.Dialog.__init__(self, None, -1, label)
 
        labels = [
            "Number of days", "Scanning interval, seconds",
            "Sending to FTP interval, seconds",
            "FTP address",
            "FTP login",
            "FTP password"]
        #posSText = []
        for i in range (0, len(labels)):
            text = wx.StaticText(self, wx.ID_ANY, labels[i], pos = (10, 10 + 60*i))
            text.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))

        i = 0
        for key in self.def_dict.keys():
            if isinstance (self.def_dict[key], list):
                temp = wx.Choice(self, wx.ID_ANY, pos = (10, 35 + 60*i), size = (380, 30),
                                 choices = [str(some) for some in self.def_dict[key]])
                temp.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
                temp.SetSelection(choices[key])
                self.Data.append(temp)
            else:
                if key != "ftp_password":
                    temp = wx.TextCtrl(
                        self, wx.ID_ANY, settings[key], pos = (10, 35 + 60*i),
                        size = (330, 30))
                else:
                    temp = wx.TextCtrl(
                        self, wx.ID_ANY, settings[key], pos = (10, 35 + 60*i),
                        size = (330, 30), style = wx.TE_PASSWORD)
                
                temp.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
                temp.SetValue(settings[key])
                self.Data.append(temp)

            i = i + 1
                
        OKButton = wx.Button(self, wx.ID_OK, "OK", pos = (140, 20 + 60*len(labels)), size = (120, 30))
        OKButton.SetDefault()
        OKButton.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
        self.SetClientSize((400, 60 + len(labels)*60))
        self.Bind(wx.EVT_CLOSE, self.NoClose)
    
    def NoClose(self, evt):
        print("No Close")

#checkframe
class ChFrame(wx.Frame):
    def __init__(
        self, parent = None, label = " ", data = {"some_Key": "some_value"}):

        self.data = data
    
        #wx.Frame.__init__(self, None, -1, label)
        wx.Frame.__init__(
            self, parent, -1, label,
            style =
            wx.MINIMIZE_BOX|wx.CAPTION|wx.SYSTEM_MENU|wx.CLOSE_BOX|
            wx.CLIP_CHILDREN|wx.MAXIMIZE_BOX|wx.RESIZE_BORDER)

        frameIcon = wx.Icon(os.getcwd() + "\\images\\WritingPNG.png")
        self.SetIcon(frameIcon)

        #self.sizer = sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer = sizer = wx.FlexGridSizer(rows = len(data.keys()), cols = 2, hgap = 6, vgap = 6)
        #for i in range (0, len(data.keys())):
            #sizer.AddGrowableRow(i, 0)
        #sizer.AddGrowableCol(0, 1)
        sizer.AddGrowableCol(1, 1)
        self.values = {}
        
        for key in data.keys():
            text = wx.StaticText(self, wx.ID_ANY, key)
            text.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
            sizer.Add(text, -1, wx.ALL | wx.EXPAND , border = 5)
          
            temp = wx.TextCtrl(self, wx.ID_ANY, data[key], style = wx.TE_READONLY | wx.TE_CENTRE)
            temp.SetFont(wx.Font(12, wx.ROMAN, wx.NORMAL, wx.NORMAL))
            temp.SetValue(data[key])
            self.values.update({key: temp})
            sizer.Add(temp, -1, wx.ALL | wx.EXPAND , border = 5)
         

        colour = wx.SystemSettings.GetColour(wx.SYS_COLOUR_MENU)
        self.SetBackgroundColour(colour)
        
        self.SetSizer(sizer)
        
        self.SetMaxSize((1000, len(data.keys()) * 60))
        self.Layout()
        self.Fit()
        self.SetClientSize(500, len(data.keys()) * 60)
        self.Show(True)

    @except_method_brief
    def UpdateData(self, upd_data = {"some_Key": "some_value"}):
        for key in self.data.keys():
            for upd_key in upd_data.keys():
                if key == upd_key:
                    self.values[key].SetValue(upd_data[key])
        
#=========================================
#=========================================
#=========================================
#=========================================        
# OsnFrame
class Main_Frame(wx.Frame):

    def __init__(self, parent, WinPos = wx.DefaultPosition, DateUser = "01.01.2021", DocDir = os.getcwd()):
        wx.Frame.__init__(
            self, parent, -1, "Список совещаний")
        #global Region
        #Region = 1
        self.WinPos = WinPos
        self.Date = Dt_to_txt(datetime.today())
        self.DocDir = DocDir

        frameIcon = wx.Icon(os.getcwd() + "\\images\\WritingPNG.png")
        self.SetIcon(frameIcon)

        #self.pub = f"Main_Frame {self.GetId()}"
        self.pub = "Main_Frame"
        pub.subscribe(self.UpdateDisplay, self.pub)
        
        #print(f"MyId (Frame) = MainFrame {self.GetId()}")
        
        self.OpenPanel()
        #self.Show(True)

        #self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        ToLog("Main Frame opened on Date = " + str(self.Date))

#=========================================
    @info_method_dec
    def OpenPanel(self):
        self.panel = Main_Panel(self, self.Date, self.DocDir, self.pub)

#==========================================
    #@except_method_dec
    def UpdateDisplay(self, mess = "someMessage"):
        #print("\t\tGOT mess = " + str(mess))
        if isinstance (mess, str):
            #print("\t\tGOT mess = " + mess)
            if mess == "RenewNow":
                self.panel.RenewCommand()
            if mess == "RenewNowFirst":
                self.panel.RenewCommand(first_time = True)
            elif mess == "ExitCmd":
                self.panel.CloseCmd(from_tray = True)
            elif mess == "Settings":
                self.panel.ShowSettings()
            elif mess == "StartStop":
                self.panel.StartStop()
            elif mess == "CheckFrame":
                self.panel.CheckFrame()
        if isinstance (mess, list):
            #print("\t\tGOT mess = " + mess[0])
            if mess[0] == "ListDate":
                self.panel.CheckDate(first_time = mess[1], date = mess[2], data = mess[3])
            elif mess[0] == "ChangedDate":
                self.panel.ChangeDate(mess[1])
            elif mess[0] == "FTPtime":
                self.panel.chData.update({"Last FTP sending time": mess[1]})
                    
#=========================================
#=========================================
#=========================================
#=========================================
# panel
class Main_Panel(wx.Panel):
    def __init__(self, parent, date, docdir, pub):
        wx.Panel.__init__(self, parent = parent)

        self.frame = parent
        self.DocDir = docdir
        self.date = date
        self.pub = pub

        self.settings = {}
        self.LoadSettings()

        #change here num of days
        self.threads = None
        self.actualData = None
        self.conns = {}
        self.pause = True
        self.RenewEvt = threading.Event()

        #chframe
        self.chFrame = None
        self.chData = {
            "Current threads": "not updated",
            "Last databases renew time": "not updated",
            "Last FTP sending time": "not updated",
            "Script started at time": str(datetime.today()),
            "This frame opened at time": "not updated"}
        self.FTPtime = None
        
        CommonVbox = wx.FlexGridSizer(rows = 2, cols = 1, hgap = 6, vgap = 6)
        CommonVbox.AddGrowableRow(0, 1)
        CommonVbox.AddGrowableRow(1, 1)
        CommonVbox.AddGrowableCol(0, 1)

        #-----------------------------------------------------------------
        Btn = wx.Button(self, wx.ID_ANY, "StartStopBtn")
        Btn.Bind(wx.EVT_BUTTON, self.OnBtn)

        #ButSize = (300, 200)
        #LoadBtn.SetSize(ButSize)
        CommonVbox.Add(Btn, -1, wx.EXPAND | wx.ALL, 4)
        #CommonVbox.SetMinSize(ButSize)

        self.SetSizer(CommonVbox)
        self.SetSize((300, 400))
        #self.Fit()


        self.frame.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        self.Show(True)

        self.Tray = IconTray(self.frame, docdir = self.DocDir, stop = self.pause, pub = self.pub)
        self.StartScanDays(days = int(self.settings["num_days"]))
        self.StartRenewThread()
        
#=============================================
    except_method_brief
    def LoadSettings(self):
        if "settings.db" not in os.listdir(self.DocDir + "\\Based"):
            print("Create new db file")
            self.CreateSettingsDB(path = self.DocDir + "\\Based\\settings.db")
            
        self.LoadFrDB(path = self.DocDir + "\\Based\\settings.db")
        
        ToLog("Settings after load = " + str(self.settings))

    @except_method_dec
    def LoadFrDB(self, path = "somepath"):
        setconn = sqlite3.connect(path)
        cursor = setconn.execute("SELECT * FROM SETTINGS WHERE show_status=(?)", (1, ))
        pair = [(row[1], row[2]) for row in cursor]
        setconn.close()

        for i in range (0, len(pair)):
            self.settings.update({pair[i][0]:pair[i][1]})

        return 

    @except_method_dec
    def CreateSettingsDB(self, path = "somepath"):
        setconn = sqlite3.connect(path)
        setconn.execute(
            "CREATE TABLE IF NOT EXISTS SETTINGS" +
            "(num INTEGER PRIMARY KEY AUTOINCREMENT," +
            "key TEXT NOT NULL," +
            "value TEXT NOT NULL," +
            "show_status BOOLEAN NOT NULL CHECK(show_status IN (0, 1))," + 
            "note1 CHAR(100)," +
            "note2 CHAR(100));")

        dictDef = {
            "num_days": "4", "time_renew": "60", "time_ftp": "1200",
            "ftp_addr": "10.135.11.177", "ftp_login": "login",
            "ftp_password": "TSA&44186"}

        for item in dictDef.keys():
            setconn.execute(
                "INSERT INTO SETTINGS VALUES (" +
                "NULL, '" + item + "', '" + dictDef[item] + "', 1, NULL, NULL)")

        setconn.commit()
        setconn.close()

    @except_method_dec
    def SaveSettings(self):
        if "settings.db" not in os.listdir(self.DocDir + "\\Based"):
            #print("Create new db file")
            self.CreateSettingsDB(path = self.DocDir + "\\Based\\settings.db")
        else:
            self.SaveDB(path = self.DocDir + "\\Based\\settings.db")
        #print("Settings after load = " + str(self.settings))

    @except_method_dec
    def SaveDB(self, path = "somepath"):
        setconn = sqlite3.connect(path)
        for item in self.settings.keys():
            ToLog(f"Updated values in settings db =  {item}: {self.settings[item]}")
            cursor = setconn.execute("UPDATE SETTINGS SET value = '" + self.settings[item] + "' WHERE key = '" + item + "'")
        setconn.commit()
        setconn.close()

    @except_method_brief
    def ShowSettings(self):
        dlg1 = dlg(label = "Settings of MonitSpiskiSS", settings = self.settings)
        if dlg1.ShowModal() == wx.ID_OK:
            self.settings.update({"num_days": dlg1.Data[0].GetStringSelection()})
            self.settings.update({"time_renew": dlg1.Data[1].GetStringSelection()})
            self.settings.update({"time_ftp": dlg1.Data[2].GetStringSelection()})
            self.settings.update({"ftp_addr": dlg1.Data[3].GetValue()})
            self.settings.update({"ftp_login": dlg1.Data[4].GetValue()})
            self.settings.update({"ftp_password": dlg1.Data[5].GetValue()})
            
        #ToLog("New settings entered = " + str(self.settings))
        #ToLog("saving settings to db")
        self.SaveSettings()
        ToLog("Stopping threads after changing settings")
        self.StopThreads(except_renew = True)
        ToLog("Starting threads with new settings")
        self.StartScanDays(days = int(self.settings["num_days"]))
        self.RenewThread.UpdateData(upd_data = self.settings)
        

    @except_method_brief
    def CheckFrame(self):
        if self.chFrame != None:
            try:
                ToLog("Destroy previous CheckFrame")
                self.chFrame.Destroy()
            except Exception:
                pass
            self.chFrame = None
        self.chData.update({"This frame opened at time": str(datetime.today())})
        self.chFrame = ChFrame(label = "Check status", data = self.chData)                    
 
    @except_method_dec  
    def OnCloseWindow(self, evt):
        self.CloseCmd()
        evt.Skip()
        sys.exit()

    @except_method_dec 
    def CloseCmd(self, from_tray = False):
        ToLog("Application closed by User's command")
        self.Show(False)
        ToLog(f"Save and close db file, now conns = {self.conns}")
        print(str(self.conns.items()))
        for key in self.conns.keys():
            self.conns[key].commit()
            self.conns[key].close()

        self.StopThreads()
       
        ToLog("Stopping LogThread")
        global threadLog
        threadLog.stop = True
        threadLog.join()

        if from_tray == True:
            self.frame.Destroy()
            sys.exit()
            
    @except_method_dec 
    def StartRenewThread(self):
        self.RenewThread = RenewThread(evt = self.RenewEvt, pub = self.pub, settingsDict = self.settings)
        self.RenewThread.setDaemon(True)
        self.RenewEvt.clear()
        self.RenewThread.start()

    @except_method_dec 
    def StopThreads(self, except_renew = False):
        if except_renew == False:
            self.RenewThread.Stop()
        if self.threads:
            for thread in self.threads:
                thread.Stop()

    @except_method_dec
    def StartStop(self):
        if self.RenewEvt.isSet():
            print("Pause Renew")
            self.Tray.stop = True
            self.RenewEvt.clear()
        else:
            print("Resume")
            self.Tray.stop = False
            self.RenewEvt.set()
            
    @except_method_dec
    def StartScanDays(self, days = 4):
        self.threads = []
        self.actualData = {}
        for day in range (0, days):
            thread = ScanDayThread(pub = self.pub, num_day = day)
            thread.setDaemon(True)
            thread.start()
            self.threads.append(thread)

    @except_method_dec
    def RenewCommand(self, first_time = False):
        if first_time == False:
            self.chData.update({"Last databases renew time": str(datetime.today())})
        if self.FTPtime:
            self.chData.update({"Last FTP sending time": self.FTPtime})                   
        #print("RENEW cur conns = " + ", ".join(self.conns.keys()))
        #print("RENEW cur datas = " + ", ".join(self.actualData.keys()))
        #print("Threads in self.threaads")
        templist = []
        if self.threads:
            for thread in self.threads:
                templist.append(thread.date)
                thread.ScanNow(first_time)
            self.chData.update({
            "Current threads": str(len(templist)) + " days watching: " + ", ".join(templist)})

        if self.chFrame:
            self.chFrame.UpdateData(upd_data = self.chData)     

    @except_method_briefer
    def CheckDate(self, first_time, date, data):
        if first_time == False:
            if datetime.strptime(date, "%d.%m.%Y") < (datetime.strptime(Dt_to_txt(datetime.today()), "%d.%m.%Y")):
                print(f"Received data from past days {date}, deleting this date from threads")
                ToLog(f"Received data from past days {date}, deleting this date from threads")
                if date in self.actualData.keys():
                    del self.actualData[date]
                if date in self.conns.keys():
                    self.conns[date].commit()
                    self.conns[date].close()
                    del self.conns[date]
            else:
                self.RenewData(first_time, date, data)
                
        else:
            self.RenewData(first_time, date, data)

    @except_method_dec
    def ChangeDate(self, date):
        text_date = Dt_to_txt(date)
        
        #compare current treads and dates
        data_list = []
        thread_list = []
        for day in range (0, int(self.settings["num_days"])):
            temp_day = date + timedelta(days = day)
            data_list.append(Dt_to_txt(temp_day))
        for day in range (0, len(self.threads)):
            thread_list.append(self.threads[day].date)
        #pop old threads
        temp = []
        for thread in self.threads:
            if datetime.strptime(thread.date, "%d.%m.%Y").date() < date.date():
                ToLog(f"Deleting ScanThread with old date {thread.date}")
                print(f"Deleting ScanThread with old date {thread.date}")
                thread.Stop()
                if thread.date in self.actualData.keys():
                    del self.actualData[thread.date]
                if thread.date in self.conns.keys():
                    self.conns[thread.date].commit()
                    self.conns[thread.date].close()
                    del self.conns[thread.date]
            else:
                self.CreateNewDB(text_date)
                temp.append(thread)
        self.threads = temp[:]
        new_datas = list(set(data_list) - set(thread_list))

        for day in new_datas:
            thread = ScanDayThread(pub = self.pub, num_day = day, date = day)
            thread.setDaemon(True)
            thread.start()
            self.threads.append(thread)
                 
    @except_method_briefer
    def RenewData(self, first_time, date, data):
        if len(self.actualData) == 0:
            self.actualData.update({date: data})
            self.SaveToDB(first_time = first_time, date = date, data = data)
        elif date in self.actualData:
            self.CompareData(date = date, data = data)
        else:
            self.actualData.update({date: data})
            self.SaveToDB(first_time = first_time, date = date, data = data)
        #ToLog(f"Now ActualData\n {self.actualData}")

    @except_method_briefer
    def CompareData(self, date, data):
        if self.actualData[date] == data:
            ToLog(f"No changes in actualData on date {date}")
        else:
            #print(", ".join(data.keys()))
            for item in data.keys():
                #if new meeting added
                if item not in self.actualData[date].keys():
                    ToLog("Meeting added:\n\t " + str(data[item]))
                    self.actualData[date].update({item: data[item]})
                    self.SaveToDB(first_time = False, date = date, item = data[item])
                else:
                    if data[item] == self.actualData[date][item]:
                        ToLog(f"No changes in item {item}")
                    else:
                        ToLog(
                            f"Changes in item {item}:\n\t previous field = " +
                            f"{self.actualData[date][item]}, \n\t new field = {data[item]}")
                        self.SaveToDB(first_time = False, date = date, item = data[item])
            #if meeting popped
            diff = list(set(self.actualData[date].keys()) - set(data.keys())) 
            if len(diff) > 0:
                self.PopMeetings(date, diff)
            self.actualData[date] = data

    @except_method_dec
    def CreateNewDB(self, date = "01.01.2024"):
        nameDB = self.DocDir + "\\Monitoring_Logs\\mchanges_" + date + ".db"
        if "mchanges_" + date + ".db" in os.listdir(self.DocDir + "\\Monitoring_Logs"):
            if date not in self.conns.keys():
            #ToLog(f"mchanges_{date}.db is already in dir {self.DocDir}\\Monitoring_Logs")
                conn = sqlite3.connect(nameDB)
                self.conns.update({date: conn})
            
        else:
            #print(f"cerate new db with date {text_date}")
            ToLog(f"creating mchanges_{date}.db in dir {self.DocDir}\\Monitoring_Logs")
            conn = sqlite3.connect(nameDB)
            self.conns.update({date: conn})
            conn.execute(
                "CREATE TABLE IF NOT EXISTS MEETING_CHANGES" +
                "(num INTEGER PRIMARY KEY AUTOINCREMENT," +
                "date_now TEXT NOT NULL," +
                "time_now TEXT NOT NULL," +
                "date TEXT NOT NULL," +
                "id TEXT NOT NULL," +
                "studia TEXT NOT NULL," +
                "initiator TEXT NOT NULL," +
                "time TEXT NOT NULL," +
                "rezhim TEXT NOT NULL," +
                "theme TEXT," +
                "participants TEXT," + 
                #"show_status BOOLEAN NOT NULL CHECK(show_status IN (0, 1)),"
                "note1 CHAR(100)," +
                "note2 CHAR(100));")
            self.conns.update({date: conn})

    @except_method_brief
    def SaveToDB(self, first_time = False, date = "01.01.2024", data = None, item = None):
        self.CreateNewDB(date = date)
        #print(f"I received data = {data}, item = {item}")
        if item:
            if first_time == False:
                self.conns[date].execute(
                    "INSERT INTO MEETING_CHANGES VALUES (" +
                    "NULL,'"  + Dt_to_txt(datetime.now()) + "','" +
                    str(datetime.now())[11:16] + "','" + date + "','" +
                    "','".join(item) + "',NULL,NULL)")
            else:
                self.conns[date].execute(
                    "INSERT INTO MEETING_CHANGES VALUES (" +
                    "NULL,'"  + Dt_to_txt(datetime.now()) + "','" +
                    str(datetime.now())[11:16] + "','" + date + "','" +
                    "','".join(item) + "','this item added when thread was started',NULL)")
                
            self.conns[date].commit()

        if data:
            for key in data.keys():
                self.SaveToDB(first_time = first_time, date = date, item = data[key])
                
    @except_method_dec
    def PopMeetings(self, date = "01.01.2024", id_list = ["111111"]):
        print("popping " + ", ".join(id_list))
        for item in id_list:
            self.conns[date].execute(
                "INSERT INTO MEETING_CHANGES VALUES (" +
                "NULL,'"  + Dt_to_txt(datetime.now()) + "','" +
                str(datetime.now())[11:16] + "','" + date + "','" +
                item + "','Deleted','Deleted','Deleted','Deleted','Deleted','Deleted',NULL,NULL)")
                #item + "','Deleted'" * 6 + ",NULL,NULL)")
            self.conns[date].commit()

    @except_method_brief
    def OnBtn(self, evt):
        ToLog("OnBtn pressed")

#=============================================
#=============================================
#=============================================
#=============================================
#RenewThread
class ScanDayThread(threading.Thread):
    def __init__(self, num_day = 0, pub = "my_pub", date = None):
        super().__init__()
        self.stop = False
        self.pub = pub
        if date:
            self.date = date
        else:
            date_tm = datetime.today() + timedelta(days = num_day)
            self.date = Dt_to_txt(date_tm)

    @except_method_brief
    def run(self):
        ToLog(f"Thread on date {self.date} started")
        while True:
            if self.stop == True:
                break
            time.sleep(1)
        ToLog(f"Thread on date {self.date} finished")

    @except_method_briefer
    def ScanNow(self, first_time = True):
        ListDate = SpisokFromDate(date = self.date)
        wx.CallAfter(
            pub.sendMessage, self.pub,
            mess = ["ListDate", first_time, self.date, ListDate])
        
    @except_method_brief        
    def Stop(self):
        ToLog(f"Thread on date {self.date} received Stop command")
        self.stop = True


#=============================================
#=============================================
#=============================================
#=============================================
#RenewThread
class RenewThread(threading.Thread):
    @except_method_dec
    def __init__(self, evt, pub = "somePub", settingsDict = {"some": "some"}):
        super().__init__()
        self.stop = False
        self.evt = evt
        self.pub = pub
        self.settings = settingsDict
        self.TimeRenew = int(self.settings["time_renew"])
        self.TimeFTP = int(self.settings["time_ftp"])
        self.cycles = self.TimeFTP // self.TimeRenew
        self.today = datetime.today().date()
        #print("at start today = " + str(self.today))
        self.evt.set()
        self.once = 0
        
    @except_method_brief
    def run(self):
        ToLog(f"Renew thread started wuth pub {self.pub} and timerenew {self.TimeRenew}")
        thread = FTPThread(self.settings, self.pub)
        thread.setDaemon(True)
        thread.start()
        self.RenewThreadCommand(first_time = True)
        startTime = time.time()
        now_cycles = 0
        #tempor
        #offset = 1
        while True:
            #print("--iter--")
            self.evt.wait()
            
            if self.stop == True:
                ToLog("RenewThread stopped by Stop event")
                break
            if time.time() - startTime < self.TimeRenew:
                #ToLog("\t\tToo early for renew, last time = " + str(time.time() - startTime))
                time.sleep(2)
                continue
            else:
                self.RenewThreadCommand()
                #if now_cycles % 2 == 0:
                    #tempor
                    #self.ChangedDate(date = (datetime.today() + timedelta(days = offset)))
                    #print("temporary incrasing ate by 1")
                    #self.once = 1
                    #offset = offset + 1

                #tempor commented
                if datetime.today().date() != self.today:
                    self.ChangedDate(date = datetime.today())

                startTime = time.time()
                now_cycles += 1
                if now_cycles >= self.cycles:
                    thread = FTPThread(self.settings, self.pub)
                    thread.setDaemon(True)
                    thread.start()
                    now_cycles = 0
                    
            #ToLog("RenewThread Finished with RenewTime = " + str(TimeRenew))

    @except_method_brief
    def Stop(self):
        ToLog(f"Renew thread received Stop command")
        self.stop = True

    @except_method_brief
    def RenewThreadCommand(self, first_time = False):
        if first_time == False:
            ToLog(f"Threads will be renewed because of time Renew = {self.TimeRenew} elapsed")
            wx.CallAfter(pub.sendMessage, self.pub, mess = "RenewNow")
        else:
            ToLog(f"Threads will be renewed because it's lust started")
            wx.CallAfter(pub.sendMessage, self.pub, mess = "RenewNowFirst")

    @except_method_dec
    def ChangedDate(self, date = datetime.today().date()):
        ToLog(f"Changed date to {date}, sending ChangedDate command")
        print(f"Changed date to {date}, sending ChangedDate command")
        self.today = date.date()
        wx.CallAfter(pub.sendMessage, self.pub, mess = ["ChangedDate", date])

    @except_method_brief
    def UpdateData(self, upd_data = {"somedata": "data"}):
        print("upd_data")
        self.settings = upd_data
        self.TimeRenew = int(self.settings["time_renew"])
        self.TimeFTP = int(self.settings["time_ftp"])
        self.cycles = self.TimeFTP // self.TimeRenew
        self.evt.set()
        thread = FTPThread(self.settings, self.pub)
        thread.setDaemon(True)
        thread.start()

#=============================================
#=============================================
#=============================================
#=============================================
#RenewThread
class FTPThread(threading.Thread):
    @except_method_brief
    def __init__(self, settingsDict = {"some": "some"}, pub = "somepub"):
        super().__init__()
        self.settings = settingsDict
        self.pub = pub

    @except_method_dec
    def run(self):
        #print(f"FTP thread started")
        ToLog(f"FTP thread started")
        self.SendToFTP(path = self.settings["ftp_addr"], login = self.settings["ftp_login"], password = self.settings["ftp_password"], key = "mchanges_")
        self.CleanOldFiles(key = "mchanges_")
        #print(f"FTP thread finished")
        ToLog(f"FTP thread fifnished")

    @except_method_dec
    def CleanOldFiles(self, key = "some_key"):
        global MonitLogDir
        for item in os.listdir(MonitLogDir):
            if item.find(key) != -1 and item.find(".db") != -1:
                name = item[item.find(key) + len(key):item.find(".db")]
                date = Dt_to_txt(datetime.today())
                if datetime.strptime(name, "%d.%m.%Y") < datetime.strptime(Dt_to_txt(datetime.today()), "%d.%m.%Y"):
                    os.remove(MonitLogDir + "\\" + item)
                    ToLog(f"Removed old file = {item}")
            else:
                ToLog(f"no remove {item}, it is not auto database file")
                #print(f"no remove {item}, it is not auto database file")
            #else:
            #    print(f"newFile = {item}")

    def SendToFTP(self, path = "10.135.11.177", login = "DB", password = "TSA&44186", key = "somekey"):
        global MonitLogDir
        try:
            namedir = str(socket.gethostname()).replace("\\", " ")
            namedir = namedir.replace(":", " ")
            ftp = ftplib.FTP(path)
            ftp.login(login, password)
            self.CreateDir(ftp, "MonitSpiskiSS\\" + namedir)
            for file in os.listdir(MonitLogDir):
                if file.find(key) != -1 and file.find(".db") != -1:
                    with open(MonitLogDir + "\\" + file, "rb") as somefile:
                        ftp.storbinary("STOR " + file, somefile)
                        somefile.close()
                else:
                    ToLog(f"not send to FTP {file}, it is not auto database file")
                    print(f"not send to FTP {file}, it is not auto database file")
            ftp.quit()
            wx.CallAfter(pub.sendMessage, self.pub, mess = ["FTPtime", str(datetime.today())])
        except Exception as Err:
            print("Error connecting to FTP")
            ToLog("Error connecting to FTP, Error code = " + str(Err))

    @info_method_dec
    def CreateDir(self, ftp, dirname):
        try:
            ftp.cwd(dirname)
        except Exception:
            ftp.mkd(dirname)
            ftp.cwd(dirname)

#=============================================
#=============================================
#=============================================
#=============================================
#ConvertThread
class ConvertThread(threading.Thread):
    @except_method_dec
    def __init__(self, path_to_file = "some_file", path_to_dir = "some_dir"):
        super().__init__()
        self.file = path_to_file
        self.dir = path_to_dir

    @except_method_brief
    def run(self):
        ToLog(f"Convert thread started wuth {self.dir}")
        self.CreateStyles()
        conn = sqlite3.connect(self.file)
        cursor = conn.execute("SELECT name FROM sqlite_master WHERE type = 'table';")
        self.tables = [table[0] for table in cursor.fetchall()]
            
        #print("found tables = " + str(self.tables))
        wb = openpyxl.Workbook()
        for table in self.tables:
            wb.create_sheet(table)
            work = wb[table]

            #read column names
            cursor = conn.execute("SELECT name, type FROM pragma_table_info('" + table + "')")
            labels = [row[0] for row in cursor]

            for i in range (0, len(labels)):
                work.cell(row = 1, column = 1 + i, value = labels[i]).font = self.fonttable
                work.cell(row = 1, column = 1 + i).alignment = self.alightable
                work.cell(row = 1, column = 1 + i).border = self.bordertable
                    
                #read data
            cursor = conn.execute("SELECT * FROM " + table)
            raw_data = [row for row in cursor]
            
            for row in range (0, len(raw_data)):
                for col in range (0, len(raw_data[row])):
                    work.cell(row = 2 + row, column = 1 + col, value = raw_data[row][col]).font = self.fonttable
                    work.cell(row = 2 + row, column = 1 + col).alignment = self.alightable
                    work.cell(row = 2 + row, column = 1 + col).border = self.bordertable

            colsize = [8, 13, 9, 14, 11, 17, 27, 13, 15, 56, 75, 27, 27]
            for col in range (0, len(colsize)):
                work.column_dimensions[get_column_letter(col + 1)].width = colsize[col]
                

        wb.remove(wb['Sheet'])
        wb.save(self.dir + "\\result.xlsx")
        conn.close()

        os.remove(self.file)
        os.startfile(os.path.realpath(self.dir + "\\result.xlsx"))                     

    @except_method_brief
    def CreateStyles(self):
        self.fonttable = Font(name="Times New Roman", size=12, bold=False)
        
        self.bordertable = Border(left=Side(border_style="thin", color="FF000000"),
                             right=Side(border_style="thin", color="FF000000"),
                             top=Side(border_style="thin", color="FF000000"),
                             bottom=Side(border_style="thin", color="FF000000"))

        self.alightable = Alignment(horizontal = "center",
                               vertical = "center",
                               text_rotation = 0,
                               wrap_text = False,
                               shrink_to_fit = True,
                               indent = 0)
         
#@except_foo_dec
def Dt_to_txt(datetime):
    return str(datetime)[8:10] + "." + str(datetime)[5:7] + "." + str(datetime)[0:4]        
#=============================================
#=============================================
#=============================================
#=============================================
# scaling bitmap
@except_foo_dec
def ScaleBitmap(bitmap, size):
    image = bitmap.ConvertToImage()
    image = image.Scale(size[0], size[1], wx.IMAGE_QUALITY_HIGH)
    return wx.Image(image).ConvertToBitmap()

#=============================================
#=============================================
#=============================================
#=============================================
# Tolog - renew log
def ToLog(message, startThread = False):
    try:
        global LogQueue
        LogQueue.put(str(datetime.today())[10:19] + "  " + str(message) + "\n")
    except Exception as Err:
        print("Error in ToLog function, Error code = " + str(Err))
        
#=============================================
#=============================================
#=============================================
#=============================================
# Thread for saving logs
class LogThread(threading.Thread):
    def __init__(self):
        super().__init__()
        self.stop = False

    def run(self):
        global LogQueue
        ToLog("LogThread started!!!")
        self.writingQueue()
        ToLog("LogThread finished!!!")

    def writingQueue(self):
        global LogQueue, LogDir
        while True:
            try:
                if LogQueue.empty():
                    if self.stop == True:
                        print("LogThreadStopped")
                        break
                    time.sleep(1)
                    continue
                else:
                    with open(LogDir + "\\" + str(datetime.today())[0:10] + ".cfg", "a") as file:
                        while not LogQueue.empty():
                            mess = LogQueue.get_nowait()
                            file.write(mess)
                            #print("Wrote to Log:\t" + mess)
                        file.close()
            except Exception as Err:
                print("Error writing to Logfile, Error code = " + str(Err))
                #raise Exception

#=============================================
#=============================================
#=============================================
#=============================================
#Copy file foo
@except_foo_dec
def CopyFile(source, dist, buffer = 1024*1024):
    ToLog("CopyFile " + source + " to " + dist + " function started")
    with open(source, "rb") as SrcFile, open(dist, "wb") as DestFile:
        while True:
            copy_buffer = SrcFile.read(buffer)
            if not copy_buffer:
                break
            DestFile.write(copy_buffer)
    ToLog("CopyFile " + source + " to " + dist + " function finished")

#=============================================
#=============================================
#=============================================
#=============================================
# ClearOldLogs
@except_foo_dec
def ClearLogs(Dir, numfiles = int(10)):
    while len(os.listdir(Dir)) >= numfiles:
        if len(os.listdir(Dir)) < numfiles:
                break
        try:
            os.remove(os.path.abspath(FindOldest(Dir)))
            ToLog(f"foo ClearLogs, dir = {Dir}: DELETING FILE " + str(FindOldest(Dir)))
        except Exception as Err:
            ToLog(f"foo ClearLogs, dir = {Dir}: Old file with logs wasn't deleted, Error code = " + str(Err))
            #raise Exception
            break

#=============================================
#=============================================
#=============================================
#=============================================   
# DeleteOldest
@except_foo_dec
def FindOldest(Dir):
    List = os.listdir(Dir)
    fullPath = [Dir + "/{0}".format(x) for x in List]
    oldestFile = min(fullPath, key = os.path.getctime)
    return oldestFile
    
#=============================================
#=============================================
#=============================================
#=============================================
@except_foo_dec
def FindMyDir(nameDir, subDirs = None):
    if "Documents" in os.listdir(os.path.expanduser("~")):
        DocDir = os.path.expanduser("~") + "\\Documents"
    else:
        os.mkdir(os.path.expanduser("~") + "\\Documents")
        DocDir = os.path.expanduser("~") + "\\Documents"
    if nameDir not in os.listdir(DocDir):
        os.mkdir(DocDir + "\\" + nameDir)
        ToLog(nameDir + "folder was Created")

    DocDir = DocDir + "\\" + nameDir
    if isinstance (subDirs, list):
        for word in subDirs:
            if word not in os.listdir(DocDir):
                os.mkdir(DocDir + "\\" + word)
                ToLog(word + " folder was Created")
    return DocDir
    #if "DataBase.db" not in os.listdir(DocDir + "\\Based"):
    #    CopyFile(os.getcwd() + "\\Based\\DataBase.db", DocDir + "\\Based\\DataBase.db")
    #else:
    #    os.remove(DocDir + "\\Based\\DataBase.db")
    #    CopyFile(os.getcwd() + "\\Based\\DataBase.db", DocDir + "\\Based\\DataBase.db")
    
#===============================================
#===============================================
#===============================================
#===============================================        
# Создание класса окна любой ошибки
@except_foo_brief
def SomeError(parent, title):
    wx.MessageBox(title, "Ошибка", wx.OK)

#===============================================
#===============================================
#===============================================
#===============================================
# List of meetings
@except_foo_dec
def SpisokFromDate(date = "03.12.2024", region = 0):
  
    # Подставляем дату в ссылку:
    ssilka = str("http://10.132.71.156/pls/ss/selector.sels.list?us=" + str(region) +
        "&str=" + date + "&wday=5")

    response = requests.get(
        f"http://10.132.71.156/pls/ss/selector.sels.list?us={region}&str={date}&wday=5")

    idsov = []
    namesov = []
    rezhimsov = []
    timesov = []
    initsov = []
    studiasov = []
    themesov = []
    spisoksov = []
    spisokuchast = []
    nomsov = 1

    # Если Управление
    if int(region) == 0:
        
        filesplit = response.text.splitlines()
        #print(str(filesplit))

        #Задаем параметры поиска
        poisk = "&us=0&sid="
        poisk2 = '''<td width=15% class=zag>Примечание</td>'''
        poisk3 = '''&nbsp;</td></tr>'''
        poisk4 = '''<td class="zag" rowspan=2>'''
        poisk5 = '''<td class="msk" rowspan=2>'''
        poisk6 = '''<a href="javascript:go(0,1,0'''
        poisk7 = '''Регион-'''
        poisk8 = '''&nbsp;'''
        poisk9 = '''&nbsp;</td><td class=norm>&nbsp'''
        nachalo = "2"
        konec = '''</td>'''

        #Обработка кода страницы и составление списков
        for i in range(0, len(filesplit)-1):
            filesplit[i] = str(filesplit[i]).strip()

            #добавление в списки разделителей - строки Начало совещания и Список участников
            if (
                filesplit[i].find('''<td width=15% class=zag>Примечание</td>''')>-1
                or
                filesplit[i]=='''&nbsp;</td></tr>'''):
                spisoksov.append(str(nomsov))
                spisokuchast.append("Список участников  "+str(nomsov))
                nomsov = nomsov + 1

            #составление списка участников конференций (необработанного)
            if filesplit[i].find('''<a href="javascript:go(0,1,0''')!=-1:
                spisokuchast.append(filesplit[i][filesplit[i].find('''">''')+2:filesplit[i].find('''</a>''')])
 
            #составление списка как в SMS
            if (
                (filesplit[i].find(poisk4)!=-1)
                or
                (filesplit[i].find(poisk5)!=-1)):
                if filesplit[i].find('''<br>''') > -1:
                    filesplit[i] = filesplit[i][:filesplit[i].find('''</td>''')+1]
                filesplit[i] = filesplit[i][filesplit[i].find(nachalo) + 2:filesplit[i].find(konec)]
                spisoksov.append(filesplit[i])
                #print("\tfrom SMS = " + str(filesplit[i]))

            #find themes
            if filesplit[i].find(poisk9)!=-1:
                filesplit[i] = filesplit[i][filesplit[i].find(poisk9) + len(poisk9) + 1:]
                spisoksov.append(filesplit[i])
                #print("\ttheme = " + str(filesplit[i]))

            #составление списка ID конференций (внутри списка SMS)
            if filesplit[i].find(poisk)!=-1:
                filesplit[i] = filesplit[i][filesplit[i].find(poisk)+10:filesplit[i].find('''>"''')-1]
                spisoksov.append(str(filesplit[i]))

        #print("begin of deparse")
        for i in range (6, len(spisoksov)):
            #print(str(spisoksov[i]))
            if (i+1)%7==0:
                studiasov.append(spisoksov[i-5])
                rezhimsov.append(spisoksov[i-4])
                timesov.append(spisoksov[i-3])
                initsov.append(spisoksov[i-2])
                themesov.append(spisoksov[i-1])
                idsov.append(spisoksov[i])

    # формируем списки с учетом отмен и проверок
    idsov1 = []
    studiasov1 = []
    themesov1 = []
    #namesov1 = []
    rezhimsov1 = []
    initsov1 = []
    timesov1 = []
    uchastsov1 = []
    #nomer = []
    #nomernach = 1
    uchastsov = []
    temp_uchast = []

    # Формируем список списков участников
    for i in range (1, len(spisokuchast)):
        if spisokuchast[i].find("Список участников")==-1:
            temp_uchast.append(spisokuchast[i])
        else:
            if len(temp_uchast)==0:
                uchastsov.append(["None"])
                temp_uchast.clear()
            else:
                uchastsov.append(temp_uchast[:])
                temp_uchast.clear()
                    
    for i in range (0, len(timesov)):
        #if (
        #    rezhimsov[i].find('''тмена''')!=-1
        #    or
        #    rezhimsov[i].find('''роверка''')!=-1):
        #    continue
        if 1 == 0:
            pass
        else:
            #nomernach = nomernach + 1
            #idsov1.append(idsov[i])
            #rezhimsov1.append(rezhimsov[i])
            #studiasov1.append(studiasov[i])
            #themesov1.append(themesov[i])
                   
            if len(timesov[i]) < 16:
                timesov[i] = timesov[i].replace("<br>-<br>", "")
                if timesov[i][-1] == ":":
                    timesov[i] = timesov[i][:-1]
                timesov1.append(timesov[i])
                             
            else:
                timesov1.append(timesov[i].replace("<br>-<br>", "-"))
            #initsov1.append(initsov[i])
            #nomer.append(str(nomernach-1))
            #if region == 0:
            #uchastsov1.append(uchastsov[i])

    itog = {}
    for item in range (0, len(idsov)):
        itog.update(
            {idsov[item]: [idsov[item], studiasov[item], initsov[item],
                           timesov1[item], rezhimsov[item], themesov[item],
                           ", ".join(uchastsov[item][:])]})
    #itog.append(idsov)
    #itog.append(studiasov)
    #itog.append(initsov)
    #itog.append(timesov1)
    #itog.append(rezhimsov)
    #itog.append(themesov)
    #itog.append(uchastsov)

    return itog

#===============================================
#===============================================
#===============================================
#=============================================== 
# Making list of meeting with some id
@except_foo_dec
def formfile(idsov):
    itogitogov = []
    for k in range (0,5):
        ssilkaSS = str(
            "http://10.132.71.156/pls/ss/selector.report.study_p?sid="+idsov+"&us="+str(k))
        
        # Запрашиваем код
        try:
            responseSS = requests.get(ssilkaSS)
        except Exception as Err:
            ToLog("Ошибка в обработке страницы списка присутствующих, код ошибки = " + str(Err))
            sys.exit()
        filesplit = responseSS.text.splitlines()

        #Задаем наши списки и номер совещания
        dolgnost = []
        fio = []
        prim = []

        #Обработка кода страницы и составление списков
        for i in range(0, len(filesplit)-1):                            
                # Формируем списки для таблицы - должность, ФИО, Примечание
            if (
                filesplit[i].find('''<tr><td colspan=3 class=z2>''')!=-1
                or
                filesplit[i].find('''<tr><td class=spr valign=top>''')!=-1):
                
                if filesplit[i].find('''<tr><td colspan=3 class=z2>''')!=-1:
                    dolgnost.append("КАБИНЕТ" + str(filesplit[i][filesplit[i].find('''z2>''')+3:filesplit[i].find('''</td>''')]))
                    fio.append("NONE")
                    prim.append("NONE")

                if filesplit[i].find('''<tr><td class=spr valign=top>''')!=-1:
                    dolgnost.append(filesplit[i+1])
                    fio.append("Новые участники:")
                    n = i
                    while filesplit[n].find('''</table></td>''')==-1:
                        n = n+1
                    for s in range (i,n):
                       
                        if filesplit[s].find('''<td class=spr>''')!=-1:
                            fio.append(
                                str(filesplit[s])[filesplit[s].find('''<td class=spr>''')+14:filesplit[s].find('''&nbsp;&nbsp''')]+
                                str("  ")+str(filesplit[s])[filesplit[s].find('''&nbsp;&nbsp''')+12:filesplit[s].find('''</td>''')])
                        elif (
                            filesplit[s+1].find('''</table></td>''')!=-1
                            and
                            filesplit[s].find('''<table width=''')!=-1):
                            fio.append("PUSTO")
                                                
                    prim.append(filesplit[n+2])

        # Преобразовываем список участников, чтобы сгруппировать их по должностям    
        fio.append(str("Новые участники"))
        fio1 = []

        for i in range (0, len(fio)):
            if fio[i] =="NONE":
                fio1.append("NONE")
            elif (fio[i].find("Новые участники")!=-1) and (i<(len(fio)-2)):
                fio1.append(" ")
                n = i+1
                while fio[n].find("Новые участники")==-1:
                    n = n+1
                for s in range (i,n):
                    if fio[s].find("Новые участники")==-1:
                        temp = fio1[-1][:]
                        fio1[-1] = temp+"/NEXT/"+fio[s][:]

        #for i in range(0, len(prim)):
        #    print ("Долж = "+dolgnost[i]+" --- ФИО = "+fio1[i]+" --- Прим = "+prim[i])

        # итог цикла
        itog = [dolgnost, fio1, prim]
        itogitogov.append(itog)

    itoglist = []
    for k in range (0, 5):
        temp = [[], [], []]

        dolg = itogitogov[k][0]
        fio = itogitogov[k][1]
        prim = itogitogov[k][2]

        for i in range (1, len(dolg)+1):
            dolg[i-1] = dolg[i-1].replace("&nbsp;", " ").strip()
            fio[i-1] = fio[i-1][7:].replace("/NEXT/"," ")
            fio[i-1] = fio[i-1].replace("PUSTO", " ")
            fio[i-1] = fio[i-1].replace("NONE", "").strip()
            prim[i-1] = prim[i-1].replace("&nbsp;"," ").strip()

            
        itoglist.append([dolg[:], fio[:], prim[:]])
    temp = itoglist
    return itoglist
    
#=============================================
#=============================================
#=============================================
#=============================================
# Определение локали!
locale.setlocale(locale.LC_ALL, "")

global LogDir, MonitLogDir, LogQueue, MyDate, threadLog
LogQueue = queue.Queue()
MyDate = "18.12.2024"
MonitOpen = False

ToLog("\n\n" + "!" * 40)
ToLog("Application started")

DocDir = FindMyDir(nameDir = "Monit_SpiskiSS_Files", subDirs = ["Script_Logs", "Monitoring_Logs", "Based", "Temp"])
LogDir = DocDir + "\\Script_Logs"
MonitLogDir = DocDir + "\\Monitoring_Logs"
ClearLogs(LogDir)
ClearLogs(MonitLogDir)

threadLog = LogThread()
threadLog.setDaemon(True)
threadLog.start()

ex = wx.App()

HelloFrame()
global Frame_Osn
Frame_Osn = Main_Frame(None, DocDir = DocDir)

ex.MainLoop()





