# -*- coding: gbk -*-
"""
Created on Thu Aug 25 10:06:41 2011

@author: 01188416
"""
import os, sys
import wx
import wx.grid
import serial
import threading
import time
import wx.gizmos
import urllib, urllib2
import sqlite3, locale, types
import json
import re
reload(sys)
sys.setdefaultencoding('gbk')

pline = ""
today = time.strftime("%Y%m%d",time.localtime())
StatData = {}
VERSION = "- 2012.12.30"

class ComDev:
    def __init__(self, comconf, waittime=5):
        self.comport = comconf["port"]
        self.baudrate = comconf["baudrate"]
        self.stopbits = comconf["stopbits"]
        self.bytesize = comconf["bytesize"]
        self.parity = comconf["parity"]
        self.waittime = waittime
        self.com = None
        self.comData = {}
        self.Open()
        
    def Open(self):
        try:
            #print self.comport,self.baudrate, self.bytesize,self.parity,self.stopbits
            self.com = serial.Win32Serial(self.comport,baudrate=self.baudrate, bytesize=self.bytesize,parity=self.parity,
                                          stopbits=self.stopbits,xonxoff=0, timeout=1)
        except:
            self.com = None
            print 'Open %s fail!' % self.comport
            
    def Close(self):
        if type(self.com) != type(None):
            self.com.close()
            self.com = None
            return True
        return False

    def ReadOnly(self):
        text = ""
        try:
            while 1:
                if type(self.com) != type(None):
                    text = text + self.com.read(1)
                    n = self.com.inWaiting()
                    if n:
                        text = text + self.com.read(n)
                    if (len(text)>2) and ord(text[-1]) == 13:
                        if len(text)>15:
                            comReadData = text[-13:-1]
                            #comReadData = comReadStr.split()[1]
                            [self.comData["dx"], self.comData["dy"], self.comData["Y"]] = comReadData.split(";")
                            frame.reSetvalue(self.comData)
                            #print self.comData,time.time()
                            text = ""
                else:
                    print "Reopen The Chroma Device!"
                    self.Open()
        except:
            print 'ReadData fail!'
            self.ReadOnly()

    def SendData(self,cmdData):
        if type(self.com) != type(None):
            try:
                self.com.write(cmdData)
                return True
            except:
                print 'SendData fail!'
                self.Close()
                return False
            return False
    def IsOpen(self):
        return type(self.com) != type(None)

class DBStorage:
    def __init__(self, path):
        self.localcharset = locale.getdefaultlocale()[1]
        self.charset = 'gbk'
        self.path = path
        if type(path) == types.UnicodeType:
            self.path = path.encode(self.charset)
        self.db = sqlite3.connect(self.path)
        self.version = '' 
        
    def close(self):
        self.db.close()
        self.db = None
        
    def query(self, sql, iszip=True):
        if type(sql) == types.UnicodeType:
            sql = sql.encode(self.charset, 'ignore')
        cur = self.db.cursor()
        cur.execute(sql)
        res = cur.fetchall()
        ret = []
        if res and iszip:
            des = cur.description
            names = [x[0] for x in des]
 
            for line in res:
                ret.append(dict(zip(names, line))) 
        else:
            ret = res 
        cur.close()
        return ret 

    def execute(self, sql, autocommit=True):
        self.db.execute(sql)
        if autocommit:
            self.db.commit()
            
    def execute_param(self, sql, param, autocommit=True):
        self.db.execute(sql, param)
        if autocommit:
            self.db.commit()

def Tvrecord(data):
    global pline, npline
    if pline =="":
        dlg = wx.SingleChoiceDialog(None, 
                '��ѡ���������壡', '����ѡ��',
               ['��������', '��װһ��', '��װ����', '��װ����', '��װ����'])
        if dlg.ShowModal() == wx.ID_OK:
            pline = dlg.GetStringSelection()
            npline = dlg.GetSelection()
        dlg.Destroy()
    #Write To Sqlite
    Sqldb  = "D:\\NxData.db"
    Creattable = 1
    if os.path.isfile(Sqldb):
        Creattable = 0
    db = DBStorage(Sqldb)
    if Creattable:
        db.execute("CREATE TABLE IF NOT EXISTS Tvnx (id integer PRIMARY KEY, \
                    sern varchar(24), xh varchar(20), bright varchar(50), \
                    power varchar(6), eei varchar(6), pline int(2), ptime int(10))")
    sql = "insert into Tvnx (sern,xh,bright,power,eei,pline,ptime) values (?,?,?,?,?,?,?)"
    db.execute_param(sql,(data["sern"], data["xh"], data["bright"], data["cost"], "%.3f" % data["eei"], npline, int(time.time())))
    db.close()
    #Write To CSV
    filename = "NxData-" + today + ".csv"
    Creattab = 1
    datafile  = "D:\\" + filename
    if os.path.isfile(datafile):
        Creattab = 0
    detxt = open(datafile, "a")
    if Creattab:
        detxt.write("�ͺ�,��������,��������,����,����,����,����,����,����,��Ч,����ʱ��\n")
    #detxt.write(data["sern"] + "," + data["bright"].replace(" ", ",") + "," + str(data["cost"]) + "," + str(time.strftime("%Y-%m-%d %X",time.localtime())) + "\n")
    detxt.write("%s,%s,%s,%s,%s,%.2f,%s\n" % (data["xh"], data["sern"], pline, data["bright"].replace(" ", ","),
                                              str(data["cost"]), data["eei"], str(time.strftime("%Y-%m-%d %X",time.localtime()))))
    detxt.close()

def UpRecord():
    global npline
    frame.upframe.disstr.SetValue("�ϴ���ʼ��")
    Sqldb  = "D:\\NxData.db"
    num = 0
    error = 0
    if os.path.isfile(Sqldb):
        db = DBStorage(Sqldb)
        nxdatas = db.query("select * from Tvnx")
        for nx in nxdatas:
            num += 1
            pdata = {"sern":nx["sern"], "xh":nx["xh"], "bright":nx["bright"],
                     "cost":nx["power"], "eei":nx["eei"]}
            if not nx.has_key("pline"):
                pdata["pline"] = npline
            else:
                pdata["pline"] = nx["pline"]
            frame.upframe.disstr.SetValue("�� %i ���ϴ���" % num)
            #frame.upframe.disstr.SetValue("�� %i �������ϴ���" % num)
            try:
                rs = urllib2.urlopen('http://192.168.10.10/cgi/tvdata/index.php/nx/add/%s' % nx["ptime"],
                                     urllib.urlencode(pdata))
            except:
                frame.showMessage("�ϴ��������Ժ����ԣ�")
                error = 1
                break
        db.close()
        if error == 0:
            os.remove(Sqldb)
    frame.upframe.disstr.SetValue("%i ���ϴ���ϣ�" % num)
    time.sleep(3)
    frame.upframe.OnClose()
    exit()

class PowerCost(wx.Frame):
    def __init__(self, xh = ""):
        wx.Frame.__init__(self, None, -1, "������������", size=(420, 240), pos =(300, 300))
        self.SetBackgroundColour('D4D0C8')
        self.zyh = xh
        if self.zyh != "":
            if Tvinfos[self.zyh].has_key("cost"):
                self.cost = int(Tvinfos[self.zyh]["cost"])
            else:
                self.cost = "80"
        else:
            self.cost = "80"
        menu = wx.Menu()
        Trig = menu.Append(-1, "Trig")
        self.Bind(wx.EVT_MENU, self.OnCilck, Trig)
        
        topLbl = wx.StaticText(self, -1, "������������(W):")
        topLbl.SetFont(wx.Font(26, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))

        nameLbl = wx.StaticText(self, -1, "����:")
        nameLbl.SetFont(wx.Font(36, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        self.name = wx.TextCtrl(self, -1, str(self.cost), size=(250, -1), style=wx.TE_PROCESS_ENTER)
        font = wx.Font(36, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.name.SetFont(font)

        copyBtn = wx.Button(self, -1, "  ȷ��  ")
        copyBtn.SetFont(wx.Font(18, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        self.Bind(wx.EVT_BUTTON, self.OnCilck, copyBtn)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(topLbl, 0, wx.ALL, 5)
        mainSizer.Add(wx.StaticLine(self), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        addrSizer = wx.FlexGridSizer(cols=2, hgap=5, vgap=5)
        addrSizer.AddGrowableCol(1)
        addrSizer.Add(nameLbl, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        addrSizer.Add(self.name)
        mainSizer.Add(addrSizer, 0, wx.EXPAND|wx.ALL, 10)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer.Add((20,20), 1)
        btnSizer.Add(copyBtn)
        btnSizer.Add((20,20), 1)

        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 10)
        self.SetSizer(mainSizer)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnCilck,self.name)
        acceltbl = wx.AcceleratorTable([(wx.ACCEL_NORMAL, 32, Trig.GetId())])
        self.SetAcceleratorTable(acceltbl)
    
    def OnCilck(self, evt):
        pcost = float(self.name.GetValue())
        if pcost<20:
            self.name.SetSelection(0,6)
            return 0
        Tvinfos[self.zyh]["cost"] = pcost
        frame.DealDisplay()
        self.Close()

class UploadData(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "��Ч���������ϴ�", size=(500, 240), pos =(300, 300))
        self.SetBackgroundColour('D4D0C8')
        menu = wx.Menu()
        Trig = menu.Append(-1, "Trig")
        self.Bind(wx.EVT_MENU, self.OnCilck, Trig)
        
        topLbl = wx.StaticText(self, -1, "��Ч���������ϴ�:")
        topLbl.SetFont(wx.Font(26, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))

        self.disstr = wx.TextCtrl(self, -1, "�����ʼ�ϴ�", size=(470, -1), style=wx.TE_READONLY|wx.TE_CENTER)
        font = wx.Font(36, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.disstr.SetFont(font)

        self.Btn = wx.Button(self, -1, "  ��ʼ  ")
        self.Btn.SetFont(wx.Font(18, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        self.Bind(wx.EVT_BUTTON, self.OnCilck, self.Btn)

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(topLbl, 0, wx.ALL, 5)
        mainSizer.Add(wx.StaticLine(self), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 5)
        addrSizer = wx.FlexGridSizer(cols=2, hgap=5, vgap=5)
        addrSizer.AddGrowableCol(1)
        addrSizer.Add(self.disstr)
        mainSizer.Add(addrSizer, 0, wx.EXPAND|wx.ALL, 10)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        btnSizer.Add((20,20), 1)
        btnSizer.Add(self.Btn)
        btnSizer.Add((20,20), 1)

        mainSizer.Add(btnSizer, 0, wx.EXPAND|wx.BOTTOM, 10)
        self.SetSizer(mainSizer)

    def OnClose(self):
        self.Close()
        exit()
        
    def OnCilck(self, event):
        Sqldb  = "D:\\NxData.db"
        if not os.path.isfile(Sqldb):
            frame.showMessage("δ�ҵ����������ļ���")
        else:
            self.Btn.Disable()
            UpDataThread = threading.Thread(target = UpRecord, args = (), name = 'UpDataThread')
            UpDataThread.start()

class MainWindow(wx.Frame):

    def __init__(self):
        self.startthread = 0
        self.cellwidth = 162
        self.comData = {"Y":0.0}
        self.running = 0
        self.runstep = 0
        self.doadd = 0
        self.tid = ""
        self.refreshFlag = 0
        self.numrows = 1
        self.remainrows = 10
        self.zyh = ""
        self.tvcostone = 0
        self.lightdis = 0
        self.totaonum = 0
        
        self.conf = {"gatenum":1.8, "clight":260.0, "tvcostone":0, "lightdis":1}
        
        
        wx.Frame.__init__(self, None, -1, 'XXXX��Ч��������  ' + VERSION, 
                          size=(950, 750), style=wx.DEFAULT_FRAME_STYLE)
        self.SetBackgroundColour('D4D0C8')
        self.statusbar = self.CreateStatusBar()
        self.statusbar.SetFieldsCount(4)
        self.statusbar.SetStatusWidths([-5, -3, -2, -2])
        self.statusbar.SetStatusText(" Power by liuwt123@gmail.com " + VERSION, 0)
        self.statusbar.SetStatusText("�ͺ�: ƽ������  ��С���� ƽ����Ч", 1)
        self.statusbar.SetStatusText("�������ޣ�" + str(self.conf["gatenum"]), 2)
        self.statusbar.SetStatusText("����������ֵ��260.0", 3)
        
        ID_COST = wx.NewId()
        ID_DISL = wx.NewId()
        menuBar = wx.MenuBar()
        menu = wx.Menu()
        setgate = menu.Append(-1, "��������")
        centerlight = menu.Append(-1, "��������")
        submenu = wx.Menu()
        tvcost = submenu.AppendCheckItem(ID_COST, "����ͬ��")
        dislight = submenu.AppendCheckItem(ID_DISL, "ʵʱ����")
        menu.AppendMenu(-1, "Ĭ�ϲ���", submenu)
        fexit = menu.Append(-1, "��  ��")
        menuBar.Append(menu, "��������")
        menu = wx.Menu()
        starttest = menu.Append(-1, "��ʼ����")
        nexttest = menu.Append(-1, "��������")
        restart = menu.Append(-1, "���²���(I)")
        menuBar.Append(menu, "������")
        menu = wx.Menu()
        updata = menu.Append(-1, "�����ϴ�")
        menuBar.Append(menu, "�����ϴ�")
        self.SetMenuBar(menuBar)
        menu = wx.Menu()
        sabout = menu.Append(-1, "�������")
        menuBar.Append(menu, "�������")
        menuBar.Enable(ID_COST, 0)
        menuBar.Enable(ID_DISL, 1)
        menuBar.Check(ID_DISL, 1)
        self.Bind(wx.EVT_MENU, self.OnSetGate, setgate)
        self.Bind(wx.EVT_MENU, self.OnCenterlight, centerlight)
        self.Bind(wx.EVT_MENU, self.OnExit, fexit)
        self.Bind(wx.EVT_MENU, self.OnEnter, starttest)
        self.Bind(wx.EVT_MENU, self.OnNext, nexttest)
        self.Bind(wx.EVT_MENU, self.OnRestart, restart)
        self.Bind(wx.EVT_MENU, self.OnRestart, restart)
        self.Bind(wx.EVT_MENU, self.OnTvcost, tvcost)
        self.Bind(wx.EVT_MENU, self.OnDislight, dislight)
        self.Bind(wx.EVT_MENU, self.OnUpdata, updata)
        self.Bind(wx.EVT_MENU, self.About, sabout)
        
        #panel = wx.Panel(self, -1)
        
        hbox = wx.BoxSizer(wx.VERTICAL)
        f0box = wx.FlexGridSizer(1, 3, 10, 10)
        f1box = wx.FlexGridSizer(1, 2, 10, 10)
        f2box = wx.FlexGridSizer(1, 5, 10, 30)
        
        #-----------------  �� �� �� �� �� �� ϵ ͳ  ------------------
        ltLabel = wx.StaticText(self, -1, 'X X X X �� Ч �� �� �� ��', style=wx.TE_CENTER)
        ltLabel.SetFont(wx.Font(36, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        
        #-----------------f0box-----------------------
        ldlabel = wx.StaticText(self, -1, 'ʵʱ����:', style=wx.TE_CENTER)
        ldlabel.SetFont(wx.Font(24, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        
        self.ldText = wx.StaticText(self, -1, '    OFF    ', style=wx.TE_CENTER)
        self.ldText.SetForegroundColour("blue")
        self.ldText.SetFont(wx.Font(24, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        
        self.statText = wx.StaticText(self, -1, '�ͺ�: ƽ������  ��С���� ƽ����Ч', style=wx.TE_CENTER)
        self.statText.SetFont(wx.Font(24, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        #-----------------f0box-----------------------
        
        #-----------------f1box-----------------------
        idlabel = wx.StaticText(self, -1, '����:', style=wx.TE_CENTER)
        idlabel.SetFont(wx.Font(40, wx.SWISS, wx.NORMAL, wx.BOLD, False, 'Tahoma'))
        
        self.idText = wx.TextCtrl(self, -1, "", size=(794, -1), style=wx.TE_PROCESS_ENTER)
        self.idText.SetForegroundColour("red")
        self.idText.SetBackgroundColour("blue")
        self.idText.SetMaxLength(20)
        font = wx.Font(40, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.idText.SetFont(font)
        #-----------------f1box-----------------------        
        
        #-----------------f2box-----------------------
        self.cell11 = wx.TextCtrl(self, -1, "000", size=(self.cellwidth, -1), style=wx.TE_READONLY|wx.TE_CENTER)
        self.cell12 = wx.TextCtrl(self, -1, "000", size=(self.cellwidth, -1), style=wx.TE_READONLY|wx.TE_CENTER)
        self.cell13 = wx.TextCtrl(self, -1, "000", size=(self.cellwidth, -1), style=wx.TE_READONLY|wx.TE_CENTER)
        self.cell14 = wx.TextCtrl(self, -1, "000", size=(self.cellwidth, -1), style=wx.TE_READONLY|wx.TE_CENTER)
        self.cell15 = wx.TextCtrl(self, -1, "000", size=(self.cellwidth, -1), style=wx.TE_READONLY|wx.TE_CENTER)
        
        font = wx.Font(32, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.cell11.SetFont(font)
        self.cell12.SetFont(font)
        self.cell13.SetFont(font)
        self.cell14.SetFont(font)
        self.cell15.SetFont(font)
        #-----------------f2box-----------------------
        
        self.colLabels = ["ID", "�� ��", "�� �� �� �� ֵ", "ƽ������", "����", "��Ч"]
        self.grid = wx.grid.Grid(self)
        self.grid.CreateGrid(1,6)
        self.grid.EnableEditing(False)
        self.grid.SetLabelFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
        
        for row in range(0, 6):
            self.grid.SetColLabelValue(row, self.colLabels[row])
        #self.grid.SetRowLabelValue(0, "")
        #self.grid.SetDefaultColSize(100)
        #self.grid.SetDefaultRowSize(36)
        #self.grid.SetColSize(1, 200)
        self.grid.SetDefaultCellFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD))
        self.grid.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        self.grid.SetRowSize(0, 25)
        self.grid.AutoSizeColumns(True)
        #self.grid.SetRowSize(1, 100)
        self.grid.SetColSize(0, 60)
        self.grid.SetColSize(1, 220)
        self.grid.SetColSize(2, 280)
        self.grid.SetColSize(3, 100)
        self.grid.SetColSize(4, 80)
        self.grid.SetColSize(5, 80)
        #self.grid.AutoSizeColumns(True)

        f0box.AddMany([(ldlabel), (self.ldText, 1, wx.EXPAND), (self.statText, 1, wx.EXPAND)])
        f1box.AddMany([(idlabel), (self.idText, 1, wx.EXPAND)])
        f2box.AddMany([(self.cell11), (self.cell12), (self.cell13), (self.cell14), (self.cell15)])
        hbox.Add(ltLabel, 0, wx.ALL | wx.EXPAND, 5)
        hbox.Add(f0box, 0, wx.ALL | wx.EXPAND, 5)
        hbox.Add(f1box, 0, wx.ALL | wx.EXPAND, 5)
        hbox.Add(f2box, 0, wx.ALL | wx.EXPAND, 5)
        hbox.Add(self.grid, 1, wx.ALL | wx.EXPAND, 15)
        self.SetSizer(hbox)

        self.Bind(wx.EVT_TEXT_ENTER, self.OnEnter,self.idText)
        #self.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        #wx.EVT_KEY_UP(self, self.OnKeyUp)
        #self.Bind(wx.EVT_CHAR, self.OnKeyUp)
        acceltbl = wx.AcceleratorTable([(wx.ACCEL_NORMAL, ord("O"), restart.GetId()),
                                        (wx.ACCEL_NORMAL, ord("I"), restart.GetId()),
                                        (wx.ACCEL_NORMAL, 32, starttest.GetId())])
        self.SetAcceleratorTable(acceltbl)
        #self.DataStat({"sern":"DC1JV123212322222223", 'avvalue':268})

    def OnEnter(self,event):
        if self.runstep:
            self.OnNext(event)
            return 1
        Tid = self.idText.GetValue()
        if len(Tid) != 20:
            self.showMessage("�������������ȷ�Ϻ����ԣ�")
            return 0
        if self.totaonum % 10 == 0:
            dlg = wx.MessageDialog(None, "��ȷ��������������������������?", '��ȷ�ϣ�', wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
            if dlg.ShowModal() != 5103:
                return 0
            dlg.Destroy()
        self.tid = Tid
        self.zyh = Tid[0:11]
        if not Tvinfos.has_key(self.zyh):
            try:
                AcRe = urllib2.Request('http://192.168.10.10/cgi/tvdata/index.php/tvinfo/get/%s/%i' % (Tid, int(time.time())))
                xhinfo = json.loads(urllib2.urlopen(AcRe).read())
                if len(xhinfo)>1:
                    Tvinfos[self.zyh] = {}
                    Tvinfos[self.zyh]["xh"] = xhinfo["exp_pcode"]
            except:
                try:
                    mfile = open('tvinfo.txt', 'r')
                    for line in mfile.readlines():
                        line = line.replace("\n", "")
                        nline = line.split()
                        if len(nline)>0:
                            Tvinfos[nline[0]] = {}
                            Tvinfos[nline[0]]["xh"] = nline[1]
                    mfile.close()
                except:
                    print "Local File Not Found!"
            if not Tvinfos.has_key(self.zyh):
                dlg = wx.TextEntryDialog(None, "����δ���ӣ����������ļ�δ���֣������������ͺţ�", '�����ͺ�����', 'LE32A910')
                if dlg.ShowModal() == wx.ID_OK:
                    Tvinfos[self.zyh] = {}
                    Tvinfos[self.zyh]["xh"] = dlg.GetValue()
                dlg.Destroy()
        if not Tvinfos[self.zyh].has_key("size"):
            tvsize = re.findall('^[A-Z]{0,3}([0-9]+)[A-Z]', Tvinfos[self.zyh]["xh"])
            Tvinfos[self.zyh]["size"] = tvsize[0]
        if self.startthread == 1:
            return 0
        else:
            self.startthread = 1
            self.runstep = 1
        self.cell11.SetBackgroundColour("yellow")
        self.cell11.SetValue("����")
        self.cell12.SetBackgroundColour("yellow")
        self.cell12.SetValue("����")
        self.cell13.SetBackgroundColour("yellow")
        self.cell13.SetValue("����")
        self.cell14.SetBackgroundColour("yellow")
        self.cell14.SetValue("����")
        self.cell15.SetBackgroundColour("yellow")
        self.cell15.SetValue("����")
        
    def OnNext(self, event):
        #print "Start:",self.running
        if self.runstep == 0:
            self.showMessage("����ɨ�����������ٽ������ݲɼ���")
            return 0
        if self.running:
            print "Running , Please Wait!"
            return 0
        self.running = 1
        avbright = self.GetAvdata()
        if self.runstep == 1:           
            self.cell11.SetBackgroundColour("green")
            self.cell11.SetValue(avbright)
            if self.conf["clight"] > float(avbright):
                self.showMessage("�������ȹ��ͣ����ص�ȷ�ϣ�")
        elif self.runstep == 2:
            self.cell12.SetBackgroundColour("green")
            self.cell12.SetValue(avbright)
        elif self.runstep == 3:
            self.cell13.SetBackgroundColour("green")
            self.cell13.SetValue(avbright)
        elif self.runstep == 4:
            self.cell14.SetBackgroundColour("green")
            self.cell14.SetValue(avbright)
        elif self.runstep == 5:
            self.cell15.SetBackgroundColour("green")
            self.cell15.SetValue(avbright)
            
            if self.conf["tvcostone"] == 0:
                self.SetCost(self.zyh)
            else:
                if not Tvinfos[self.zyh].has_key("cost"):
                    self.SetCost(self.zyh)
                else:
                    self.DealDisplay()
            #print self.idText.GetInsertionPoint()
        self.running = 0
        if self.runstep != 5:
            self.runstep = self.runstep + 1
        else:
            self.runstep = 0
        
    def showMessage(self, msg):
        dlg = wx.MessageDialog(None, msg, '�쳣��', wx.OK | wx.ICON_EXCLAMATION)
        result = dlg.ShowModal()
        dlg.Destroy()
    
    def OnTvcost(self, event):
        if self.conf["tvcostone"]:
            self.conf["tvcostone"] = 0
            print "Usb Different PowerCost"
        else:
            self.conf["tvcostone"] = 1
            print "Usb Same PowerCost"
    
    def OnUpdata(self, event):
        self.upframe = UploadData()
        self.upframe.Show()
        
    def OnDislight(self, event):
        if self.conf["lightdis"]:
            self.conf["lightdis"] = 0
            self.ldText.SetLabel("   OFF")
        else:
            self.conf["lightdis"] = 1
        
    def OnExit(self, event):
        self.Close()
    
    def OnRestart(self,event):
        self.running = 0
        self.runstep = 0
        self.startthread = 0
        print "Reset"
    
    def SetCost(self, xh = ""):
        PowerCost(xh).Show()
    
    def About(self, event):
        msg = "�������:\n\n"
        msg += "�ռ����ȡ��������ݸ����ͺųߴ������Чֵ��\n\n"
        msg += "���ⷴ����liuwt123@gmail.com\n\n"
        msg += "������£�http://192.168.79.188/iscsky/\n\n"
        msg += "                    http://192.168.10.10/iscsky/"
        dlg = wx.MessageDialog(None, msg, 'ʹ�ð���', wx.OK | wx.ICON_INFORMATION)
        result = dlg.ShowModal()
        dlg.Destroy()    
        
    def OnSetGate(self, event):
        dlg = wx.TextEntryDialog(None, "�趨��Ч��������ֵ��", '��Ч����ֵ�趨', '1.8')
        if dlg.ShowModal() == wx.ID_OK:
            self.conf["gatenum"] = float(dlg.GetValue())
            self.statusbar.SetStatusText("�������� : " + str(self.conf["gatenum"]), 2)
        dlg.Destroy()
    
    def OnCenterlight(self, event):
        dlg = wx.TextEntryDialog(None, "�趨�����������ƣ�", '������������', '260')
        if dlg.ShowModal() == wx.ID_OK:
            self.conf["clight"] = float(dlg.GetValue())
            self.statusbar.SetStatusText("������������ : " + str(self.conf["clight"]), 3)
        dlg.Destroy()
    
    def GetAvdata(self):
        self.refreshFlag = 0
        time.sleep(0.3)
        tinve = 0
        while not self.refreshFlag:
            if tinve > 3:
                self.showMessage("δ�ܻ�ȡ���������ݸ��£��������������Ӽ�������������")
                self.refreshFlag = 1
                tinve = 0
            time.sleep(0.25)
            tinve += 1
        try:
            nbright = float(self.comData["Y"])
        except:
            nbright = 0.0
        return str(nbright)
        
    def DealDisplay(self):
        num1 = self.cell11.GetValue()
        num2 = self.cell12.GetValue()
        num3 = self.cell13.GetValue()
        num4 = self.cell14.GetValue()
        num5 = self.cell15.GetValue()
        bright =  num1 + " " + num2 + " " + num3 + " " + num4 + " " + num5
        avvalue = (float(num1) + float(num2) + float(num3) + float(num4) + float(num5))/5
        powercost = Tvinfos[self.zyh]["cost"]
        tvm = Tvtsdata[Tvinfos[self.zyh]["size"]]
        eei = avvalue * tvm / (powercost - 10) / 1.1
        if self.doadd:
            self.grid.InsertRows(0, 1)
            self.numrows += 1
        if eei < self.conf["gatenum"]:
            for c in range(0,6):
                self.grid.SetCellBackgroundColour(0, c, "red")
        
        self.grid.SetRowSize(0, 25)
        self.grid.SetCellValue(0, 0, str(self.totaonum))
        self.grid.SetCellValue(0, 1, self.tid)
        self.grid.SetCellValue(0, 2, bright)
        self.grid.SetCellValue(0, 3, "%.2f" % avvalue)
        self.grid.SetCellValue(0, 4, str(powercost) + "W")
        self.grid.SetCellValue(0, 5, "%.2f" % eei)
        self.doadd = 1
        self.totaonum += 1
        if self.numrows >self.remainrows and self.numrows % self.remainrows == 6:
            self.grid.DeleteRows(self.remainrows, 10)
        self.startthread = 0
        Newnum = int(self.tid[-4:]) + 1
        self.idText.SetValue(self.tid[0:16] + '%04d' % Newnum)
        self.idText.SetInsertionPoint(20)
        data = {}
        data["sern"] = self.tid
        data["bright"] = bright
        data["pline"] = pline
        data["avvalue"] = avvalue
        data["cost"] = powercost
        data["eei"] = eei
        data["xh"] = Tvinfos[self.zyh]["xh"]
        Tvrecord(data)
        self.DataStat(data)

    def DataStat(self, data):
        tvid = data["sern"][0:11]
        if StatData.has_key(tvid):
            StatData[self.zyh]["TY"] += data["avvalue"]
            StatData[self.zyh]["TP"] += data["eei"]
            StatData[self.zyh]["TN"] += 1
            if StatData[self.zyh]["LMIN"] > data["avvalue"]:
                StatData[self.zyh]["LMIN"] = data["avvalue"]
        else:
            StatData[self.zyh] = {}
            StatData[self.zyh]["TY"] = data["avvalue"]
            StatData[self.zyh]["TP"] = data["eei"]
            StatData[self.zyh]["TN"] = 1
            StatData[self.zyh]["LMIN"] = data["avvalue"]
        dislist = (StatData[tvid]["TY"]/StatData[tvid]["TN"], StatData[tvid]["LMIN"], StatData[tvid]["TP"]/StatData[tvid]["TN"])
        disstr = Tvinfos[self.zyh]["xh"] + ": %.2f %.2f %.2f" % dislist
        self.statusbar.SetStatusText(disstr, 1)
        self.statText.SetLabel(disstr)
    
    def reSetvalue(self, data):
        self.refreshFlag = 1
        self.comData = data
        if self.conf["lightdis"]:
            self.ldText.SetLabel(data["Y"])
        

Tvtsdata = {"19":	0.094417, "22":0.128416, "23":0.14517699, "24":0.15285, 
            "26":0.1864, "32":0.2737, "37":0.3774, "39":0.4102, "40":0.4412,
            "42":	0.4865, "43":0.49889, "46":0.583, "47":0.608, "48":0.624985,
            "50":0.6745, "55":0.823, "58":0.7686, "65":1.1478 }
Tvinfos = {}

if __name__ == "__main__":
    app = wx.PySimpleApp()
    frame = MainWindow()
    frame.Show()
    Chromacomconf = {"port":"COM1", "baudrate":9600, "bytesize":8, "parity":"N", "stopbits":2}
    ChromaCom = ComDev(Chromacomconf)
    ChromaReadThread = threading.Thread(target = ChromaCom.ReadOnly, args = (), name = 'ChromaReadThread')
    ChromaReadThread.setDaemon(True)
    ChromaReadThread.start()
    app.MainLoop()
    ChromaCom.Close()