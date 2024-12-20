import tkinter as tk
from tkinter import *  # 导入tkinter模块的所有内容
import time
import os
import pyvisa
import win32com.client
from tkinter import messagebox
from tkinter import filedialog
import xlwings as xw


class EasyExcel:
    """A utility to make it easier to get at Excel.  Remembering
    to save the data is your problem, as is  error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('ket.Application')
        if filename:
            self.filename = filename
            print(filename)
            self.xlBook = self.xlApp.Workbooks.Open(filename)
            self.xlApp.Visible = True
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, PictureName, Range, left_offset, Top_offset, Width, Heigth):
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        cell = sht.Range(Range)
        sht.Shapes.AddPicture(PictureName, LinkToFile=False, SaveWithDocument=True, Left=cell.Left + left_offset,
                              Top=cell.Top + Top_offset,
                              Width=Width, Height=Heigth)

        # sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self):
        shts = self.xlBook.Worksheets
        shts(1).Copy(None, shts(1))


class OscMPO5series:
    def __init__(self, address):
        address = address.strip()
        address = address.rstrip()
        self.osc = rm.open_resource(address)

    def state(self, state):
        if state == 'run':
            self.osc.write('DIS:PERS:RESET')  # clear
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE RUN')
        elif state == 'single':
            self.osc.write('ACQUIRE:STOPAFTER SEQUENCE')  # 按下single
            self.osc.write('ACQUIRE:STATE 1')
        elif state == 'stop':
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE STOP')

    def measure(self, measNum, channel, type1):
        self.osc.write('MEASUREMENT:ADDNEW "MEAS%d"' % measNum)
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE1 %s' % (measNum, channel))
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, type1))  # 设置测量内容
        self.osc.write('MEASUrement:MEAS%d:DISPlaystat:ENABle ON' % measNum)

    def measure_1(self, measNum, channel, TYPE):
        self.osc.write('MEASUREMENT:ADDNEW "MEAS%d"' % measNum)
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, TYPE))
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE %s' % (measNum, channel))



    def measOff(self, measNum):
        self.osc.write('MEASU:DEL "MEAS%d"' % measNum)
        # self.osc.write('MEASUrement:ANNOTate AUTO')

    def makeDir(self, dir1):
        self.osc.write('FILESystem:MKDir "%s"' % dir1)

    def export(self, temp1, dir1):
        self.osc.write('SAV:IMAG "%s.%s"' % (dir1, temp1))

    def readfile(self, dir1):
        self.osc.write('FILESYSTEM:READFILE "%s"' % dir1)

    def persistence(self, state):
        self.osc.write('DISplay:PERSistence %s' % state)  # 关闭累积

    def cursor(self, state):
        self.osc.write('CURSOR:STATE %s' % state)  # 关闭cursor

    def hormode(self, state):
        self.osc.write('HOR:MODE %s' % state)  # 设置 Horizontal格式
        self.osc.write('HOR:MODE:%s:CONFIGure HORIZ' % state)
        self.osc.write('DISplay:WAVEView:GRIDTYPE FIXED')  # 设置 Horizontal格式
        self.osc.write('DISplay:WAVEView1:VIEWStyle OVErlay')

    def horpos(self, num):
        self.osc.write('HORIZONTAL:POSITION %d' % num)  # 水平位置

    def coupling(self, channel, state):
        self.osc.write('%s:COUP %s' % (channel, state))

    def number(self, number):
        num = 0
        while num <= number:
            time.sleep(1)
            num = self.osc.query('ACQuire:NUMAC?')
            if MSO5 == 1:
                num = num[15:]
            num = int(num)

    def record(self, num):
        num = num * 1.25
        self.osc.write('HOR:MODE:RECO %d' % num)

    def query(self, query):
        self.osc.query('%s' % query)
        return self.osc.query('%s' % query)

    def write(self, write):
        self.osc.write('%s' % write)

    def scale(self, channel, num):
        self.osc.write('%s:SCALE %.3f' % (channel, num))

    def channel(self, ch1, ch2, ch3, ch4, math1, math2):
        if math1 == 'ON':
            self.osc.write('MATH:ADDNEW "MATH1"')
        else:
            self.osc.write('MATH:DELETE "MATH1"')
        if math2 == 'ON':
            self.osc.write('MATH:ADDNEW "MATH2"')
        else:
            self.osc.write('MATH:DELETE "MATH2"')
        self.osc.write('SELECT:CH4 %s' % ch4)
        self.osc.write('SELECT:CH3 %s' % ch3)
        self.osc.write('SELECT:CH2 %s' % ch2)
        self.osc.write('SELECT:CH1 %s' % ch1)

    def label(self, channel, name, xi, y):
        self.osc.write('%s:LABel:NAMe "%s"' % (channel, name))  # 设置label
        xi_new = 348 * xi - 174
        y_new = 94 * y
        self.osc.write('%s:LABel:XPOS %.1f' % (channel, xi_new))
        self.osc.write('%s:LABel:YPOS %.1f' % (channel, y_new))
        self.osc.write('%s:LABel:FONT:BOLD OFF' % channel)
        self.osc.write('%s:LABel:FONT:ITALic OFF' % channel)
        self.osc.write('%s:LABel:FONT:SIZE 14' % channel)
        self.osc.write('%s:LABel:FONT:UNDERline OFF' % channel)

    def chanset(self, channel, pos, offset, bandwidth, scale):
        self.osc.write('%s:POS %.1f' % (channel, pos))  # 竖直位置
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:BANDWIDTH %s' % (channel, bandwidth))  # 一通道带宽设置为20MHz
        self.osc.write('%s:SCALE %.3f' % (channel, scale))  # 设置一通道的scale

    def trigger(self, mode, channel, slope, level):
        self.osc.write('TRIGGER:A:MODE %s' % mode)
        self.osc.write('TRIGGER:A:EDGE:SOURCE %s' % channel)
        self.osc.write('TRIGGER:A:EDGE:SLOPE %s' % slope)  # 设置触发频道和形式
        self.osc.write('TRIGGER:A:LEVEL:%s %.3f' % (channel, level))

    def math(self, channel, define, offset, pos, scale):
        self.osc.write('MATH:%s:DEFINE "%s"' % (channel, define))
        self.osc.write('MATH:%s:VERT:AUTOSC OFF' % channel)
        time.sleep(1)
        self.osc.write('MATH:%s:OFFSET %.1f' % (channel, offset))  # 设置offset
        time.sleep(1)
        self.osc.write('DISplay:WAVEView1:MATH:%s:VERTICAL:POSITION %.1f' % (channel, pos))  # 设置math通道的position
        time.sleep(1)
        self.osc.write('DISplay:WAVEView1:MATH:%s:VERTICAL:SCALE %.1f' % (channel, scale))
        # self.osc.write('MATH:ADDNEW "%s"' % channel)  # 开启math通道

    def readraw(self, file_path):
        data = self.osc.read_raw()
        data_temp = open(file_path, 'wb')
        data_temp.write(data)
        data_temp.close()


class El63600:
    def __init__(self, address):
        address = address.strip()
        address = address.rstrip()
        self.rm = pyvisa.ResourceManager()
        self.el = self.rm.open_resource(address)

    def mode(self, tpye):
        self.el.write('MODE %s' % tpye)

    def state(self, state):
        self.el.write('LOAD %s' % state)

    def dynamic(self, channel, load_max, time1):
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        self.el.write('MODE CCDH')  # 设置动态模式
        self.el.write('CURR:DYN:L1 %.2f' % (0.8 * load_max))  # 设置负载的上下电流值
        self.el.write('CURR:DYN:L2 %.2f' % (0.2 * load_max))
        self.el.write('CURR:DYN:T1 %.2fms' % time1)
        self.el.write('CURR:DYN:T2 %.2fms' % time1)
        self.el.write('CURR:DYN:FALL MAX')
        self.el.write('CURR:DYN:RISE MAX')
        self.el.write('CURR:DYN:REP 0')

    def static(self, channel, rise, load):
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        if ocp_spec <= 0.6:
            self.el.write('MODE CCL')
        elif ocp_spec <= 6:
            self.el.write('MODE CCM')
        else:
            self.el.write('MODE CCH')
        self.el.write('CURR:STAT:RISE %s' % rise)
        self.el.write('CURR:STAT:FALL %s' % rise)
        self.el.write('CURR:STAT:L1 %.2f' % load)

    def query(self, query):
        self.el.query('%s' % query)

    def write(self, write):
        self.el.write('%s' % write)

    def short(self, state):
        self.el.write('LOAD:SHOR %s' % state)


class OscDPO7000C:
    def __init__(self, address):
        address = address.strip()
        address = address.rstrip()
        self.osc = rm.open_resource(address)

    def state(self, state):
        if state == 'run':
            self.osc.write('DIS:PERS:RESET')  # clear
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE RUN')
        elif state == 'single':
            self.osc.write('ACQUIRE:STOPAFTER SEQUENCE')  # 按下single
            self.osc.write('ACQUIRE:STATE 1')
        elif state == 'stop':
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE STOP')
        else:
            print('状态设置失败')

    def measure(self, measNum, channel, type1):
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE1 %s' % (measNum, channel))
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, type1))  # 设置测量内容
        self.osc.write('MEASUrement:MEAS%d:STATE ON' % measNum)
        self.osc.write('MEASUrement:ANNOTation:STATE MEAS%d' % measNum)

    def measOff(self, measNum):
        self.osc.write('MEASUrement:MEAS%d:STATE OFF' % measNum)
        self.osc.write('MEASUrement:ANNOTate AUTO')

    def makeDir(self, dir1):
        self.osc.write('FILESystem:MKDir "%s"' % dir1)

    def export(self, temp1, dir1):
        self.osc.write('EXPort:FORMat %s' % temp1)
        self.osc.write('EXPORT:FILENAME "%s"' % dir1)  # 保存图片
        self.osc.write('EXPort STARt')

    def readfile(self, dir1):
        self.osc.write('FILESYSTEM:READFILE "%s"' % dir1)

    def persistence(self, state):
        self.osc.write('DISplay:PERSistence %s' % state)  # 关闭累积

    def cursor(self, state):
        self.osc.write('CURSOR:STATE %s' % state)  # 关闭cursor

    def hormode(self, state):
        self.osc.write('HOR:MODE %s' % state)  # 设置 Horizontal格式

    def horpos(self, num):
        self.osc.write('HORIZONTAL:POSITION %d' % num)  # 水平位置

    def coupling(self, channel, state):
        self.osc.write('%s:COUP %s' % (channel, state))

    def number(self, number):
        num = 0
        while num <= number:
            time.sleep(1)
            num = self.osc.query('ACQuire:NUMAC?')
            num = int(num)

    def record(self, num):
        self.osc.write('HOR:MODE:RECO %d' % num)

    def query(self, query):
        self.osc.query('%s' % query)
        return self.osc.query('%s' % query)

    def write(self, write):
        self.osc.write('%s' % write)

    def readraw(self, file_path):
        data = self.osc.read_raw()
        data_temp = open(file_path, 'wb')
        data_temp.write(data)
        data_temp.close()

    def scale(self, channel, num):
        self.osc.write('%s:SCALE %.3f' % (channel, num))

    def channel(self, ch1, ch2, ch3, ch4, math1, math2):
        self.osc.write('SELECT:MATH2 %s' % math2)
        self.osc.write('SELECT:MATH1 %s' % math1)
        self.osc.write('SELECT:CH4 %s' % ch4)
        self.osc.write('SELECT:CH3 %s' % ch3)
        self.osc.write('SELECT:CH2 %s' % ch2)
        self.osc.write('SELECT:CH1 %s' % ch1)

    def label(self, channel, name, xi, y):
        self.osc.write('%s:LABel:NAMe "%s"' % (channel, name))  # 设置label
        self.osc.write('%s:LABel:XPOS %.1f' % (channel, xi))
        self.osc.write('%s:LABel:YPOS %.1f' % (channel, y))

    def chanset(self, channel, pos, offset, bandwidth, scale):
        self.osc.write('%s:POS %.1f' % (channel, pos))  # 竖直位置
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:BANDWIDTH %s' % (channel, bandwidth))  # 一通道带宽设置为20MHz
        self.osc.write('%s:SCALE %.3f' % (channel, scale))  # 设置一通道的scale

    def trigger(self, mode, channel, slope, level):
        self.osc.write('TRIGGER:A:MODE %s' % mode)
        self.osc.write('TRIGGER:A:EDGE:SOURCE %s' % channel)
        self.osc.write('TRIGGER:A:EDGE:SLOPE %s' % slope)  # 设置触发频道和形式
        self.osc.write('TRIGGER:A:LEVEL %.2f' % level)

    def math(self, channel, define, offset, pos, scale):
        self.osc.write('%s:DEFINE "%s"' % (channel, define))
        self.osc.write('%s:VERT:AUTOSC OFF' % channel)
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:VERTICAL:POSITION %.2f' % (channel, pos))  # 设置math通道的position
        self.osc.write('%s:VERTICAL:SCALE %.2f' % (channel, scale))
        self.osc.write('SELECT:%s ON' % channel)  # 开启math通道


class OscDPO5104B:
    def __init__(self, address):
        address = address.strip()
        address = address.rstrip()
        self.osc = rm.open_resource(address)

    def state(self, state):
        if state == 'run':
            self.osc.write('DIS:PERS:RESET')  # clear
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE RUN')
        elif state == 'single':
            self.osc.write('ACQUIRE:STOPAFTER SEQUENCE')  # 按下single
            self.osc.write('ACQUIRE:STATE 1')
        elif state == 'stop':
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE STOP')
        else:
            print('状态设置失败')

    def measure(self, measNum, channel, type1):
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE1 %s' % (measNum, channel))
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, type1))  # 设置测量内容
        self.osc.write('MEASUrement:MEAS%d:STATE ON' % measNum)
        self.osc.write('MEASUrement:ANNOTation:STATE MEAS%d' % measNum)

    def measOff(self, measNum):
        self.osc.write('MEASUrement:MEAS%d:STATE OFF' % measNum)
        self.osc.write('MEASUrement:ANNOTate AUTO')

    def makeDir(self, dir1):
        self.osc.write('FILESystem:MKDir "%s"' % dir1)

    def export(self, temp1, dir1):
        self.osc.write('EXPort:FORMat %s' % temp1)
        self.osc.write('EXPORT:FILENAME "%s"' % dir1)  # 保存图片
        self.osc.write('EXPort STARt')

    def readfile(self, dir1):
        self.osc.write('FILESYSTEM:READFILE "%s"' % dir1)

    def persistence(self, state):
        self.osc.write('DISplay:PERSistence %s' % state)  # 关闭累积

    def cursor(self, state):
        self.osc.write('CURSOR:STATE %s' % state)  # 关闭cursor

    def hormode(self, state):
        self.osc.write('HOR:MODE %s' % state)  # 设置 Horizontal格式

    def horpos(self, num):
        self.osc.write('HORIZONTAL:POSITION %d' % num)  # 水平位置

    def coupling(self, channel, state):
        self.osc.write('%s:COUP %s' % (channel, state))

    def number(self, number):
        num = 0
        while num <= number:
            time.sleep(1)
            num = self.osc.query('ACQuire:NUMAC?')
            num = int(num)

    def record(self, num):
        self.osc.write('HOR:MODE:RECO %d' % num)

    def query(self, query):
        self.osc.query('%s' % query)
        return self.osc.query('%s' % query)

    def write(self, write):
        self.osc.write('%s' % write)

    def readraw(self, file_path):
        data = self.osc.read_raw()
        data_temp = open(file_path, 'wb')
        data_temp.write(data)
        data_temp.close()

    def scale(self, channel, num):
        self.osc.write('%s:SCALE %.3f' % (channel, num))

    def channel(self, ch1, ch2, ch3, ch4, math1, math2):
        self.osc.write('SELECT:MATH2 %s' % math2)
        self.osc.write('SELECT:MATH1 %s' % math1)
        self.osc.write('SELECT:CH4 %s' % ch4)
        self.osc.write('SELECT:CH3 %s' % ch3)
        self.osc.write('SELECT:CH2 %s' % ch2)
        self.osc.write('SELECT:CH1 %s' % ch1)

    def label(self, channel, name, xi, y):
        self.osc.write('%s:LABel:NAMe "%s"' % (channel, name))  # 设置label
        self.osc.write('%s:LABel:XPOS %.1f' % (channel, xi))
        self.osc.write('%s:LABel:YPOS %.1f' % (channel, y))

    def chanset(self, channel, pos, offset, bandwidth, scale):
        self.osc.write('%s:POS %.1f' % (channel, pos))  # 竖直位置
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:BANDWIDTH %s' % (channel, bandwidth))  # 一通道带宽设置为20MHz
        self.osc.write('%s:SCALE %.3f' % (channel, scale))  # 设置一通道的scale

    def trigger(self, mode, channel, slope, level):
        self.osc.write('TRIGGER:A:MODE %s' % mode)
        self.osc.write('TRIGGER:A:EDGE:SOURCE %s' % channel)
        self.osc.write('TRIGGER:A:EDGE:SLOPE %s' % slope)  # 设置触发频道和形式
        self.osc.write('TRIGGER:A:LEVEL %.2f' % level)

    def math(self, channel, define, offset, pos, scale):
        self.osc.write('%s:DEFINE "%s"' % (channel, define))
        self.osc.write('%s:VERT:AUTOSC OFF' % channel)
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:VERTICAL:POSITION %.2f' % (channel, pos))  # 设置math通道的position
        self.osc.write('%s:VERTICAL:SCALE %.2f' % (channel, scale))
        self.osc.write('SELECT:%s ON' % channel)  # 开启math通道


class El6312A:
    def __init__(self, address):
        address = address.strip()
        address = address.rstrip()
        self.rm = pyvisa.ResourceManager()
        self.el = self.rm.open_resource(address)

    def mode(self, tpye):
        self.el.write('MODE %s' % tpye)

    def state(self, state):
        self.el.write('LOAD %s' % state)

    def dynamic(self, channel, load_max, time1):
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        self.el.write('MODE CCDH')  # 设置动态模式
        self.el.write('CURR:DYN:L1 %.2f' % (0.8 * load_max))  # 设置负载的上下电流值
        self.el.write('CURR:DYN:L2 %.2f' % (0.2 * load_max))
        self.el.write('CURR:DYN:T1 %.2fms' % time1)
        self.el.write('CURR:DYN:T2 %.2fms' % time1)
        self.el.write('CURR:DYN:FALL MAX')
        self.el.write('CURR:DYN:RISE MAX')

    def static(self, channel, rise, load):
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        if ocp_spec <= 0.6:
            self.el.write('MODE CCL')
        elif ocp_spec <= 6:
            self.el.write('MODE CCM')
        else:
            self.el.write('MODE CCH')
        self.el.write('CURR:STAT:RISE %s' % rise)
        self.el.write('CURR:STAT:FALL %s' % rise)
        self.el.write('CURR:STAT:L1 %.2f' % load)

    def query(self, query):
        self.el.query('%s' % query)

    def write(self, write):
        self.el.write('%s' % write)

    def short(self, state):
        self.el.write('CURR:STAT:L1 0')
        self.el.write('LOAD %s' % state)
        time.sleep(1)
        self.el.write('LOAD:SHOR %s' % state)


def mkdir(path):
    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        print(path + ' 创建成功')
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(path + ' 目录已存在')
        return False

def mkdir(path):
    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        print(path + ' 创建成功')
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(path + ' 目录已存在')
        return False

def savepic(name):
    mkpath = '%s/POL Test Pictures/%s/%s' % (pic_path, temp, entry.get())
    file_path = r'%s/POL Test Pictures/%s/%s/%s.PNG' % (pic_path, temp, entry.get(), name)
    mkdir(mkpath)
    osc.makeDir('C:\\POL Test Pictures')
    osc.makeDir('C:\\POL Test Pictures\\%s' % temp)
    osc.makeDir('C:\\POL Test Pictures\\%s\\%s' % (temp, entry.get()))
    osc.export('PNG', 'C:\\POL Test Pictures\\%s\\%s\\%s' % (temp, entry.get(), name))  # 保存图片
    time.sleep(2)
    osc.readfile('C:/POL Test Pictures/%s/%s/%s.PNG' % (temp, entry.get(), name))
    time.sleep(3)
    osc.readraw(file_path)
    time.sleep(2)

def go():  # 处理事件，*args表示可变参数
    global dy, ld_max, freq, ocp_spec, temp, jitter, osc, el, xls, vin, file_path
    xls = EasyExcel(file_path)
    temp = xls.getCell('Test Summary', 6, 3)
    dy = xls.getCell(entry.get(), 5, 10)
    ld_max = xls.getCell(entry.get(), 28, 3)
    freq = xls.getCell(entry.get(), 160, 3) / 0.9
    ocp_spec = xls.getCell(entry.get(), 292, 3)
    vin = xls.getCell(entry.get(), 5, 11)
    print(dy)
    print(freq)
    print(ld_max)
    print(ocp_spec)
    EnValue2.set(dy)

def instrument():
    global osc, el, rm, MSO5, DPO7000, DPO5104B
    rm = pyvisa.ResourceManager()
    insadd = rm.list_resources()
    print(insadd)
    DPO7000 = 0
    DPO5104B = 0
    MSO5 = 0
    CH6310 = 0
    CH63600 = 0
    for addr in insadd:
        str0 = addr.find('GPIB')
        str01 = addr.find('USB')
        if str0 != -1 or str01 != -1:
            ins = rm.open_resource(addr)
            insinf = ins.query('*IDN?')
            insinf = insinf.upper()
            print(insinf)
            str1 = insinf.find('TEKTRONIX,DPO7')
            if str1 != -1:
                print('该仪器型号为TEKTRONIX,DPO7000系列示波器，设备连接成功')
                print('地址为' + addr)
                osc = OscDPO7000C(addr)
                DPO7000 = 1
            str2 = insinf.find('TEKTRONIX,MSO')
            if str2 != -1:
                print('该仪器型号为TEKTRONIX,MSO4/5/6系列示波器，设备连接成功')
                print('地址为' + addr)
                osc = OscMPO5series(addr)
                MSO5 = 1
            str3 = insinf.find('CHROMA,631')
            if str3 != -1:
                print('该仪器型号为Chroma,6310系列电子负载，设备连接成功')
                print('地址为' + addr)
                el = El6312A(addr)
                CH6310 = 1
            str4 = insinf.find('CHROMA,63600')
            if str4 != -1:
                print('该仪器型号为Chroma,63600系列电子负载，设备连接成功')
                print('地址为' + addr)
                el = El63600(addr)
                CH63600 = 1
            str5 = insinf.find('TEKTRONIX,DPO5')
            if str5 != -1:
                print('该仪器型号为TEKTRONIX,DPO5000系列示波器，设备连接成功')
                print('地址为' + addr)
                osc = OscDPO5104B(addr)
                DPO5104B = 1
    oscstate = DPO7000 or MSO5 or DPO5104B
    elstate = CH6310 or CH63600
    if oscstate and elstate:
        messagebox.showinfo(title='仪器连接', message='示波器和电子负载均已正确连接')
    elif oscstate:
        messagebox.showerror(title='仪器连接', message='电子负载连接错误，请检查')
    elif elstate:
        messagebox.showerror(title='仪器连接', message='示波器连接错误，请检查')
    else:
        messagebox.showerror(title='仪器连接', message='示波器和电子负载均连接错误，请检查')

def select_excel_path():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    print(file_path)
    if file_path:
        EnValue3.set(file_path)

def select_pic_path():
    global pic_path
    pic_path = filedialog.askdirectory()
    print(pic_path)
    if pic_path:
        EnValue4.set(pic_path)

def count():
    global counter, countmode
    if countmode == 'ON':
        timestr = '{:02}:{:02}'.format(*divmod(counter, 60))
        display.config(text=str(timestr))
        counter += 1
        display.after(1000, count)
    else:
        pass

def common_set():
    osc.state('stop')
    el.state('OFF')
    osc.persistence('OFF')  # 关闭累积
    osc.cursor('OFF')  # 关闭cursor
    # osc.hormode('MAN')  # 设置 Horizontal格式
    osc.coupling('CH1', 'DC')
    osc.coupling('CH2', 'DC')
    osc.coupling('CH3', 'DC')
    osc.coupling('CH4', 'DC')
    osc.channel('OFF', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
    osc.measOff(1)
    osc.measOff(2)
    osc.measOff(3)
    osc.measOff(4)
    osc.measOff(5)
    osc.measOff(6)
    osc.measOff(7)
    osc.measOff(8)
    if DPO7000 == 1:
        osc.write('HORIZONTAL:ROLL OFF')
    if DPO5104B == 1:
        osc.write('HORIZONTAL:ROLL OFF')
    osc.state('run')


def measure1():  # 仅测量CH1
    osc.measure(1, 'CH1', 'MAXIMUM')
    osc.measure(2, 'CH1', 'MINIMUM')
    osc.measure(3, 'CH1', 'RMS')
    osc.measure(4, 'CH1', 'PK2PK')
    if MSO5 == 1:
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')


def measure2():  # 仅测量CH1 4
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH1', 'RMS')
        osc.measure_1(4, 'CH1', 'PK2PK')
        osc.measure_1(5, 'CH4', 'MAXIMUM')
        osc.measure_1(6, 'CH4', 'MINIMUM')
        osc.measure_1(7, 'CH4', 'FREQUENCY')
        osc.measure_1(8, 'CH4', 'PDUTY')
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS5:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS6:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS7:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS8:DISPlaystat:ENABle OFF')
    else:
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
        osc.measure(3, 'CH1', 'RMS')
        osc.measure(4, 'CH1', 'PK2PK')
        osc.measure(5, 'CH4', 'MAXIMUM')
        osc.measure(6, 'CH4', 'MINIMUM')
        osc.measure(7, 'CH4', 'FREQUENCY')
        osc.measure(8, 'CH4', 'PDUTY')


def measure3():
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH2', 'MAXIMUM')
        osc.measure_1(4, 'CH2', 'MINIMUM')
        osc.measure_1(5, 'CH3', 'MAXIMUM')
        osc.measure_1(6, 'CH3', 'MINIMUM')
        osc.measure_1(7, 'CH4', 'MAXIMUM')
        osc.measure_1(8, 'CH4', 'MINIMUM')
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS5:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS6:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS7:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS8:DISPlaystat:ENABle OFF')
    else:
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
        osc.measure(3, 'CH2', 'MAXIMUM')
        osc.measure(4, 'CH2', 'MINIMUM')
        osc.measure(5, 'CH3', 'MAXIMUM')
        osc.measure(6, 'CH3', 'MINIMUM')
        osc.measure(7, 'CH4', 'MAXIMUM')
        osc.measure(8, 'CH4', 'MINIMUM')


def measure4_1():  # 仅测量CH1
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
    else:
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
    if MSO5 == 1:
        osc.measure_1(3, 'CH1', 'RISETIME')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
    if DPO7000 == 1:
        osc.measure(3, 'CH1', 'RISE')
    if DPO5104B == 1:
        osc.measure(3, 'CH1', 'RISE')


def measure4_2():  # 仅测量CH1
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
    else:
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
    if MSO5 == 1:
        osc.measure_1(3, 'CH1', 'FALLTIME')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
    if DPO7000 == 1:
        osc.measure(3, 'CH1', 'FALL')
    if DPO5104B == 1:
        osc.measure(3, 'CH1', 'FALL')


def measure5():
    osc.measure(1, 'CH4', 'MAXIMUM')
    osc.measure(2, 'CH4', 'MINIMUM')
    osc.measure(3, 'CH4', 'FREQUENCY')
    osc.measure(4, 'CH4', 'PDUTY')
    osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')


def measure5_1():
    osc.measure(1, 'CH4', 'MAXIMUM')
    osc.measure(2, 'CH4', 'MINIMUM')
    osc.measure(3, 'CH4', 'PWIDTH')
    osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')


def measure6():
    osc.measure(1, 'CH4', 'MAXIMUM')
    osc.measure(2, 'CH4', 'MINIMUM')
    if MSO5 == 1:
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')


def measure9():
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH3', 'MAXIMUM')
        osc.measure_1(4, 'CH3', 'MINIMUM')
        osc.measure_1(5, 'CH4', 'MAXIMUM')
        osc.measure_1(6, 'CH5', 'MINIMUM')
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS5:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS6:DISPlaystat:ENABle OFF')
    else:
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
        osc.measure(3, 'CH3', 'MAXIMUM')
        osc.measure(4, 'CH3', 'MINIMUM')
        osc.measure(5, 'CH4', 'MAXIMUM')
        osc.measure(6, 'CH4', 'MINIMUM')





def scale1():  # 按照输出电压的电压值调整横向scale
    osc.write('HORIZONTAL:MODE AUTO')
    osc.write('HORIZONTAL:MODE:SCALE 2e-6')
    osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')


def scale2():  # 按照输出电压的电压值调整横向scale
    osc.write('HORIZONTAL:MODE AUTO')
    osc.write('HORIZONTAL:MODE:SCALE 1e-1')
    osc.write('HORIZONTAL:MODE:SAMPLERATE 1e6')


def scale2_1():  # 按照输出电压的电压值调整横向scale
    osc.write('HORIZONTAL:MODE MANual')
    osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
    postime = dy * 1000000 / (vin * freq)
    if postime >= 2000:
        osc.record(100000)
    elif postime >= 1000:
        osc.record(80000)
    elif postime >= 500:
        osc.record(40000)
    elif postime >= 200:
        osc.record(20000)
    else:
        osc.record(10000)


def scale2_2():  # 按照输出电压的电压值调整横向scale
    osc.write('HORIZONTAL:MODE MANual')
    postime = dy * 1000000 / (vin * freq)
    if postime >= 2000:
        osc.record(50000)
    elif postime >= 1000:
        osc.record(40000)
    elif postime >= 500:
        osc.record(20000)
    elif postime >= 200:
        osc.record(10000)
    else:
        osc.record(5000)


def scale3():  # 按照电压pk2pk的大小自动调整纵向scale
    rpk = osc.query('MEASUrement:MEAS4:MAX?')
    if MSO5 == 1:
        rpk = rpk[26:]
    rpk = float(rpk)
    rpk_1 = rpk * 10000
    rpk_1 = int(rpk_1)
    if rpk_1 in range(0, 410):
        osc.scale('CH1', 10E-03)
    elif rpk_1 in range(410, 810):
        osc.scale('CH1', 20E-03)
    elif rpk_1 in range(810, 1210):
        osc.scale('CH1', 30E-03)
    elif rpk_1 in range(1210, 1610):
        osc.scale('CH1', 40E-03)
    elif rpk_1 in range(1610, 2010):
        osc.scale('CH1', 50E-03)
    elif rpk_1 in range(2010, 2410):
        osc.scale('CH1', 60E-03)
    elif rpk_1 in range(2410, 2810):
        osc.scale('CH1', 70E-03)
    elif rpk_1 in range(2810, 3210):
        osc.scale('CH1', 80E-03)
    elif rpk_1 in range(3210, 3610):
        osc.scale('CH1', 90E-03)
    else:
        print("Out of Output Voltage Range!")


def tl1_channel_set():
    osc.horpos(50)  # 水平位置
    osc.chanset('CH1', 0, dy, '20.0000E+06', 10E-02)
    osc.label('CH1', entry.get(), 1, 6)  # 设置label
    osc.trigger('AUTO', 'CH1', 'RISE', dy)
    if MSO5 == 1:
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')


def tl1_1_channel_set():
    osc.horpos(50)  # 水平位置
    osc.chanset('CH1', 0, dy, '20.0000E+06', 10E-02)
    osc.label('CH1', entry.get(), 1, 6)  # 设置label
    osc.trigger('AUTO', 'CH1', 'RISE', dy)


def tl2_1_channel_set():
    osc.horpos(40)  # 水平位置
    osc.chanset('CH1', 2, dy, '20.0000E+06', 10E-02)
    ldstep = int(ld_max / 3)
    osc.chanset('CH4', -4, 0, '20.0000E+06', ldstep)
    osc.label('CH1', entry.get(), 2, 4)  # 设置label
    osc.label('CH4', "Iout", 2, 10)  # 设置label
    osc.trigger('AUTO', 'CH4', 'RISE', ld_max / 2)
    if MSO5 == 1:
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')
    if infinite_off.get() == 'True':
        osc.persistence('OFF')  # 开启累积
    else:
        osc.persistence('INFPersist')  # 开启累积


def tl2_2_channel_set():
    osc.horpos(40)  # 水平位置
    osc.chanset('CH1', 2, dy, '20.0000E+06', 10E-02)
    ldstep = int(ld_max / 3)
    osc.chanset('CH4', -4, 0, '20.0000E+06', ldstep)
    osc.label('CH1', entry.get(), 1, 4)  # 设置label
    osc.label('CH4', "Iout", 2, 10)  # 设置label
    osc.trigger('AUTO', 'CH4', 'RISE', ld_max / 2)
    if infinite_off.get() == 'True':
        osc.persistence('OFF')  # 开启累积
    else:
        osc.persistence('INFPersist')  # 开启累积


def tl2_3_channel_set():
    osc.horpos(40)  # 水平位置
    osc.chanset('CH1', 2, dy, '20.0000E+06', 10E-02)
    ldstep = int(ld_max / 3)
    osc.chanset('CH4', -4, 0, '20.0000E+06', ldstep)
    osc.label('CH1', entry.get(), 1, 4)  # 设置label
    osc.label('CH4', "Iout", 2, 10)  # 设置label
    osc.trigger('AUTO', 'CH4', 'RISE', ld_max / 2)
    if infinite_off.get() == 'True':
        osc.persistence('OFF')  # 开启累积
    else:
        osc.persistence('INFPersist')  # 开启累积


def tl3_channel_set():
    osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
    if dy >= 5:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 1.5)
    elif dy >= 3.3:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 1)
    elif dy >= 2:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.7)
    elif dy >= 1.5:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.5)
    else:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.4)
    osc.chanset('CH2', -1, 0, '20.0000E+06', 1)
    osc.chanset('CH3', -3, 0, '20.0000E+06', 1)
    osc.chanset('CH4', -0, 0, '20.0000E+06', 3)
    osc.label('CH1', entry.get(), 1, 9)  # 设置label
    osc.label('CH2', "EN", 1.5, 9)  # 设置label
    osc.label('CH3', "PG", 2, 9)  # 设置label
    osc.label('CH4', "VIN", 2.5, 9)  # 设置label
    if MSO5 == 1:
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH2:STATE OFF')
        osc.write('DISplay:GLObal:CH3:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH2:STATE ON')
        osc.write('DISplay:GLObal:CH3:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')


def tl4_channel_set():
    osc.horpos(40)  # 水平位置
    if dy >= 5:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 1)
    elif dy >= 3.3:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.7)
    elif dy >= 2:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.5)
    elif dy >= 1.5:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.4)
    else:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.3)
    osc.label('CH1', entry.get(), 1, 9)  # 设置label
    if MSO5 == 1:
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')


def tl5_channel_set():
    if vin >= 10:
        osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
        osc.label('CH4', "PHASE", 2, 8)
    else:
        osc.chanset('CH4', -1, 0, '500.0000E+06', 2)
        osc.label('CH4', "PHASE", 2, 8)
    osc.trigger('NORMAL', 'CH4', 'RISE', 6)
    # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
    # osc.persistence('INFPersist')  # 开启累积
    el.write('CHAN 1')  # 选择相应的通道
    if MSO5 == 1:
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')

def tl5_jitter_set():
    global jitter
    jitter_max = osc.query('MEASUrement:MEAS3:MAX?')  # pos wid
    jitter_max = float(jitter_max)
    jitter_min = osc.query('MEASUrement:MEAS3:MINI?')  # pos wid
    jitter_min = float(jitter_min)
    jitter = jitter_max - jitter_min
    jitter = int(jitter * 1000000000)


def tl6_channel_set_1():
    osc.horpos(10)  # 水平位置
    # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
    if vin >= 10:
        osc.chanset('CH4', -2, 0, 'FULL', 3)
        # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
        osc.trigger('NORMAL', 'CH4', 'RISE', 6)
    else:
        osc.chanset('CH4', -2, 0, 'FULL', 2)
        # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
        osc.trigger('NORMAL', 'CH4', 'RISE', 3)
    # osc.label('CH3', "VIN_FET", 1.5, 2.5)
    osc.label('CH4', "HS_Vds", 1, 7)
    # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
    # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
    if infinite_off_6.get() == 'True':
        osc.persistence('OFF')  # 关闭累积
    else:
        osc.persistence('INFPersist')  # 开启累积



def tl0():
    root0 = Toplevel()
    root0.title('T-0 DMM&Scope Offset Record')  # 设置tl在宽和高
    root0.geometry('340x200')
    root0.transient(root)  # 为了区别root和tl，我们向tl中添加了一个Label
    Label(root0,
          text='测试前请校准探头，请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    Button(root0, text="开始测试", command=t00).grid(row=1, column=0, padx=5, pady=20)
    quit11 = Button(root0, text='退出测试', command=root0.destroy, activeforeground='white', activebackground='red')
    quit11.grid(row=1, column=1, padx=5, pady=20)  # 退出按钮的设计
    root0.attributes("-topmost", 1)


def tl1():
    root1 = Toplevel()
    root1.title('T-1 DC Regulation+Ripple&Noise Test')  # 设置tl在宽和高
    root1.geometry('340x200')
    root1.transient(root)  # 为了区别root和tl，我们向tl中添加了一个Label
    Label(root1, text='请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    Button(root1, text="开始测试", command=t01).grid(row=1, column=0, padx=5, pady=20)
    quit11 = Button(root1, text='退出测试', command=root1.destroy, activeforeground='white', activebackground='red')
    quit11.grid(row=1, column=1, padx=5, pady=20)  # 退出按钮的设计
    root1.attributes("-topmost", 1)


def tl2():
    global infinite_off
    root2 = Toplevel()
    root2.title('T-2 Loading Transient Response Test')  # 设置tl在宽和高
    root2.geometry('360x330')  # 为了区别root和tl，我们向tl中添加了一个Label
    group2 = LabelFrame(root2, text='单项测试', padx=5, pady=5)
    group2.grid(row=3, rowspan=2, column=0, columnspan=3, padx=50, pady=15)
    Label(root2, text='请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=3, padx=20, pady=20)
    theButton21 = Button(group2, text="运行 Low Frequency Test", command=t02_1)  # 按下按钮 执行 t02_1函数
    theButton21.grid(row=3, column=0, sticky=E + W, padx=40, pady=5)
    theButton22 = Button(group2, text="运行 Mid Frequency Test", command=t02_2)  # 按下按钮 执行 t02_2函数
    theButton22.grid(row=4, column=0, sticky=E + W, padx=40, pady=5)
    theButton23 = Button(group2, text="运行 High Frequency Test", command=t02_3)  # 按下按钮 执行 t02_3函数
    theButton23.grid(row=8, column=0, sticky=E + W, padx=40, pady=5)
    theButton24 = Button(root2, text="开始测试", command=t02_4)  # 按下按钮 执行 t02_4函数
    theButton24.grid(row=1, column=0, padx=40, pady=5)
    infinite_off = StringVar()
    infinite = Checkbutton(root2, text='关闭累积', variable=infinite_off, onvalue="True", offvalue="False",
                           state="normal")
    infinite.grid(row=2, column=1)
    infinite_off.set("False")
    quit21 = Button(root2, text='退出测试', command=root2.destroy, activeforeground='white', activebackground='red')
    quit21.grid(row=1, column=2, padx=40, pady=5)  # 退出按钮的设计
    root2.attributes("-topmost", 1)


def tl3():
    root3 = Toplevel()
    root3.title('T-3 Power Up & Down Sequence Measurement')  # 设置tl在宽和高
    root3.geometry('400x400')  # 为了区别root和tl，我们向tl中添加了一个Label
    group3 = LabelFrame(root3, text='单项测试', padx=5, pady=5)
    group3.grid(row=2, rowspan=2, column=0, columnspan=3, padx=60, pady=15)
    Label(root3, text="请使用探头1连接示波器的一通道和待测VR的输出端，使用探头2连接示波器的二通道和待测VR的EN信号，"
                      "使用探头3连接示波器的三通道和待测VR的PG信号，使用探头4连接示波器的四通道和待测VR的VIN信号，"
                      "单击“开始测试”进行测试。", wraplength=300,
          anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    theButton35 = Button(root3, text="开始测试", command=t03_5)  # 按下按钮 执行 t03_5函数
    theButton35.grid(row=1, column=0, padx=60, pady=5)
    quit31 = Button(root3, text='退出测试', command=root3.destroy, activeforeground='white', activebackground='red')
    quit31.grid(row=1, column=1, padx=50, pady=5)  # 退出按钮的设计
    theButton31 = Button(group3, text="运行 Power Up Sequence with NO Load", command=t03_1)  # 按下按钮 执行 t03_1函数
    theButton31.grid(row=1, column=0, sticky=E + W, padx=5, pady=5)
    theButton32 = Button(group3, text="运行 Power Down Sequence with NO Load", command=t03_2)  # 按下按钮 执行 t03_2函数
    theButton32.grid(row=2, column=0, sticky=E + W, padx=5, pady=5)
    theButton33 = Button(group3, text="运行 Power Up Sequence with Max Load", command=t03_3)  # 按下按钮 执行 t03_3函数
    theButton33.grid(row=3, column=0, sticky=E + W, padx=5, pady=5)
    theButton34 = Button(group3, text="运行 Power Down Sequence with Max Load", command=t03_4)  # 按下按钮 执行 t03_4函数
    theButton34.grid(row=4, column=0, sticky=E + W, padx=5, pady=5)
    root3.attributes("-topmost", 1)


def tl4():
    root4 = Toplevel()
    root4.title('T-4 OVS & UDS Sequence Measurement')  # 设置tl在宽和高
    root4.geometry('340x350')
    group4 = LabelFrame(root4, text='单项测试', padx=5, pady=5)
    group4.grid(row=2, rowspan=2, column=0, columnspan=3, padx=60, pady=15)
    Label(root4, text='请将探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    theButton41 = Button(group4, text="运行 Overshoot with NO Load", command=t04_1)  # 按下按钮 执行 t04_1函数
    theButton41.grid(row=1, column=0, sticky=E + W, padx=5, pady=5)
    theButton42 = Button(group4, text="运行 Undershoot with NO Load", command=t04_2)  # 按下按钮 执行 t04_2函数
    theButton42.grid(row=2, column=0, sticky=E + W, padx=5, pady=5)
    theButton43 = Button(group4, text="运行 Overshoot with Max Load", command=t04_3)  # 按下按钮 执行 t04_3函数
    theButton43.grid(row=3, column=0, sticky=E + W, padx=5, pady=5)
    theButton44 = Button(group4, text="运行 Undershoot with Max Load", command=t04_4)  # 按下按钮 执行 t04_4函数
    theButton44.grid(row=4, column=0, sticky=E + W, padx=5, pady=5)
    theButton45 = Button(root4, text="开始测试", command=t04_5)  # 按下按钮 执行 t04_5函数
    theButton45.grid(row=1, column=0, padx=60, pady=5)
    quit41 = Button(root4, text='退出测试', command=root4.destroy, activeforeground='white', activebackground='red')
    quit41.grid(row=1, column=1, padx=50, pady=5)  # 退出按钮的设计
    root4.attributes("-topmost", 1)


def tl5():
    root5 = Toplevel()
    root5.title('T-5 Switching Fre. & Jitter Measurement')  # 设置tl在宽和高
    root5.geometry('355x400')  # 为了区别root和tl，我们向tl中添加了一个Label
    Label(root5, text='请将探头一端连接到示波器的四通道，另一端连接到待测VR的SW信号，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    group5 = LabelFrame(root5, text='单项测试', padx=5, pady=5)
    group5.place(x=25, y=150, width=305, height=200)
    theButton52 = Button(group5, text="        运行 Switching Freq with Light Load", command=t05_1)  # 按下按钮 执行 t05_1函数
    theButton52.grid(row=1, column=0, sticky=E + W, padx=5, pady=5)
    theButton53 = Button(group5, text="    运行 Jitter with Light Load", command=t05_2)  # 按下按钮 执行 t05_2函数
    theButton53.grid(row=2, column=0, sticky=E + W, padx=5, pady=5)
    theButton54 = Button(group5, text="    运行 Switching Freq with Heavy Load", command=t05_3)  # 按下按钮 执行 t05_3函数
    theButton54.grid(row=3, column=0, sticky=E + W, padx=5, pady=5)
    theButton55 = Button(group5, text="    运行 Jitter with Heavy Load", command=t05_4)  # 按下按钮 执行 t05_4函数
    theButton55.grid(row=4, column=0, sticky=E + W, padx=5, pady=5)
    theButton56 = Button(root5, text="开始测试", command=t05_5)  # 按下按钮 执行 t05_5函数
    theButton56.place(x=50, y=105, width=60, height=30)
    quit51 = Button(root5, text='退出测试', command=root5.destroy, activeforeground='white', activebackground='red')
    quit51.place(x=225, y=105, width=60, height=30)
    root5.attributes("-topmost", 1)


def tl6():
    global infinite_off_6
    root6 = Toplevel()
    root6.title('T-6 Power MOSFET Gate/Phase Nodes Measurement')  # 设置tl在宽和高
    root6.geometry('390x410')
    group6 = LabelFrame(root6, text='单项测试', padx=5, pady=5)
    group6.grid(row=3, rowspan=2, column=0, columnspan=3, padx=20, pady=15)
    Label(root6, text='请使用探头1连接示波器的三通道和待测VR的Vin_FET信号，'
                      '使用探头2连接示波器的四通道和待测VR的SW信号，单击“开始测试”进行测试。', wraplength=310, anchor='w') \
        .grid(row=0, column=0, columnspan=3, padx=20, pady=20)
    theButton61 = Button(group6, text="运行 MOSFET Switching HS-Vds with NO Load", command=t06_1)  # 按下按钮 执行 t05_1函数
    theButton61.grid(row=1, column=0, sticky=E + W, padx=30, pady=5)
    theButton62 = Button(group6, text="运行 MOSFET Switching LS-Vds with NO Load", command=t06_3)  # 按下按钮 执行 t05_2函数
    theButton62.grid(row=2, column=0, sticky=E + W, padx=30, pady=5)
    theButton63 = Button(group6, text="运行 MOSFET Switching HS-Vds with Max Load", command=t06_2)  # 按下按钮 执行 t05_1函数
    theButton63.grid(row=3, column=0, sticky=E + W, padx=30, pady=5)
    theButton64 = Button(group6, text="运行 MOSFET Switching LS-Vds with Max Load", command=t06_4)  # 按下按钮 执行 t05_2函数
    theButton64.grid(row=4, column=0, sticky=E + W, padx=30, pady=5)
    theButton65 = Button(root6, text="开始测试", command=t06_5)  # 按下按钮 执行 t05_5函数
    theButton65.grid(row=1, column=0, padx=10, pady=5)
    infinite_off_6 = StringVar()
    infinite = Checkbutton(root6, text='关闭累积', variable=infinite_off_6,
                           onvalue="True", offvalue="False", state="normal")
    infinite.grid(row=2, column=1)
    infinite_off_6.set("False")
    quit61 = Button(root6, text='退出测试', command=root6.destroy, activeforeground='white', activebackground='red')
    quit61.grid(row=1, column=2, padx=10, pady=5)  # 退出按钮的设计
    root6.attributes("-topmost", 1)


def tl7():
    root7 = Toplevel()
    root7.title('T-7 Bode Plots Measurement')  # 设置tl在宽和高
    root7.geometry('400x300')  # 为了区别root和tl，我们向tl中添加了一个Label
    Label(root7, text='该项测试正在开发中，敬请期待...').pack()
    root7.attributes("-topmost", 1)


def tl8():
    root8 = Toplevel()
    root8.title('T-8 Efficiency Measurement')  # 设置tl在宽和高
    root8.geometry('400x300')  # 为了区别root和tl，我们向tl中添加了一个Label
    Label(root8, text='该项测试正在开发中，敬请期待...').pack()
    root8.attributes("-topmost", 1)


def tl9():
    global ocpmode
    root9 = Toplevel()
    root9.title('T-9 OCP & SCP Test')  # 设置tl在宽和高
    root9.geometry('350x420')
    group9 = LabelFrame(root9, text='单项测试', padx=5, pady=5)
    group9.grid(row=2, rowspan=2, column=0, columnspan=3, padx=20, pady=15)
    Label(root9, text='请使用探头1连接示波器的一通道和待测VR的输出端，探头3连接示波器的三通道和待测VR的PG信号输出端，使用电流探棒1连接示波器的四通道和待测'
                      'VR的输出电流线缆，单击测试项进行测试。', wraplength=300, anchor='w').grid(row=0,
                                                                                               column=0, columnspan=3,
                                                                                               padx=20, pady=20)
    Label(root9, text='OCP模式：').grid(row=1, column=0, sticky=E)
    ocpmode = IntVar()
    ocpmode.set(1)
    Radiobutton(root9, text='Latch', variable=ocpmode, value=0).grid(row=1, column=1)
    Radiobutton(root9, text='Hiccup', variable=ocpmode, value=1).grid(row=1, column=2)
    theButton91 = Button(group9, text="运行 Slow OCP Test", command=t09_1)  # 按下按钮 执行 t05_1函数
    theButton91.grid(row=1, column=0, sticky=E + W, padx=30, pady=5)
    theButton92 = Button(group9, text="运行 Fast OCP Test", command=t09_2)  # 按下按钮 执行 t05_1函数
    theButton92.grid(row=2, column=0, sticky=E + W, padx=30, pady=5)
    theButton93 = Button(group9, text="运行 SCP before Power on Test", command=t09_3)  # 按下按钮 执行 t05_1函数
    theButton93.grid(row=3, column=0, sticky=E + W, padx=30, pady=5)
    theButton94 = Button(group9, text="运行 SCP after Power on Test", command=t09_4)  # 按下按钮 执行 t05_1函数
    theButton94.grid(row=4, column=0, sticky=E + W, padx=30, pady=5)
    root9.attributes("-topmost", 1)


def tl10():
    global display, counter
    root10 = Toplevel()
    root10.title('T-10 Thermal Test')  # 设置tl在宽和高
    root10.geometry('340x200')
    root10.transient(root)
    Label(root10, text='请将电子负载通过负载线连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    Label(root10, text='测试累积时间：', anchor='w').grid(row=1, column=0, padx=5, pady=10)
    display = Label(root10, text='00:00', anchor='w')
    display.grid(row=1, column=1, padx=5, pady=10)
    button = Button(root10, text="开始测试", command=t10)
    button.grid(row=2, column=0, padx=5, pady=5)
    quit11 = Button(root10, text='停止测试', command=t10_1, activeforeground='white', activebackground='red')
    quit11.grid(row=2, column=1, padx=5, pady=5)  # 退出按钮的设计
    counter = 0
    root10.attributes("-topmost", 1)


def tl11():
    xls.save()
    xls.close()


def t00():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl1_channel_set()
        measure1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 2e-5') # 调整水平刻度ms
        osc.write('HORIZONTAL:POSITION 50') # 调整垂直位置T
        tdc = xls.getCell(entry.get(), 8, 4)
        el.static(1, 'MAX', tdc)  # 选择相应的通道
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = trigger_rms[24:]
        print(trigger_rms)
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = trigger_rms[24:]
        trigger_rms = float(trigger_rms)
        time.sleep(1)
        osc.trigger('AUTO', 'CH1', 'RISE', trigger_rms)
        time.sleep(1)
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl1_channel_set()
        measure1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-3')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e8')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 50')
        tdc = xls.getCell(entry.get(), 8, 4)
        el.static(1, 'MAX', tdc)  # 选择相应的通道
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        print(trigger_rms)
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = float(trigger_rms)
        time.sleep(1)
        osc.trigger('AUTO', 'CH1', 'RISE', trigger_rms)
        time.sleep(1)

    osc.state('run')  # 示波器开始采样
    osc.number(300)
    osc.state('stop')  # 示波器停止采样
    time.sleep(2)
    RMSwindow = messagebox.askquestion(title='程序执行完毕',
                                       message='程序已执行完毕，请确认波形是否正确，如果正确请使用六位半数字万用表测量输出端电压并在表格中填写！失败请点击否')
    if RMSwindow == 'yes':
        rms = osc.query('MEASUrement:MEAS3:MEAN?')
        if MSO5 == 1:
            rms = rms[24:]
        rms = float(rms)
        xls.setCell(entry.get(), 8, 8, rms)
        time.sleep(1)
        savepic('T0')  # 保存图片
        a0 = r'%s/POL Test Pictures/%s/%s/T0.png' % (pic_path, temp, entry.get())
        if v.get() == 1:
            xls.addPicture(entry.get(), a0, 'N1', 25, 0, 337, 212)
        else:
            xls.addPicture(entry.get(), a0, 'N1', 25, 0, 337, 212)
        xls.save()
        el.state('OFF')
    else:
        el.state('OFF')
    time.sleep(1)


def t01():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl1_channel_set()
        measure1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-5')
        osc.write('HORIZONTAL:POSITION 50')
        ldsp = 0.2 * ld_max
        i = 0
        load = 0
        block = 23
        while block <= 28:
            el.static(1, 'MAX', load)  # 选择相应的通道
            el.state('ON')
            osc.state('run')  # 示波器开始采样
            time.sleep(2)
            trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
            trigger_rms = trigger_rms[24:]
            trigger_rms = float(trigger_rms)
            osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
            time.sleep(5)
            scale3()
            time.sleep(3)
            trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
            trigger_rms = trigger_rms[24:]
            trigger_rms = float(trigger_rms)
            time.sleep(2)
            osc.trigger('AUTO', 'CH1', 'RISE', trigger_rms)
            time.sleep(1)
            osc.state('run')  # 示波器开始采样
            osc.number(300)
            osc.state('stop')  # 示波器停止采样
            time.sleep(1)
            el.state('OFF')
            time.sleep(2)
            if i == 0:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-1')  # 保存图片
                    a1_1 = r'%s/POL Test Pictures/%s/%s/T1-1.png' % (pic_path, temp, entry.get())
                    # a1_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Static Test with No Load.png'
                    if v.get() == 1:
                        xls.addPicture(entry.get(), a1_1, 'M18', 36, 10, 368, 237)
                    else:
                        xls.addPicture(entry.get(), a1_1, 'M18', 36, 10, 368, 237)
                else:
                    break
            elif i == 1:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-2')  # 保存图片
                else:
                    break
            elif i == 2:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-3')  # 保存图片
                else:
                    break
            elif i == 3:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-4')  # 保存图片
                    a1_2 = r'%s/POL Test Pictures/%s/%s/T1-4.png' % (pic_path, temp, entry.get())
                    # a1_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Static Test with Mid Load.png'
                    if v.get() == 1:
                        xls.addPicture(entry.get(), a1_2, 'U18', 36, 10, 358, 238)
                    else:
                        xls.addPicture(entry.get(), a1_2, 'U18', 36, 10, 358, 238)
                else:
                    break
            elif i == 4:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-5')  # 保存图片
                else:
                    break
            elif i == 5:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-6')  # 保存图片
                    a1_3 = r'%s/POL Test Pictures/%s/%s/T1-6.png' % (pic_path, temp, entry.get())
                    # a1_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Static Test with Heavy Load.png'
                    if v.get() == 1:
                        xls.addPicture(entry.get(), a1_3, 'AC18', 36, 10, 358, 238)
                    else:
                        xls.addPicture(entry.get(), a1_3, 'AC18', 36, 10, 358, 238)
                else:
                    break
            else:
                pass
            load = load + ldsp
            block = block + 1
            i = i + 1
            time.sleep(1)
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl1_1_channel_set()
        measure1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-5')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 50')
        ldsp = 0.2 * ld_max
        i = 0
        load = 0
        block = 23
        while block <= 28:
            el.static(1, 'MAX', load)  # 选择相应的通道
            el.state('ON')
            osc.state('run')  # 示波器开始采样
            time.sleep(2)
            trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
            trigger_rms = float(trigger_rms)
            osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
            time.sleep(5)
            scale3()
            time.sleep(3)
            trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
            trigger_rms = float(trigger_rms)
            time.sleep(2)
            osc.trigger('AUTO', 'CH1', 'RISE', trigger_rms)
            time.sleep(1)
            osc.state('run')  # 示波器开始采样
            osc.number(300)
            osc.state('stop')  # 示波器停止采样
            time.sleep(1)
            el.state('OFF')
            time.sleep(2)
            if i == 0:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-1')  # 保存图片
                    a1_1 = r'%s/POL Test Pictures/%s/%s/T1-1.png' % (pic_path, temp, entry.get())
                    # a1_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Static Test with No Load.png'
                    if v.get() == 1:
                        xls.addPicture(entry.get(), a1_1, 'M18', 36, 10, 368, 237)
                    else:
                        xls.addPicture(entry.get(), a1_1, 'M18', 36, 10, 368, 237)
                else:
                    break
            elif i == 1:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                else:
                    break
            elif i == 2:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                else:
                    break
            elif i == 3:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-2')  # 保存图片
                    a1_2 = r'%s/POL Test Pictures/%s/%s/T1-2.png' % (pic_path, temp, entry.get())
                    # a1_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Static Test with Mid Load.png'
                    if v.get() == 1:
                        xls.addPicture(entry.get(), a1_2, 'U18', 36, 10, 358, 238)
                    else:
                        xls.addPicture(entry.get(), a1_2, 'U18', 36, 10, 358, 238)
                else:
                    break
            elif i == 4:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                else:
                    break
            elif i == 5:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                if ocpwindow == 'yes':
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = float(rms)
                    xls.setCell(entry.get(), block, 4, rms)
                    time.sleep(1)
                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)
                    savepic('T1-3')  # 保存图片
                    a1_3 = r'%s/POL Test Pictures/%s/%s/T1-3.png' % (pic_path, temp, entry.get())
                    # a1_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Static Test with Heavy Load.png'
                    if v.get() == 1:
                        xls.addPicture(entry.get(), a1_3, 'AC18', 36, 10, 358, 238)
                    else:
                        xls.addPicture(entry.get(), a1_3, 'AC18', 36, 10, 358, 238)
                else:
                    break
            else:
                pass
            load = load + ldsp
            block = block + 1
            i = i + 1
            time.sleep(1)
    xls.save()
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')


def t02_1():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        tl2_1_channel_set()
        measure2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 2e-3')
        osc.write('HORIZONTAL:POSITION 55')
        el.dynamic(1, ld_max, 2.5)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = trigger_rms[24:]
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        osc.write('DISPLAY:PERSISTENCE:RESET')  # clear
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
        el.static(1, 'MAX', 0)  # 选择相应的通道
        time.sleep(1)
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        tl2_1_channel_set()
        measure2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 2e-3')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 4e7')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 55')
        el.dynamic(1, ld_max, 2.5)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        osc.write('DISPLAY:PERSISTENCE:RESET')  # clear
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
        el.static(1, 'MAX', 0)  # 选择相应的通道
        time.sleep(1)
    window = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if window == 'yes':
        ch1max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch1max = ch1max[26:]
        ch1max = float(ch1max)
        xls.setCell(entry.get(), 55, 3, ch1max)
        time.sleep(1)
        ch1min = osc.query('MEASUrement:MEAS2:MINI?')
        if MSO5 == 1:
            ch1min = ch1min[26:]
        ch1min = float(ch1min)
        xls.setCell(entry.get(), 55, 4, ch1min)
        time.sleep(1)
        savepic('T2-1')
        a2_1 = r'%s/POL Test Pictures/%s/%s/T2-1.png' % (pic_path, temp, entry.get())
        # a2_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Dynamic Test with Low Frequency.png'
        # print(a2_1)
        if v.get() == 1:
            xls.addPicture(entry.get(), a2_1, 'F40', 36, 10, 352, 250)
        else:
            xls.addPicture(entry.get(), a2_1, 'F40', 36, 10, 352, 250)


def t02_2():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        tl2_1_channel_set()
        measure2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 4e-4')
        osc.write('HORIZONTAL:POSITION 55')
        el.dynamic(1, ld_max, 0.5)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = trigger_rms[24:]
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        osc.write('DISPLAY:PERSISTENCE:RESET')  # clear
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
        el.static(1, 'MAX', 0)  # 选择相应的通道
        time.sleep(1)
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        tl2_2_channel_set()
        measure2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-4')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 3e8')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 55')
        el.dynamic(1, ld_max, 0.5)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        osc.write('DISPLAY:PERSISTENCE:RESET')  # clear
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
        el.static(1, 'MAX', 0)  # 选择相应的通道
        time.sleep(1)
    window = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if window == 'yes':
        ch1max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch1max = ch1max[26:]
        ch1max = float(ch1max)
        xls.setCell(entry.get(), 55, 15, ch1max)
        time.sleep(1)
        ch1min = osc.query('MEASUrement:MEAS2:MINI?')
        if MSO5 == 1:
            ch1min = ch1min[26:]
        ch1min = float(ch1min)
        xls.setCell(entry.get(), 55, 16, ch1min)
        time.sleep(1)
        savepic('T2-2')  # 保存图片
        a2_2 = r'%s/POL Test Pictures/%s/%s/T2-2.png' % (pic_path, temp, entry.get())
        # a2_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Dynamic Test with Mid Frequency.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a2_2, 'R40', 36, 10, 352, 250)
        else:
            xls.addPicture(entry.get(), a2_2, 'R40', 36, 10, 352, 250)


def t02_3():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        tl2_1_channel_set()
        measure2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 40e-6')
        osc.write('HORIZONTAL:POSITION 55')
        el.dynamic(1, ld_max, 0.05)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = trigger_rms[24:]
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        osc.write('DISPLAY:PERSISTENCE:RESET')  # clear
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
        el.static(1, 'MAX', 0)  # 选择相应的通道
        time.sleep(1)
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        tl2_3_channel_set()
        measure2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 40e-6')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 2e9')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 55')
        el.dynamic(1, ld_max, 0.05)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        time.sleep(5)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
        trigger_rms = float(trigger_rms)
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
        time.sleep(3)
        scale3()
        time.sleep(2)
        osc.write('DISPLAY:PERSISTENCE:RESET')  # clear
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
        el.static(1, 'MAX', 0)  # 选择相应的通道
        time.sleep(1)
    window = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if window == 'yes':
        ch1max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch1max = ch1max[26:]
        ch1max = float(ch1max)
        xls.setCell(entry.get(), 55, 27, ch1max)
        time.sleep(1)
        ch1min = osc.query('MEASUrement:MEAS2:MINI?')
        if MSO5 == 1:
            ch1min = ch1min[26:]
        ch1min = float(ch1min)
        xls.setCell(entry.get(), 55, 28, ch1min)
        time.sleep(1)
        savepic('T2-3')  # 保存图片
        a2_3 = r'%s/POL Test Pictures/%s/%s/T2-3.png' % (pic_path, temp, entry.get())
        # a2_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Dynamic Test with High Frequency.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a2_3, 'AD40', 36, 10, 352, 250)
        else:
            xls.addPicture(entry.get(), a2_3, 'AD40', 36, 10, 352, 250)


def t02_4():
    t02_1()
    time.sleep(1)
    t02_2()
    time.sleep(1)
    t02_3()
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')


def t03_1():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 4e-2')
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH3', 'RISE', tri_level)
        osc.write('HORIZONTAL:POSITION 60')
        osc.state('single')
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.horpos(40)  # 水平位置
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e7')  # 设置波形采样频率和scale
        tri_level = float(2)
        osc.trigger('NORMAL', 'CH3', 'RISE', tri_level)
        osc.write('HORIZONTAL:POSITION 50')
        osc.state('single')
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路上电确认',
                                     message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
    if powerup == 'yes':
        time.sleep(2)
        savepic('T3-1')  # 保存图片
        a3_1 = r'%s/POL Test Pictures/%s/%s/T3-1.png' % (pic_path, temp, entry.get())
        # a3_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Power Up Sequence with NO Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a3_1, 'F67', 36, 10, 362, 220)
        else:
            xls.addPicture(entry.get(), a3_1, 'F67', 36, 10, 362, 220)
        return 1
    else:
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        return 0


def t03_2():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 4e-1')
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH3', 'FALL', tri_level)
        osc.write('HORIZONTAL:POSITION 35')
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.horpos(25)  # 水平位置
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 2.5e6')  # 设置波形采样频率和scale
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH3', 'FALL', tri_level)
        osc.write('HORIZONTAL:POSITION 35')
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路下电确认',
                                     message='请进行电路下电，下电成功请点击是确认存图，失败请点击否')
    if powerup == 'yes':
        time.sleep(2)
        savepic('T3-2')  # 保存图片
        a3_2 = r'%s/POL Test Pictures/%s/%s/T3-2.png' % (pic_path, temp, entry.get())
        # a3_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Power Down Sequence with NO Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a3_2, 'F86', 36, 10, 362, 220)
        else:
            xls.addPicture(entry.get(), a3_2, 'F86', 36, 10, 362, 220)
        return 1
    else:
        messagebox.showerror(title='错误', message='电路下电失败，请退出重试')
        return 0


def t03_3():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 4e-2')
        el.static(1, 'MIN', ld_max)
        el.write('CONF:LVP ON')
        el.state('ON')
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH3', 'RISE', tri_level)
        osc.write('HORIZONTAL:POSITION 60')
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.horpos(40)  # 水平位置左移
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e7')  # 设置波形采样频率和scale
        el.static(1, 'MIN', ld_max)
        el.write('CONF:LVP ON')
        el.state('ON')
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH1', 'RISE', tri_level)
        osc.write('HORIZONTAL:POSITION 50')
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路上电确认',
                                     message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
    if powerup == 'yes':
        el.state('OFF')
        el.write('CONF:LVP OFF')
        time.sleep(2)
        savepic('T3-3')  # 保存图片
        a3_3 = r'%s/POL Test Pictures/%s/%s/T3-3.png' % (pic_path, temp, entry.get())
        # a3_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Power Up Sequence with Max Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a3_3, 'R67', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a3_3, 'R67', 36, 10, 349, 223)
        return 1
    else:
        el.state('OFF')
        el.write('CONF:LVP OFF')
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        return 0


def t03_4():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 4e-1')
        el.static(1, 'MIN', ld_max)
        el.write('CONF:LVP OFF')
        el.state('ON')
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH3', 'FALL', tri_level)
        osc.write('HORIZONTAL:POSITION 35')
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.horpos(25)  # 水平位置左移
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 2.5e6')  # 设置波形采样频率和scale
        el.static(1, 'MIN', ld_max)
        el.state('ON')
        tri_level = float(2.48)
        osc.trigger('NORMAL', 'CH3', 'FALL', tri_level)
        osc.write('HORIZONTAL:POSITION 35')
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路下电确认',
                                     message='请进行电路下电，下电成功请点击是确认存图，失败请点击否')
    if powerup == 'yes':
        el.state('OFF')
        time.sleep(2)
        savepic('T3-4')  # 保存图片
        a3_4 = r'%s/POL Test Pictures/%s/%s/T3-4.png' % (pic_path, temp, entry.get())
        # a3_4 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Power Down Sequence with Max Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a3_4, 'R86', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a3_4, 'R86', 36, 10, 349, 223)
        return 1
    else:
        el.state('OFF')
        messagebox.showerror(title='错误', message='电路下电失败，请退出重试')
        return 0


def t03_5():
    if t03_1() == 0:
        return 0
    time.sleep(1)
    if t03_2() == 0:
        return 0
    time.sleep(1)
    if t03_3() == 0:
        return 0
    time.sleep(1)
    if t03_4() == 0:
        return 0
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')


def t04_1():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')
        tl4_channel_set()
        measure4_1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-3')
        osc.write('HORIZONTAL:POSITION 50')
        tri_level = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'RISE', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl4_channel_set()
        measure4_1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-3')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e8')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 50')
        tri_level = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'RISE', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路上电确认', message='请进行电路上电，上电成功请点击是，失败请点击否')
    if powerup == 'yes':
        time.sleep(2)
        savepic('T4-1')  # 保存图片
        a4_1 = r'%s/POL Test Pictures/%s/%s/T4-1.png' % (pic_path, temp, entry.get())
        # a4_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Overshoot with NO Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a4_1, 'F108', 36, 10, 361, 223)
        else:
            xls.addPicture(entry.get(), a4_1, 'F108', 36, 10, 361, 223)
        ch1max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch1max = ch1max[26:]
        ch1max = float(ch1max)
        xls.setCell(entry.get(), 111, 3, ch1max)
        return 1
    else:
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        return 0


def t04_2():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')
        tl4_channel_set()
        osc.write('CH1:OFFSET 0')  # 设置offset
        measure4_2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e6')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 30')
        tri_level = float(0.2 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl4_channel_set()
        osc.write('CH1:OFFSET 0')  # 设置offset
        measure4_2()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-1')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e6')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 30')
        tri_level = float(0.2 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路下电确认', message='请进行电路下电，下电成功请点击是，失败请点击否')
    if powerup == 'yes':
        time.sleep(2)
        savepic('T4-2')  # 保存图片
        a4_2 = r'%s/POL Test Pictures/%s/%s/T4-2.png' % (pic_path, temp, entry.get())
        # a4_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Undershoot with NO Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a4_2, 'F127', 36, 10, 363, 223)
        else:
            xls.addPicture(entry.get(), a4_2, 'F127', 36, 10, 363, 223)
        ch1min = osc.query('MEASUrement:MEAS2:MINI?')
        if MSO5 == 1:
            ch1min = ch1min[26:]
        ch1min = float(ch1min)
        xls.setCell(entry.get(), 130, 3, ch1min)
        return 1
    else:
        messagebox.showerror(title='错误', message='电路下电失败，请退出重试')
        return 0


def t04_3():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl4_channel_set()
        measure4_1()
        el.static(1, 'MIN', ld_max)
        el.write('CONF:LVP ON')
        el.state('ON')
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-3')
        osc.write('HORIZONTAL:POSITION 50')
        tri_level = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'RISE', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl4_channel_set()
        measure4_1()
        el.static(1, 'MIN', ld_max)
        el.write('CONF:LVP ON')
        el.state('ON')
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-3')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e8')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 50')
        tri_level = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'RISE', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路上电确认', message='请进行电路上电，上电成功请点击是，失败请点击否')
    if powerup == 'yes':
        el.write('LOAD OFF')
        el.write('CONF:LVP OFF')
        time.sleep(2)
        savepic('T4-3')  # 保存图片
        a4_3 = r'%s/POL Test Pictures/%s/%s/T4-3.png' % (pic_path, temp, entry.get())
        # a4_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Overshoot with Max Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a4_3, 'R108', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a4_3, 'R108', 36, 10, 349, 223)
        ch1max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch1max = ch1max[26:]
        ch1max = float(ch1max)
        xls.setCell(entry.get(), 111, 15, ch1max)
        return 1
    else:
        el.write('LOAD OFF')
        el.write('CONF:LVP OFF')
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        return 0


def t04_4():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl4_channel_set()
        osc.write('CH1:OFFSET 0')  # 设置offset
        measure4_2()
        el.static(1, 'MIN', ld_max)
        el.write('CONF:LVP OFF')
        el.state('ON')
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-3')
        osc.write('HORIZONTAL:POSITION 40')
        tri_level = float(0.3 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    else:
        common_set()
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')
        tl4_channel_set()
        osc.write('CH1:OFFSET 0')  # 设置offset
        measure4_2()
        el.static(1, 'MIN', ld_max)
        el.state('ON')
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-4')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 2e8')  # 设置波形采样频率和scale
        osc.write('HORIZONTAL:POSITION 40')
        tri_level = float(0.3 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', tri_level)
        osc.state('single')  # 设置单步触发
        time.sleep(2)  # 设置延时为主板上下电作准备
    powerup = messagebox.askquestion(title='电路下电确认', message='请进行电路下电，下电成功请点击是，失败请点击否')
    if powerup == 'yes':
        el.write('LOAD OFF')
        time.sleep(1)
        savepic('T4-4')  # 保存图片
        a4_4 = r'%s/POL Test Pictures/%s/%s/T4-4.png' % (pic_path, temp, entry.get())
        # a4_4 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Undershoot with Max Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a4_4, 'R127', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a4_4, 'R127', 36, 10, 349, 223)
        ch1min = osc.query('MEASUrement:MEAS2:MINI?')
        if MSO5 == 1:
            ch1min = ch1min[26:]
        ch1min = float(ch1min)
        xls.setCell(entry.get(), 130, 15, ch1min)
        return 1
    else:
        el.write('LOAD OFF')
        messagebox.showerror(title='错误', message='电路下电失败，请退出重试')
        return 0


def t04_5():
    if t04_1() == 0:
        return 0
    time.sleep(1)
    if t04_2() == 0:
        return 0
    time.sleep(1)
    if t04_3() == 0:
        return 0
    time.sleep(1)
    if t04_4() == 0:
        return 0
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')


def t05_1():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        tl5_channel_set()
        measure5()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 2e-6')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        # load = 0.2 * ld_max
        # el.static(1, 'MAX', load)
        # el.state('ON')
        el.state('OFF')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    else:
        common_set()
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        tl5_channel_set()
        measure5()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 2e-6')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        # load = 0.2 * ld_max
        # el.static(1, 'MAX', load)
        # el.state('ON')
        el.state('OFF')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if powerup == 'yes':
        # el.state('OFF')
        ch4freq = osc.query('MEASUrement:MEAS3:MEAN?')
        if MSO5 == 1:
            ch4freq = ch4freq[24:]
        ch4freq = float(ch4freq)
        ch4freq = ch4freq / 1000
        xls.setCell(entry.get(), 152, 3, ch4freq)
        time.sleep(1)
        savepic('T5-1')
        a5_1 = r'%s/POL Test Pictures/%s/%s/T5-1.png' % (pic_path, temp, entry.get())
        # a5_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Switching Freq with Light Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a5_1, 'F149', 36, 10, 365, 225)
        else:
            xls.addPicture(entry.get(), a5_1, 'F149', 36, 10, 365, 225)


def t05_2():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')
        tl5_channel_set()
        measure5_1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(10)  # 水平位置
        osc.persistence('INFPersist')  # 开启累积
        scale2_1()
        # lload = 0.2 * ld_max
        # el.static(1, 'MAX', lload)
        # el.state('ON')
        el.state('OFF')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        osc.write('DISplay:WAVEView1:CURSor:CURSOR1:STATE ON')
        osc.write('DISplay:WAVEView1:CURSor:CURSOR1:ASOUrce CH4')
        osc.write('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:FUNCTION VBArs')
        # osc.write('CURSOR:LINESTYLE DASHed')
    else:
        common_set()
        tl5_channel_set()
        measure5_1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(10)  # 水平位置
        osc.persistence('INFPersist')  # 开启累积
        scale2_1()
        # lload = 0.2 * ld_max
        # el.static(1, 'MAX', lload)
        el.state('OFF')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        osc.write('CURSOR:STATE ON')
        osc.write('CURSOR:SOURCE 4')
        osc.write('CURSOR:FUNCTION VBArs')
        osc.write('CURSOR:LINESTYLE DASHed')
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请手动设置Cursors，请点击是确认波形是否正确！')
    if powerup == 'yes':
        # el.state('OFF')
        if MSO5 == 1:
            b = osc.query('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:AXPOSITION?')
            b = b[51:]
            print(b)
            b = float(b) * 1000000000
            print(b)
            c = osc.query('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:BXPOSITION?')
            c = c[51:]
            print(c)
            c = float(c) * 1000000000
            print(c)
            xls.setCell(entry.get(), 173, 3, b)
            xls.setCell(entry.get(), 173, 4, c)
            time.sleep(1)
            poswidth = osc.query('MEASUrement:MEAS3:MEAN?')
            poswidth = poswidth[24:]
            poswidth = float(poswidth)
            poswidth = poswidth * 1000000000
        else:
            text = osc.query('CURSOR:VBARS?')
            print(text)
            split_text = text.split(";")
            if len(split_text) == 3:
                a, b, c = split_text
                b = float(b.strip()) * 1000000000
                c = float(c.strip()) * 1000000000

                print("a:", a.strip())
                print("b:", b)
                print("c:", c)
            else:
                print("分割后的项数量不是三个")
            xls.setCell(entry.get(), 173, 3, b)
            xls.setCell(entry.get(), 173, 4, c)
            time.sleep(1)
            poswidth = float(osc.query('MEASUrement:MEAS3:MEAN?')) * 1000000000
        xls.setCell(entry.get(), 175, 3, poswidth)
        savepic('T5-2')
        a5_2 = r'%s/POL Test Pictures/%s/%s/T5-2.png' % (pic_path, temp, entry.get())
        # a5_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Jitter with Light Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a5_2, 'F168', 36, 10, 365, 225)
        else:
            xls.addPicture(entry.get(), a5_2, 'F168', 36, 10, 365, 225)


def t05_3():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        tl5_channel_set()
        measure5()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        scale1()
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    else:
        common_set()
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        tl5_channel_set()
        measure5()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        scale1()
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if powerup == 'yes':
        el.write('LOAD OFF')
        ch4freq = osc.query('MEASUrement:MEAS3:MEAN?')
        if MSO5 == 1:
            ch4freq = ch4freq[24:]
        ch4freq = float(ch4freq)
        ch4freq = ch4freq / 1000
        xls.setCell(entry.get(), 152, 15, ch4freq)
        time.sleep(1)
        savepic('T5-3')
        a5_3 = r'%s/POL Test Pictures/%s/%s/T5-3.png' % (pic_path, temp, entry.get())
        # a5_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Switching Freq with Heavy Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a5_3, 'R149', 36, 10, 348, 222)
        else:
            xls.addPicture(entry.get(), a5_3, 'R149', 36, 10, 348, 222)


def t05_4():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')
        tl5_channel_set()
        measure5_1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(10)  # 水平位置
        osc.persistence('INFPersist')  # 开启累积
        scale2_1()
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        osc.write('DISplay:WAVEView1:CURSor:CURSOR1:STATE ON')
        osc.write('DISplay:WAVEView1:CURSor:CURSOR1:ASOUrce CH4')
        osc.write('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:FUNCTION VBArs')
        # osc.write('CURSOR:LINESTYLE DASHed')
    else:
        common_set()
        tl5_channel_set()
        measure5_1()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(10)  # 水平位置
        osc.persistence('INFPersist')  # 开启累积
        scale2_1()
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.write('LOAD OFF')
        osc.write('CURSOR:SOURCE 4')
        osc.write('CURSOR:FUNCTION VBArs')
        osc.write('CURSOR:LINESTYLE DASHed')
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请手动设置Cursors，请点击是确认波形是否正确！')
    if powerup == 'yes':
        el.state('OFF')
        if MSO5 == 1:
            b = osc.query('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:AXPOSITION?')
            b = b[51:]
            print(b)
            b = float(b) * 1000000000
            print(b)
            c = osc.query('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:BXPOSITION?')
            c = c[51:]
            print(c)
            c = float(c) * 1000000000
            print(c)
            xls.setCell(entry.get(), 173, 15, b)
            xls.setCell(entry.get(), 173, 16, c)
            time.sleep(1)
            poswidth = osc.query('MEASUrement:MEAS3:MEAN?')
            poswidth = poswidth[24:]
            poswidth = float(poswidth)
            poswidth = poswidth * 1000000000
        else:
            text = osc.query('CURSOR:VBARS?')
            print(text)
            split_text = text.split(";")
            if len(split_text) == 3:
                a, b, c = split_text
                b = float(b.strip()) * 1000000000
                c = float(c.strip()) * 1000000000

                print("a:", a.strip())
                print("b:", b)
                print("c:", c)
            else:
                print("分割后的项数量不是三个")
            xls.setCell(entry.get(), 173, 15, b)
            xls.setCell(entry.get(), 173, 16, c)
            time.sleep(1)
            poswidth = float(osc.query('MEASUrement:MEAS3:MEAN?')) * 1000000000
        xls.setCell(entry.get(), 175, 15, poswidth)
        savepic('T5-4')
        a5_4 = r'%s/POL Test Pictures/%s/%s/T5-4.png' % (pic_path, temp, entry.get())
        # a5_4 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Jitter with Heavy Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a5_4, 'R168', 36, 10, 348, 222)
        else:
            xls.addPicture(entry.get(), a5_4, 'R168', 36, 10, 348, 222)


def t05_5():
    t05_1()
    time.sleep(1)
    t05_2()
    time.sleep(1)
    t05_3()
    time.sleep(1)
    t05_4()
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')


def t06_1():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_004.SET"')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "HS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        # scale1()
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    else:
        common_set()
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "HS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        # scale1()
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if powerup == 'yes':
        ch4max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch4max = ch4max[27:]
        ch4max = float(ch4max)
        xls.setCell(entry.get(), 198, 4, ch4max)
        time.sleep(1)
        savepic('T6-1')
        a6_1 = r'%s/POL Test Pictures/%s/%s/T6-1.png' % (pic_path, temp, entry.get())
        # a6_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\MOSFET Switching HS-Vds with NO Load'
        if v.get() == 1:
            xls.addPicture(entry.get(), a6_1, 'F190', 36, 10, 363, 224)
        else:
            xls.addPicture(entry.get(), a6_1, 'F190', 36, 10, 363, 224)


def t06_2():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_004.SET"')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "HS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        # scale1()
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    else:
        common_set()
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "HS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        # scale1()
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if powerup == 'yes':
        ch4max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch4max = ch4max[27:]
        ch4max = float(ch4max)
        xls.setCell(entry.get(), 198, 16, ch4max)
        savepic('T6-2')
        a6_2 = r'%s/POL Test Pictures/%s/%s/T6-2.png' % (pic_path, temp, entry.get())
        # a6_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\MOSFET Switching with LS-Vds with NO Load'
        if v.get() == 1:
            xls.addPicture(entry.get(), a6_2, 'R190', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a6_2, 'R190', 36, 10, 349, 223)
    el.state('OFF')


def t06_3():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_004.SET"')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "LS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    else:
        common_set()
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "LS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if powerup == 'yes':
        ch4max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch4max = ch4max[27:]
        ch4max = float(ch4max)
        xls.setCell(entry.get(), 217, 4, ch4max)
        savepic('T6-3')
        a6_3 = r'%s/POL Test Pictures/%s/%s/T6-3.png' % (pic_path, temp, entry.get())
        # a6_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\MOSFET Switching HS-Vds with Max Load'
        if v.get() == 1:
            xls.addPicture(entry.get(), a6_3, 'F209', 36, 10, 363, 224)
        else:
            xls.addPicture(entry.get(), a6_3, 'F209', 36, 10, 363, 224)


def t06_4():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_004.SET"')
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "LS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
    else:
        common_set()
        osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        osc.horpos(50)  # 水平位置
        # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
        if vin >= 10:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)
        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "LS_Vds", 1, 7)
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale
        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        measure6()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 5e-7')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')
        el.static(1, 'MAX', ld_max)
        el.state('ON')
        osc.state('run')  # 示波器开始采样
        osc.number(300)
        osc.state('stop')  # 示波器停止采样
        time.sleep(1)
        el.state('OFF')
    powerup = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
    if powerup == 'yes':
        ch4max = osc.query('MEASUrement:MEAS1:MAX?')
        if MSO5 == 1:
            ch4max = ch4max[27:]
        ch4max = float(ch4max)
        xls.setCell(entry.get(), 217, 16, ch4max)
        savepic('T6-4')
        a6_4 = r'%s/POL Test Pictures/%s/%s/T6-4.png' % (pic_path, temp, entry.get())
        # a6_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\MOSFET Switching LS-Vds with Max Load'
        if v.get() == 1:
            xls.addPicture(entry.get(), a6_4, 'R209', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a6_4, 'R209', 36, 10, 349, 223)
    el.state('OFF')


def t06_5():
    t06_1()
    time.sleep(1)
    t06_2()
    time.sleep(1)
    t06_3()
    time.sleep(1)
    t06_4()
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')


def t09_1():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_003.SET"')
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(25)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        scale_ld = float(ocp_spec / 1.5)
        scale_ld = round(scale_ld)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2.5, 9.5)  # 设置label
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH3:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH3:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        # tri_level1 = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH3', 'FALL', 1)
        time.sleep(5)
    else:
        common_set()
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(25)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        scale_ld = ocp_spec / 1.5
        scale_ld = round(scale_ld)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2.5, 9.5)  # 设置label
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e7')  # 设置波形采样频率和scale
        # tri_level1 = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', 1)
        time.sleep(5)
    osc.state('single')  # 设置单步触发
    i = 0
    ld_ocp = ld_max
    ldmax_step = 0.005 * ocp_spec
    while ld_ocp <= ocp_spec * 1.5:
        ld_ocp = ld_max + ldmax_step * i
        i = i + 1
        el.static(1, 'MAX', ld_ocp)
        el.state('ON')
        time.sleep(1)
        tri_state = osc.query('ACQUIRE:STATE?')
        if MSO5 == 1:
            tri_state = tri_state[15:]
        tri_state = int(tri_state)
        # el.state('OFF')
        # time.sleep(2)
        if tri_state != 1:
            ocpwindow = messagebox.askquestion(title='OCP测试确认', message='OCP测试成功，请点击是确认存图，失败请点击否')
            if ocpwindow == 'yes':
                el.state('OFF')
                ld_ocp = osc.query('MEASUrement:MEAS5:MAX?')
                if MSO5 == 1:
                    ld_ocp = ld_ocp[27:]
                ld_ocp = float(ld_ocp)
                # ld_ocp = float(ld_ocp)
                xls.setCell(entry.get(), 286, 4, ld_ocp)
                savepic('T9-1')
                a9_1 = r'%s/POL Test Pictures/%s/%s/T9-1.png' % (pic_path, temp, entry.get())
                # a9_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\OCP Test.png'
                if v.get() == 1:
                    xls.addPicture(entry.get(), a9_1, 'F281', 36, 10, 363, 223)
                else:
                    xls.addPicture(entry.get(), a9_1, 'F281', 36, 10, 363, 223)
                break
            else:
                el.state('OFF')
                break
        elif ld_ocp >= ocp_spec * 1.5:
            el.state('OFF')
            messagebox.showerror(title='错误', message='电路OCP功能未设置或设置过高，请检查并退出')


def t09_2():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_003.SET"')
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(25)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        scale_ld = ocp_spec / 1.5
        scale_ld = round(scale_ld)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2.5, 9.5)  # 设置label
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH3:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH3:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        # tri_level1 = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH3', 'FALL', 1)
        time.sleep(5)
    else:
        common_set()
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(25)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        scale_ld = ocp_spec / 1.5
        scale_ld = round(scale_ld)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2.5, 9)  # 设置label
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e7')  # 设置波形采样频率和scale
        # tri_level1 = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', 1)
        time.sleep(5)
    osc.state('single')  # 设置单步触发
    time.sleep(5)
    i = 0
    ld_ocp = ocp_spec + 0.5
    el.static(1, 'MAX', ld_ocp)
    el.state('ON')
    time.sleep(5)
    tri_state = osc.query('ACQUIRE:STATE?')
    if MSO5 == 1:
        tri_state = tri_state[15:]
    tri_state = int(tri_state)
    # el.state('OFF')
    # time.sleep(2)
    if tri_state != 1:
        ocpwindow = messagebox.askquestion(title='OCP测试确认', message='OCP测试成功，请点击是确认存图，.')
        if ocpwindow == 'yes':
            el.state('OFF')
            ld_ocp = osc.query('MEASUrement:MEAS5:MAX?')
            if MSO5 == 1:
                ld_ocp = ld_ocp[27:]
            ld_ocp = float(ld_ocp)
            xls.setCell(entry.get(), 286, 16, ld_ocp)
            savepic('T9-2')
            a9_2 = r'%s/POL Test Pictures/%s/%s/T9-2.png' % (pic_path, temp, entry.get())
            # a9_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\OCP Test.png'
            if v.get() == 1:
                xls.addPicture(entry.get(), a9_2, 'R281', 36, 10, 349, 223)
            else:
                xls.addPicture(entry.get(), a9_2, 'R281', 36, 10, 349, 223)
    elif ld_ocp > ocp_spec * 2.4:
        el.state('OFF')
        messagebox.showerror(title='错误', message='电路OCP功能未设置或设置过高，请检查并退出')

def t09_3():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_003.SET"')
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(35)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        scale_ld = round(ocp_spec * 1.5)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2, 9)  # 设置label
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH3:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH3:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.trigger('NORMAL', 'CH4', 'RISE', 5)
        osc.state('single')  # 设置单步触发
        time.sleep(5)
        el.static(1, 'MAX', 0)
        el.short('OFF')
        time.sleep(2)
        el.short('ON')
        time.sleep(5)
    else:
        common_set()
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(35)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        scale_ld = round(ocp_spec * 1.5)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2, 9)  # 设置label
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e7')  # 设置波形采样频率和scale
        osc.trigger('NORMAL', 'CH4', 'RISE', 5)
        osc.state('single')  # 设置单步触发
        el.static(1, 'MAX', 0)
        el.short('OFF')
        time.sleep(2)
        el.short('ON')
        time.sleep(5)
    powerup = messagebox.askquestion(title='电路上电确认',
                                     message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
    if powerup == 'yes':
        el.short('OFF')
        ld_short = osc.query('MEASUrement:MEAS5:MAX?')
        if MSO5 == 1:
            ld_short = ld_short[27:]
        ld_short = float(ld_short)
        xls.setCell(entry.get(), 304, 4, ld_short)
        savepic('T9-3')
        a9_3 = r'%s/POL Test Pictures/%s/%s/T9-3.png' % (pic_path, temp, entry.get())
        # a9_2 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\SCP before Power on Test.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a9_3, 'F299', 36, 10, 363, 223)
        else:
            xls.addPicture(entry.get(), a9_3, 'F299', 36, 10, 363, 223)
    else:
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        el.short('OFF')


def t09_4():
    if MSO5 == 1:
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_003.SET"')
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(40)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        scale_ld = round(ocp_spec * 1.5)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2, 9)  # 设置label
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH3:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH3:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        # tri_level1 = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', 1)
        osc.state('single')  # 设置单步触发
        time.sleep(5)
        el.static(1, 'MAX', 0)
        el.short('OFF')
        time.sleep(2)
        el.short('ON')
        time.sleep(5)
    else:
        common_set()
        if pg == 0:
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
        osc.horpos(40)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)
        scale_ld = round(ocp_spec * 1.5)
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        osc.chanset('CH3', -1, 0, '20.0000E+06', 2)
        osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)
        osc.label('CH1', entry.get(), 1, 9)  # 设置label
        osc.label('CH3', "PG", 1.5, 9)  # 设置label
        osc.label('CH4', "Iout", 2, 9)  # 设置label
        measure9()
        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 1e-2')
        osc.write('HORIZONTAL:MODE:SAMPLERATE 1e7')  # 设置波形采样频率和scale
        # tri_level1 = float(0.5 * dy)
        osc.trigger('NORMAL', 'CH1', 'FALL', 1)
        osc.state('single')  # 设置单步触发
        el.static(1, 'MAX', 0)
        el.short('OFF')
        time.sleep(2)
        el.short('ON')
        time.sleep(5)
    powerup = messagebox.askquestion(title='电路上电确认',
                                     message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
    if powerup == 'yes':
        el.short('OFF')
        ld_short = osc.query('MEASUrement:MEAS5:MAX?')
        if MSO5 == 1:
            ld_short = ld_short[27:]
        ld_short = float(ld_short)
        xls.setCell(entry.get(), 304, 16, ld_short)
        savepic('T9-4')
        a9_4 = r'%s/POL Test Pictures/%s/%s/T9-4.png' % (pic_path, temp, entry.get())
        # a9_3 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\SCP after Power on Test.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a9_4, 'R299', 36, 10, 349, 223)
        else:
            xls.addPicture(entry.get(), a9_4, 'R299', 36, 10, 349, 223)
    else:
        el.short('OFF')
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')


def t10():
    global countmode
    tdc = xls.getCell(entry.get(), 324, 3)
    el.static(1, 'MAX', tdc)
    el.state('ON')
    countmode = 'ON'
    count()


def t10_1():
    global countmode
    countmode = 'OFF'
    el.state('OFF')




global dy, ld_max, freq, ocp_spec, temp, jitter, osc, el, xls, vin, display, counter, \
    infinite_off_6, infinite_off, ocpmode, osc_addr, el_addr, MSO5, EnValue3, rm, DPO7000, countmode, file_path, pic_path
root = Tk()  # 初始框的声明
root.title('Suma Power Test')
root.resizable(False, False)
root.geometry('730x495')  # 设置初始框的大小
# root.focus_force() # Take Focus
visa_dll = 'c:/windows/system32/visa32.dll'
time.sleep(1)
image_file = PhotoImage(file='suma.png')
image = Label(root, image=image_file)
image.grid(row=0, column=0, columnspan=3, padx=30, pady=20)
Label(root, text='SheetName:').grid(row=3, column=0, sticky=E)
EnValue1 = StringVar()
EnValue2 = StringVar()
EnValue3 = StringVar()
EnValue4 = StringVar()
entry = Entry(root, show=None, width=20, textvariable=EnValue1)
entry.grid(row=3, column=1, columnspan=2)
Label(root, text='输出电压：').grid(row=4, column=0, sticky=E)
Entry(root, show=None, width=10, textvariable=EnValue2, state='readonly').grid(row=4, column=1)
Label(root, text='V').grid(row=4, column=1, sticky=E)
Entry(root, show=None, textvariable=EnValue3, state='readonly').place(x=130, y=120, width=200, height=30)
Entry(root, show=None, textvariable=EnValue4, state='readonly').place(x=130, y=160, width=200, height=30)
v = IntVar()
v.set(1)
# Label(root, text='选择DPI：').grid(row=3, column=0, sticky=E)
# Radiobutton(root, text='100%', variable=v, value=0).grid(row=3, column=1)
# Radiobutton(root, text='125%', variable=v, value=1).grid(row=3, column=2)
Label(root, text='工作模式：').grid(row=6, column=0, sticky=E)
x = IntVar()
x.set(0)
Radiobutton(root, text='COT模式', variable=x, value=0).grid(row=6, column=1)
Radiobutton(root, text='PWM模式', variable=x, value=1).grid(row=6, column=2)
Label(root, text='PG信号：').grid(row=7, column=0, sticky=E)
pg = IntVar()
pg.set(1)
Radiobutton(root, text='不输出', variable=pg, value=0).grid(row=7, column=2)
Radiobutton(root, text='输出', variable=pg, value=1).grid(row=7, column=1)
theButton = Button(root, text="选择测试报告路径", command=select_excel_path)  # 按下按钮 执行instrument函数
theButton.place(x=10, y=120, width=115, height=30)
theButton = Button(root, text="选择保存图片路径", command=select_pic_path)  # 按下按钮 执行instrument函数
theButton.place(x=10, y=160, width=115, height=30)
theButton = Button(root, text="连接仪器", command=instrument)  # 按下按钮 执行instrument函数
theButton.place(x=10, y=430, width=80, height=30)
theButton11 = Button(root, text="读取表格", command=go)  # 按下按钮 执行go函数
theButton11.place(x=120, y=430, width=80, height=30)
theButton12 = Button(root, text='保存并退出表格', command=tl11, activeforeground='white', activebackground='red')
theButton12.place(x=230, y=430, width=100, height=30)  # 退出按钮的设计
group = LabelFrame(root, text='POL测试项', padx=5, pady=5)
group.grid(row=0, rowspan=12, column=3, padx=30, pady=15)
theButton00 = Button(group, text="T-0                 DMM&Scope Offset Record              ", command=tl0)  # 按下按钮 打开tl0界面
theButton00.grid(row=1, column=3, sticky=E + W, padx=5, pady=5)
theButton01 = Button(group, text="T-1             DC Regulation+Ripple&Noise Test         ", command=tl1)  # 按下按钮 打开tl1界面
theButton01.grid(row=2, column=3, sticky=E + W, padx=5, pady=5)
theButton02 = Button(group, text="T-2             Loading Transient Response Test           ", command=tl2)  # 按下按钮 打开tl2界面
theButton02.grid(row=3, column=3, sticky=E + W, padx=5, pady=5)
theButton03 = Button(group, text="T-3    Power Up & Down Sequence Measurement     ", command=tl3)  # 按下按钮 打开tl3界面
theButton03.grid(row=4, column=3, sticky=E + W, padx=5, pady=5)
theButton04 = Button(group, text="T-4         OVS & UDS Sequence Measurement          ", command=tl4)  # 按下按钮 打开tl4界面
theButton04.grid(row=5, column=3, sticky=E + W, padx=5, pady=5)
theButton05 = Button(group, text="T-5          Switching Fre. & Jitter Measurement          ", command=tl5)  # 按下按钮 打开tl5界面
theButton05.grid(row=6, column=3, sticky=E + W, padx=5, pady=5)
theButton06 = Button(group, text="T-6 Power MOSFET Gate/Phase Nodes Measurement", command=tl6)  # 按下按钮 打开tl6界面
theButton06.grid(row=7, column=3, sticky=E + W, padx=5, pady=5)
theButton07 = Button(group, text="T-7               Bode Plots Measurement(TBD)             ", command=tl7)  # 按下按钮 打开tl7界面
theButton07.grid(row=8, column=3, sticky=E + W, padx=5, pady=5)
theButton08 = Button(group, text="T-8                 Efficiency Measurement(TBD)              ", command=tl8)  # 按下按钮 打开tl8界面
theButton08.grid(row=9, column=3, sticky=E + W, padx=5, pady=5)
theButton09 = Button(group, text="T-9                    OCP & SCP & OVP Test                   ", command=tl9)  # 按下按钮 打开tl9界面
theButton09.grid(row=10, column=3, sticky=E + W, padx=5, pady=5)
theButton10 = Button(group, text="T-10                       Thermal Test                             ", command=tl10)  # 按下按钮 打开tl10界面
theButton10.grid(row=11, column=3, sticky=E + W, padx=5, pady=5)
root.mainloop()
