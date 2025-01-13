# 项目名称
"中科硬件实验室 POL电源自动化测试"

# —————————————————————————————分割线———————————————————————————————
# 项目所需包的导入

import tkinter as tk   # 导入tkinter的全部模块, 并重命名为tk, 使用 tk.Label  tk.Button
# tkinter用于创建图形用户界面（GUI), 构建桌面应用程序
# 具有图形化界面，提供窗口、按钮、标签、文本框等控件

import time             # 提供时间相关的功能
# time.sleep(seconds)  暂停执行指定的秒数 给予反应时间

import math

import os               # 提供与操作系统交互的功能，例如文件或目录操作
# os.path.exists(path) 检查路径是否存在
# os.makedirs(path)    递归地创建目录

import pyvisa           # 用于与仪器通信，指在测试和测量应用中
# rm = pyvisa.ResourceManager()
# 创建一个 ResourceManager 实例，这个实例用于管理和访问连接到计算机的仪器
# osc = rm.open_resource('GPIB0::14::INSTR')  # 连接到指定的仪器
# response = osc.query('*IDN?')  # 发送查询命令并获取响应


import win32com.client  # 和windows应用程序交互，指 excel
# excel = win32com.client.Dispatch("Excel.Application")   # 启动Excel应用程序
# workbook = excel.Workbooks.Add()                        # 新建一个工作簿
# sheet = workbook.Sheets(1)                              # 访问第一个工作表
# sheet.Cells(1, 1).Value = 'Hello, World!'               # 在单元格中写入数据
# workbook.SaveAs('example.xlsx')                         # 保存工作簿
# excel.Application.Quit()                                # 退出Excel应用程序


from tkinter import messagebox  # 用于显示对话框消息，如警告、错误信息、确认对话框等。
from tkinter import filedialog  # 提供文件选择对话框，允许用户选择文件或目录
# messagebox.showinfo(title='仪器连接', message='示波器和电子负载均已正确连接')
# messagebox.showerror(title='仪器连接', message='电子负载连接错误，请检查')
# file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
# 弹出一个文件选择对话框，选择指定类型的excel文件保存在指定地址
# pic_path = filedialog.askdirectory()
# 显示一个目录选择对话框，让用户选择一个目录，返回用户选择的目录路径

import xlwings as xw    # 导入了 xlwings 库并将其别名为 xw# xlwings 是一个用于在 Python 中操作 Excel 的库  （好像没用上）

# —————————————————————————————分割线———————————————————————————————

#  仪器使用说明
class EasyExcel:
    """
    创建一个 Excel实例  excel = EasyExcel("C:\\Users\\Seir\\Desktop\\测试文档POL.xlsx")
    保存工作簿 excel.save(filename)
    关闭工作簿 excel.close()
    获取单元格值 excel.getCell(sheet, row, col)
    设置单元格值 excel.setCell(sheet, row, col)
    获取范围的值 excel.getRange(sheet, row1, col1, row2, col2)
    添加图片 excel.addPicture(self, sheet, PictureName, Range, left_offset, Top_offset, Width, Height)
    excel.addPicture(entry.get(), a0, 'N1', 25, 0, 337, 212)
    复制工作表 excel.cpSheet()
    """

    def __init__(self, filename=None):
        """
        初始化 参数: filename（可选）用于指定打开的 Excel 文件。如果未提供 filename，则创建一个新的空工作簿
        如果文件名存在，使用旧文件名并输出它，打开指定的工作簿并将excel应用程序设置为可见
        如果未提供filename则创建一个新工作簿，并记录新文件名
        """
        self.xlApp = win32com.client.Dispatch('excel.Application')
        # 创建一个 Excel 应用程序实例 (self.xlApp)  ket是特定的标识符
        if filename:
            self.filename = filename
            print(filename)
            self.xlBook = self.xlApp.Workbooks.Open(filename)
            self.xlApp.Visible = True
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):
        """
        保存当前的工作簿 参数: newfilename（可选）用于指定保存时的文件名。如果不提供 newfilename，则使用当前文件名保存
        如果提供了新文件名，则将工作簿保存问新文件，并更新文件名
        如果没有提供新文件名，则以旧文件名保存工作簿
        """
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        # 关闭当前工作簿并释放 Excel 应用程序对象
        self.xlBook.Close(SaveChanges=0)  # 参数为0表示不保存更改！！！
        self.xlApp.Quit()       # 先关闭再删除对象
        del self.xlApp

    def getCell(self, sheet: object, row: object, col: object) -> object:
        # 获取指定工作表中指定单元格的值。sheet=指定工作表，row=行，col=列
        sht = self.xlBook.Worksheets(sheet)  # sht是工作表的局部变量，可以访问其中的单元格、范围、图形
        sht.Activate()  # 激活工作表
        return sht.Cells(row, col).Value  # 返回指定单元格的值

    def setCell(self, sheet, row, col, value):
        # 设置指定工作表中指定单元格的值 同 getCell. 直接设置不返回值
        # 参数 value 为要设置的新值
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        # 获取指定工作表中指定范围的单元格值
        # 参数sheet为指定工作表，row1,col1为起始行列，row2,col2为终止行列
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
        # 返回指定范围内所有单元格的值

    def addPicture(self, sheet, PictureName, Range, left_offset, Top_offset, Width, Height):
        # 在指定工作表的指定位置添加图片
        # sheet为指定工作表，picturename为图片文件名，range为基准单元格范围，left_offset和 top_offset为图片的位置偏移量，width和 height为图片的宽度和高度
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        cell = sht.Range(Range)  # 获取指定范围的单元格对象
        sht.Shapes.AddPicture(PictureName, LinkToFile=False, SaveWithDocument=True, Left=cell.Left + left_offset,
                              Top=cell.Top + Top_offset,
                              Width=Width, Height=Height)
        """
        插入图片的方法 示例
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)
        picturename:图片文件路径，linktofile=false:图片不会链接到原文件，而是嵌入到文档中
        SaveWithDocument=True: 图片会随文档保存
        Left=cell.Left + left_offset: 图片左边距，相对单元格左边距加上偏移量
        Top=cell.Top + Top_offset: 图片上边距，相对单元格上边距加上偏移量
        Width=Width: 图片的宽度
        """

    def cpSheet(self):
        # 在一个工作簿中复制第一个工作表，并将新工作表插入到第一个工作表之前
        shts = self.xlBook.Worksheets  # 获取工作簿中的所有工作表
        shts(1).Copy(None, shts(1))
        """
        sht(1) 获取工作簿中的第一个工作表，
        None表示指定复制的位置为空(即在当前工作簿中进行复制)
        shts(1) 指定要复制到的位置，即原始工作表的前面
        """

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
        # self.el.write('CURR:STAT:L1 0')
        self.el.write('LOAD %s' % state)
        time.sleep(1)
        self.el.write('LOAD:SHOR %s' % state)

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

class Data_Acquisition:
    global date, DAQ973A, Data_Acquisition_id

    def __init__(self, address):
        address = address.strip()                   # 处理前后两端空白字符（包括空格和换行符）
        self.rm = pyvisa.ResourceManager()
        self.Data_Acquisition = self.rm.open_resource(address)

    def _detect_model(self):                    # 检查是否是电子负载
        model_info = self.Data_Acquisition.query('*IDN?')     # 通过发送指令查询设备的标识字符串
        if 'DAQ973A' in model_info:             # 检查响应字符串中是否包含 'EL63600'
            return 'DAQ973A'
        elif 'DAQ34970A' in model_info:
            return 'DAQ34970A'
        # elif '数据采集仪型号' in model_info:
        #     return '数据采集仪型号'
        return 'Unknown'

    def Channel_Set(self):
        self.Data_Acquisition.write('CONF:VOLT:DC 10,0.0001,(@101)')
        self.Data_Acquisition.write('CONF:VOLT:DC 100,0.0001,(@102)')
        self.Data_Acquisition.write('CONF:CURR:DC 1,0.0001,(@103)')
        self.Data_Acquisition.write('CONF:VOLT:DC 10,0.0001,(@104)')

    def Scan_Channel(self):
        self.Data_Acquisition.write('ROUT:SCAN (@101,102,103,104)')
        self.Data_Acquisition.write('INIT')

    def Read_Date(self):
        d = self.Data_Acquisition.query('FETC?')
        d = d.strip('\n')
        d = d.split(",")
        return d

class DCsource:
    def __init__(self, address):
        """
        初始化 DC Source 类，连接到指定地址
        :param address: SCPI 地址字符串
        """
        address = address.strip()
        address = address.rstrip()
        self.rm = pyvisa.ResourceManager()
        self.source = self.rm.open_resource(address)

    def mode(self, type_):
        """
        设置直流电源的模式
        :param type_: 模式类型，例如 'CV'（恒压）、'CC'（恒流）
        """
        self.write(f"CONF:{type_}")

    def output(self, state):
        """
        控制直流电源输出开关
        :param state: 'ON' 或 'OFF'
        """
        self.write(f"CONF:OUTP {state}")

    def set_voltage(self, voltage):
        """
        设置输出电压
        :param voltage: 输出电压值 (单位: V)
        """
        self.write(f"SOUR:VOLT {voltage:.3f}")

    def set_current(self, current):
        """
        设置输出电流
        :param current: 输出电流值 (单位: A)
        """
        self.write(f"SOUR:CURR {current:.3f}")

    def query_voltage(self):
        """
        查询输出电压
        :return: 当前输出电压值 (单位: V)
        """
        return float(self.query("SOUR:VOLT?"))

    def query_current(self):
        """
        查询输出电流
        :return: 当前输出电流值 (单位: A)
        """
        return float(self.query("SOUR:CURR?"))

    def set_protection(self, ovp=None, ocp=None):
        """
        设置保护参数（过压保护 OVP、过流保护 OCP）
        :param ovp: 过压保护值 (单位: V)
        :param ocp: 过流保护值 (单位: A)
        """
        if ovp is not None:
            self.write(f"SOUR:VOLT:PROT:HIGH {ovp:.3f}")
        if ocp is not None:
            self.write(f"SOUR:CURR:PROT:HIGH {ocp:.3f}")

    def enable_protection(self):
        """
        启用保护功能
        """
        self.write("SYST:PROT:STAT ON")

    def disable_protection(self):
        """
        禁用保护功能
        """
        self.write("SYST:PROT:STAT OFF")

    def set_slew_rate(self, voltage=None, current=None):
        """
        设置电压或电流的上升/下降速率
        :param voltage: 电压上升速率 (单位: V/ms)
        :param current: 电流上升速率 (单位: A/ms)
        """
        if voltage is not None:
            self.write(f"SOUR:VOLT:SLEW {voltage:.3f}")
        if current is not None:
            self.write(f"SOUR:CURR:SLEW {current:.3f}")

    def write(self, command):
        """
        发送自定义 SCPI 指令
        :param command: SCPI 指令字符串
        """
        self.source.write(command)

    def query(self, command):
        """
        查询自定义 SCPI 指令
        :param command: SCPI 指令字符串
        :return: 查询结果字符串
        """
        return self.source.query(command)

    def close(self):
        """
        关闭设备连接
        """
        self.source.close()

# —————————————————————————————分割线———————————————————————————————
#  函数使用说明
global volt, ld_max, freq, ocp_spec, temp, osc, el, xls, vin, display, counter, \
    infinite_off_6, infinite_off_2, ocpmode, osc_addr, el_addr, MSO5, EnValue3, rm, DPO7000, \
    DP05104B, EL6310, EL63600, countmode, file_path, pic_path, entry
# 全局变量声明  通用变量


def common_set():  # 对示波器和电子负载设备进行一系列常见的初始化设置
    osc.state('stop')           # 停止示波器的采集
    el.state('OFF')             # 关闭电子负载

    osc.persistence('OFF')      # 关闭示波器的持久性显示和光标
    osc.cursor('OFF')
    # osc.hormode('MAN')                                        # 设置示波器的水平模式 Horizontal Mode MAN表示手动模式

    for channel in ['CH1', 'CH2', 'CH3', 'CH4']:                # 为所有通道设置耦合方式为 直流 DC
        osc.coupling(channel, 'DC')

    osc.channel('OFF', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')       # 关闭所有通道
    for i in range(1, 9):                                       # 关闭所有测量项
        osc.measOff(i)
    if DPO7000 == 1 or DPO5104B == 1:                 # 关闭水平滚动
        osc.write('HORIZONTAL:ROLL OFF')
    osc.state('run')                                            # 启动示波器的采集

def control_dc_source(voltage, current, output_state):

    # 设置为恒压模式并打开输出
    dc.mode('CV')
    dc.set_voltage(voltage)  # 设置电压
    dc.set_current(current)  # 设置电流
    dc.output(output_state)  # 设置输出开关

    if ocp_spec is not None:
        dc.set_protection(ocp=ocp_spec)  # 只传递OCP，不传递OVP


    # 查询当前电压和电流
    voltage = dc.query_voltage()
    current = dc.query_current()
    print(f'Voltage: {voltage} V, Current: {current} A')

def count():        # 计时器功能，每秒钟更新一次显示内容,在指定的条件下继续计时
    global counter, countmode                                   # 全局变量 counter 时间, countmode 控制计时器开关
    if countmode == 'ON':                                       # 计时器状态判断
        timestr = f'{counter // 60:02}:{counter % 60:02}'       # 将秒数转换为分钟和秒数
        display.config(text=str(timestr))                       # 将格式化后的时间字符串显示在界面上
        counter += 1                                            # 增加计时器的秒数计数器 counter的值
        display.after(1000, count)                              # 计时器每秒更新一次

def generate_sequence():
    # 效率测试里面的计算步长
    step = ld_max / 10
    # 生成从0到最大值的数列，每步增加step
    sequence = [round(i * step, 2) for i in range(11)]  # 11表示包括0到最大值共11个数
    return sequence

def generate_excel_address(col, row):
    """将列号和行号转换为 Excel 单元格地址，例如 'AA152'."""
    letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"

def get_ocp_spec():
    """
    使用 Excel 的 Find 方法快速找到 'OCP' 的位置，并获取其左下角的值。
    :return: OCP 单元格左下角的值
    :raises ValueError: 如果未找到指定工作表或 OCP
    """
    # 获取用户输入的表名
    sheet_name = entry.get()

    try:
        # 确保工作表存在
        sheet = xls.xlBook.Worksheets(sheet_name)
    except Exception:
        raise ValueError(f"Sheet '{sheet_name}' does not exist in the Excel file.")

    # 使用 Excel 的 Find 方法查找 "OCP"（全词匹配）
    cell = sheet.Cells.Find(
        What="OCP",  # 要查找的值
        LookIn=-4163,  # -4163 表示查找单元格的值 (xlValues)
        LookAt=1,  # 1 表示完全匹配 (xlWhole)
        SearchOrder=1,  # 1 表示按行查找 (xlByRows)
        SearchDirection=1,  # 1 表示从上到下查找 (xlNext)
        MatchCase=True  # 区分大小写
    )

    if cell:
        # 找到 "OCP"，计算左下角位置
        target_row = cell.Row + 1
        target_col = cell.Column - 1

        # 确保目标单元格在有效范围内
        if target_row <= sheet.UsedRange.Rows.Count and target_col > 0:
            return xls.getCell(sheet_name, target_row, target_col)

    # 如果未找到 "OCP"
    raise ValueError("'OCP' not found in the specified sheet.")

def go():  # 从 Excel 文件中提取数据并更新相关全局变量
    # 处理事件，*args表示可变参数
    global volt, ld_max, freq, ocp_spec, temp, xls, vin, Iin, file_path           # 全局变量声明
    # 从 EasyExcel 中获取数据并去除空格
    xls = EasyExcel(file_path)
    temp = xls.getCell('Test Summary', 6, 3).strip()  # 去除前后空格
    volt = xls.getCell(entry.get(), 5, 10)  # 去除前后空格
    ld_max = xls.getCell(entry.get(), 28, 3)  # 去除前后空格
    cell_value = xls.getCell(entry.get(), 160, 3)  # 去除前后空格

    # 对 freq 进行计算并格式化为 KHz
    freq_value = float(cell_value) / 0.9
    freq = f"{freq_value} KHz"

    # 获取 ocp_spec 的值并去除空格
    ocp_spec = get_ocp_spec()

    # 获取 vin 的值并去除空格
    vin = xls.getCell(entry.get(), 5, 11)  # 去除前后空格

    Iin = math.ceil((volt * ocp_spec) / (0.75 * vin)) + 1
    print(f"volt: {volt}")  # 输出 volt 的值
    print(f"ld_max: {ld_max}")  # 输出 ld_max 的值
    print(f"freq: {freq}")  # 输出 freq 的值
    print(f"ocp_spec: {ocp_spec}")  # 输出 ocp_spec 的值
    print(f"vin: {vin}")  # 输出 ocp_spec 的值
    print(f"Iin: {Iin}")  # 输出 ocp_spec 的值

    EnValue2.set(volt)    # 更新UI组件

def instrument():       # 用于记录仪器连接
    global osc, el, rm, da, dc, MSO5, DPO7000, DPO5104B, EL6312A, EL63600,EL63640A, DAQ973A, DAQ34970A, Data_Acquisition, DC62024P


    rm = pyvisa.ResourceManager()               # 创建一个资源管理器实例，用于查询连接的仪器
    instrument_list = rm.list_resources()                # 列出所有可用的资源地址
    print(instrument_list)


    DPO7000 = 0
    DPO5104B = 0
    MSO5 = 0
    EL6312A = 0
    EL63600 = 0
    EL63212E = 0
    EL63640A = 0
    DAQ973A = 0
    DAQ34970A = 0
    DC62024P = 0


    # 初始化设备表示，所有设备的标志默认为0.即未连接

    for address in instrument_list:                             # 遍历所有资源地址
        if 'GPIB' in address or 'USB' in address:         # 检查仪器地址是否有效
            ins = rm.open_resource(address, timeout=20000)            # 打开指定地址的仪器资源
            device_id = ins.query('*IDN?').upper()     # 查询仪器的标识信息并大写返回
            print(device_id)                           # 打印仪器的标识信息

            if 'TEKTRONIX,DPO7' in device_id:
                print('TEKTRONIX,DPO7000系列示波器连接成功，地址为' + address)
                osc = OscDPO7000C(address)
                DPO7000 = 1
            elif 'TEKTRONIX,MSO' in device_id:
                print('TEKTRONIX,MSO4/5/6系列示波器连接成功，地址为' + address)
                osc = OscMPO5series(address)
                MSO5 = 1
            elif 'CHROMA,6312A' in device_id:
                print('Chroma,6312系列电子负载连接成功，地址为' + address)
                el = El6312A(address)
                EL6312A = 1
            elif 'CHROMA,63212E' in device_id:
                print('Chroma,6312系列电子负载连接成功，地址为' + address)
                el = El6312A(address)
                EL63212E = 1
            elif 'CHROMA,63600' in device_id:
                print('Chroma,63600系列电子负载连接成功，地址为' + address)
                el = El63600(address)
                EL63600 = 1
            elif 'CHROMA,63640' in device_id:
                print('CHROMA,63640A系列电子负载连接成功, 地址为' + address)
                el = El63600(address)
                EL63640A = 1
            elif 'TEKTRONIX,DPO5' in device_id:
                print('TEKTRONIX,DPO5000系列示波器连接成功，地址为' + address)
                osc = OscDPO5104B(address)
                DPO5104B = 1
            elif 'DAQ973A' in device_id:
                print('Keysight Technologies,DAQ973A数据采集仪连接成功，地址为' + address)
                da = Data_Acquisition(address)
                da.Channel_Set()
                DAQ973A = 1
            elif '34970A' in device_id:
                print('Keysight Technologies,DAQ34970A数据采集仪连接成功，地址为' + address)
                da = Data_Acquisition(address)
                da.Channel_Set()
                DAQ34970A = 1
            elif 'CHROMA,62024P-80-60' in device_id:  # 匹配 DC62014P 的标识符
                print('DC62014P 电源设备连接成功，地址为' + address)
                dc = (DCsource(address))  # 创建 DCSource 类实例
                DC62024P = 1  # 更新标志变量



    oscstate = DPO7000 or MSO5 or DPO5104B  # 检查是否至少有一个示波器和一个电子负载已经连接
    elstate = EL6312A or EL63600 or EL63640A or EL63212E
    data_acquisition = DAQ973A or DAQ34970A
    dcstate = DC62024P  # 检查 DC62024P 是否连接

    connected_devices = []
    if data_acquisition:
        connected_devices.append("数据采集仪")
    else:
        print("没有连接到数据采集仪。")
    if oscstate:
        connected_devices.append("示波器")
    if elstate:
        connected_devices.append("电子负载")
    if dcstate:
        connected_devices.append("DCsource")

    # 如果有已连接的设备，显示连接的信息
    if connected_devices:
        devices_message = "和".join(connected_devices) + "已正确连接"
        auto_close_messagebox(root, title='仪器连接', message=devices_message, timeout=1000)

def mkdir(path):
    """处理地址并创建目录。"""
    # 去除首尾空格并去掉尾部的斜杠
    path = path.strip().rstrip("\\")

    # 判断路径是否存在并创建目录
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"{path} 创建成功")
        return True
    else:
        print(f"{path} 目录已存在")
        return False

def process(type):
    # 非常好用的数据处理函数，用于对各种特殊要求的数据进行处理，并针对 05 系列示波器进行区分
    commands = {
        'tri': 'ACQUIRE:STATE?',
        'ch4freq': 'MEASUREMENT:MEAS3:MEAN?',
        'poswidth': 'MEASUREMENT:MEAS3:MEAN?',
        'rms': 'MEASUREMENT:MEAS3:MEAN?',
        'ch1max': 'MEASUREMENT:MEAS1:MAX?',
        'ch1min': 'MEASUREMENT:MEAS2:MINI?',
        'ch4max': 'MEASUREMENT:MEAS1:MAX?',
        'ch3max': 'MEASUREMENT:MEAS3:MAX?',
        'ch_m1_max': 'MEASUREMENT:MEAS5:MAX?',
        'rpk': 'MEASUREMENT:MEAS4:MAX?',
        'ld_ocp': 'MEASUREMENT:MEAS5:MAX?',
        'ld_short': 'MEASUREMENT:MEAS5:MAX?',
        'pk2pk': 'MEASUREMENT:MEAS4:MAX?'
    }

    start_indices = {
        'tri': 15,
        'ch4freq': 24,
        'poswidth': 24,
        'rms': 24,
        'ch1max': 26,
        'ch1min': 26,
        'ch4max': 27,
        'ch3max': 27,
        'ch_m1_max': 27,
        'rpk': 26,
        'ld_ocp': 27,
        'ld_short': 27,
        'pk2pk': 27
    }

    command = commands[type]  # 读取命令类型
    result = osc.query(command)  # 执行命令


    if MSO5 == 1:
        start_index = start_indices[type]  # 读取命令获得位数需求
        if type == 'tri':
            print(f"Original result for 'tri': {result}")  # 输出 'tri' 的原始结果
            # 提取数字部分
            result_numeric = ''.join(filter(str.isdigit, result))
            print(f"Numeric part extracted from 'tri': {result_numeric}")
            measurement_value = int(result_numeric)  # 将数字部分转换为整数
        else:
            trimmed_result = result[start_index:]
            print(f"Trimmed result: {trimmed_result}")  # 调试打印
            measurement_value = float(trimmed_result)
            print(f"measurement_value: {trimmed_result}")  # 调试打印

        # 对某些特定类型进行额外的处理
    else:
        measurement_value = float(result)

    if type == 'ch4freq':
        measurement_value = measurement_value / 1000
    if type == 'poswidth':
        measurement_value = measurement_value * 1000000000
    if type == 'rpk':
        measurement_value = int(measurement_value * 10000)
    if type == 'pk2pk':  # 对pk2pk进行三位小数扩大
        measurement_value = measurement_value * 1000
    return measurement_value

def refresh(channels=None, electronic=False, delay=1):
    if channels:
        # 刷新通道显示      注意通道必须是列表形式
        for channel in channels:
            osc.write(f'DISplay:GLObal:{channel}:STATE OFF')

        time.sleep(delay)

        for channel in channels:
            osc.write(f'DISplay:GLObal:{channel}:STATE ON')

    elif electronic:        # 刷新电子负载
        el.short('OFF')

        time.sleep(delay)

        el.short('ON')

def scale():  # 用于 test0 1 2 的设置
    rpk = process('rpk')

    # 定义范围及对应的尺度
    ranges_and_scales = [
        (410, 10E-03),
        (810, 20E-03),
        (1210, 30E-03),
        (1610, 40E-03),
        (2010, 50E-03),
        (2410, 60E-03),
        (2810, 70E-03),
        (3210, 80E-03),
        (3610, 90E-03)
    ]

    # 找到适合的尺度并应用
    for upper_limit, scale in ranges_and_scales:
        if rpk < upper_limit:
            osc.scale('CH1', scale)
            print(f"Applied scale: {scale}")  # 打印应用的尺度
            break
    else:
        # 处理 rpk 超过范围的情况
        print("输出电压超出范围！")

def auto_close_messagebox(win, title, message, timeout=5000):
    """
    在指定的主窗口上显示一个自动关闭的消息框，timeout 毫秒后自动关闭。

    参数:
    win - 父窗口
    title - 消息框标题
    message - 消息框内容
    timeout - 自动关闭的时间（毫秒），默认 5000 毫秒
    """
    # 获取父窗口的大小和位置
    win_width = win.winfo_width()
    win_height = win.winfo_height()

    # 获取屏幕的大小
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # 计算消息框的居中位置
    x_offset = (screen_width - 300) // 2  # 300 是消息框的宽度
    y_offset = (screen_height - 100) // 2  # 100 是消息框的高度

    # 创建一个新的顶层窗口（消息框）
    msg_box = tk.Toplevel(win)
    msg_box.title(title)
    msg_box.geometry(f"300x100+{x_offset}+{y_offset}")  # 设置居中位置
    msg_box.grab_set()  # 模态窗口，阻止对主窗口的输入

    # 在窗口中添加消息文本
    tk.Label(msg_box, text=message, wraplength=250).pack(pady=20)

    # 设置自动关闭
    msg_box.after(timeout, msg_box.destroy)

def set_cursor():
    if MSO5 == 1:
        osc.write('DISplay:WAVEView1:CURSor:CURSOR1:STATE ON')

        # 设置游标源
        osc.write('DISplay:WAVEView1:CURSor:CURSOR1:ASOUrce CH4')

        # 设置游标功能
        osc.write('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:FUNCTION VBArs')
    else:
        print("其他示波器光标开启")
        osc.write('CURSOR:STATE ON')
        osc.write('CURSOR:SOURCE 4')
        osc.write('CURSOR:FUNCTION VBArs')
        osc.write('CURSOR:LINESTYLE DASHed')

def set_cursor_position(cursor_type):       # test5中 求时间位置
    if cursor_type == 'a':
        command = 'DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:AXPOSITION?'
    elif cursor_type == 'b':
        command = 'DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:BXPOSITION?'
    position = osc.query(command)
    position = position[51:]  # 去除多余的前缀或单位信息
    position = float(position) * 1000000000  # 转换为纳秒
    return position

def set_excel_path():    # 主窗口选择文件路径
    global file_path, EnValue3
    file_path = tk.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    # 弹出一个文件对话框，让用户选择一个文件，类型是xlsx或者xls
    print(f"测试报告路径：{file_path}")
    if file_path:           # 如果文件路径正确，更新文件路径的值到EnValues3控件
        EnValue3.set(file_path)

def set_horizontal_mode(scale_ms05=None, scale_other=None, position=None, samplerate=None, recordlength=None, mode='AUTO'):
    # 设置示波器的水平模式、时间尺度和采样率
    # 也可以只设置水平位置和关闭注释
    osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
    print("注释已关闭")  # 提示：关闭测量注释

    # 设置水平模式
    if mode == 'AUTO':
        osc.write('HORIZONTAL:MODE AUTO')
        print("水平模式已设置为：AUTO")
    elif mode == 'MANUAL':
        osc.write('HORIZONTAL:MODE MANUAL')
        print("水平模式已设置为：MANUAL")

    # 设置水平刻度
    if position:
        osc.write(f'HORIZONTAL:POSITION {position}')
        print(f"水平位置已设置为：{position}")

    # 设置采样频率
    if samplerate:
        osc.write(f'HORIZONTAL:MODE:SAMPLERATE {samplerate}')  # 设置采样频率
        print(f"采样频率已设置为：{samplerate}")

    if MSO5 == 1:
        if scale_ms05:
            osc.write(f'HORIZONTAL:MODE:SCALE {scale_ms05}')  # 设置水平刻度
            print(f"水平刻度（MSO5）已设置为：{scale_ms05}")
    else:
        if scale_other:
            osc.write(f'HORIZONTAL:MODE:SCALE {scale_other}')  # 设置水平刻度
            print(f"水平刻度（其他示波器）已设置为：{scale_other}")
    # 设置水平位置


    if recordlength:
        osc.write(f'HORIZONTAL:MODE:RECORDLENGTH {recordlength}')  # 设置采样频率
        print(f"采样频率已设置为：{recordlength}")

    time.sleep(1)

def set_load(channel=None, mode=None, load_value=None, lvp_on=True, load=True,):
    # 设置电源开关，以及LVP模式开关
    if load:
        print("Load is ON.")

        if lvp_on:
            # 设置电子负载的通道和参数
            el.static(channel, mode, load_value)
            time.sleep(0.5)
            print(f"Channel: {channel}, Mode: {mode}, Load Value: {load_value}")



            # 打开电子负载
            el.state('ON')
            print("Electronic load state set to ON.")

        else:
            # 设置默认负载模式
            el.static(9, 'MIN', ld_max)
            el.write('CONF:LVP OFF')
            print("Low Voltage Protection (LVP) is OFF. Setting load to minimum.")

            # 打开电子负载
            el.state('ON')
            print("Electronic load state set to ON with default settings.")

    else:
        # 关闭电子负载和LVP模式
        el.state('OFF')
        el.write('CONF:LVP OFF')
        print("Load is OFF. Low Voltage Protection (LVP) is OFF.")

def set_osc():      # 示波器采样操作
    osc.state('run')  # 示波器开始采样
    osc.number(300)
    osc.state('stop')  # 示波器停止采样
    time.sleep(1)

def set_pic_path():      # 主窗口  选择图片保存路径
    global  pic_path, EnValue4
    pic_path = tk.filedialog.askdirectory()    # 弹出对话框选择目录，存储于pic_path变量中
    print(f"保存图片路径：{pic_path}")
    if pic_path:            # 如果文件路径正确，更新文件路径的值到EnValues4控件
        EnValue4.set(pic_path)

def set_record_length(mode='FINE'):          # 主窗口设置postime以及模式
    global volt,vin,freq
    if isinstance(freq, str) and "KHz" in freq:
        # 处理 freq 为字符串时，去除 "KHz" 并转换为浮点数
        freq = float(freq.replace("KHz", "").strip())
    elif isinstance(freq, float):
        # 如果 freq 已经是 float 类型，直接使用
        pass
    else:
        raise ValueError("freq 必须是带有 'KHz' 单位的字符串或浮动类型")
    postime= volt * 1000000 / (vin * freq)    # 根据计算得到的记录时间调整记录长度

    if mode == 'FINE':
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
    elif mode == 'COARSE':
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

def set_trigger(mode, channel, edge, level_option=None):        # 触发设置
    if level_option == 0.2:
        tri_level = float(0.2 * volt)
    elif level_option == 0.5:
        tri_level = float(0.5 * volt)
    else:
        tri_level = float(2.48)
    osc.trigger(mode, channel, edge, tri_level)

def start_countdown(minutes):
    total_seconds = minutes * 60
    update_countdown(total_seconds)

def update_countdown(remaining_seconds):
    """更新倒计时"""
    global timer_id
    minutes, seconds = divmod(remaining_seconds, 60)
    if countdown_label.winfo_exists():  # 确保控件存在
        countdown_label.config(text=f"倒计时：{minutes:02}:{seconds:02}")

    if remaining_seconds > 0:
        # 每秒更新一次倒计时
        timer_id = root.after(1000, update_countdown, remaining_seconds - 1)
    else:
        # 倒计时结束后执行操作
        ask_test10()

def ask_test10():
    response = messagebox.askquestion(title='程序提醒', message='六分钟已到，请问TEST10测试是否完成？')
    if response == 'yes':
        print("用户确认TEST10已完成")
        root1.destroy()
        test2('ALL')
    else:
        print("用户尚未完成TEST10")

def on_close():
    """退出程序的安全处理"""
    global timer_id
    if timer_id is not None:
        root.after_cancel(timer_id)  # 取消定时器任务
        timer_id = None
    root1.destroy()  # 销毁窗口

def test_save(num, i, picture_cell, left_offset, top_offset, width, height, add=True):
    """
    导出图片并插入 Excel
    :param num: 数字编号，用于生成图片名称
    :param i: 序号，用于生成图片名称
    :param picture_cell: 插入图片的目标单元格
    :param left_offset: 图片插入位置的左偏移量
    :param top_offset: 图片插入位置的上偏移量
    :param width: 图片宽度
    :param height: 图片高度
    :param add: 是否插入图片到 Excel，默认为 True
    """
    # 构造图片名称
    name = f"T{num}-{i}"
    print(f"[调试] 生成的图片名称: {name}")

    # 构造路径
    # 去除 entry.get() 和 temp 中的空格
    picture_dir = os.path.normpath(f"{pic_path}/POL Test Pictures/{temp.strip()}/{entry.get().strip()}")
    picture_path = os.path.normpath(f"{picture_dir}/{name}.PNG")
    osc_dir = os.path.normpath(f'C:/POL Test Pictures/{temp.strip()}/{entry.get().strip()}')

    # 确保本地目录存在
    try:
        mkdir(picture_dir)  # 自定义目录创建函数
        print(f"[成功] 本地目录已创建或存在：{picture_dir}")
    except Exception as e:
        print(f"[错误] 创建本地目录时发生异常：{e}")
        return

    # 确保示波器目录存在
    try:
        osc.makeDir('C:\\POL Test Pictures')
        osc.makeDir(f'C:\\POL Test Pictures\\{temp}')
        osc.makeDir(f'C:\\POL Test Pictures\\{temp}\\{entry.get()}')
        print("[成功] 示波器目录已创建或存在")
    except Exception as e:
        print(f"[错误] 创建示波器目录时发生异常：{e}")
        return

    # 导出图片
    try:
        osc.export('PNG', f'{osc_dir}/{name}')
        print(f"[成功] 图片已导出到示波器路径：{osc_dir}/{name}")
    except Exception as e:
        print(f"[错误] 导出图片时发生异常：{e}")
        return

    time.sleep(3)  # 等待示波器导出完成

    # 读取示波器文件
    saved_image_path = os.path.normpath(f'{osc_dir}/{name}.PNG')
    try:
        osc.readfile(saved_image_path)
        print(f"[成功] 已从示波器读取图片文件：{saved_image_path}")
    except Exception as e:
        print(f"[错误] 从示波器读取图片文件时发生异常：{e}")
        return

    time.sleep(3)  # 等待读取完成

    # 读取原始文件数据到本地
    try:
        osc.readraw(picture_path)
        print(f"[成功] 已从示波器保存图片到本地路径：{picture_path}")
    except Exception as e:
        print(f"[错误] 保存图片到本地时发生异常：{e}")
        return

    time.sleep(2)

    # 插入图片到 Excel
    if add:
        if os.path.exists(picture_path):
            try:
                print(f"[调试] 正在添加图片到 Excel: {picture_path}")
                xls.addPicture(entry.get(), picture_path, picture_cell, left_offset, top_offset, width, height)
                print(f"[成功] 图片成功插入到 Excel: {picture_path}")
            except Exception as e:
                print(f"[错误] 插入图片到 Excel 时发生异常：{e}")
        else:
            print(f"[警告] 图片未找到，无法插入到 Excel: {picture_path}")


# —————————————————————————————分割线———————————————————————————————

# 测量方法 & 通道设置
def measure1():  # 仅测量CH1 MAX MIN RMS PK2PK
    # 参数 1：第一次测量  CH1：被测量的通道  MAXIMUM：测量最大值
    osc.measure(1, 'CH1', 'MAXIMUM')
    osc.measure(2, 'CH1', 'MINIMUM')
    osc.measure(3, 'CH1', 'RMS')
    osc.measure(4, 'CH1', 'PK2PK')

    if MSO5 == 1:
        for i in range(1, 5):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')

def measure2():
    """测量 CH1 和 CH4 的特定参数，处理不同设备配置的统计显示"""
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH1', 'RMS')
        osc.measure_1(4, 'CH1', 'PK2PK')
        osc.measure_1(5, 'CH4', 'MAXIMUM')
        osc.measure_1(6, 'CH4', 'MINIMUM')
        osc.measure_1(7, 'CH4', 'FREQUENCY')
        osc.measure_1(8, 'CH4', 'PDUTY')
        for i in range(1, 9):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')
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
        """测量 CH1、CH2、CH3 和 CH4 的最大值和最小值"""
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH2', 'MAXIMUM')
        osc.measure_1(4, 'CH2', 'MINIMUM')
        osc.measure_1(5, 'CH3', 'MAXIMUM')
        osc.measure_1(6, 'CH3', 'MINIMUM')
        osc.measure_1(7, 'CH4', 'MAXIMUM')
        osc.measure_1(8, 'CH4', 'MINIMUM')
        # 使用循环关闭所有统计显示
        for i in range(1, 9):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')

    # 根据设备类型选择是否关闭统计显示
    else:
        # 使用循环关闭所有统计显示
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
        osc.measure(3, 'CH2', 'MAXIMUM')
        osc.measure(4, 'CH2', 'MINIMUM')
        osc.measure(5, 'CH3', 'MAXIMUM')
        osc.measure(6, 'CH3', 'MINIMUM')
        osc.measure(7, 'CH4', 'MAXIMUM')
        osc.measure(8, 'CH4', 'MINIMUM')

def measure4(trend=None):
    """用于测量 CH1 的 MAX、MIN、RISETIME 和 RISE"""
    if MSO5 == 1:
        # 测量 CH1 的 MAXIMUM 和 MINIMUM
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        if trend == 'rise':
            osc.measure(3, 'CH1', 'RISETIME')
        elif trend == 'fall':
            osc.measure(3, 'CH1', 'FALLTIME')

        # 关闭统计显示
        for i in range(1, 4):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')
    else:
        # 测量 CH1 的 MAXIMUM 和 MINIMUM
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')

    # 根据设备类型测量 RISE
    if DPO7000 == 1 or DPO5104B == 1:
        if trend == 'rise':
            osc.measure(3, 'CH1', 'RISE')
        elif trend == 'fall':
            osc.measure(3, 'CH1', 'FALL')

def measure5(type=None):  # 测量CH4的MAX MIN FRE PDUTY
    if type == 'FRE':
        # 测量 CH4 的 MAX、MIN、FREQUENCY 和 PDUTY
        osc.measure(1, 'CH4', 'MAXIMUM')
        osc.measure(2, 'CH4', 'MINIMUM')
        osc.measure(3, 'CH4', 'FREQUENCY')
        osc.measure(4, 'CH4', 'PDUTY')
        # 关闭所有测量显示
        for i in range(1, 5):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')

    elif type == 'PWIDTH':
        # 测量 CH4 的 MAX、MIN 和 PWIDTH
        osc.measure(1, 'CH4', 'MAXIMUM')
        osc.measure(2, 'CH4', 'MINIMUM')
        osc.measure(3, 'CH4', 'PWIDTH')
        # 关闭所有测量显示
        for i in range(1, 4):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')

def measure6():  # 测量CH4的 MAX MIN
    global table_type
    if MSO5 == 1:
        if table_type == '表格1':
            osc.measure_1(1, 'CH4', 'MAXIMUM')
            osc.measure_1(2, 'CH4', 'MINIMUM')
            for i in range(1, 3):
                osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')
        if table_type == '表格2':
            osc.measure_1(1, 'CH4', 'MAXIMUM')
            osc.measure_1(2, 'CH4', 'MINIMUM')
            osc.measure_1(3, 'CH3', 'MAXIMUM')
            osc.measure_1(4, 'CH3', 'MINIMUM')
            osc.measure_1(5, 'MATH1', 'MAXIMUM')
            osc.measure_1(6, 'MATH1', 'MINIMUM')
            for i in range(1, 7):
                osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')


    else:
        if table_type == '表格1':
            osc.measure_1(1, 'CH4', 'MAXIMUM')
            osc.measure_1(2, 'CH4', 'MINIMUM')
            for i in range(1, 3):
                osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')
        if table_type == '表格2':
            osc.measure(1, 'CH4', 'MAXIMUM')
            osc.measure(2, 'CH4', 'MINIMUM')
            osc.measure(3, 'CH3', 'MAXIMUM')
            osc.measure(4, 'CH3', 'MINIMUM')
            osc.measure(5, 'MATH1', 'MAXIMUM')
            osc.measure(6, 'MATH1', 'MINIMUM')
            for i in range(1, 7):
                osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')

def measure9():
    # 若是05示波器 测量CH1和 CH3的 MAX MIN  CH4:MAX  CH5:MIN
    # 若不是     则测量 CH1 CH3 CH4的 MAX MIN
    if MSO5 == 1:
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH3', 'MAXIMUM')
        osc.measure_1(4, 'CH3', 'MINIMUM')
        osc.measure_1(5, 'CH4', 'MAXIMUM')
        osc.measure_1(6, 'CH4', 'MINIMUM')
        for i in range(1, 7):
            osc.write(f'MEASUREMENT:MEAS{i}:DISPlaystat:ENABle OFF')
    else:
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
        osc.measure(3, 'CH3', 'MAXIMUM')
        osc.measure(4, 'CH3', 'MINIMUM')
        osc.measure(5, 'CH4', 'MAXIMUM')
        osc.measure(6, 'CH4', 'MINIMUM')

def tl1_channel_set():
    osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')  # 只开启通道1
    osc.chanset('CH1', 0, volt, '20.0000E+06', 10E-02)      # 设置通道1的位置
    osc.label('CH1', entry.get(), 1, 6)                   # 设置label
    osc.trigger('AUTO', 'CH1', 'RISE', volt)                # 触发设置

    if MSO5 == 1:  # 如果是 MS05示波器，关闭通道1的显示再打开（刷新）
        refresh(['CH1'])

def tl2_channel_set():
    # 只开启通道1和通道4
    osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
    print("通道1和通道4已开启。")

    # 设置通道1的相关参数
    osc.chanset('CH1', 2, volt, '20.0000E+06', 2e-2)
    print(f"CH1 设置：位置=2, 缩放尺度={volt}, 时间基准=20.0000E+06, 时间尺度=0.01s")

    # 设置通道4的纵向尺度
    ldstep = float(ld_max / 3)  # 假设ld_max在其他地方已定义
    osc.chanset('CH4', -4, 0, '20.0000E+06', ldstep)
    print(f"CH4 设置：位置=-4, 偏移=0, 时间基准=20.0000E+06, 纵向尺度={ldstep}")

    # 设置通道标签
    osc.label('CH1', entry.get(), 1, 4)
    print(f"CH1 标签设置为: {entry.get()}")

    osc.label('CH4', "Iout", 2, 10)
    print("CH4 标签设置为: Iout")

    # 设置自动触发模式
    osc.trigger('AUTO', 'CH4', 'RISE', ld_max / 2)
    time.sleep(0.1)  # 添加0.1秒的延迟
    print(f"通道4触发设置为自动模式，触发电平={ld_max / 2}")

    # 刷新通道显示
    if MSO5 == 1:  # 假设MSO5在其他地方已定义
        refresh(['CH1', 'CH4'])
        print("刷新通道显示。")

    # 开启无限累计模式
    osc.persistence('INFPersist')
    print("已开启无限累计模式。")

def tl3_channel_set():  # 根据电压 volt 的值设置示波器的四个通道，并调整它们的显示设置
    osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')           # 打开通道1到4，关闭通道5和6

    if volt >= 5:  # 根据电压值 dy调整通道1的设置
        osc.chanset('CH1', -2, 0, '20.0000E+06', 1.5)
    elif volt >= 3.3:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 1)
    elif volt >= 2:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.7)
    elif volt >= 1.5:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.5)
    else:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.4)

    osc.chanset('CH2', -1, 0, '20.0000E+06', 1)  # 设置通道2、3、4
    osc.chanset('CH3', -3, 0, '20.0000E+06', 1)
    osc.chanset('CH4', -0, 0, '20.0000E+06', 1)

    osc.label('CH1', entry.get(), 1, 9)  # 设置label
    osc.label('CH2', "EN", 1.5, 9)
    osc.label('CH3', "PG", 2, 9)
    osc.label('CH4', "VIN", 2.5, 9)

    if MSO5 == 1:  # 刷新通道显示
        refresh(['CH1', 'CH2', 'CH3', 'CH4'])

def tl4_channel_set():  # 据电压 volt 的值调整示波器的 CH1 通道设置
    osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')        # 只开启通道1

    if volt >= 5:  # 同上
        osc.chanset('CH1', -3, 0, '20.0000E+06', 1)
    elif volt >= 3.3:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.7)
    elif volt >= 2:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.5)
    elif volt >= 1.5:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.4)
    elif volt >= 1:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.3)
    else:
        osc.chanset('CH1', -3, 0, '20.0000E+06', 0.2)

    osc.label('CH1', entry.get(), 1, 9)  # 设置label

    if MSO5 == 1:  # 刷新显示
        refresh(['CH1'])

def tl5_channel_set():  # 根据 vin 的值来配置示波器的 CH4 通道
    osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')        # 只开启通道四
    if vin >= 10:  # 同上
        osc.chanset('CH4', -3, 0, '500.0000E+06', 2)
        osc.label('CH4', "PHASE", 2, 8)
    else:
        osc.chanset('CH4', -3, 0, '500.0000E+06', 2)
        osc.label('CH4', "PHASE", 2, 8)
    osc.trigger('NORMAL', 'CH4', 'RISE', volt / 2)
    #  触发模式：NORMAL  通道：CH4  触发条件：RISE  触发位置：6
    # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率
    # osc.persistence('INFPersist')  # 开启累积
    el.write('CHAN 1')  # 选择电子负载的通道为1
    if MSO5 == 1:  # 刷新通道的显示
        refresh(['CH4'])

def tl6_channel_set():  # 配置示波器的水平位置、通道设置、触发、标签和累积模式
    osc.channel('OFF', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
    if vin >= 10:  # 同上
        if table_type == '表格1':
            osc.chanset('CH4', -3, 0, 'FULL', 2)
        if table_type == '表格2':
            osc.chanset('CH3', 0, vin, 'FULL', 0.6)
            osc.chanset('CH4', 1, 0, 'FULL', 5)
            math_channel_set('MATH1','CH3','CH4',0, pos= -4, scale = 5)
        osc.trigger('NORMAL', 'CH4', 'RISE', 6)
    else:
        if table_type == '表格1':
            osc.chanset('CH4', -3, 0, 'FULL', 2)
        if table_type == '表格2':
            osc.chanset('CH3', 0, vin, 'FULL', 0.6)
            osc.chanset('CH4', 1, 0, 'FULL', 5)
            math_channel_set('MATH1','CH3','CH4',0, pos= -4, scale = 5)
        osc.trigger('NORMAL', 'CH4', 'RISE', 3)

    osc.persistence('INFPersist')  # 开启累积
    refresh('CH4')
    if table_type == '表格2':
        osc.label('CH3', "P12V_FET", 1, 3)
        osc.label('CH3',"VDS_LMOS", 1, 7)
        osc.label('Math1',"VDS_HMOS", 1, -7)
        refresh(['CH3','CH4','Math1'])

def tl9_channel_set():
    osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')
    # 计算垂直比例尺度并设置通道
    scale_v = float(volt / 2)
    osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)  # 设置 CH1
    osc.chanset('CH3', -2, 0, '20.0000E+06', 2)       # 设置 CH3
    scale_ld = float(ocp_spec / 1.5)
    scale_ld = round(scale_ld)  # 计算 CH4 的比例尺度
    osc.chanset('CH4', -3, 0, '20.0000E+06', scale_ld)  # 设置 CH4

    # 设置通道标签
    osc.label('CH1', entry.get(), 1, 9)  # 设置 CH1 标签
    osc.label('CH3', "PG", 1.5, 9)      # 设置 CH3 标签
    osc.label('CH4', "Iout", 2.5, 9.5)  # 设置 CH4 标签


    if MSO5 == 1:
        refresh(['CH1', 'CH3', 'CH4'], False, 1)

def math_channel_set(channel, source1, source2, offset, pos, scale):
    # 添加数学通道
    osc.write(f'MATH:ADDNew "{channel}"')  # 添加 MATH1

    # 设置源为 CH3 - CH4
    osc.write(f'MATH:{channel}:SOURCE1 {source1}')  # 设置第一个源为 CH3
    osc.write(f'MATH:{channel}:SOURCE2 {source2}')  # 设置第二个源为 CH4

    osc.write(f'{channel}:TYPE BASIC')  # 设置数学通道类型为 Basic
    osc.write(f'{channel}:FUNCTION SUBTRACT')  # 设置运算为减法 CH3 - CH4

    # 设置 offset, position, scale
    if MSO5 == 1:
        time.sleep(1)
        osc.write(f'MATH:{channel}:OFFSET {offset:.1f}')  # 设置 offset

        time.sleep(1)
        osc.write(f'DISplay:WAVEView1:MATH:{channel}:VERTICAL:POSITION {pos:.1f}')  # 设置 position

        time.sleep(1)
        osc.write(f'DISplay:WAVEView1:MATH:{channel}:VERTICAL:SCALE {scale:.1f}')  # 设置 scale
    else:
        osc.write(f'{channel}:OFFSET {offset:.2f}')  # 设置 offset

        osc.write(f'{channel}:VERTICAL:POSITION {pos:.2f}')  # 设置 position

        osc.write(f'{channel}:VERTICAL:SCALE {scale:.2f}')  # 设置 scale

    # 确认设置是否生效


# —————————————————————————————分割线———————————————————————————————

# 测试 UI部分  每个测试对应的窗口设计及函数执行
def t0_win():  # T-0 DMM&Scope Offset Record
    global root0
    root0 = tk.Toplevel()  # 创建一个新的顶层窗口root0，独立于主窗口root，作为当前测试窗口
    root0.title('T-0 DMM&Scope Offset Record')  # # 设置窗口标题
    root0.geometry('300x200')  # 设置新窗口尺寸
    root0.transient(root)  # 将root0作为root的临时窗口,前者关闭不影响后者
    tk.Label(root0, text='测试前请校准探头，给电路上电, \n请将差分探头一端连接到示波器的一通道，\n另一端连接到待测VR的输出端, 单击“开始测试”进行测试。',
             wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    # 在root0中创建文本标签，设置最大宽度为300(超过自动换行), 文本对齐到标签的左边, 将标签放到窗口的第0行第0列, 占据两列, 设置内边距

    tk.Button(root0, text="开始测试", command=test0).grid(row=1, column=0, padx=5, pady=20)

    tk.Button(root0, text='退出测试', command=root0.destroy, activeforeground='white', activebackground='red').grid(row=1, column=1, padx=5, pady=20)
    root0.attributes("-topmost", 1)  # 将root0设置为总在最前面显示

def t1_win():  # T-1 DC Regulation+Ripple&Noise  Test
    global countdown_label, root1
    root1 = tk.Toplevel()                                      # 创建一个新的顶层窗口root1，独立于主窗口root，作为当前测试窗口
    root1.title('T-1 DC Regulation+Ripple&Noise Test')      # 设置窗口标题
    root1.geometry('340x200')                               # 设置新窗口尺寸
    root1.transient(root)                                   # 将root1作为root的临时窗口,前者关闭不影响后者
    root1.attributes("-topmost", 1)                         # 将root1设置为总在最前面显示
    tk.Label(root1, text='请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)

    # 倒计时标签
    countdown_label = tk.Label(root1, text="倒计时：6:00", font=("Arial", 12))
    countdown_label.grid(row=2, column=1, padx=10, pady=10)

    # 按钮放在同一行
    tk.Button(root1, text="开始测试", command=test1).grid(row=1, column=0, padx=5, pady=5)
    tk.Button(root1, text="开始计时", command=lambda: start_countdown(6)).grid(row=2, column=0, padx=5, pady=5)
    tk.Button(root1, text='退出测试', command=on_close, activeforeground='white', activebackground='red').grid(row=1, column=1, columnspan=2, padx=5, pady=5)

def t2_win():
    global root2
    root2 = tk.Toplevel()                                             # 创建一个新的顶层窗口 root2
    root2.title('T-2 Loading Transient Response Test')                # 设置窗口标题
    root2.geometry('360x330')                                         # 设置窗口尺寸
    root2.transient(root)                                             # 将 root2 作为 root 的临时窗口
    root2.attributes("-topmost", 1)                                   # root2 最前显示

    # 提示标签
    tk.Label(root2, text='请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
             wraplength=300, anchor='w').grid(row=0, column=0, columnspan=3, padx=20, pady=20)

    # 分组框
    group2 = tk.LabelFrame(root2, text='单项测试', padx=5, pady=5)
    group2.grid(row=3, rowspan=2, column=0, columnspan=3, padx=50, pady=15)

    # 低频、中频和高频测试按钮，调用 test2 函数传递不同模式
    tk.Button(group2, text="运行 Low Frequency Test", command=lambda: test2('Low')).grid(row=3, column=0, sticky=tk.E + tk.W, padx=40, pady=5)
    tk.Button(group2, text="运行 Mid Frequency Test", command=lambda: test2('Mid')).grid(row=4, column=0, sticky=tk.E + tk.W, padx=40, pady=5)
    tk.Button(group2, text="运行 High Frequency Test", command=lambda: test2('High')).grid(row=5, column=0, sticky=tk.E + tk.W, padx=40, pady=5)

    # 开始测试和退出按钮
    tk.Button(root2, text="开始测试", command=lambda: test2('ALL')).grid(row=1, column=0, padx=40, pady=5)
    tk.Button(root2, text='退出测试', command=root2.destroy, activeforeground='white', activebackground='red').grid(row=1, column=2, padx=40, pady=5)

def t3_win():               # T-3 Power Up & Down Sequence Measurement
    global root3
    root3 = tk.Toplevel()                                                  # 创建一个新的顶层窗口root3，独立于主窗口root，作为当前测试窗口
    root3.title('T-3 Power Up & Down Sequence Measurement')             # 设置窗口标题
    root3.geometry('400x450')                                           # 设置新窗口尺寸
    root3.transient(root)                                               # 将root3作为root的临时窗口,前者关闭不影响后者
    root3.attributes("-topmost", 1)                                     # root3 最前显示
    tk.Label(root3, text="请使用探头1连接示波器的一通道和待测VR的输出端，使用探头2连接示波器的二通道和待测VR的EN信号，"
                      "使用探头3连接示波器的三通道和待测VR的PG信号，使用探头4连接示波器的四通道和待测VR的VIN信号，"
                      "单击“开始测试”进行测试。", wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)

    tk.Button(root3, text="开始测试", command=lambda:test3('ALL')).grid(row=1, column=0, padx=5, pady=20)
    tk.Button(root3, text='退出测试', command=root3.destroy, activeforeground='white', activebackground='red').grid(
        row=1, column=1, padx=5, pady=20)

    group3 = tk.LabelFrame(root3, text='单项测试', padx=5, pady=5)
    group3.grid(row=2, column=0, columnspan=2, padx=60, pady=15)

    tk.Button(group3, text="运行 Power Up Sequence with NO Load", command=lambda: test3('upno')).grid(row=0, column=0, sticky=tk.E + tk.W, padx=5, pady=5)
    tk.Button(group3, text="运行 Power Down Sequence with NO Load", command=lambda: test3('downno')).grid(row=1, column=0, sticky=tk.E + tk.W, padx=5, pady=5)
    tk.Button(group3, text="运行 Power Up Sequence with Max Load", command=lambda: test3('upmax')).grid(row=2, column=0, sticky=tk.E + tk.W, padx=5, pady=5)
    tk.Button(group3, text="运行 Power Down Sequence with Max Load", command=lambda: test3('downmax')).grid(row=3, column=0, sticky=tk.E + tk.W, padx=5, pady=5)

def t4_win():  # T-4 OVS & UDS Sequence Measurement
    global root4
    root4 = tk.Toplevel()  # 创建新的顶层窗口
    root4.title('T-4 OVS & UDS Sequence Measurement')  # 设置窗口标题
    root4.geometry('360x350')  # 设置新窗口尺寸
    root4.transient(root)  # 将root4设置为root的临时窗口
    root4.attributes("-topmost", 1)  # 保证root4窗口在最前面显示

    # 添加说明文本
    tk.Label(root4, text='请将探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
             wraplength=300, anchor='w').grid(row=0, column=0, columnspan=3, padx=20, pady=20)

    # 创建 LabelFrame 来组织测试按钮
    group4 = tk.LabelFrame(root4, text='单项测试', padx=5, pady=5)
    group4.grid(row=1, column=0, columnspan=3, padx=40, pady=15)

    # 按钮绑定到测试函数，使用不同的测试模式
    tk.Button(group4, text="运行 Overshoot with NO Load", command=lambda: test4('ON')).grid(row=1, column=0,
                                                                                            sticky=tk.E + tk.W, padx=40,
                                                                                            pady=5)
    tk.Button(group4, text="运行 Undershoot with NO Load", command=lambda: test4('UN')).grid(row=2, column=0,
                                                                                             sticky=tk.E + tk.W,
                                                                                             padx=40, pady=5)
    tk.Button(group4, text="运行 Overshoot with Max Load", command=lambda: test4('OM')).grid(row=3, column=0,
                                                                                             sticky=tk.E + tk.W,
                                                                                             padx=40, pady=5)
    tk.Button(group4, text="运行 Undershoot with Max Load", command=lambda: test4('UM')).grid(row=4, column=0,
                                                                                              sticky=tk.E + tk.W,
                                                                                              padx=40, pady=5)

    theButton45 = tk.Button(root4, text="开始测试", command=lambda: test4('ALL'))
    theButton45.grid(row=5, column=0, padx=60, pady=5)

    quit41 = tk.Button(root4, text='退出测试', command=root4.destroy, activeforeground='white', activebackground='red')
    quit41.grid(row=5, column=1, padx=50, pady=5)  # 退出按钮的设计

    root4.mainloop()  # 启动事件循环

def t5_win():  # T-5 Switching Fre. & Jitter Measurement
    global root5, phase_value, multi_phase_enabled
    root5 = tk.Toplevel()
    root5.title('T-5 Switching Fre. & Jitter Measurement')
    root5.geometry('380x430')
    root5.transient(root)
    root5.attributes("-topmost", 1)

    tk.Label(root5, text='请将探头一端连接到示波器的四通道，另一端连接到待测VR的SW信号，单击“开始测试”进行测试。',
             wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)

    # 创建 "Phase 个数" 标签
    tk.Label(root5, text="Phase 个数：").grid(row=2, column=0, pady=10, sticky=tk.E)

    # Phase 选项列表
    phase_options = ['1', '2', '3', '4', '5', '6']
    phase_value = tk.StringVar()
    phase_value.set('1')  # 默认值为 1

    # 创建下拉框，初始时禁用
    phase_dropdown = tk.OptionMenu(root5, phase_value, *phase_options)
    phase_dropdown.grid(row=2, column=1, columnspan=1, padx=5, pady=10, sticky=tk.W)
    phase_dropdown.config(state="disabled")  # 初始禁用下拉框

    # 创建 "多项测试" 复选框，启用时才允许选择 Phase
    multi_phase_enabled = tk.StringVar(value="False")
    multi_phase_checkbox = tk.Checkbutton(
        root5, text='多项测试', variable=multi_phase_enabled, onvalue="True",
        offvalue="False", state="normal",
        command=lambda: phase_dropdown.config(
            state="normal" if multi_phase_enabled.get() == "True" else "disabled"
        )
    )
    multi_phase_checkbox.grid(row=2, column=0, columnspan=1, padx=90, pady=10, sticky=tk.W)

    # 单项测试按钮
    group5 = tk.LabelFrame(root5, text='单项测试', padx=5, pady=5)
    group5.grid(row=3, column=0, columnspan=3, padx=40, pady=15)

    tk.Button(group5, text="运行 Switching Freq with Light Load", command=lambda: test5('sl')).grid(
        row=1, column=0, sticky=tk.E + tk.W, padx=40, pady=5
    )
    tk.Button(group5, text="运行 Switching Freq with Heavy Load", command=lambda: test5('sh')).grid(
        row=2, column=0, sticky=tk.E + tk.W, padx=40, pady=5
    )
    tk.Button(group5, text="运行 Jitter with Light Load", command=lambda: test5('jl')).grid(
        row=3, column=0, sticky=tk.E + tk.W, padx=40, pady=5
    )
    tk.Button(group5, text="运行 Jitter with Heavy Load", command=lambda: test5('jh')).grid(
        row=4, column=0, sticky=tk.E + tk.W, padx=40, pady=5
    )

    # 开始测试与退出按钮
    tk.Button(root5, text="开始测试", command=lambda: test5('ALL')).grid(row=4, column=0, padx=5, pady=20)
    tk.Button(root5, text='退出测试', command=root5.destroy, activeforeground='white', activebackground='red').grid(
        row=4, column=1, padx=5, pady=20
    )

def t6_win():                   # T-6 Power MOSFET Gate/Phase Nodes Measurement
    global root6, phase_value, multi_phase_enabled
    root6 = tk.Toplevel()                                                  # 创建一个新的顶层窗口root6，独立于主窗口root，作为当前测试窗口
    root6.title('T-6 Power MOSFET Gate/Phase Nodes Measurement')        # 设置窗口标题
    root6.geometry('410x430')                                           # 设置新窗口尺寸
    root6.transient(root)                                               # 将root6作为root的临时窗口,前者关闭不影响后者
    root6.attributes("-topmost", 1)                                     # root6 最前显示
    tk.Label(root6, text='请使用探头1连接示波器的三通道和待测VR的Vin_FET信号，'
                      '使用探头2连接示波器的四通道和待测VR的SW信号，单击“开始测试”进行测试。', wraplength=310, anchor='w') \
        .grid(row=0, column=0, columnspan=3, padx=20, pady=20)

    tk.Label(root6, text="Phase 个数：").grid(row=2, column=0, pady=10, sticky=tk.E)

    # 选项列表
    phase_options = ['1', '2', '3', '4', '5', '6']
    phase_value = tk.StringVar()
    phase_value.set('1')  # 默认值为 1

    # 创建下拉框，初始时禁用
    phase_dropdown = tk.OptionMenu(root6, phase_value, *phase_options)
    phase_dropdown.grid(row=2, column=1, columnspan=1, padx=5, pady=10, sticky=tk.W)
    phase_dropdown.config(state="disabled")  # 初始禁用下拉框

    # 创建 "多项测试" 复选框，启用时才允许选择 phase
    multi_phase_enabled = tk.StringVar()
    multi_phase_checkbox = tk.Checkbutton(root6, text='多项测试', variable=multi_phase_enabled, onvalue="True",
                                          offvalue="False", state="normal",
                                          command=lambda: phase_dropdown.config(
                                              state="normal" if multi_phase_enabled.get() == "True" else "disabled"))
    multi_phase_checkbox.grid(row=2, column=0, columnspan=1, padx=90, pady=10, sticky=tk.W)
    multi_phase_enabled.set("False")  # 默认状态为 "False"（未启用多项 Phase 模式）


    multi_phase_rb_no = tk.Radiobutton(root6, text="否", variable=multi_phase_enabled, value=False)
    multi_phase_rb_no.grid(row=3, column=2, padx=10, pady=10, sticky=tk.W)
    group6 = tk.LabelFrame(root6, text='单项测试', padx=5, pady=5)
    group6.grid(row=3, rowspan=2, column=0, columnspan=3, padx=20, pady=15)
    tk.Button(group6, text="运行 MOSFET Switching HS-Vds with NO Load", command=lambda: test6('hsno')).grid(row=1, column=0,
                                                                                                    sticky=tk.E + tk.W,
                                                                                                    padx=40,
                                                                                                    pady=5)
    tk.Button(group6, text="运行 MOSFET Switching HS-Vds with Max Load", command=lambda: test6('hsmax')).grid(row=2, column=0,
                                                                                                    sticky=tk.E + tk.W,
                                                                                                    padx=40, pady=5)
    tk.Button(group6, text="运行 MOSFET Switching LS-Vds with NO Load", command=lambda: test6('lsno')).grid(row=3, column=0,
                                                                                            sticky=tk.E + tk.W,
                                                                                            padx=40, pady=5)
    tk.Button(group6, text="运行 MOSFET Switching LS-Vds with Max Load", command=lambda: test6('lsmax')).grid(row=4, column=0,
                                                                                            sticky=tk.E + tk.W,
                                                                                            padx=40, pady=5)

    tk.Button(root6, text="开始测试", command=lambda: test6('ALL')).grid(row=1, column=0, padx=5, pady=20)
    tk.Button(root6, text='退出测试', command=root6.destroy, activeforeground='white', activebackground='red').grid(
        row=1, column=1, padx=5, pady=20)

def t7_win():
    root7 = tk.Toplevel()
    root7.title('T-7 Bode Plots Measurement')
    root7.geometry('340x200')
    root7.transient(root)
    tk.Label(root7, text='该项测试正在开发中，敬请期待...', wraplength=310, anchor='w') \
        .grid(row=0, column=0, columnspan=3, padx=20, pady=20)
    root7.attributes("-topmost", 1)

    tk.Button(root7, text="开始测试", command=test0).grid(row=1, column=0, padx=5, pady=20)
    tk.Button(root7, text='退出测试', command=root7.destroy, activeforeground='white', activebackground='red').grid(
        row=1, column=1, padx=5, pady=20)

def t8_win():
    global root8 ,choice
    root8 = tk.Toplevel()
    root8.title('T-8 Efficiency Measurement')
    root8.geometry('340x200')
    root8.transient(root)
    tk.Label(root8, text='请选择测量模式，\n如需测量近端，则将VOUT1,VIN,Iin连接到电路上，\n如需测试远端，则将VOUT2连接到电路上', wraplength=310, anchor='w') \
        .grid(row=0, column=0, columnspan=3, padx=20, pady=20)
    root8.attributes("-topmost", 1)

    # 创建 IntVar 变量用于保存选择状态
    choice = tk.IntVar(value=0)  # 默认选中近端（值为0）

    # 创建单选框
    tk.Label(root8, text='测量模式：').grid(row=2, column=0, sticky=tk.E)
    tk.Radiobutton(root8, text='近端测试', variable=choice, value=0).grid(row=2, column=1)
    tk.Radiobutton(root8, text='远端测试', variable=choice, value=1).grid(row=2, column=2)


    tk.Button(root8, text="开始测试", command=test8).grid(row=3, column=0, padx=5, pady=20)
    tk.Button(root8, text='退出测试', command=root8.destroy, activeforeground='white', activebackground='red').grid(
        row=3, column=2, padx=5, pady=20)

def t9_win():                   # T-9 OCP & SCP Test
    global root9
    root9 = tk.Toplevel()                                              # 创建一个新的顶层窗口root9，独立于主窗口root，作为当前测试窗口
    root9.title('T-9 OCP & SCP Test')                               # 设置窗口标题
    root9.geometry('350x400')                                       # 设置新窗口尺寸
    root9.transient(root)                                           # 将root9作为root的临时窗口,前者关闭不影响后者
    root9.attributes("-topmost", 1)                                 # root6 最前显示
    tk.Label(root9, text='请使用探头1连接示波器的一通道和待测VR的输出端，探头3连接示波器的三通道和待测VR的PG信号输出端，使用电流探棒1连接示波器的四通道和待测'
                      'VR的输出电流线缆，测试前请先上电，单击开始测试项进行测试。', wraplength=300, anchor='w').grid(row=0, column=0, columnspan=3, padx=20, pady=20)
    tk.Button(root9, text="开始测试", command=lambda: test9('ALL')).grid(row=1, column=0, padx=5, pady=20)
    tk.Button(root9, text='退出测试', command=root9.destroy, activeforeground='white', activebackground='red').grid(
        row=1, column=2, padx=5, pady=20)
    group9 = tk.LabelFrame(root9, text='单项测试', padx=5, pady=5)
    group9.grid(row=2, rowspan=2, column=0, columnspan=3, padx=20, pady=15)



    theButton91 = tk.Button(group9, text="运行 Slow OCP Test", command=lambda: test9("slowocp"))  # 执行 test9("slowocp")
    theButton91.grid(row=1, column=0, sticky=tk.E + tk.W, padx=30, pady=5)

    theButton92 = tk.Button(group9, text="运行 Fast OCP Test", command=lambda: test9("fastocp"))  # 执行 test9("fastocp")
    theButton92.grid(row=2, column=0, sticky=tk.E + tk.W, padx=30, pady=5)

    theButton93 = tk.Button(group9, text="运行 SCP before Power on Test",
                            command=lambda: test9("scpbefore"))  # 执行 test9("scpbefore")
    theButton93.grid(row=3, column=0, sticky=tk.E + tk.W, padx=30, pady=5)

    theButton94 = tk.Button(group9, text="运行 SCP after Power on Test",
                            command=lambda: test9("scpafter"))  # 执行 test9("scpafter")
    theButton94.grid(row=4, column=0, sticky=tk.E + tk.W, padx=30, pady=5)


def main_window():      # 主窗口设计
    global root, EnValue1, EnValue2, EnValue3, EnValue4, entry
    root = tk.Tk()  # 创建主窗口实例
    root.title('Suma Power Test')  # 设置窗口标题
    root.resizable(False, False)  # 禁止调整窗口大小
    root.geometry('730x495')  # 设置窗口大小

    # visa_dll = 'c:/windows/system32/visa32.dll'
    # 加载图片
    image_file = tk.PhotoImage(file='suma.png')
    image = tk.Label(root, image=image_file)
    image.grid(row=0, column=0, columnspan=3, padx=40, pady=20)

    # 设置输入框和标签
    tk.Label(root, text='SheetName:').grid(row=4, column=0, sticky=tk.E)
    EnValue1 = tk.StringVar()
    EnValue1.set('P3V3_OCP1')
    entry = tk.Entry(root, show=None, width=20, textvariable=EnValue1)
    entry.grid(row=4, column=1, columnspan=2)

    tk.Label(root, text='输出电压：').grid(row=5, column=0, sticky=tk.E)
    EnValue2 = tk.StringVar()
    # EnValue2.set(xls.getCell('P3V3_AUX', 5, 10))
    tk.Entry(root, show=None, width=10, textvariable=EnValue2, state='readonly').grid(row=5, column=1)
    tk.Label(root, text='V').grid(row=5, column=1, sticky=tk.E)


    EnValue3 = tk.StringVar()
    EnValue4 = tk.StringVar()
    tk.Entry(root, show=None, textvariable=EnValue3, state='readonly').place(x=140, y=150, width=200, height=30)
    tk.Entry(root, show=None, textvariable=EnValue4, state='readonly').place(x=140, y=190, width=200, height=30)

    # 设置工作模式的单选按钮


    # 设置按钮
    tk.Button(root, text="选择测试报告路径", command=set_excel_path).place(x=10, y=150, width=115, height=30)
    tk.Button(root, text="选择保存图片路径", command=set_pic_path).place(x=10, y=190, width=115, height=30)
    tk.Button(root, text="连接仪器", command=instrument).place(x=70, y=370, width=80, height=30)
    tk.Button(root, text="读取表格", command=go).place(x=210, y=370, width=80, height=30)

    # 创建 LabelFrame 并添加测试按钮
    group = tk.LabelFrame(root, text='POL测试项', padx=5, pady=5)
    group.grid(row=0, rowspan=12, column=3, padx=30, pady=15)

    test_buttons = [
        ("T-0 DMM&Scope Offset Record", t0_win),
        ("T-1 DC Regulation+Ripple&Noise Test", t1_win),
        ("T-2 Loading Transient Response Test", t2_win),
        ("T-3 Power Up & Down Sequence Measurement", t3_win),
        ("T-4 OVS & UDS Sequence Measurement", t4_win),
        ("T-5 Switching Fre. & Jitter Measurement", t5_win),
        ("T-6 Power MOSFET Gate/Phase Nodes Measurement", t6_win),
        ("T-7 Bode Plots Measurement(TBD)", t7_win),
        ("T-8 Efficiency Measurement", t8_win),
        ("T-9 OCP & SCP & OVP Test", t9_win)
    ]

    for i, (text, command) in enumerate(test_buttons):
        tk.Button(group, text=text, command=command).grid(row=i + 1, column=3, sticky=tk.E + tk.W, padx=5, pady=5)

    root.mainloop()  # 启动 Tkinter 的事件循环

# —————————————————————————————分割线———————————————————————————————
# 实验测试函数        每个测试具体的实验步骤

def test0():
    print("TEST0测试开始")
    control_dc_source(vin, Iin, 'ON')
    if MSO5 == 1:
        print("恢复出厂设置...")
        osc.write('FACTORY')
        print("加载设置文件...")
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')

    print("执行公共设置...")
    common_set()

    print("设置通道...")
    tl1_channel_set()

    print("测量...")
    measure1()

    print("设置水平模式...")
    set_horizontal_mode(1e-3, 1e-3, 50, 2e8, 2e6, mode='MANUAL')
    tdc = xls.getCell(entry.get(), 8, 4)
    print(f"获取TD值: {tdc}")
    el.static(9, 'MAX', tdc)  # 选择相应的通道
    el.state('ON')

    print("开始采样...")
    osc.state('run')
    time.sleep(5)

    print("处理RMS值...")
    rms = process('rms')
    print(f"设置偏移: {rms:.2f}")
    osc.write(f'CH1:OFFSET {rms:.2f}')

    time.sleep(3)
    print("设置尺度...")
    scale()
    time.sleep(2)

    print("再次处理RMS值...")
    rms = process('rms')
    print("触发示波器...")
    osc.trigger('AUTO', 'CH1', 'RISE', rms)
    time.sleep(1)

    print("停止采样...")
    osc.state('stop')
    print("重置显示...")
    osc.write('DISPLAY:PERSISTENCE:RESET')
    set_osc()

    print("获取最后的RMS值...")
    rms = process('rms')
    xls.setCell(entry.get(), 8, 8, rms)
    time.sleep(1)

    print("保存测试数据...")
    test_save(0, 1, 'N1', 25, 0, 337, 212)
    xls.save()  # 保存 Excel 文件

    print("TEST0示波器测试结束")
    time.sleep(1)
    inquery = messagebox.askquestion(title='程序执行完毕', message='TEST0示波器测试结束，'
                    '使用六位半数字万用表校准电压值来完成TEST0测试, 如需继续进行TEST1测试请点击是')
    if inquery == 'yes':
        xls.save()  # 保存 Excel 文件
        el.state('OFF')
        root0.destroy()
        t1_win()
        test1()
    else:
        return 0

def test1():
    print("TEST1测试开始")
    control_dc_source(vin, Iin, 'ON')
    if MSO5 == 1:
        print("恢复到出厂设置并加载预设配置文件")
        osc.write('FACTORY')  # 恢复到出厂设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')  # 加载预设的配置文件

    print("应用通用设置")
    common_set()
    print("配置通道1")
    tl1_channel_set()
    print("启动测量")
    measure1()

    print("设置水平模式")
    set_horizontal_mode(1e-3, 1e-3, 50, 2e8, 2e6, mode='MANUAL')

    ldsp = 0.2 * ld_max
    i = 1
    load = 0
    block = 23

    while block <= 28:
        print(f"开始处理块 {block}")
        el.static(9, 'MAX', load)
        el.state('ON')
        osc.state('run')

        time.sleep(2)
        rms = process('rms')
        osc.write(f'CH1:OFFSET {rms:.2f}')

        time.sleep(3)
        print("调整比例")
        scale()

        time.sleep(3)
        rms = process('rms')
        time.sleep(2)
        osc.trigger('AUTO', 'CH1', 'RISE', rms)

        time.sleep(1)
        osc.state('stop')

        osc.write('DISPLAY:PERSISTENCE:RESET')  # 清除显示
        print("重置显示")
        set_osc()

        time.sleep(2)
        rms = process('rms')
        print(f"记录块 {block} 的 RMS 值: {rms}")
        xls.setCell(entry.get(), block, 4, rms)

        time.sleep(1)
        pk2pk = process('pk2pk')
        print(f"记录块 {block} 的峰峰值: {pk2pk}")
        time.sleep(1)
        xls.setCell(entry.get(), block, 5, pk2pk)

        image_index = {1: 'M18', 4: 'U18', 6: 'AC18'}
        picture_cell = image_index.get(i, 0)  # 如果 i 不在 image_index 中，默认使用 0

        if i in [1, 4, 6]:
            print(f"保存测试图像 {picture_cell}")
            test_save(1, i, picture_cell, 36, 10, 358, 238)
        elif i in [2, 3, 5]:
            print(f"保存测试图像 {picture_cell}（无附加信息）")
            test_save(1, i, picture_cell, 36, 10, 358, 238, add=False)

        load += ldsp
        block += 1
        i += 1
        time.sleep(1)
    print("获取最后的RMS值...")
    rms = process('rms')
    xls.setCell(entry.get(), 8, 8, rms)
    time.sleep(1)

    print("保存 Excel 文件")
    print("TEST1测试结束")
    xls.save()  # 保存 Excel 文件
    time.sleep(1)
    auto_close_messagebox(root, '程序执行完毕', 'TEST1测试完成，现在开始计时,六分钟后完成test10温度的拍摄')
    start_countdown(6)

def test2(type):
    print("TEST2测试开始")
    control_dc_source(vin, Iin, 'ON')
    test_modes = {
        'Low': {
            'horizontal_params': (2e-3, 2e-3, 55, 2e8),
            'dynamic_param': 2.5,
            'index': 1
        },
        'Mid': {
            'horizontal_params': (5e-4, 5e-4, 55, 2e8),
            'dynamic_param': 0.5,
            'index': 2
        },
        'High': {
            'horizontal_params': (50e-6, 50e-6, 55, 2e8),
            'dynamic_param': 0.05,
            'index': 3
        }
    }
    if type != 'ALL':
        mode_params = test_modes[type]
        horizontal_params = mode_params['horizontal_params']
        dynamic_param = mode_params['dynamic_param']

    if MSO5 == 1:
        osc.write('FACTORY')  # 恢复到出厂设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')  # 加载预设的配置文件
        print("加载预设的配置文件完成")

    # 通用设置
    common_set()
    print("通用设置已完成")

    measure2()
    print("测量设置已完成")

    tl2_channel_set()
    print("通道设置已完成")

    i = 1
    while i <= 3:
        if type == 'ALL':
            print(f"第 {i} 次测试开始")  # 测试开始提示

            # 根据不同情况设置模式
            if i == 1:
                set_horizontal_mode(2e-3, 2e-3, 55, 2e8)
                print('正在设置水平位置')
                el.dynamic(1, ld_max, 2.5)
                print('正在设置电子负载模式')
                time.sleep(2)
                print("模式 1 设置完成")
            elif i == 2:
                set_horizontal_mode(5e-4, 5e-4, 55, 2e8)
                el.dynamic(1, ld_max, 0.5)
                time.sleep(2)
                print("模式 2 设置完成")
            elif i == 3:
                set_horizontal_mode(50e-6, 50e-6, 55, 2e8)
                el.dynamic(1, ld_max, 0.05)
                time.sleep(2)
                print("模式 3 设置完成")


        else:
            i = mode_params['index']
            print(f"执行单项测试：{type} 模式开始")
            set_horizontal_mode(*horizontal_params)
            el.dynamic(1, ld_max, dynamic_param)
            time.sleep(2)
            print(f"{type} 模式设置完成")

        el.state('ON')
        time.sleep(2)
        print("负载已开启")

        osc.state('run')  # 示波器开始采样
        print("示波器开始试采样")

        scale()
        print("比例设置完成")

        osc.write('DISPLAY:PERSISTENCE:RESET')  # 清除显示
        print("显示持久性已清除")


        osc.state('stop')
        set_osc()
        print("示波器采样结束")
        time.sleep(1)

        el.state('OFF')
        el.static(9, 'MAX', 0)  # 选择相应通道
        print("负载关闭")
        time.sleep(1)

    # 获取和记录最大值
        y1 = {1: 3, 2: 15, 3: 27}
        picture_position = y1.get(i)
        ch1max = process('ch1max')
        xls.setCell(entry.get(), 55, picture_position, ch1max)
        print(f"记录 CH1 最大值: {ch1max} 于位置 {picture_position}")
        time.sleep(1)
        # 获取和记录最小值

        y2 = {1: 4, 2: 16, 3: 28}
        picture_position = y2.get(i)
        ch1min = process('ch1min')
        xls.setCell(entry.get(), 55, picture_position, ch1min)
        print(f"记录 CH1 最小值: {ch1min} 于位置 {picture_position}")
        time.sleep(1)

    # 保存测试结果图像
        image_index = {1: 'F40', 2: 'R40', 3: 'AD40'}
        picture_cell = image_index.get(i, 0)
        test_save(2, i, picture_cell, 36, 10, 352, 250)
        print(f"保存图像于单元格 {picture_cell}")
        print(f"i的值为{i}")
        if type == 'ALL':
            i += 1
        else:
            break

    xls.save()  # 保存 Excel 文件
    print("Excel 文件已保存")
    print("TEST2测试结束")
    auto_close_messagebox(root,'程序执行完毕', 'TEST2测试完成，若继续进行TEST4测试，请进行电路下电')
    control_dc_source(vin, Iin, 'OFF')
    if root:
        root2.destroy()
    test4('ALL')

def test3(type):
    print("TEST3测试开始")
    test_modes = {
        'upno': {
            'horizontal_params': (8e-3, 8e-3, 60, 2e8),  # 水平模式参数
            'trigger': 'RISE',  # 触发边沿
            'load': None,  # 不使用负载参数
            'index': 1  # 模式索引
        },
        'downno': {
            'horizontal_params': (8e-3, 1e-2, 60, 2e8),  # 水平模式参数
            'trigger': 'FALL',  # 触发边沿
            'load': None,  # 不使用负载参数
            'index': 2  # 模式索引
        },
        'upmax': {
            'horizontal_params': (8e-3, 1e-2, 60, 2e8),  # 水平模式参数
            'trigger': 'RISE',  # 触发边沿
            'load': (1, 'MIN', ld_max, True),  # 负载设置参数
            'index': 3  # 模式索引
        },
        'downmax': {
            'horizontal_params': (8e-3, 1e-2, 60, 2e8),  # 水平模式参数
            'trigger': 'FALL',  # 触发边沿
            'load': (1, 'MIN', ld_max, False),  # 负载设置参数
            'index': 4  # 模式索引
        }
    }

    if type != 'ALL':
        mode_params = test_modes[type]
        horizontal_params = mode_params['horizontal_params']  # 获取水平设置
        trigger_type = mode_params['trigger']  # 获取触发类型（RISE 或 FALL）
        load_params = mode_params['load']  # 获取负载设置

        # 设置触发
        osc.trigger('NORMAL', 'CH1', trigger_type, volt / 2)  # 根据模式类型设置触发条件

    if MSO5 == 1:
        print("恢复到出厂设置")
        osc.write('FACTORY')  # 恢复到出厂设置
        print("加载预设的配置文件")
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')  # 加载预设的配置文件
        time.sleep(3)

    print("进行通用设置")
    common_set()

    print("进行通道设置")
    tl3_channel_set()

    print("进行测量设置")
    measure3()
    time.sleep(3)
    i = 1
    while i <= 4:
        print(f"第 {i} 次测试开始")  # 提示测试开始
        if type == 'ALL':
            if i == 1:
                set_horizontal_mode(8e-3, 8e-3, 60, 2e8)
                osc.trigger('NORMAL', 'CH1', 'RISE',volt / 2)
            elif i == 2:
                set_horizontal_mode(8e-3, 8e-3, 60, 2e8)
                osc.trigger('NORMAL', 'CH1', 'FALL',volt / 2)
            elif i == 3:
                set_horizontal_mode(8e-3, 8e-3, 60, 2e8)
                set_load(1, 'MIN', ld_max, lvp_on=True)
                osc.trigger('NORMAL', 'CH1', 'RISE',volt / 2)
            elif i == 4:
                set_horizontal_mode(8e-3, 8e-3, 60, 2e8)
                set_load(1, 'MIN', ld_max, lvp_on=False)
                osc.trigger('NORMAL', 'CH1', 'FALL',volt / 2)
        else:
            i = mode_params['index']
            print(f"执行单项测试：{type} 模式开始")
            set_horizontal_mode(*horizontal_params)
            if load_params is not None:
                set_load(*load_params)
            osc.trigger('NORMAL', 'CH1', trigger_type, volt / 2)  # 根据模式类型设置触发条件
        print("设置单步触发")
        osc.state('single')  # 设置单步触发
        print("等待 3 秒，以准备主板上下电")
        time.sleep(3)  # 设置延时为主板上下电作准备
        if i == 1 or i == 3:
            control_dc_source(vin, Iin, 'ON')
        if i == 2 or i == 4:
            control_dc_source(vin, Iin, 'OFF')

        time.sleep(1)
        # 保存测试结果图像
        image_index = {1: 'F67', 2: 'F86', 3: 'R67', 4: 'R86'}
        picture_cell = image_index.get(i, 0)
        print(f"保存测试图像于单元格 {picture_cell}")
        test_save(3, i, picture_cell, 36, 10, 362, 220)
        print(f"第 {i} 次测试完成\n")  # 提示测试完成
        if type == 'ALL':
            i += 1
        else:
            break

    print("保存 Excel 文件")
    xls.save()  # 保存 Excel 文件

    # 完成提示
    messagebox.showinfo(title='程序执行完毕', message='TEST3测试完成，如需进行TEST5测试，请更换CH4通道为TDP0500无源单端探头点接电路')
    root3.destroy()
    print("TEST3测试结束")

def test4(type):
    print("TEST4测试开始")

    test_modes = {
        'ON': {
            'measure': 'RISE',
            'horizontal_params': (1e-3, 1e-3, 50, 1e8),
            'trigger': 'RISE',  # 只保留 RISE
            'process': 'ch1max',
            'xls_position': (111, 3),
            'load': None,  # No load setup for ON
            'index': 1  # 添加索引值
        },
        'UN': {
            'measure': 'FALL',
            'horizontal_params': (1e-2, 1e-1, 30, 1e6),
            'trigger': 'FALL',  # 只保留 FALL
            'process': 'ch1min',
            'xls_position': (130, 3),
            'load': None,  # No load setup for UN
            'index': 2  # 添加索引值
        },
        'OM': {
            'measure': 'RISE',
            'horizontal_params': (1e-3, 1e-3, 50, 1e8),
            'trigger': 'RISE',  # 只保留 RISE
            'process': 'ch1max',
            'xls_position': (111, 15),
            'load': (1, 'MIN', ld_max, True),  # Set load for OM
            'index': 3  # 添加索引值
        },
        'UM': {
            'measure': 'FALL',
            'horizontal_params': (1e-3, 1e-1, 40, 1e6),
            'trigger': 'FALL',  # 只保留 FALL
            'process': 'ch1min',
            'xls_position': (130, 3),
            'load': (1, 'MIN', ld_max, False),  # Set load for UM
            'index': 4  # 添加索引值
        }
    }
    if type != 'ALL':
        mode_params = test_modes[type]
        horizontal_params = mode_params['horizontal_params']  # 获取水平设置
        trigger_type = mode_params['trigger']  # 获取触发类型（RISE 或 FALL）
        process_type = mode_params['process']  # 获取测量类型
        xls_position = mode_params['xls_position']  # 获取Excel保存位置
        load_params = mode_params['load']  # 获取负载设置

    if MSO5 == 1:
        osc.write('FACTORY')  # 恢复到出厂设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')  # 加载预设的配置文件

    common_set()
    tl4_channel_set()

    i = 1
    while i <= 4:
        if type == 'ALL':
            print(f"第 {i} 次测试开始")  # 测试开始提示
            if i == 1:
                measure4('RISE')
                set_horizontal_mode(1e-3, 1e-3, 50, 1e8)
                osc.trigger('NORMAL', 'CH1', 'RISE', volt / 2)
                osc.state('single')  # 设置单步触发
                time.sleep(2)  # 设置延时为主板上下电作准备


            elif i == 2:
                measure4('FALL')
                set_horizontal_mode(1e-2, 1e-2, 30, 1e6)
                osc.trigger('NORMAL', 'CH1', 'FALL', volt / 2)
                osc.state('single')  # 设置单步触发
                time.sleep(2)  # 设置延时为主板上下电作准备


            elif i == 3:
                measure4('RISE')
                set_load(1, 'MIN', ld_max, lvp_on=True)
                set_horizontal_mode(1e-3, 1e-3, 50, 1e8)
                osc.trigger('NORMAL', 'CH1', 'RISE', volt / 2)
                osc.state('single')  # 设置单步触发
                time.sleep(2)  # 设置延时为主板上下电作准备


            elif i == 4:
                measure4('FALL')
                set_load(1, 'MIN', ld_max, lvp_on=False)
                set_horizontal_mode(1e-2, 1e-2, 40, 1e6)
                osc.trigger('NORMAL', 'CH1', 'FALL', volt / 2)
                osc.state('single')  # 设置单步触发
                time.sleep(2)  # 设置延时为主板上下电作准备

        else:
            i = mode_params['index']
            measure4(mode_params['measure'])
            if load_params is not None:
                set_load(*load_params)
            # 设置水平模式
            set_horizontal_mode(*horizontal_params)

            # 设置触发
            osc.trigger('NORMAL', 'CH1', trigger_type, volt / 2)  # 根据模式类型设置触发条件

            # 设置单步触发
            osc.state('single')
            time.sleep(2)

        if i == 1 or i == 3:
            control_dc_source(vin, Iin, 'ON')
        if i == 2 or i == 4:
            control_dc_source(vin, Iin, 'OFF')
        time.sleep(3)
        if type != 'ALL':
            ch_value = process(process_type)
            xls.setCell(entry.get(), *xls_position, ch_value)
        else:
            if i == 1:
                ch1max = process('ch1max')
                xls.setCell(entry.get(), 111, 3, ch1max)

            elif i == 2:
                ch1min = process('ch1min')
                xls.setCell(entry.get(), 130, 3, ch1min)
            if i == 3:
                set_load(load=False)
                print("关闭电子负载的低电压保护 LVP模式")
                ch1max = process('ch1max')
                xls.setCell(entry.get(), 111, 15, ch1max)
            elif i == 4:
                el.state('OFF')
                print("关闭负载")
                ch1min = process('ch1min')
                xls.setCell(entry.get(), 130, 15, ch1min)
        time.sleep(3)
        # 保存测试结果图像
        image_index = {1: 'F108', 2: 'F127', 3: 'R108', 4: 'R127'}
        picture_cell = image_index.get(i, 0)
        print(f"保存测试图像于单元格 {picture_cell}")
        test_save(4, i, picture_cell, 36, 10, 361, 223)
        print(f"第 {i} 次测试完成\n")  # 提示测试完成
        print(f"i的值为{i}")
        if type == 'ALL':
            i += 1
        else:
            break


    print("保存 Excel 文件")
    xls.save()  # 保存 Excel 文件

    print("TEST4测试结束")
    auto_close_messagebox(root, '程序执行完毕', 'TEST4测试完成')
    if root4:
        root4.destroy()
    else:
        return 1

def test5(type):
    global root5, table_type, phase, multi_phase_enabled
    if multi_phase_enabled.get() == "True":
        phase = int(phase_value.get())  # 假设 phase_value 是界面上的输入框变量
        print(f'phase的值为{phase}')
        table_type = '表格2'
        print('表格2')
    else:
        table_type = '表格1'
        print('表格1')
    print("TEST5测试开始")
    test_modes = {
        'sl': {
            'measure': 'FRE',
            'process': 'ch4freq',
            'xls_position': (152, 3),
            'load': None,  # No load setup for ON
            'cursor': None,
            'index': 1  # 添加索引值
        },
        'sh': {
            'measure': 'FRE',
            'process': 'ch4freq',
            'xls_position': (152, 15),
            'load': (9, 'MIN', ld_max),  # No load setup for UN
            'cursor': None,
            'index': 2  # 添加索引值
        },
        'jl': {
            'measure': 'PWIDTH',
            'process': 'poswidth',
            'xls_position': (175, 3),
            'load': None,  # Set load for OM
            'cursor': True,
            'index': 3  # 添加索引值
        },
        'jh': {
            'measure': 'PWIDTH',
            'process': 'poswidth',
            'xls_position': (175, 15),
            'load': (9, 'MIN', ld_max),  # Set load for UM
            'cursor': True,
            'index': 4  # 添加索引值
        }
    }
    if type !='ALL':
        print('进行单项测试')
        mode_params = test_modes[type]
        i = mode_params['index']
        process_type = mode_params['process']  # 获取测量类型
        xls_position = mode_params['xls_position']  # 获取Excel保存位置
        load_params = mode_params['load']  # 获取负载设置
        cursor_params = mode_params['cursor']
    else:
        print('进行连续四项测试或者多项测试')
        i = 1

    if MSO5 == 1:
        osc.write('FACTORY')  # 恢复到出厂设置
        if i == 1 or i == 2:
            osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')  # 加载预设的配置文件
        if i == 3 or i == 4:
            osc.write('RECALL:SETUP "C:/Tektronix/Tek000_002.SET"')  # 加载预设的配置文件

    common_set()
    tl5_channel_set()
    set_horizontal_mode(5e-7, 5e-7, 10, 1e10)

    if table_type == '表格1':
        while i <= 4:
            if i == 3 or i == 4:
                osc.persistence('INFPersist')  # 开启累积
            print(f"第 {i} 次测试开始")  # 测试开始提示
            if type == 'ALL':

                if i == 1:
                    measure5('FRE')
                    set_osc()
                    connect = messagebox.askquestion(title='电路接switch',
                                                     message='请进行电路连接，连接成功请点击是确认存图，失败请点击否')
                    if connect == 'yes':
                        print('数据处理')
                        ch4freq = process('ch4freq')
                        xls.setCell(entry.get(), 152, 3, ch4freq)
                    else:
                        return 0

                elif i == 2:
                    measure5('FRE')
                    el.static(9, 'MAX', ld_max-1)
                    el.state('ON')
                    set_osc()
                    connect = messagebox.askquestion(title='电路接switch',
                                                     message='请进行电路连接，连接成功请点击是确认存图，失败请点击否')
                    if connect == 'yes':
                        el.state('OFF')
                        ch4freq = process('ch4freq')
                        xls.setCell(entry.get(), 152, 15, ch4freq)
                    else:
                        return 0

                elif i == 3:
                    measure5('PWIDTH')
                    set_record_length()
                    set_cursor()
                    set_osc()
                    connect = messagebox.askquestion(title='电路接switch',
                                                     message='请进行电路连接，连接成功请点击是确认存图，失败请点击否')
                    if connect == 'yes':
                        if MSO5:
                            a = set_cursor_position('a')
                            print(f"a的position是{a}")
                            b = set_cursor_position('b')
                            print(f"b的position是{b}")
                            xls.setCell(entry.get(), 173, 3, a)
                            xls.setCell(entry.get(), 173, 4, b)
                            time.sleep(1)
                            poswidth = process('poswidth')
                        else:
                            text = osc.query('CURSOR:VBArs?')
                            print(f'{text}')
                            split_text = text.split(";")
                            a, b, c = split_text
                            b = float(b.strip()) * 1000000000
                            c = float(c.strip()) * 1000000000
                            xls.setCell(entry.get(), 173, 3, b)
                            xls.setCell(entry.get(), 173, 4, c)
                            poswidth = process('poswidth')
                        xls.setCell(entry.get(), 175, 3, poswidth)
                    else:
                        return 0

                elif i == 4:
                    measure5('PWIDTH')
                    set_record_length()
                    el.static(9, 'MAX', ld_max-1)
                    el.state('ON')
                    set_cursor()
                    set_osc()
                    connect = messagebox.askquestion(title='电路接switch',
                                                     message='请进行电路连接，连接成功请点击是确认存图，失败请点击否')
                    if connect == 'yes':
                        el.state('OFF')
                        if MSO5:
                            print("MSO5 模式已启用。")
                            a = set_cursor_position('a')
                            b = set_cursor_position('b')
                            print(f"在 MSO5 模式下设置的光标位置：a = {a}, b = {b}")

                            xls.setCell(entry.get(), 173, 15, a)
                            xls.setCell(entry.get(), 173, 16, b)
                            time.sleep(1)

                            poswidth = process('poswidth')
                            print(f"在 MSO5 模式下处理的 poswidth 值：{poswidth}")
                        else:
                            text = osc.query('CURSOR:VBARS?')
                            print(f"从示波器查询得到的光标位置数据：{text}")

                            split_text = text.split(";")
                            a, b, c = split_text
                            b = float(b.strip()) * 1000000000
                            c = float(c.strip()) * 1000000000
                            print(f"解析并转换后的光标位置：b = {b} 纳秒, c = {c} 纳秒")

                            xls.setCell(entry.get(), 173, 15, b)
                            xls.setCell(entry.get(), 173, 16, c)

                            poswidth = process('poswidth')
                            print(f"处理后的 poswidth 值：{poswidth}")

                        xls.setCell(entry.get(), 175, 15, poswidth)
                        print("最终的 poswidth 值已写入 Excel 表格。")
                    else:
                        return 0
            else:  # 如果是单项测试
                measure5(mode_params['measure'])
                if load_params is not None:
                    el.state('ON')
                    el.static(*load_params)
                if cursor_params:
                    set_cursor()
                set_osc()
                connect = messagebox.askquestion(title='图片保存确认',
                                                 message='若波形正确请点击确认存图，失败请点击否')
                if connect == 'yes':
                    if load_params is not None:
                        el.state('OFF')
                    if type == 'sl' or type == 'sh':
                        ch_value = process(process_type)
                    cursor_mapping = {
                        'jl': {'col_a': 3, 'col_b': 4},
                        'jh': {'col_a': 15, 'col_b': 16}
                    }

                    if type in cursor_mapping:
                        col_a, col_b = cursor_mapping[type]['col_a'], cursor_mapping[type]['col_b']

                        if MSO5:
                            a = set_cursor_position('a')
                            b = set_cursor_position('b')
                            xls.setCell(entry.get(), 173, col_a, a)
                            xls.setCell(entry.get(), 173, col_b, b)
                            time.sleep(1)
                            ch_value = process(process_type)
                        else:
                            osc.write('CURSOR:LINESTYLE DASHed')
                            text = osc.query('CURSOR:VBArs?')
                            split_text = text.split(";")
                            a, b, c = split_text
                            a = float(b.strip()) * 1e9
                            b = float(c.strip()) * 1e9
                            xls.setCell(entry.get(), 173, col_a, a)
                            xls.setCell(entry.get(), 173, col_b, b)
                            ch_value = process(process_type)

                    xls.setCell(entry.get(), *xls_position, ch_value)
            time.sleep(1)
            image_index = {1: 'F149', 2: 'R149', 3: 'F168', 4: 'R168'}
            picture_cell = image_index.get(i, 0)
            test_save(5, i, picture_cell, 36, 10, 365, 225)
            print(f"第 {i} 次测试完成\n")  # 提示测试完成
            print(f"i的值为{i}")
            if type == 'ALL':
                i += 1
            else:
                break

    elif table_type == '表格2':
        base_col = 3  # 基准列
        base_row = 152  # 基准行，201对应 ch3max
        base_picture_col = 6  # 第一张图片的基准位置
        col_offset = base_col + (phase - 1) * 24
        picture_offset = base_picture_col + (phase - 1) * 24  # 每个 phase 增加 24
        row1 = 149  # 第一张图片的行
        row2 = 168  # 第二张图片的行

        while i <= 4:
            print(f"第 {i} 次测试开始")  # 测试开始提示
            if i == 3 or i == 4:
                osc.persistence('INFPersist')  # 开启累积
            # 测量设置
            if i == 1 or i == 2:
                measure5('FRE')  # 测量频率
            elif i == 3 or i == 4:
                measure5('PWIDTH')  # 测量脉宽
                set_record_length()
                set_cursor()

            set_osc()  # 设置示波器

            # 拉载操作（i == 2 和 i == 4 都需要拉载）
            if i == 2 or i == 4:
                el.static(9, 'MAX', ld_max)  # 拉载处理
                el.state('ON')  # 开启电源
                print("拉载开始，负载已开启")

                # 电路连接确认
            connect = messagebox.askquestion(
                title='电路接switch',
                message='请进行电路连接，连接成功请点击是确认存图，失败请点击否'
            )

            if connect == 'yes':
                if i == 1:
                    # 测量频率
                    ch4freq = process('ch4freq')
                    xls.setCell(entry.get(), base_row, col_offset, ch4freq)
                    print(f'保存数据到列 {col_offset}: ch4freq={ch4freq}')

                    # 保存第一张图片
                    cell_address_1 = generate_excel_address(picture_offset, row1)
                    test_save(5, phase * 4 - 3, cell_address_1, 10, 10, 362, 220)
                    print(f"第一张图片保存到单元格 {cell_address_1}")

                elif i == 2:
                    # 测量频率
                    ch4freq = process('ch4freq')
                    xls.setCell(entry.get(), base_row, col_offset + 12, ch4freq)
                    print(f'保存数据到列 {col_offset + 12}: ch4freq={ch4freq}')

                    # 保存第二张图片
                    cell_address_2 = generate_excel_address(picture_offset + 12, row1)
                    test_save(5, phase * 4 - 2, cell_address_2, 10, 10, 362, 220)
                    print(f"第二张图片保存到单元格 {cell_address_2}")

                elif i == 3:
                    # 测量脉宽并计算游标
                    if MSO5:
                        a = set_cursor_position('a')
                        print(f"a的position是{a}")
                        b = set_cursor_position('b')
                        print(f"b的position是{b}")
                        val1 = a  # 第一种情况写入 a
                        val2 = b  # 第一种情况写入 b
                        time.sleep(1)
                        poswidth = process('poswidth')

                    else:
                        text = osc.query('CURSOR:VBArs?')
                        print(f'游标返回文本: {text}')
                        split_text = text.split(";")
                        _, b, c = split_text
                        b = float(b.strip()) * 1e9  # 转换为纳秒
                        c = float(c.strip()) * 1e9  # 转换为纳秒
                        val1 = b  # 第二种情况写入 b
                        val2 = c  # 第二种情况写入 c
                        time.sleep(1)
                        poswidth = process('poswidth')

                    # 保存数据
                    xls.setCell(entry.get(), base_row + 21, col_offset, val1)  # 写入 173 行，基准列
                    xls.setCell(entry.get(), base_row + 21, col_offset + 1, val2)  # 写入 173 行，基准列
                    xls.setCell(entry.get(), base_row + 23, col_offset, poswidth)  # 写入 175 行，基准列


                    # 保存第一张图片
                    cell_address_1 = generate_excel_address(picture_offset, row2)
                    test_save(5, phase * 4 - 1, cell_address_1, 10, 10, 362, 220)  # 保存第一张图片
                    print(f"第一张图片保存到单元格 {cell_address_1}")

                elif i == 4:

                    # 测量脉宽并计算游标
                    if MSO5:
                        a = set_cursor_position('a')  # 获取游标a的值
                        print(f"a的position是{a}")
                        b = set_cursor_position('b')  # 获取游标b的值
                        print(f"b的position是{b}")
                        val1 = a  # 第二种情况写入 b
                        val2 = b  # 第二种情况写入 c
                        time.sleep(1)
                        poswidth = process('poswidth')
                    else:
                        # 如果不是MSO5，通过示波器查询游标值
                        text = osc.query('CURSOR:VBArs?')  # 查询游标的返回值
                        print(f'游标返回文本: {text}')
                        split_text = text.split(";")
                        _, b, c = split_text
                        b = float(b.strip()) * 1e9  # 转换为纳秒
                        c = float(c.strip()) * 1e9  # 转换为纳秒
                        val1 = b  # 第二种情况写入 b
                        val2 = c  # 第二种情况写入 c
                        time.sleep(1)
                        poswidth = process('poswidth')

                    # 保存数据
                    xls.setCell(entry.get(), base_row + 21, col_offset + 12, val1)  # 写入 173 行，基准列 + 12
                    xls.setCell(entry.get(), base_row + 21, col_offset + 13, val2)  # 写入 173 行，基准列 + 12
                    xls.setCell(entry.get(), base_row + 23, col_offset + 12, poswidth)  # 写入 175 行，基准列 + 12


                    # 保存第二张图片
                    cell_address_2 = generate_excel_address(picture_offset + 12, row2)
                    test_save(5, phase * 4 , cell_address_2, 10, 10, 362, 220)  # 保存第二张图片
                    print(f"第二张图片保存到单元格 {cell_address_2}")

            else:
                messagebox.showerror(title='错误', message='波形错误，图片保存失败，请退出重试')
                return 0
            i += 1
            # 结束时关闭电源
            el.state('OFF')  # 关闭电源

    xls.save()  # 保存 Excel 文件
    print("TEST5测试结束")

    auto_close_messagebox(root5, '程序执行完毕', 'TEST5测试完成，如需继续进行TEST6测试，请更换CH3通道为TDP0500连接VIN')

    root5.destroy()

def test6(type):
    global root6, table_type, phase, multi_phase_enabled

    if multi_phase_enabled.get() == "True":
        phase = int(phase_value.get())  # 假设 phase_value 是界面上的输入框变量
        print(f'phase的值为{phase}')
        table_type = '表格2'
        print('表格2')
    else:
        table_type = '表格1'
        print('表格1')
    print("TEST6测试开始")
    test_modes = {
        'hsno': {
            'load': None,  # No load setup for ON
            'index': 1  # 添加索引值
        },
        'hsmax': {
            'load': (1, 'MAX', ld_max),  # No load setup for UN
            'index': 2  # 添加索引值
        },
        'lsno': {
            'load': None,  # Set load for OM
            'index': 3  # 添加索引值
        },
        'lsmax': {
            'load': (1, 'MAX', ld_max),  # Set load for UM
            'index': 4  # 添加索引值
        }
    }
    if type !='ALL':
        mode_params = test_modes[type]
        load_params = mode_params['load']  # 获取负载设置
    if MSO5 == 1:
        osc.write('FACTORY')  # 恢复到出厂设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_004.SET"')  # 加载预设的配置文件
    common_set()
    tl6_channel_set()
    measure6()
    set_horizontal_mode(5e-7, 5e-7, 50, 1e10)


    i = 1
    if table_type == '表格1':
        while i <= 4:
            if type == 'ALL':
                if i == 2 or i == 4:
                    el.static(9, 'MAX', ld_max)
                    el.state('ON')
                    time.sleep(2)
            else:
                i = mode_params['index']
                if load_params is not None:
                    set_load(*load_params)
            set_osc()
            connect = messagebox.askquestion(title='图片保存确认',
                                             message='若波形正确请点击确认存图，失败请点击否')
            if connect == 'yes':
                ch4max = process('ch4max')
                if i == 1:
                    xls.setCell(entry.get(), 198, 4, ch4max)
                elif i == 2:
                    xls.setCell(entry.get(), 198, 16, ch4max)
                elif i == 3:
                    xls.setCell(entry.get(), 217, 4, ch4max)
                elif i == 4:
                    xls.setCell(entry.get(), 217, 16, ch4max)
                el.state('OFF')
                time.sleep(1)
                image_index = {1: 'F190', 2: 'R190', 3: 'F209', 4: 'R209'}
                picture_cell = image_index.get(i, 0)
                test_save(6, i, picture_cell, 36, 10, 362, 220)
                if type == 'ALL':
                    i += 1
                else:
                    break
            else:
                messagebox.showerror(title='错误', message='波形错误，图片保存失败，请退出重试')
                return 0

    elif table_type == '表格2':
        base_col = 4  # 基准列
        base_row = 201  # 基准行，201对应 ch3max
        base_picture_col = 6  # 第一张图片的基准位置
        # 根据 phase 计算具体的偏移量
        col_offset = base_col + (phase - 1) * 24
        picture_offset = base_picture_col + (phase - 1) * 24  # 每个 phase 增加12

        set_osc()
        connect = messagebox.askquestion(title='图片保存确认',
                                         message='若波形正确请点击确认存图，失败请点击否')
        if connect == 'yes':
            # 获取数据值
            ch3max = process('ch3max')
            ch4max = process('ch4max')
            ch1_m1_max = process('ch_m1_max')

            # 设置单元格值，按照对应的 phase 位置写入数据
            xls.setCell(entry.get(), base_row, col_offset, ch3max)
            xls.setCell(entry.get(), base_row + 1, col_offset, ch4max)
            xls.setCell(entry.get(), base_row + 2, col_offset, ch1_m1_max)

            # 输出提示信息
            print(f'保存数据到列 {col_offset}: ch3max={ch3max}, ch4max={ch4max}, ch1_m1_max={ch1_m1_max}')

            row, col = 190, picture_offset  # 第一张图片的行列坐标
            letters = ""
            while col > 0:
                col, remainder = divmod(col - 1, 26)
                letters = chr(65 + remainder) + letters
            cell_address_1 = f"{letters}{row}"  # 生成单元格地址，例如 "BB190"

            test_save(6, phase * 2 - 1, cell_address_1, 10, 10, 362, 220)  # 保存第一张图片

            # 输出图片保存提示
            print(f'第一张图片保存到单元格 {cell_address_1}')
            print('第一张图片和数据保存完毕')
        else:
            messagebox.showerror(title='错误', message='波形错误，图片保存失败，请退出重试')
            return 0

        el.static(9, 'MAX', ld_max)
        el.state('ON')
        time.sleep(2)
        print('MAX_load拉载成功')

        set_osc()
        ch3max = process('ch3max')
        ch4max = process('ch4max')
        ch1_m1_max = process('ch_m1_max')

        # 设置下一个位置的值
        xls.setCell(entry.get(), base_row, col_offset + 12, ch3max)
        xls.setCell(entry.get(), base_row + 1, col_offset + 12, ch4max)
        xls.setCell(entry.get(), base_row + 2, col_offset + 12, ch1_m1_max)

        # 输出提示信息
        print(f'保存数据到列 {col_offset + 12}: ch3max={ch3max}, ch4max={ch4max}, ch1_m1_max={ch1_m1_max}')

        row, col = 190, picture_offset + 12  # 第二张图片的行列坐标
        letters = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            letters = chr(65 + remainder) + letters
        cell_address_2 = f"{letters}{row}"  # 生成单元格地址，例如 "BC190"
        test_save(6, phase * 2, cell_address_2, 10, 10, 362, 220)  # 保存第二张图片

        # 输出图片保存提示
        print(f'第二张图片保存到单元格 {cell_address_2}')
        print('第二张图片和数据保存完毕')

        el.state('OFF')
        time.sleep(1)

    xls.save()  # 保存 Excel 文件
    print("TEST6测试结束")
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')
    root6.destroy()

def test8():
    test_currents = generate_sequence()
    global table_type
    if xls.getCell(entry.get(), 189, 8) == 'MOSFET Switching with NO Load':
        table_type = '表格2'
        print('表格2')
    else:
        table_type = '表格1'
        print('表格1')

    if table_type == '表格1':
        row_start = 264
        row_end = 274
    elif table_type == '表格2':
        row_start = 245
        row_end = 255

    control_dc_source(vin, Iin, 'ON')
    # 获取 choice 的值，决定测量模式
    choice_value = choice.get()

    row = row_start
    for test_current in test_currents:
        if row > row_end:
            break

        # 设置电流
        el.static(9, 'MAX', test_current)
        el.state('ON')
        time.sleep(5)

        # 扫描并读取数据
        da.Scan_Channel()
        date = da.Read_Date()

        # 更新测量数据的顺序
        Measured_Vout1 = float(date[0])
        Measured_Vin = float(date[1])
        Measured_Iin = float(date[2]) * 1000
        Measured_Vout2 = float(date[3])

        print("{:.4f}\t{:.4f}\t{:.4f}\t{:.4f}".format(Measured_Vout1, Measured_Vin, Measured_Iin, Measured_Vout2))

        # 根据 choice_value 选择要写入的数据列
        if choice_value == 0:  # 近端模式：VOUT1, VIN, IIN
            data_to_write = [Measured_Vout1, Measured_Vin, Measured_Iin]
        else:  # 远端模式：VOUT1, VIN, VOUT2
            data_to_write = [Measured_Vout1, Measured_Vin, Measured_Vout2]

        # 写入 Excel 表格
        col_offset = 4  # 从第4列开始写入
        for value in data_to_write:
            xls.setCell(entry.get(), row, col_offset, value)
            col_offset += 1  # 列递增

        row += 1
        el.state('OFF')
    xls.save()  # 保存 Excel 文件
    time.sleep(1)
    print("TEST8测试结束")
    root8.destroy()
    control_dc_source(vin, Iin, 'OFF')

def test9(type):
    print("TEST9测试开始")
    global table_type
    if xls.getCell(entry.get(), 189, 8) == 'MOSFET Switching with NO Load':
        table_type = '表格2'
        print('表格2')
    else:
        table_type = '表格1'
        print('表格1')

    if MSO5 == 1:
        osc.write('FACTORY')  # 恢复到出厂设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_003.SET"')  # 加载预设的配置文件
        print("已恢复出厂设置并加载配置文件")

    common_set()
    tl9_channel_set()
    measure9()

    if table_type == '表格1':
        image_index = {1: 'F281', 2: 'R281', 3: 'F299', 4: 'R299'}
        cell_positions = {
            1: (286, 4),
            2: (286, 16),
            3: (304, 4),
            4: (304, 16),
        }

    if table_type == '表格2':
        image_index = {1: 'F262', 2: 'R262', 3: 'F280', 4: 'R280'}
        cell_positions = {
            1: (267, 4),
            2: (267, 16),
            3: (285, 4),
            4: (285, 16),
        }

    i = 1
    while i <= 4:
        if type == 'ALL':
            print(f"正在执行第 {i} 次测试...")  # 每次循环开始提示
            if i == 1:
                print("设置水平模式和触发条件，开始单步触发...")
                set_horizontal_mode(1e-2, 1e-2, 25, 1e7)
                osc.trigger('NORMAL', 'CH1', 'FALL', volt / 2)
                control_dc_source(vin, Iin, 'ON')

                time.sleep(3)
                osc.state('single')  # 设置单步触发
                j = 0
                ld_ocp = ld_max  # 初始负载电流设为最大值
                ldmax_step = 0.05 * ocp_spec
                print(f"初始负载电流：{ld_ocp}")  # 显示初始负载电流
                while ld_ocp <= ocp_spec * 1.5:  # 直到负载电流超过标准的1.5倍
                    ld_ocp = ld_max + ldmax_step * j
                    j = j + 1
                    print(f"当前负载电流：{ld_ocp}")  # 显示当前负载电流
                    el.static(9, 'MAX', ld_ocp)
                    el.state('ON')
                    time.sleep(1)
                    tri = process('tri')
                    print(f"调用 'tri' 后的返回值：{tri}")  # 显示 'tri' 返回值
                    if tri != 1:
                        control_dc_source(vin, Iin, 'OFF')
                        el.state('OFF')
                        ld_ocp = process('ld_ocp')
                        print(f"调用 'ld_ocp' 后的返回值：{ld_ocp}")  # 显示 'ld_ocp' 返回值
                        row, col=cell_positions.get(i,(0,0))
                        xls.setCell(entry.get(), row, col, ld_ocp)
                        picture_cell = image_index.get(i, 0)
                        test_save(9, i, picture_cell, 36, 10, 349, 223)
                        print(f"第 {i} 次测试完成，数据已保存")
                        break

                    elif ld_ocp >= ocp_spec * 1.5:  # 电流超载
                        el.state('OFF')
                        print("电流超载，测试停止")
                        break  # 如果超载，停止测试

            elif i == 3:
                print("设置水平模式和触发条件，开始单步触发...")
                control_dc_source(vin, Iin, 'OFF')
                set_horizontal_mode(1e-2, 1e-2, 35, 1e7)
                osc.trigger('NORMAL', 'CH4', 'RISE', volt / 2)
                time.sleep(3)
                osc.state('single')  # 设置单步触发
                el.static(9, 'MAX', 0)
                el.state('ON')
                el.short('ON')
                time.sleep(2)
                control_dc_source(vin, Iin, 'ON')
                time.sleep(2)
                ld_short = process('ld_short')
                row, col = cell_positions.get(i, (0, 0))
                xls.setCell(entry.get(), row, col, ld_short)
                picture_cell = image_index.get(i, 0)
                test_save(9, i, picture_cell, 36, 10, 349, 223)
                print(f"第 {i} 次测试完成，数据已保存")
                el.state('OFF')
                el.short('OFF')


            elif i == 4:
                print("设置水平模式和触发条件，开始单步触发...")
                control_dc_source(vin, Iin, 'OFF')
                set_horizontal_mode(1e-2, 1e-2, 40, 1e7)
                osc.trigger('NORMAL', 'CH1', 'FALL', volt / 2)
                time.sleep(3)
                osc.state('single')  # 设置单步触发
                control_dc_source(vin, Iin, 'ON')
                el.static(9, 'MAX', ld_max)
                el.state('ON')
                time.sleep(2)
                el.short('ON')
                time.sleep(2)
                ld_short = process('ld_short')
                row, col = cell_positions.get(i, (0, 0))
                xls.setCell(entry.get(), row, col, ld_short)
                picture_cell = image_index.get(i, 0)
                test_save(9, i, picture_cell, 36, 10, 349, 223)
                print(f"第 {i} 次测试完成，数据已保存")
                el.short('OFF')
                el.state('OFF')



        elif type =='slowocp':
            i = 1
            print("设置水平模式和触发条件，开始单步触发...")
            set_horizontal_mode(1e-2, 1e-2, 25, 1e7)
            osc.trigger('NORMAL', 'CH1', 'FALL', volt / 2)
            control_dc_source(vin, Iin, 'ON')
            time.sleep(3)
            osc.state('single')  # 设置单步触发
            j = 0
            ld_ocp = ld_max  # 初始负载电流设为最大值
            ldmax_step = 0.05 * ocp_spec
            print(f"初始负载电流：{ld_ocp}")  # 显示初始负载电流
            while ld_ocp <= ocp_spec * 1.5:  # 直到负载电流超过标准的1.5倍
                ld_ocp = ld_max + ldmax_step * j
                j = j + 1
                print(f"当前负载电流：{ld_ocp}")  # 显示当前负载电流
                el.static(9, 'MAX', ld_ocp)
                el.state('ON')
                time.sleep(1)
                tri = process('tri')
                print(f"调用 'tri' 后的返回值：{tri}")  # 显示 'tri' 返回值
                if tri != 1:
                    control_dc_source(vin, Iin, 'ON')
                    el.state('OFF')
                    ld_ocp = process('ld_ocp')
                    row, col = cell_positions.get(i, (0, 0))
                    xls.setCell(entry.get(), row, col, ld_ocp)
                    picture_cell = image_index.get(i, 0)
                    test_save(9, i, picture_cell, 36, 10, 349, 223)
                    print(f"第 {i} 次测试完成，数据已保存")
                    break

                elif ld_ocp >= ocp_spec * 1.5:  # 电流超载
                    el.state('OFF')
                    print("电流超载，测试停止")
                    break  # 如果超载，停止测试

        elif type== 'scpbefore':
            i = 3
            print("设置水平模式和触发条件，开始单步触发...")
            control_dc_source(vin, Iin, 'OFF')
            set_horizontal_mode(1e-2, 1e-2, 35, 1e7)
            osc.trigger('NORMAL', 'CH4', 'RISE', Iin / 2)
            time.sleep(3)
            osc.state('single')  # 设置单步触发
            el.static(9, 'MAX', 0)
            el.short('ON')
            time.sleep(3)
            print('短路')
            control_dc_source(vin, Iin, 'ON')
            time.sleep(3)
            ld_short = process('ld_short')
            row, col = cell_positions.get(i, (0, 0))
            xls.setCell(entry.get(), row, col, ld_short)
            picture_cell = image_index.get(i, 0)
            test_save(9, i, picture_cell, 36, 10, 349, 223)
            print(f"第 {i} 次测试完成，数据已保存")
            el.short('OFF')


        elif type == 'scpafter':
            i = 4
            print("设置水平模式和触发条件，开始单步触发...")
            control_dc_source(vin, Iin, 'OFF')
            set_horizontal_mode(1e-2, 1e-2, 40, 1e7)
            osc.trigger('NORMAL', 'CH1', 'FALL', volt / 2)
            time.sleep(3)
            osc.state('single')  # 设置单步触发

            el.state('OFF')
            el.short('OFF')
            control_dc_source(vin, Iin, 'ON')
            el.static(9, 'MAX', ld_max)
            el.state('ON')
            time.sleep(3)
            el.short('ON')
            time.sleep(2)
            ld_short = process('ld_short')
            row, col = cell_positions.get(i, (0, 0))
            xls.setCell(entry.get(), row, col, ld_short)
            picture_cell = image_index.get(i, 0)
            test_save(9, i, picture_cell, 36, 10, 349, 223)
            print(f"第 {i} 次测试完成，数据已保存")
            el.short('OFF')
            el.state('OFF')

        if type == 'ALL':
            i += 1
        else:
            break

    xls.save()  # 保存 Excel 文件
    print("TEST9测试结束")
    auto_close_messagebox(root9, title='程序执行完毕',message='TEST9执行完毕,如需进行TEST3测试，请更换CH4通道为高压差分探头连接VIN，CH2添加单端无源探头连接EN，并进行电路下电')
    control_dc_source(vin, Iin, 'OFF')
    root9.destroy()


# xls = EasyExcel("E:\\SUMA\\SUMA POWER TEST\\POL.xlsx")
main_window()






