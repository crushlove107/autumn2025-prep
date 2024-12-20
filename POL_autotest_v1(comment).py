# 项目名称

"中科硬件实验室 POL电源自动化测试"

# —————————————————————————————分割线———————————————————————————————


# 项目所需包的导入

from tkinter import *  # 可以直接使用tkinter的模块名
import tkinter as tk   # 将tkinter重命名为tk，方便引用————————并未使用

# tkinter用于创建图形用户界面（GUI), 构建桌面应用程序
# 具有图形化界面，提供窗口、按钮、标签、文本框等控件

import time             # 提供时间相关的功能
import os               # 提供与操作系统交互的功能，例如文件或目录操作
import pyvisa           # 用于与仪器通信，指在测试和测量应用中
import win32com.client  # 和windows应用程序交互，指 excel


"""常用功能
time.sleep(seconds)  暂停执行指定的秒数
os.path.exists(path) 检查路径是否存在
os.makedirs(path)    递归地创建目录
rm = pyvisa.ResourceManager() 
创建一个 ResourceManager 实例，这个实例用于管理和访问连接到计算机的仪器
osc = rm.open_resource('GPIB0::14::INSTR')  # 连接到指定的仪器
response = osc.query('*IDN?')  # 发送查询命令并获取响应

excel = win32com.client.Dispatch("Excel.Application")  # 启动Excel应用程序
workbook = excel.Workbooks.Add()  # 新建一个工作簿
sheet = workbook.Sheets(1)  # 访问第一个工作表
sheet.Cells(1, 1).Value = 'Hello, World!'  # 在单元格中写入数据
workbook.SaveAs('example.xlsx')  # 保存工作簿
excel.Application.Quit()  # 退出Excel应用程序
"""
from tkinter import messagebox  # 用于显示对话框消息，如警告、错误信息、确认对话框等。
from tkinter import filedialog  # 提供文件选择对话框，允许用户选择文件或目录

"""
messagebox.showinfo(title='仪器连接', message='示波器和电子负载均已正确连接')
messagebox.showerror(title='仪器连接', message='电子负载连接错误，请检查')
ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
弹出一个文件选择对话框，选择指定类型的excel文件保存在指定地址
pic_path = filedialog.askdirectory()
显示一个目录选择对话框，让用户选择一个目录，返回用户选择的目录路径
"""

import xlwings as xw    # 导入了 xlwings 库并将其别名为 xw
# xlwings 是一个用于在 Python 中操作 Excel 的库
"""
app = xw.App(visible=True)  # 启动Excel应用程序
wb = app.books.add()  # 新建一个工作簿
sheet = wb.sheets[0]  # 访问第一个工作表
sheet.range('A1').value = 'Hello, World!'  # 在单元格中写入数据
wb.save('example.xlsx')  # 保存工作簿
app.quit()  # 退出Excel应用程序

"""


# —————————————————————————————分割线———————————————————————————————

#  工具和仪器说明
class EasyExcel:
    """
    创建一个 Excel实例  excel = EasyExcel(r"C:\\Users\\Seir\Desktop\\0402电阻表.xlsx")
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
        # 初始化 参数: filename（可选）用于指定打开的 Excel 文件。如果未提供 filename，则创建一个新的空工作簿
        self.xlApp = win32com.client.Dispatch('ket.Application')
        # 创建一个 Excel 应用程序实例 (self.xlApp)
        if filename:
            self.filename = filename
            print(filename)
            self.xlBook = self.xlApp.Workbooks.Open(filename)
            self.xlApp.Visible = True
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''
        """
        如果文件名存在，使用旧文件名并输出它，打开指定的工作簿并将excel应用程序设置为可见
        如果未提供filename则创建一个新工作簿，并记录新文件名
        """

    def save(self, newfilename=None):
        # 保存当前的工作簿 参数: newfilename（可选）用于指定保存时的文件名。如果不提供 newfilename，则使用当前文件名保存
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()
        """
        如果提供了新文件名，则将工作簿保存问新文件，并更新文件名
        如果没有提供新文件名，则以旧文件名保存工作簿
        """

    def close(self):
        # 关闭当前工作簿并释放 Excel 应用程序对象
        self.xlBook.Close(SaveChanges=0)  # 参数为0表示不保存更改！！！
        del self.xlApp

    def getCell(self, sheet: object, row: object, col: object) -> object:
        # 获取指定工作表中指定单元格的值。sheet=指定工作表，row=行，col=列
        sht = self.xlBook.Worksheets(sheet)  # sht是工作表的局部变量，可以访问其中的单元格、范围、图形
        sht.Activate()  # 激活工作表
        return sht.Cells(row, col).Value  # 返回指定单元格的值

    def setCell(self, sheet, row, col, value):
        # 设置指定工作表中指定单元格的值 同getCell. 直接设置不返回值
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

    def addPicture(self, sheet, PictureName, Range, left_offset, Top_offset, Width, Heigth):
        # 在指定工作表的指定位置添加图片
        # sheet为指定工作表，picturename为图片文件名，range为基准单元格范围，left_offset和top_offset为图片的位置偏移量，width和height为图片的宽度和高度
        sht = self.xlBook.Worksheets(sheet)
        sht.Activate()
        cell = sht.Range(Range)  # 获取指定范围的单元格对象
        sht.Shapes.AddPicture(PictureName, LinkToFile=False, SaveWithDocument=True, Left=cell.Left + left_offset,
                              Top=cell.Top + Top_offset,
                              Width=Width, Height=Heigth)
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


class OscMPO5series:  # 控制OscMPO5 示波器 包括设置测量参数、触发、通道配置、导出图像和读取数据
    def __init__(self, address):  # 连接到示波器
        address = address.strip()
        address = address.rstrip()  # 去掉地址前后的空白字符
        self.osc = rm.open_resource(address)  # 根据给定的地址打开示波器资源

    def state(self, state):  # 控制示波器的运行状态 run, single, stop
        if state == 'run':
            self.osc.write('DIS:PERS:RESET')  # clear 清除示波器旧的显示内容
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')  # 满足条件后停止收集
            self.osc.write('ACQUIRE:STATE RUN')  # 设置示波器为 "运行" 状态
        elif state == 'single':
            self.osc.write('ACQUIRE:STOPAFTER SEQUENCE')  # single 单次触发模式，采集一次数据后停止
            self.osc.write('ACQUIRE:STATE 1')  # 1 表示单次触发模式
        elif state == 'stop':
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')  # stop 设置停止条件
            self.osc.write('ACQUIRE:STATE STOP')  # 设置示波器为停止状态

    def measure(self, measNum, channel, type1):  # 用于设置测量参数
        # measNum 测量编号   channel 测量通道   type1 测量类型.
        self.osc.write('MEASUREMENT:ADDNEW "MEAS%d"' % measNum)
        # 添加一个新的测量项，命名为 MEAS{measNum}
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE1 %s' % (measNum, channel))
        # 设置第 measNum 个测量项的源通道，让测量项从指定的通道获取数据
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, type1))  # 设置测量内容
        # 设置测量项的类型 (峰值、均值等)
        self.osc.write('MEASUrement:MEAS%d:DISPlaystat:ENABle ON' % measNum)
        # 启用第 measNum 个测量项的显示状态

    def measure_1(self, measNum, channel, TYPE):
        #  measNum 测量编号   channel 测量通道   type1 测量类型
        self.osc.write('MEASUREMENT:ADDNEW "MEAS%d"' % measNum)
        # 添加一个新的测量项，命名为 MEAS{measNum}
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, TYPE))
        # # 设置测量项的类型 (峰值、均值等)
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE %s' % (measNum, channel))
        # 启用第 measNum 个测量项的显示状态

    def measOff(self, measNum):  # 关闭指定的测量
        self.osc.write('MEASU:DEL "MEAS%d"' % measNum)
        # 删除指定的测量项
        # self.osc.write('MEASUrement:ANNOTate AUTO')   自动显示或处理测量注释

    def makeDir(self, dir1):  # 在示波器的文件系统中创建一个目录
        self.osc.write('FILESystem:MKDir "%s"' % dir1)

    def export(self, temp1, dir1):  # 导出示波器当前图像到指定路径
        self.osc.write('SAV:IMAG "%s.%s"' % (dir1, temp1))
        # dir1 文件的路径或名称, temp1 文件的拓展名 png或 jpg

    def readfile(self, dir1):  # 读取示波器上指定路径的文件
        self.osc.write('FILESYSTEM:READFILE "%s"' % dir1)

    def persistence(self, state):  # 控制显示设置，分别用于设置波形的持续性显示和光标的状态
        # state表示持久性显示的状态如 ON OFF ACQ
        self.osc.write('DISplay:PERSistence %s' % state)

    def cursor(self, state):
        self.osc.write('CURSOR:STATE %s' % state)  # 控制光标的状态 ON 或 OFF

    def hormode(self, state):  # 设置水平（时间轴）显示模式
        # 参数 state表示水平显示模式的设置如  NORMAL ROLL
        self.osc.write('HOR:MODE %s' % state)  # 设置 Horizontal格式
        self.osc.write('HOR:MODE:%s:CONFIGure HORIZ' % state)  # 具体设置
        self.osc.write('DISplay:WAVEView:GRIDTYPE FIXED')  # 设置波形视图显示样式为叠加模式
        self.osc.write('DISplay:WAVEView1:VIEWStyle OVErlay')  # 设置视图样式为叠加

    def horpos(self, num):  # 设置波形的水平位置
        self.osc.write('HORIZONTAL:POSITION %d' % num)  # 水平位置

    def coupling(self, channel, state):  # 设置通道的耦合方式
        # channel:CH1 CH2   state: AC DC
        self.osc.write('%s:COUP %s' % (channel, state))

    def number(self, number):  # 查询示波器已获取的采集数量，等待达到指定数量
        num = 0
        while num <= number:
            time.sleep(1)  # 每次循环暂停一秒，以便示波器有时间完成收集
            num = self.osc.query('ACQuire:NUMAC?')
            if MSO5 == 1:
                num = num[15:]  # 如果是05型号示波器，去掉前十五个字符
            num = int(num)  # 数据转换为整数型

    def record(self, num):  # 设置记录长度
        num = num * 1.25  # 将记录长度值增加25%
        self.osc.write('HOR:MODE:RECO %d' % num)

    def query(self, query):  # 发送查询或写入命令到示波器
        self.osc.query('%s' % query)
        return self.osc.query('%s' % query)  # 返回示波器的响应数据

    def write(self, write):  # 发送写入命令到示波器
        self.osc.write('%s' % write)  # 仅发送命令，不获取返回值

    def scale(self, channel, num):  # 设置通道的垂直比例
        # channel:CH1 CH2   num:0.5(V/div)
        self.osc.write('%s:SCALE %.3f' % (channel, num))
        # 将通道的垂直比例设置为num 保留三位小数

    def channel(self, ch1, ch2, ch3, ch4, math1, math2):
        # 控制各通道（包括数学通道）的开启和关闭
        if math1 == 'ON':
            self.osc.write('MATH:ADDNEW "MATH1"')
        else:
            self.osc.write('MATH:DELETE "MATH1"')
        if math2 == 'ON':
            self.osc.write('MATH:ADDNEW "MATH2"')
        else:
            self.osc.write('MATH:DELETE "MATH2"')
            # 观察数学通道的状态，选择创建或删除
        self.osc.write('SELECT:CH4 %s' % ch4)
        self.osc.write('SELECT:CH3 %s' % ch3)
        self.osc.write('SELECT:CH2 %s' % ch2)
        self.osc.write('SELECT:CH1 %s' % ch1)
        # 设置四个通道的状态

    def label(self, channel, name, xi, y):  # 设置通道的标签及其位置和样式
        # channel 通道标识  name 标签文本  xi 水平轴位置比例 yi 垂直轴位置比例
        self.osc.write('%s:LABel:NAMe "%s"' % (channel, name))  # 设置标签为name
        xi_new = 348 * xi - 174
        y_new = 94 * y
        # 转换为示波器具体位置值
        self.osc.write('%s:LABel:XPOS %.1f' % (channel, xi_new))
        self.osc.write('%s:LABel:YPOS %.1f' % (channel, y_new))
        # 设置标签的水平和垂直位置
        self.osc.write('%s:LABel:FONT:BOLD OFF' % channel)
        self.osc.write('%s:LABel:FONT:ITALic OFF' % channel)
        self.osc.write('%s:LABel:FONT:UNDERline OFF' % channel)
        # 设置标签非粗体 非斜体 关闭下划线
        self.osc.write('%s:LABel:FONT:SIZE 14' % channel)
        # 设置字体大小为14

    def chanset(self, channel, pos, offset, bandwidth, scale):  # 配置通道的设置，例如垂直位置、偏移、带宽和比例
        # channel 通道标识符 pos 垂直位置 offset 垂直偏移量 bandwidth 带宽 scale 垂直比例
        self.osc.write('%s:POS %.1f' % (channel, pos))  # 竖直位置
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:BANDWIDTH %s' % (channel, bandwidth))  # 通道带宽设置为20MHz
        self.osc.write('%s:SCALE %.3f' % (channel, scale))  # 设置通道的scale

    def trigger(self, mode, channel, slope, level):  # 设置触发器的参数，例如触发模式、通道、斜率和触发电平
        # mode 触发模式: AUTO NORMAL SINGLE， channel 通道标识符， slope 斜率， level 触发电平
        self.osc.write('TRIGGER:A:MODE %s' % mode)
        # 设置触发模式：EDGE、PULSE
        self.osc.write('TRIGGER:A:EDGE:SOURCE %s' % channel)
        # 设置触发通道
        self.osc.write('TRIGGER:A:EDGE:SLOPE %s' % slope)
        # 设置触发边沿类型：RISE、FALL、BOTH
        self.osc.write('TRIGGER:A:LEVEL:%s %.3f' % (channel, level))
        # 设置触发电平

    def math(self, channel, define, offset, pos, scale):  # 设置数学通道的定义、垂直位置和比例
        # channel 通道标识符, define 通道定义, offset 垂直偏移量, pos 垂直位置, scale 垂直比例
        self.osc.write('MATH:%s:DEFINE "%s"' % (channel, define))
        self.osc.write('MATH:%s:VERT:AUTOSC OFF' % channel)
        # 关闭数学通道的自动垂直缩放
        time.sleep(1)
        self.osc.write('MATH:%s:OFFSET %.1f' % (channel, offset))  # 设置offset
        time.sleep(1)
        self.osc.write('DISplay:WAVEView1:MATH:%s:VERTICAL:POSITION %.1f' % (channel, pos))  # 设置math通道的position
        time.sleep(1)
        self.osc.write('DISplay:WAVEView1:MATH:%s:VERTICAL:SCALE %.1f' % (channel, scale))
        # self.osc.write('MATH:ADDNEW "%s"' % channel)  # 开启math通道

    def readraw(self, file_path):  # 读取原始数据并保存到指定文件路径
        data = self.osc.read_raw()  # 从示波器缓冲区获取二进制数据
        data_temp = open(file_path, 'wb')  # 以二进制形式wb打开指定路径文件
        data_temp.write(data)  # 将二进制数据写入到文件中
        data_temp.close()  # 关闭文件


class OscDPO7000C:  # 控制 Tektronix DPO7000C 系列示波器 设置状态、执行测量、调整通道设置、触发设置等。
    def __init__(self, address):  # 连接到示波器
        address = address.strip()  # 同示波器05
        address = address.rstrip()
        self.osc = rm.open_resource(address)

    def state(self, state):  # 控制示波器的运行状态
        if state == 'run':  # 运行示波器并重置显示
            self.osc.write('DIS:PERS:RESET')  # clear
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE RUN')
        elif state == 'single':  # 触发单次采集
            self.osc.write('ACQUIRE:STOPAFTER SEQUENCE')  # 按下single
            self.osc.write('ACQUIRE:STATE 1')
        elif state == 'stop':  # 停止采集
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE STOP')
        else:
            print('状态设置失败')

    def measure(self, measNum, channel, type1):  # 测量的编号、通道、频率、周期
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE1 %s' % (measNum, channel))
        # 设置测量的来源通道 measNum:测量编号  channel:通道标识符
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, type1))
        # 设置测量类型 type1:FREQ频率、PER周期、VPP峰值
        self.osc.write('MEASUrement:MEAS%d:STATE ON' % measNum)
        # 启用测量项
        self.osc.write('MEASUrement:ANNOTation:STATE MEAS%d' % measNum)
        # 启用测量器支持，显示测量结果

    def measOff(self, measNum):  # 关闭指定编号的测量功能，并将注释设置为自动模式
        self.osc.write('MEASUrement:MEAS%d:STATE OFF' % measNum)
        self.osc.write('MEASUrement:ANNOTate AUTO')

    def makeDir(self, dir1):  # 在示波器的文件系统中创建一个新的目录 dir1
        self.osc.write('FILESystem:MKDir "%s"' % dir1)

    def export(self, temp1, dir1):  # 导出示波器数据
        self.osc.write('EXPort:FORMat %s' % temp1)
        # 设置导出文件的格式 PNG、CSV、WFM
        self.osc.write('EXPORT:FILENAME "%s"' % dir1)
        # 设置导出文件的保存路径和文件名 dir1:C:/data/scope_image.png
        self.osc.write('EXPort STARt')
        # 启动导出过程

    def readfile(self, dir1):  # 读取指定路径 dir1 的文件内容
        self.osc.write('FILESYSTEM:READFILE "%s"' % dir1)

    def persistence(self, state):  # 设置示波器的显示持久性状态
        self.osc.write('DISplay:PERSistence %s' % state)  # 关闭累积 OFF

    def cursor(self, state):  # 打开或关闭示波器上的光标显示
        self.osc.write('CURSOR:STATE %s' % state)  # 关闭cursor OFF

    def hormode(self, state):  # 设置水平模式
        self.osc.write('HOR:MODE %s' % state)  # 设置 Horizontal格式

    def horpos(self, num):  # 调整水平位置（时间轴）的位置
        self.osc.write('HORIZONTAL:POSITION %d' % num)  # 水平位置

    def coupling(self, channel, state):  # 设置特定通道的耦合状态
        self.osc.write('%s:COUP %s' % (channel, state))

    def number(self, number):  # 查询示波器采集的波形数量，直到达到指定的数量 number
        num = 0
        while num <= number:
            time.sleep(1)
            num = self.osc.query('ACQuire:NUMAC?')
            num = int(num)

    def record(self, num):  # 设置记录长度
        self.osc.write('HOR:MODE:RECO %d' % num)

    def query(self, query):  # 发送查询命令到示波器并返回结果
        self.osc.query('%s' % query)
        return self.osc.query('%s' % query)

    def write(self, write):  # 发送指令到示波器，但不返回结果
        self.osc.write('%s' % write)

    def scale(self, channel, num):  # 设置指定通道的垂直比例
        self.osc.write('%s:SCALE %.3f' % (channel, num))

    def channel(self, ch1, ch2, ch3, ch4, math1, math2):
        # 选择示波器通道 ch1, ch2, ch3, ch4 的显示状态，并设置 math1, math2 数学通道的显示状态
        self.osc.write('SELECT:MATH2 %s' % math2)
        self.osc.write('SELECT:MATH1 %s' % math1)
        self.osc.write('SELECT:CH4 %s' % ch4)
        self.osc.write('SELECT:CH3 %s' % ch3)
        self.osc.write('SELECT:CH2 %s' % ch2)
        self.osc.write('SELECT:CH1 %s' % ch1)

    def label(self, channel, name, xi, y):
        # 给指定通道添加标签 name，并设置标签的横向位置 xi 和纵向位置 y
        self.osc.write('%s:LABel:NAMe "%s"' % (channel, name))  # 设置label
        self.osc.write('%s:LABel:XPOS %.1f' % (channel, xi))
        self.osc.write('%s:LABel:YPOS %.1f' % (channel, y))

    def chanset(self, channel, pos, offset, bandwidth, scale):
        # 配置指定通道的参数，如位置 pos、偏移量 offset、带宽 bandwidth 和垂直比例 scale
        self.osc.write('%s:POS %.1f' % (channel, pos))  # 竖直位置
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:BANDWIDTH %s' % (channel, bandwidth))  # 通道带宽设置为20MHz
        self.osc.write('%s:SCALE %.3f' % (channel, scale))  # 设置通道的scale

    def trigger(self, mode, channel, slope, level):
        # 配置触发设置 触发模式、触发通道、触发的边沿类型、触发电平
        self.osc.write('TRIGGER:A:MODE %s' % mode)
        self.osc.write('TRIGGER:A:EDGE:SOURCE %s' % channel)
        self.osc.write('TRIGGER:A:EDGE:SLOPE %s' % slope)  # 设置触发频道和形式
        self.osc.write('TRIGGER:A:LEVEL %.2f' % level)

    def math(self, channel, define, offset, pos, scale):
        # 设置数学通道的运算定义 define，并配置其偏移 offset、位置 pos 和比例 scale
        self.osc.write('%s:DEFINE "%s"' % (channel, define))
        self.osc.write('%s:VERT:AUTOSC OFF' % channel)
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:VERTICAL:POSITION %.2f' % (channel, pos))  # 设置math通道的position
        self.osc.write('%s:VERTICAL:SCALE %.2f' % (channel, scale))
        self.osc.write('SELECT:%s ON' % channel)  # 开启math通道

    def readraw(self, file_path):  # 从示波器读取原始数据，并将其保存到 file_path中
        data = self.osc.read_raw()
        data_temp = open(file_path, 'wb')
        data_temp.write(data)
        data_temp.close()


class OscDPO5104B:  # 控制 Tektronix DPO5104B 示波器 设置示波器的状态、进行测量、设置触发和通道
    def __init__(self, address):  # 连接到示波器
        address = address.strip()  # 同示波器05
        address = address.rstrip()
        self.osc = rm.open_resource(address)

    def state(self, state):  # 控制示波器的采集状态
        if state == 'run':  # 启动示波器并清除显示持久性
            self.osc.write('DIS:PERS:RESET')  # clear
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE RUN')
        elif state == 'single':  # 进行单次采集
            self.osc.write('ACQUIRE:STOPAFTER SEQUENCE')  # 按下single
            self.osc.write('ACQUIRE:STATE 1')
        elif state == 'stop':  # 停止采集
            self.osc.write('ACQUIRE:STOPAFTER RUNSTOP')
            self.osc.write('ACQUIRE:STATE STOP')
        else:
            print('状态设置失败')

    def measure(self, measNum, channel, type1):
        # 配置和启用测量功能 测量编号 选择测量的通道 测量类型（例如，电压、频率等）
        self.osc.write('MEASUREMENT:MEAS%d:SOURCE1 %s' % (measNum, channel))
        self.osc.write('MEASUREMENT:MEAS%d:TYPE %s' % (measNum, type1))  # 设置测量内容
        self.osc.write('MEASUrement:MEAS%d:STATE ON' % measNum)
        self.osc.write('MEASUrement:ANNOTation:STATE MEAS%d' % measNum)

    def measOff(self, measNum):  # 关闭指定编号的测量 自动注释
        self.osc.write('MEASUrement:MEAS%d:STATE OFF' % measNum)
        self.osc.write('MEASUrement:ANNOTate AUTO')

    def makeDir(self, dir1):  # 在示波器的文件系统中创建一个新的目录 dir1
        self.osc.write('FILESystem:MKDir "%s"' % dir1)

    def export(self, temp1, dir1):  # 导出示波器数据为指定格式 temp1，并将文件保存到路径 dir1
        self.osc.write('EXPort:FORMat %s' % temp1)
        self.osc.write('EXPORT:FILENAME "%s"' % dir1)  # 保存图片
        self.osc.write('EXPort STARt')

    def readfile(self, dir1):  # 读取并返回指定路径 dir1 下的文件内容
        self.osc.write('FILESYSTEM:READFILE "%s"' % dir1)

    def persistence(self, state):  # 设置或关闭显示持久性
        self.osc.write('DISplay:PERSistence %s' % state)  # 关闭累积

    def cursor(self, state):  # 控制光标显示的开关
        self.osc.write('CURSOR:STATE %s' % state)  # 关闭cursor

    def hormode(self, state):  # 设置水平模式
        self.osc.write('HOR:MODE %s' % state)  # 设置 Horizontal格式

    def horpos(self, num):  # 设置水平位置（时间轴）的位置
        self.osc.write('HORIZONTAL:POSITION %d' % num)  # 水平位置

    def coupling(self, channel, state):  # 设置通道的耦合方式
        self.osc.write('%s:COUP %s' % (channel, state))

    def number(self, number):  # 查询并等待示波器采集完成指定数量的波形
        num = 0
        while num <= number:
            time.sleep(1)
            num = self.osc.query('ACQuire:NUMAC?')
            num = int(num)

    def record(self, num):  # 设置记录长度
        self.osc.write('HOR:MODE:RECO %d' % num)

    def query(self, query):  # 发送查询命令并返回结果
        self.osc.query('%s' % query)
        return self.osc.query('%s' % query)

    def write(self, write):  # 读取原始数据，并将其保存到指定文件路径 file_path
        self.osc.write('%s' % write)

    def scale(self, channel, num):  # 设置指定通道的垂直比例（SCALE），以 num 为单位
        self.osc.write('%s:SCALE %.3f' % (channel, num))

    def channel(self, ch1, ch2, ch3, ch4, math1, math2):  # 设置通道和数学运算通道的开关状态
        self.osc.write('SELECT:MATH2 %s' % math2)
        self.osc.write('SELECT:MATH1 %s' % math1)
        self.osc.write('SELECT:CH4 %s' % ch4)
        self.osc.write('SELECT:CH3 %s' % ch3)
        self.osc.write('SELECT:CH2 %s' % ch2)
        self.osc.write('SELECT:CH1 %s' % ch1)

    def label(self, channel, name, xi, y):  # 设置指定通道的标签，包括标签名、X 轴位置和 Y 轴位置
        self.osc.write('%s:LABel:NAMe "%s"' % (channel, name))  # 设置label
        self.osc.write('%s:LABel:XPOS %.1f' % (channel, xi))
        self.osc.write('%s:LABel:YPOS %.1f' % (channel, y))

    def chanset(self, channel, pos, offset, bandwidth, scale):  # 配置通道的竖直位置、偏移量、带宽和比例
        self.osc.write('%s:POS %.1f' % (channel, pos))  # 竖直位置
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:BANDWIDTH %s' % (channel, bandwidth))  # 一通道带宽设置为20MHz
        self.osc.write('%s:SCALE %.3f' % (channel, scale))  # 设置一通道的scale

    def trigger(self, mode, channel, slope, level):
        self.osc.write('TRIGGER:A:MODE %s' % mode)  # 触发模式（边沿、脉冲等）
        self.osc.write('TRIGGER:A:EDGE:SOURCE %s' % channel)  # 触发通道
        self.osc.write('TRIGGER:A:EDGE:SLOPE %s' % slope)  # 触发边沿（上升、下降）
        self.osc.write('TRIGGER:A:LEVEL %.2f' % level)  # 触发电平

    def math(self, channel, define, offset, pos, scale):
        # 设置数学运算通道的定义和垂直设置，如偏移、位置和比例
        self.osc.write('%s:DEFINE "%s"' % (channel, define))
        self.osc.write('%s:VERT:AUTOSC OFF' % channel)
        self.osc.write('%s:OFFSET %.2f' % (channel, offset))  # 设置offset
        self.osc.write('%s:VERTICAL:POSITION %.2f' % (channel, pos))  # 设置math通道的position
        self.osc.write('%s:VERTICAL:SCALE %.2f' % (channel, scale))
        self.osc.write('SELECT:%s ON' % channel)  # 开启math通道

    def readraw(self, file_path):
        data = self.osc.read_raw()
        data_temp = open(file_path, 'wb')
        data_temp.write(data)
        data_temp.close()


class El63600:  # 控制 Chroma 63600 系列电子负载 设置电子负载的工作模式、动态和静态负载条件、短路状态
    def __init__(self, address):  # 连接到电子负载
        address = address.strip()  # 处理前后两端空白字符（包括空格和换行符）
        address = address.rstrip()
        self.rm = pyvisa.ResourceManager()  # 保存resource manager对象，管理仪器
        self.el = self.rm.open_resource(address)  # 保存电子负载的资源对象

    def mode(self, tpye):  # 设置电子负载的工作模式
        self.el.write('MODE %s' % tpye)

    def state(self, state):  # 控制电子负载的状态 ON" 或 "OFF
        self.el.write('LOAD %s' % state)

    def dynamic(self, channel, load_max, time1):
        # 设置动态负载模式 选择需要配置的通道  最大电流值  时间参数（上升和下降时间）
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        self.el.write('MODE CCDH')  # 设置动态模式 CCDH
        self.el.write('CURR:DYN:L1 %.2f' % (0.8 * load_max))  # 设置负载的上下电流值
        self.el.write('CURR:DYN:L2 %.2f' % (0.2 * load_max))
        self.el.write('CURR:DYN:T1 %.2fms' % time1)  # 设置上升和下降时间
        self.el.write('CURR:DYN:T2 %.2fms' % time1)
        self.el.write('CURR:DYN:FALL MAX')  # 设置动态模式的上下时间为最大值
        self.el.write('CURR:DYN:RISE MAX')
        self.el.write('CURR:DYN:REP 0')  # 设置动态模式的重复次数为 0

    def static(self, channel, rise, load):
        # 设置静态负载模式  选择需要配置的通道 设置电流的上升时间 设置负载电流值
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        if ocp_spec <= 0.6:  # 根据 ocp_spec 的值，选择合适的负载模式
            self.el.write('MODE CCL')
        elif ocp_spec <= 6:
            self.el.write('MODE CCM')
        else:
            self.el.write('MODE CCH')
        self.el.write('CURR:STAT:RISE %s' % rise)
        self.el.write('CURR:STAT:FALL %s' % rise)
        # 设置静态负载模式中的电流上升下降时间
        self.el.write('CURR:STAT:L1 %.2f' % load)
        # 设置静态负载模式中的负载电流值

    def query(self, query):  # 发送查询命令到电子负载并返回响应结果
        self.el.query('%s' % query)

    def write(self, write):  # 发送写入命令到电子负载 不返回
        self.el.write('%s' % write)

    def short(self, state):  # 控制电子负载的短路状态
        self.el.write('LOAD:SHOR %s' % state)


class El6312A:  # 控制电子负载仪器 设置负载模式、执行动态和静态测试、查询设备状态等。
    def __init__(self, address):  # 连接到指定地址的电子负载设备
        address = address.strip()  # 同上 EL636
        address = address.rstrip()
        self.rm = pyvisa.ResourceManager()
        self.el = self.rm.open_resource(address)

    def mode(self, tpye):  # 设置电子负载的工作模式
        self.el.write('MODE %s' % tpye)

    def state(self, state):  # 启动或停止电子负载
        self.el.write('LOAD %s' % state)

    def dynamic(self, channel, load_max, time1):  # 配置并启动动态负载测试
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        self.el.write('MODE CCDH')  # 设置动态模式
        self.el.write('CURR:DYN:L1 %.2f' % (0.8 * load_max))  # 设置负载的上下电流值
        self.el.write('CURR:DYN:L2 %.2f' % (0.2 * load_max))
        self.el.write('CURR:DYN:T1 %.2fms' % time1)  # 动态测试的时间参数，控制电流的上升和下降时间
        self.el.write('CURR:DYN:T2 %.2fms' % time1)
        self.el.write('CURR:DYN:FALL MAX')
        self.el.write('CURR:DYN:RISE MAX')

    def static(self, channel, rise, load):  # 配置并启动静态负载测试
        self.el.write('CHAN %d' % channel)  # 选择相应的通道
        if ocp_spec <= 0.6:
            self.el.write('MODE CCL')  # 根据 ocp_spec 值选择相应的电流模式
        elif ocp_spec <= 6:
            self.el.write('MODE CCM')
        else:
            self.el.write('MODE CCH')
        self.el.write('CURR:STAT:RISE %s' % rise)  # 电流的上升和下降速率
        self.el.write('CURR:STAT:FALL %s' % rise)
        self.el.write('CURR:STAT:L1 %.2f' % load)  # 静态负载的电流值

    def query(self, query):  # 发送查询命令并返回设备的响应
        self.el.query('%s' % query)

    def write(self, write):  # 发送控制命令至电子负载设备
        self.el.write('%s' % write)

    def short(self, state):  # 执行短路测试
        self.el.write('CURR:STAT:L1 0')  # 设置负载的电流值为 0
        self.el.write('LOAD %s' % state)  # 控制电子负载的开关状态
        time.sleep(1)  # 确保负载状态的切换完成，然后再进行短路设置
        self.el.write('LOAD:SHOR %s' % state)  # 设置电子负载的短路状态


def mkdir(path):  # 处理地址并创建目录
    # 去除 path 首尾空格
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
    # 生成保存测试图片的文件路径，创建相应的目录结构，然后将图片从仪器（如示波器）保存到指定路径
    mkpath = '%s/POL Test Pictures/%s/%s' % (pic_path, temp, entry.get())  # 目录路径，用来存放测试图片
    file_path = r'%s/POL Test Pictures/%s/%s/%s.PNG' % (pic_path, temp, entry.get(), name)  # 最终的图片文件路径
    mkdir(mkpath)  # 调用自定义的 mkdir 函数，创建图片存放的目录（如果目录不存在）
    osc.makeDir('C:\\POL Test Pictures')
    osc.makeDir('C:\\POL Test Pictures\\%s' % temp)
    osc.makeDir('C:\\POL Test Pictures\\%s\\%s' % (temp, entry.get()))
    osc.export('PNG', 'C:\\POL Test Pictures\\%s\\%s\\%s' % (temp, entry.get(), name))  # 保存图片
    time.sleep(2)
    osc.readfile('C:/POL Test Pictures/%s/%s/%s.PNG' % (temp, entry.get(), name))  # 读取已保存的图片
    time.sleep(3)
    osc.readraw(file_path)  # 读取原始文件数据
    time.sleep(2)


def count():  # 计时器功能，每秒钟更新一次显示内容,在指定的条件下继续计时
    global counter, countmode  # 全局变量 counter 时间, countmode 控制计时器开关
    if countmode == 'ON':  # 计时器状态判断
        timestr = '{:02}:{:02}'.format(*divmod(counter, 60))  # 将秒数转换为分钟和秒数
        display.config(text=str(timestr))  # 将格式化后的时间字符串显示在界面上
        counter += 1  # 增加计时器的秒数计数器 counter的值
        display.after(1000, count)  # 计时器每秒更新一次
    else:  # 计时器未开启则不执行操作
        pass


# —————————————————————————————分割线———————————————————————————————

# 测量方法 & 通道设置
def measure1():  # 仅测量CH1 MAX MIN RMS PK2PK
    osc.measure(1, 'CH1', 'MAXIMUM')
    #  参数 1：第一次测量  CH1：被测量的通道  MAXIMUM：测量最大值
    osc.measure(2, 'CH1', 'MINIMUM')
    osc.measure(3, 'CH1', 'RMS')  # 测量均方根值
    osc.measure(4, 'CH1', 'PK2PK')  # 测量峰峰值
    if MSO5 == 1:  # 关闭测量的统计显示
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        # 关闭第一次测量的统计显示功能
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')


def measure2():  # 仅测量CH1 4
    if MSO5 == 1:  # CH1 MAX MIN RMS PK2PK    CH2 MAX MIN FREQUENCY PDUTY占空比
        osc.measure_1(1, 'CH1', 'MAXIMUM')
        # 执行通道1上的最大值测量，并将其标记为测量1
        osc.measure_1(2, 'CH1', 'MINIMUM')
        osc.measure_1(3, 'CH1', 'RMS')
        osc.measure_1(4, 'CH1', 'PK2PK')
        osc.measure_1(5, 'CH4', 'MAXIMUM')
        osc.measure_1(6, 'CH4', 'MINIMUM')
        osc.measure_1(7, 'CH4', 'FREQUENCY')  # 频率测量
        osc.measure_1(8, 'CH4', 'PDUTY')  # 占空比测量
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        # 关闭测量1的统计显示功能
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS5:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS6:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS7:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS8:DISPlaystat:ENABle OFF')
    else:  # 如果是05则只测量不显示统计，反之显示统计信息
        osc.measure(1, 'CH1', 'MAXIMUM')
        osc.measure(2, 'CH1', 'MINIMUM')
        osc.measure(3, 'CH1', 'RMS')
        osc.measure(4, 'CH1', 'PK2PK')
        osc.measure(5, 'CH4', 'MAXIMUM')
        osc.measure(6, 'CH4', 'MINIMUM')
        osc.measure(7, 'CH4', 'FREQUENCY')
        osc.measure(8, 'CH4', 'PDUTY')


def measure3():  # 同上，测量CH1、CH2、CH3、CH4 的 MAX MIN
    if MSO5 == 1:  # 根据示波器的类型选择是否开启统计显示
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


def measure4_1():  # 用于测量 CH1的 MAX、MIN、RISETIME、RISE
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


def measure4_2():  # 测量CH1 MAX MIN FALLTIME FALL
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


def measure5():  # 测量CH4的MAX MIN FRE PDUTY
    osc.measure(1, 'CH4', 'MAXIMUM')
    osc.measure(2, 'CH4', 'MINIMUM')
    osc.measure(3, 'CH4', 'FREQUENCY')
    osc.measure(4, 'CH4', 'PDUTY')
    osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')  # 关闭统计显示
    osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS4:DISPlaystat:ENABle OFF')


def measure5_1():  # 测量 CH4的 MAX MIN PWIDTH
    osc.measure(1, 'CH4', 'MAXIMUM')
    osc.measure(2, 'CH4', 'MINIMUM')
    osc.measure(3, 'CH4', 'PWIDTH')
    osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')
    osc.write('MEASUrement:MEAS3:DISPlaystat:ENABle OFF')


def measure6():  # 测量CH4的 MAX MIN
    osc.measure(1, 'CH4', 'MAXIMUM')
    osc.measure(2, 'CH4', 'MINIMUM')
    if MSO5 == 1:
        osc.write('MEASUrement:MEAS1:DISPlaystat:ENABle OFF')
        osc.write('MEASUrement:MEAS2:DISPlaystat:ENABle OFF')


def measure9():
    # 若是05示波器 测量CH1和 CH3的 MAX MIN  CH4:MAX  CH5:MIN
    # 若不是     则测量 CH1 CH3 CH4的 MAX MIN
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


def common_set():  # 对示波器和电子负载设备进行一系列常见的初始化设置
    osc.state('stop')  # 停止示波器的采集
    el.state('OFF')  # 关闭电子负载
    osc.persistence('OFF')  # 关闭示波器的持久性显示和光标
    osc.cursor('OFF')
    # osc.hormode('MAN')  # 设置示波器的水平模式 Horizontal Mode MAN表示手动模式
    # 为所有通道设置耦合方式为 直流 DC
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
        osc.write('HORIZONTAL:ROLL OFF')  # 关闭水平滚动
    if DPO5104B == 1:
        osc.write('HORIZONTAL:ROLL OFF')
    osc.state('run')  # 启动示波器的采集


def scale1():  # 设置示波器的水平模式为自动模式，并配置横向尺度和采样率
    osc.write('HORIZONTAL:MODE AUTO')  # 自动模式
    osc.write('HORIZONTAL:MODE:SCALE 2e-6')  # 设置时间尺度为 2 微秒/格
    osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置采样率为 10 GS/s


def scale2():  # 同上
    osc.write('HORIZONTAL:MODE AUTO')
    osc.write('HORIZONTAL:MODE:SCALE 1e-1')
    osc.write('HORIZONTAL:MODE:SAMPLERATE 1e6')


def scale2_1():  # 设置示波器为手动模式，并根据计算得到的记录时间调整记录长度
    osc.write('HORIZONTAL:MODE MANual')  # 手动模式
    osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 调整采样率
    postime = dy * 1000000 / (vin * freq)  # 计算记录时间 电压*1000000/输入电压*频率
    if postime >= 2000:  # 根据计算的 postime 调整记录长度，以适应长时间的波形记录
        osc.record(100000)
    elif postime >= 1000:
        osc.record(80000)
    elif postime >= 500:
        osc.record(40000)
    elif postime >= 200:
        osc.record(20000)
    else:
        osc.record(10000)


def scale2_2():  # 按照计算得到的记录时间调整横向scale  更精细、更短时间的记录
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


def scale3():  # 在不同幅度的信号测量中自动调整示波器设置，优化信号显示
    # 根据通道1的峰峰值（pk2pk）自动调整纵向尺度
    rpk = osc.query('MEASUrement:MEAS4:MAX?')  # 查询Pk2pk的最大值
    if MSO5 == 1:
        rpk = rpk[26:]  # 如果是05示波器，需要从第26个字符开始
    rpk = float(rpk)  # 字符串转换为浮点数
    rpk_1 = rpk * 10000
    rpk_1 = int(rpk_1)  # 放大10000倍转换为整数，处理峰峰值的量级，适合于范围判断
    if rpk_1 in range(0, 410):  # 根据rpk_1(峰峰值的整数表示) 调整通道一的纵向尺度
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
    else:  # 如果 rpk_1 超过范围 打印提示
        print("Out of Output Voltage Range!")


def tl1_channel_set():
    osc.horpos(50)  # 水平位置 50%
    osc.chanset('CH1', 0, dy, '20.0000E+06', 10E-02)  # 设置通道1的位置
    osc.label('CH1', entry.get(), 1, 6)  # 设置label
    osc.trigger('AUTO', 'CH1', 'RISE', dy)  # 触发设置
    if MSO5 == 1:  # 如果是05示波器，关闭通道1的显示再打开（刷新）
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')


def tl1_1_channel_set():  # 同上
    osc.horpos(50)  # 水平位置
    osc.chanset('CH1', 0, dy, '20.0000E+06', 10E-02)
    osc.label('CH1', entry.get(), 1, 6)  # 设置label
    osc.trigger('AUTO', 'CH1', 'RISE', dy)


def tl2_1_channel_set():  #
    osc.horpos(40)  # 水平位置
    osc.chanset('CH1', 2, dy, '20.0000E+06', 10E-02)
    # CH1 通道标识符 2:垂直位置  dy:垂直缩放尺度  20.0000E+06 时间基准  10E-0.2 时间尺度0.01s
    ldstep = int(ld_max / 3)  # 设置通道4的纵向尺度
    osc.chanset('CH4', -4, 0, '20.0000E+06', ldstep)
    osc.label('CH1', entry.get(), 2, 4)  # 设置label
    osc.label('CH4', "Iout", 2, 10)
    osc.trigger('AUTO', 'CH4', 'RISE', ld_max / 2)
    # 设置自动触发模式 CH4
    if MSO5 == 1:  # 刷新通道显示
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')
    if infinite_off.get() == 'True':
        osc.persistence('OFF')  # 关闭累积显示
    else:
        osc.persistence('INFPersist')  # 开启无限累计模式


def tl2_2_channel_set():
    osc.horpos(40)  # 水平位置
    osc.chanset('CH1', 2, dy, '20.0000E+06', 10E-02)
    ldstep = int(ld_max / 3)
    osc.chanset('CH4', -4, 0, '20.0000E+06', ldstep)
    osc.label('CH1', entry.get(), 1, 4)  # 设置label
    osc.label('CH4', "Iout", 2, 10)  # 设置label
    osc.trigger('AUTO', 'CH4', 'RISE', ld_max / 2)
    if infinite_off.get() == 'True':
        osc.persistence('OFF')  # 关闭累积
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
        osc.persistence('OFF')  # 关闭累积
    else:
        osc.persistence('INFPersist')  # 开启累积


def tl3_channel_set():  # 根据电压 dy 的值设置示波器的四个通道，并调整它们的显示设置
    osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
    if dy >= 5:  # 根据电压值 dy调整通道1的设置
        osc.chanset('CH1', -2, 0, '20.0000E+06', 1.5)
    elif dy >= 3.3:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 1)
    elif dy >= 2:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.7)
    elif dy >= 1.5:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.5)
    else:
        osc.chanset('CH1', -2, 0, '20.0000E+06', 0.4)
    osc.chanset('CH2', -1, 0, '20.0000E+06', 1)  # 设置通道2、3、4
    osc.chanset('CH3', -3, 0, '20.0000E+06', 1)
    osc.chanset('CH4', -0, 0, '20.0000E+06', 3)
    osc.label('CH1', entry.get(), 1, 9)  # 设置label
    osc.label('CH2', "EN", 1.5, 9)
    osc.label('CH3', "PG", 2, 9)
    osc.label('CH4', "VIN", 2.5, 9)
    if MSO5 == 1:  # 刷新通道显示
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        osc.write('DISplay:GLObal:CH2:STATE OFF')
        osc.write('DISplay:GLObal:CH3:STATE OFF')
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')
        osc.write('DISplay:GLObal:CH2:STATE ON')
        osc.write('DISplay:GLObal:CH3:STATE ON')
        osc.write('DISplay:GLObal:CH4:STATE ON')


def tl4_channel_set():  # 据电压 dy 的值调整示波器的 CH1 通道设置
    osc.horpos(40)  # 水平位置
    if dy >= 5:  # 同上
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
    if MSO5 == 1:  # 刷新显示
        osc.write('DISplay:GLObal:CH1:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH1:STATE ON')


def tl5_channel_set():  # 根据 vin 的值来配置示波器的 CH4 通道
    if vin >= 10:  # 同上
        osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
        osc.label('CH4', "PHASE", 2, 8)
    else:
        osc.chanset('CH4', -1, 0, '500.0000E+06', 2)
        osc.label('CH4', "PHASE", 2, 8)
    osc.trigger('NORMAL', 'CH4', 'RISE', 6)
    #  触发模式：NORMAL  通道：CH4  触发条件：RISE  触发位置：6
    # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率
    # osc.persistence('INFPersist')  # 开启累积
    el.write('CHAN 1')  # 选择电子负载的通道为1
    if MSO5 == 1:  # 刷新通道的显示
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')


def tl5_jitter_set():  # 从示波器中测量和计算信号的抖动
    global jitter  # 声明全局变量jitter
    jitter_max = osc.query('MEASUrement:MEAS3:MAX?')  # 查询通道3的最大值
    jitter_max = float(jitter_max)  # 转换为浮点数
    jitter_min = osc.query('MEASUrement:MEAS3:MINI?')  # 查询通道3的最小值
    jitter_min = float(jitter_min)
    jitter = jitter_max - jitter_min  # 计算抖动=最大值-最小值
    jitter = int(jitter * 1000000000)  # 转换为纳秒


def tl6_channel_set_1():  # 配置示波器的水平位置、通道设置、触发、标签和累积模式
    osc.horpos(10)  # 水平位置
    # osc.chanset('CH3', 2, vin, 'FULL', 0.3)
    if vin >= 10:  # 同上
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


def go():  # 从 Excel 文件中提取数据并更新相关全局变量
    # 处理事件，*args表示可变参数
    global dy, ld_max, freq, ocp_spec, temp, jitter, osc, el, xls, vin, file_path
    # 全局变量声明
    xls = EasyExcel(file_path)
    # 用 EasyExcel类的实例xls来打开指定路径的 Excel文件
    temp = xls.getCell('Test Summary', 6, 3)
    # 从 Test Summary工作表中的第六行第三列读取数据，存储在temp中，以下同上
    dy = xls.getCell(entry.get(), 5, 10)
    ld_max = xls.getCell(entry.get(), 28, 3)
    freq = xls.getCell(entry.get(), 160, 3) / 0.9
    ocp_spec = xls.getCell(entry.get(), 292, 3)
    vin = xls.getCell(entry.get(), 5, 11)
    print(dy)
    print(ld_max)
    print(freq)
    print(ocp_spec)
    EnValue2.set(dy)    # 更新UI组件


def instrument():
    global osc, el, rm, MSO5, DPO7000, DPO5104B
    # 仪器的全局变量声明
    rm = pyvisa.ResourceManager()
    # 创建一个资源管理器实例，用于查询连接的仪器
    insadd = rm.list_resources()
    # 列出所有可用的资源地址
    print(insadd)
    DPO7000 = 0
    DPO5104B = 0
    MSO5 = 0
    CH6310 = 0
    CH63600 = 0
    # 初始化设备表示，所有设备的标志默认为0.即未连接
    for addr in insadd:             # 遍历每一个仪器地址
        str0 = addr.find('GPIB')    # 检查当前地址是否包括 GPIB 或 USB
        str01 = addr.find('USB')    # 没有找到则返回 -1
        if str0 != -1 or str01 != -1:   # 如果地址包含 GPIB USB 则为有效仪器地址
            ins = rm.open_resource(addr)    # 创建仪器接口对象，打开仪器资源
            insinf = ins.query('*IDN?')     # 查询仪器型号
            insinf = insinf.upper()         # 身份信息转换为大写，方便查找比较
            print(insinf)                   # 打印仪器身份信息
            str1 = insinf.find('TEKTRONIX,DPO7')    # 查找身份信息中是否包含DP07
            if str1 != -1:
                print('该仪器型号为TEKTRONIX,DPO7000系列示波器，设备连接成功')
                print('地址为' + addr)     # 添加并打印地址
                osc = OscDPO7000C(addr)   # 初始化示波器对象
                DPO7000 = 1     # 表示DP70示波器已经连接
            str2 = insinf.find('TEKTRONIX,MSO')     # 同上
            if str2 != -1:
                print('该仪器型号为TEKTRONIX,MSO4/5/6系列示波器，设备连接成功')
                print('地址为' + addr)
                osc = OscMPO5series(addr)
                MSO5 = 1
            str3 = insinf.find('CHROMA,631')        # 更新电子负载 同上示波器
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
    oscstate = DPO7000 or MSO5 or DPO5104B  # 检查是否至少有一个示波器和一个电子负载已经连接
    elstate = CH6310 or CH63600
    if oscstate and elstate:    # 检查是否同时成功连接了示波器和电子负载
        messagebox.showinfo(title='仪器连接', message='示波器和电子负载均已正确连接')
    elif oscstate:
        messagebox.showerror(title='仪器连接', message='电子负载连接错误，请检查')
    elif elstate:
        messagebox.showerror(title='仪器连接', message='示波器连接错误，请检查')
    else:
        messagebox.showerror(title='仪器连接', message='示波器和电子负载均连接错误，请检查')


# —————————————————————————————分割线———————————————————————————————

# 测试部分
def tl0():  # T-0 DMM&Scope Offset Record
    root0 = Toplevel()      # # 创建一个新的顶层窗口root0，独立于主窗口root，作为当前测试窗口
    root0.title('T-0 DMM&Scope Offset Record')      # # 设置窗口标题
    root0.geometry('340x200')       # 设置新窗口尺寸
    root0.transient(root)           # 将root0作为root的临时窗口,前者关闭不影响后者
    Label(root0,
          text='测试前请校准探头，请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    # 在root1中创建文本标签，设置最大宽度为300(超过自动换行), 文本对齐到标签的左边, 将标签放到窗口的第0行第0列, 占据两列, 设置内边距
    Button(root0, text="开始测试", command=t00).grid(row=1, column=0, padx=5, pady=20)
    # 创建一个按钮，文本内容为开始测试，被点击时调用 t00()函数进行测试
    quit11 = Button(root0, text='退出测试', command=root0.destroy, activeforeground='white', activebackground='red')
    # 创建 退出测试按钮，点击按钮时关闭 root1 窗口，激活按钮时(悬停)变成白色，被激活后变成红色
    quit11.grid(row=1, column=1, padx=5, pady=20)  # 退出按钮的设计
    root0.attributes("-topmost", 1)     # 将root0设置为总在最前面显示

def tl1():  # T-1 DC Regulation+Ripple&Noise  Test
    root1 = Toplevel()  # 创建一个新的顶层窗口root1，独立于主窗口root，作为当前测试窗口
    root1.title('T-1 DC Regulation+Ripple&Noise Test')  # 设置窗口标题
    root1.geometry('340x200')       # 设置新窗口尺寸
    root1.transient(root)           # 将root1作为root的临时窗口,前者关闭不影响后者
    Label(root1, text='请将差分探头一端连接到示波器的一通道，另一端连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    # 在root1中创建文本标签，设置最大宽度为300(超过自动换行), 文本对齐到标签的左边, 将标签放到窗口的第0行第0列, 占据两列, 设置内边距
    Button(root1, text="开始测试", command=t01).grid(row=1, column=0, padx=5, pady=20)
    # 创建一个按钮，文本内容为开始测试，被点击时调用 t01()函数进行测试
    quit11 = Button(root1, text='退出测试', command=root1.destroy, activeforeground='white', activebackground='red')
    # 创建 退出测试按钮，点击按钮时关闭 root1 窗口，激活按钮时(悬停)变成白色，被激活后变成红色
    quit11.grid(row=1, column=1, padx=5, pady=20)  # 退出按钮布局设计
    root1.attributes("-topmost", 1)     # 将root1设置为总在最前面显示

def tl2():
    global infinite_off     # 全局变量声明 控制累积模式的状态
    root2 = Toplevel()
    root2.title('T-2 Loading Transient Response Test')
    root2.geometry('360x330')
    group2 = LabelFrame(root2, text='单项测试', padx=5, pady=5)
    # 创建一个LabelFrame容器 group2，用于分组单项测试按钮，标题是单项测试
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
    quit21 = Button(root2, text='退出测试', command=root2.destroy, activeforeground='white', activebackground='red')
    quit21.grid(row=1, column=2, padx=40, pady=5)  # 退出按钮的设计

    infinite_off = StringVar()      # 创建一个StringVar类型的变量，用于跟踪复选框的状态
    infinite = Checkbutton(root2, text='关闭累积', variable=infinite_off, onvalue="True", offvalue="False",
                           state="normal")
    # 创建一个复选框，文本内容为关闭累积，复选框的状态由inifite_off 控制，当复选框被勾选时值为True，反之False
    infinite.grid(row=2, column=1)
    infinite_off.set("False")       # 复选框初始值设置为 False，默认未勾选
    root2.attributes("-topmost", 1)     # root2 最前显示

def tl3():
    root3 = Toplevel()
    root3.title('T-3 Power Up & Down Sequence Measurement')
    root3.geometry('400x400')
    group3 = LabelFrame(root3, text='单项测试', padx=5, pady=5)
    group3.grid(row=2, rowspan=2, column=0, columnspan=3, padx=60, pady=15)
    Label(root3, text="请使用探头1连接示波器的一通道和待测VR的输出端，使用探头2连接示波器的二通道和待测VR的EN信号，"
                      "使用探头3连接示波器的三通道和待测VR的PG信号，使用探头4连接示波器的四通道和待测VR的VIN信号，"
                      "单击“开始测试”进行测试。", wraplength=300,
          anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    theButton31 = Button(group3, text="运行 Power Up Sequence with NO Load", command=t03_1)  # 按下按钮 执行 t03_1函数
    theButton31.grid(row=1, column=0, sticky=E + W, padx=5, pady=5)
    theButton32 = Button(group3, text="运行 Power Down Sequence with NO Load", command=t03_2)  # 按下按钮 执行 t03_2函数
    theButton32.grid(row=2, column=0, sticky=E + W, padx=5, pady=5)
    theButton33 = Button(group3, text="运行 Power Up Sequence with Max Load", command=t03_3)  # 按下按钮 执行 t03_3函数
    theButton33.grid(row=3, column=0, sticky=E + W, padx=5, pady=5)
    theButton34 = Button(group3, text="运行 Power Down Sequence with Max Load", command=t03_4)  # 按下按钮 执行 t03_4函数
    theButton34.grid(row=4, column=0, sticky=E + W, padx=5, pady=5)
    root3.attributes("-topmost", 1)

    theButton35 = Button(root3, text="开始测试", command=t03_5)  # 按下按钮 执行 t03_5函数
    theButton35.grid(row=1, column=0, padx=60, pady=5)
    quit31 = Button(root3, text='退出测试', command=root3.destroy, activeforeground='white', activebackground='red')
    quit31.grid(row=1, column=1, padx=50, pady=5)  # 退出按钮的设计

def tl4():
    root4 = Toplevel()
    root4.title('T-4 OVS & UDS Sequence Measurement')
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
    root5.title('T-5 Switching Fre. & Jitter Measurement')
    root5.geometry('355x400')
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
    global infinite_off_6       # 同tl2
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
    quit61 = Button(root6, text='退出测试', command=root6.destroy, activeforeground='white', activebackground='red')
    quit61.grid(row=1, column=2, padx=10, pady=5)  # 退出按钮的设计

    infinite_off_6 = StringVar()
    infinite = Checkbutton(root6, text='关闭累积', variable=infinite_off_6,
                           onvalue="True", offvalue="False", state="normal")
    infinite.grid(row=2, column=1)
    infinite_off_6.set("False")
    root6.attributes("-topmost", 1)

def tl7():
    root7 = Toplevel()
    root7.title('T-7 Bode Plots Measurement')
    root7.geometry('400x300')
    Label(root7, text='该项测试正在开发中，敬请期待...').pack()
    root7.attributes("-topmost", 1)

def tl8():
    root8 = Toplevel()
    root8.title('T-8 Efficiency Measurement')
    root8.geometry('400x300')
    Label(root8, text='该项测试正在开发中，敬请期待...').pack()
    root8.attributes("-topmost", 1)

def tl9():
    global ocpmode      # 全局变量声明 ocpmode 示波器模式
    root9 = Toplevel()                      # 创建一个子窗口
    root9.title('T-9 OCP & SCP Test')       # 设置窗口标题
    root9.geometry('350x420')               # 设置窗口大小
    group9 = LabelFrame(root9, text='单项测试', padx=5, pady=5)
    group9.grid(row=2, rowspan=2, column=0, columnspan=3, padx=20, pady=15)
    # 创建带标题的框架容器，标题为单项测试
    Label(root9, text='请使用探头1连接示波器的一通道和待测VR的输出端，探头3连接示波器的三通道和待测VR的PG信号输出端，使用电流探棒1连接示波器的四通道和待测'
                      'VR的输出电流线缆，单击测试项进行测试。', wraplength=300, anchor='w').grid(row=0,column=0, columnspan=3,padx=20, pady=20)
    Label(root9, text='OCP模式：').grid(row=1, column=0, sticky=E)

    ocpmode = IntVar()  # 创建一个IntVar类型的变量，存储单选按钮的选中状态
    ocpmode.set(1)      # 默认选中值为1，即默认选择 Hiccip模式
    Radiobutton(root9, text='Latch', variable=ocpmode, value=0).grid(row=1, column=1)
    Radiobutton(root9, text='Hiccup', variable=ocpmode, value=1).grid(row=1, column=2)
    # 创建两个单选按钮，共用ocpmode变量,一0表示 Latch模式，1表示 Hiccip模式

    theButton91 = Button(group9, text="运行 Slow OCP Test", command=t09_1)  # 按下按钮 执行 t05_1函数
    theButton91.grid(row=1, column=0, sticky=E + W, padx=30, pady=5)
    theButton92 = Button(group9, text="运行 Fast OCP Test", command=t09_2)  # 按下按钮 执行 t05_1函数
    theButton92.grid(row=2, column=0, sticky=E + W, padx=30, pady=5)
    theButton93 = Button(group9, text="运行 SCP before Power on Test", command=t09_3)  # 按下按钮 执行 t05_1函数
    theButton93.grid(row=3, column=0, sticky=E + W, padx=30, pady=5)
    theButton94 = Button(group9, text="运行 SCP after Power on Test", command=t09_4)  # 按下按钮 执行 t05_1函数
    theButton94.grid(row=4, column=0, sticky=E + W, padx=30, pady=5)
    root9.attributes("-topmost", 1)     # 窗口置顶

def tl10():
    global display, counter
    root10 = Toplevel()
    root10.title('T-10 Thermal Test')
    root10.geometry('340x200')
    root10.transient(root)  # 主窗口关闭时，子窗口随之关闭

    Label(root10, text='请将电子负载通过负载线连接到待测VR的输出端，单击“开始测试”进行测试。',
          wraplength=300, anchor='w').grid(row=0, column=0, columnspan=2, padx=20, pady=20)
    Label(root10, text='测试累积时间：', anchor='w').grid(row=1, column=0, padx=5, pady=10)
    display = Label(root10, text='00:00', anchor='w')
    display.grid(row=1, column=1, padx=5, pady=10)
    # 显示测试累积时间的标签，初始值为 00:00

    button = Button(root10, text="开始测试", command=t10)
    button.grid(row=2, column=0, padx=5, pady=5)
    quit11 = Button(root10, text='停止测试', command=t10_1, activeforeground='white', activebackground='red')
    quit11.grid(row=2, column=1, padx=5, pady=5)  # 退出按钮的设计
    counter = 0
    root10.attributes("-topmost", 1)

def tl11():     # 保留测试
    xls.save()
    xls.close()

def t00():
    if MSO5 == 1:                   # 对于MS05系列示波器
        osc.write('FACTORY')        # 将示波器恢复到出产设置
        common_set()                # 调用函数进行常规设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_000.SET"')     # 加载预设的示波器配置文件
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')        # 只开启通道1
        tl1_channel_set()   # 对通道1进行设置和测量
        measure1()

        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')   # 关闭测量注释
        osc.write('HORIZONTAL:MODE AUTO')               # 设置水平模式为自动
        osc.write('HORIZONTAL:MODE:SCALE 2e-5')         # 调整水平刻度ms
        osc.write('HORIZONTAL:POSITION 50')             # 调整水平位置50
        tdc = xls.getCell(entry.get(), 8, 4)   # 从 Excel文件中获取测试数据

        el.static(1, 'MAX', tdc)  # 设置电子负载的状态和测试条件
        el.state('ON')            # 开启电子负载
        osc.state('run')          # 示波器开始采样
        time.sleep(5)             # 等待示波器采样完成

        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')  # 查询测量值的平均值
        trigger_rms = trigger_rms[24:]                      # 提取有效测量数据
        print(trigger_rms)
        trigger_rms = float(trigger_rms)            # 获取测量数据浮点数
        osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置示波器通道1的偏移值
        time.sleep(3)

        scale3()            # 进行尺度设置
        time.sleep(2)
        trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')  # 重复查询
        trigger_rms = trigger_rms[24:]
        trigger_rms = float(trigger_rms)
        time.sleep(1)
        osc.trigger('AUTO', 'CH1', 'RISE', trigger_rms)
        time.sleep(1)
    else:               # 如果不是 MS05系列示波器，其他默认设置
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

    osc.state('run')    # 示波器开始采样
    osc.number(300)     # 采样点数设置为300
    osc.state('stop')   # 示波器停止采样
    time.sleep(2)       # 确保数据被保存

    RMSwindow = messagebox.askquestion(title='程序执行完毕',message='程序已执行完毕，请确认波形是否正确，如果正确请使用六位半数字万用表测量输出端电压并在表格中填写！失败请点击否')
    # 弹出对话框，询问用户测试是否成功
    if RMSwindow == 'yes':
        rms = osc.query('MEASUrement:MEAS3:MEAN?')  # 从示波器中获取 RMS值
        if MSO5 == 1:
            rms = rms[24:]
        rms = float(rms)    # MS05数据转换
        xls.setCell(entry.get(), 8, 8, rms)     # RMS值写入 Excel指定单元格
        time.sleep(1)
        savepic('T0')  # 保存图片
        a0 = r'%s/POL Test Pictures/%s/%s/T0.png' % (pic_path, temp, entry.get())
        # 构造图片保存路径
        if v.get() == 1:    # 将图片添加到指定单元格N1,并设置图片的位置和尺寸
            xls.addPicture(entry.get(), a0, 'N1', 25, 0, 337, 212)
        else:
            xls.addPicture(entry.get(), a0, 'N1', 25, 0, 337, 212)
        xls.save()          # 保存 Excel文件更改
        el.state('OFF')     # 关闭电子负载
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
        # ld是负载最大值，ldsp是每次负载增加的步长，初始化操作

        while block <= 28:  # block 23-28
            el.static(1, 'MAX', load)  # 选择电子负载通道1，设为最大值并打开
            el.state('ON')
            osc.state('run')  # 示波器开始采样
            time.sleep(2)
            trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
            trigger_rms = trigger_rms[24:]
            trigger_rms = float(trigger_rms)
            # 查询示波器的测量3的均值,处理返回的字符串

            osc.write('CH1:OFFSET %.2f' % trigger_rms)  # 设置offset
            time.sleep(5)
            scale3()
            time.sleep(3)
            trigger_rms = osc.query('MEASUrement:MEAS3:MEAN?')
            trigger_rms = trigger_rms[24:]
            trigger_rms = float(trigger_rms)
            time.sleep(2)
            osc.trigger('AUTO', 'CH1', 'RISE', trigger_rms)
            # 配置示波器的触发设置
            time.sleep(1)
            osc.state('run')  # 示波器开始采样
            osc.number(300)
            osc.state('stop')  # 示波器停止采样
            time.sleep(1)
            el.state('OFF')
            time.sleep(2)

            if i == 0:
                ocpwindow = messagebox.askquestion(title='测试确认', message='测试成功，请点击是确认波形是否正确！')
                # 对话框询问波形数据是否正确, 是则返回 ocpwindow的值为yes, 否则为no
                if ocpwindow == 'yes':  # 波形数据正确则执行以下代码块
                    rms = osc.query('MEASUrement:MEAS3:MEAN?')
                    rms = rms[24:]
                    rms = float(rms)    # 查询示波器测量通道3的均值 RMS
                    xls.setCell(entry.get(), block, 4, rms) # RMS 值写入
                    time.sleep(1)

                    pk2pk = osc.query('MEASUrement:MEAS4:MAX?')
                    pk2pk = pk2pk[27:]      # 查询示波器测量通道4的峰峰值PK-PK
                    pk2pk = float(pk2pk)
                    pk2pk_excel = pk2pk / 0.001
                    time.sleep(1)
                    xls.setCell(entry.get(), block, 5, pk2pk_excel)

                    savepic('T1-1')  # 保存图片
                    a1_1 = r'%s/POL Test Pictures/%s/%s/T1-1.png' % (pic_path, temp, entry.get())
                    # pic_path 图片目录根路径   temp 临时存储 entry.get()输入值
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

            elif i == 2:    # 同 i==1
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

            elif i == 3:    # 同 i==0
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

            elif i == 4:        # 同 i==2
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

            elif i == 5:        # 同 i==3
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
            load = load + ldsp  # 负载迭代
            block = block + 1  # 测试阶段迭代
            i = i + 1  # 测试轮数 i 迭代
            time.sleep(1)
    else:       # 非 MS05系列示波器
        common_set()        # 同上
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

        while block <= 28:  # block 23-28
            el.static(1, 'MAX', load)  # 选择电子负载通道1，设为最大值并打开
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

            if i == 0:    # i=0 3 5 相同  i=1 2 4 相同
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
            load = load + ldsp  # 负载迭代
            block = block + 1   # 测试阶段迭代
            i = i + 1           # 测试轮数 i 迭代
            time.sleep(1)
    xls.save()  # 保存 Excel 文件 到指定位置
    messagebox.showinfo(title='程序执行完毕', message='程序已执行完毕，点击确定结束测试')
    # 弹出信息框显示 测试已完成

def t02_1():
    if MSO5 == 1:   # 使用 MS05系列示波器
        osc.write('FACTORY')    # 以下同t01
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

        # 获取 RMS 值并设置 offset
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
    if window == 'yes':         # 用户确认测试结果
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
    if MSO5 == 1:       # 同01
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
    if MSO5 == 1:       # 同01
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
    # 发送信息 程序执行完毕

def t03_1():
    if MSO5 == 1:       # 示波器05设置
        osc.write('FACTORY')    # 恢复出厂设置
        common_set()            # 常规设置
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')     # 加载预设文件
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')   # 打开1-4通道，关闭5-6通道
        tl3_channel_set()   # 通道设置
        measure3()          # 测量设置

        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')   # 关闭测量注释显示
        osc.write('HORIZONTAL:MODE AUTO')               # 水平时间轴模式为自动
        osc.write('HORIZONTAL:MODE:SCALE 4e-2')         # 水平时间尺度为4e-2
        tri_level = float(2.48)                             # 定义触发电平为2.48
        osc.trigger('NORMAL', 'CH3', 'RISE', tri_level)     # 设置示波器的触发模式
        # 参数设置  模式：NORMAL   通道：CH3   趋势：RISE   触发电平大小：tri_level
        osc.write('HORIZONTAL:POSITION 60')                 # 设置水平时间位置为60%
        osc.state('single')                                 # 设置为single模式
        time.sleep(2)  # 设置延时为主板上下电作准备

    else:       # 不使用 05示波器
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

    powerup = messagebox.askquestion(title='电路上电确认',message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
    # 弹出消息框，询问用户是否已经成功上电，结果存储在powerup变量中
    if powerup == 'yes':
        time.sleep(2)
        savepic('T3-1')  # 保存图片
        a3_1 = r'%s/POL Test Pictures/%s/%s/T3-1.png' % (pic_path, temp, entry.get())
        # a3_1 = 'D:\\POL Test Pictures\\' + temp + '\\' + entry.get() + r'\Power Up Sequence with NO Load.png'
        if v.get() == 1:
            xls.addPicture(entry.get(), a3_1, 'F67', 36, 10, 362, 220)
        else:
            xls.addPicture(entry.get(), a3_1, 'F67', 36, 10, 362, 220)
        return 1    # 上电成功
    else:
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        return 0    # 上电失败

def t03_2():
    if MSO5 == 1:   # 同 t03_1
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

    powerup = messagebox.askquestion(title='电路下电确认',message='请进行电路下电，下电成功请点击是确认存图，失败请点击否')
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
    if MSO5 == 1:       # 同 t03_1
        osc.write('FACTORY')
        common_set()
        osc.write('RECALL:SETUP "C:/Tektronix/Tek000_001.SET"')
        osc.channel('ON', 'ON', 'ON', 'ON', 'OFF', 'OFF')
        tl3_channel_set()
        measure3()

        osc.write('MEASUREMENT:ANNOTATION:STATE OFF')
        osc.write('HORIZONTAL:MODE AUTO')
        osc.write('HORIZONTAL:MODE:SCALE 4e-2')
        el.static(1, 'MIN', ld_max)     # 设置电子负载的通道和参数
        # 参数设置  选择通道：1  模式：MIN   负载值：ld_max
        el.write('CONF:LVP ON')     # 打开电子负载的低电压保护 LVP模式
        el.state('ON')              # 打开电子负载
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

    powerup = messagebox.askquestion(title='电路上电确认',message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
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
    if MSO5 == 1:           # 同 t3_3
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

    powerup = messagebox.askquestion(title='电路下电确认',message='请进行电路下电，下电成功请点击是确认存图，失败请点击否')
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
        tri_level = float(0.5 * dy)     # 计算触发电平
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
        ch1max = osc.query('MEASUrement:MEAS1:MAX?')    # 查询示波器上通道1的最大值
        if MSO5 == 1:
            ch1max = ch1max[26:]    # 处理05示波器的返回值格式
        ch1max = float(ch1max)      # 将字符串类型的 ch1max转换为浮点数类型
        xls.setCell(entry.get(), 111, 3, ch1max)    # 将ch1max值写入 Excel指定单元格
        return 1    # 操作成功

    else:
        messagebox.showerror(title='错误', message='电路上电失败，请退出重试')
        return 0    # 上电失败

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
        osc.channel('ON', 'OFF', 'OFF', 'OFF', 'OFF', 'OFF')    # 只开启通道1
        tl4_channel_set()
        measure4_1()

        el.static(1, 'MIN', ld_max)     # 电子负载设置
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
        ch4freq = osc.query('MEASUrement:MEAS3:MEAN?')      # 查询MEAS3的测量均值(频率)
        if MSO5 == 1:
            ch4freq = ch4freq[24:]
        ch4freq = float(ch4freq)
        ch4freq = ch4freq / 1000        # 数据格式转换
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
            # 查询游标1的 X轴位置
            b = b[51:]
            print(b)
            b = float(b) * 1000000000
            print(b)        # 数值转换为纳秒 ns
            c = osc.query('DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:SCREEN:BXPOSITION?')
            c = c[51:]
            print(c)
            c = float(c) * 1000000000
            print(c)        # 同上打印游标2的X轴位置，单位为纳秒
            xls.setCell(entry.get(), 173, 3, b)
            xls.setCell(entry.get(), 173, 4, c) # 游标位置写入单元格
            time.sleep(1)

            poswidth = osc.query('MEASUrement:MEAS3:MEAN?')
            poswidth = poswidth[24:]
            poswidth = float(poswidth)
            poswidth = poswidth * 1000000000     # 获取脉宽平均值并处理

        else:
            text = osc.query('CURSOR:VBARS?')       # 查询两个垂直游标的位置
            print(text)         # 返回的字符串包含两个游标的位置和游标之间的时间差
            split_text = text.split(";")    # 按照分号分割返回的字符串，存入列表
            if len(split_text) == 3:        # 分割后的列表是否包含三个元素
                a, b, c = split_text    # a b 游标1 2 位置 c 游标之间时间差
                b = float(b.strip()) * 1000000000
                c = float(c.strip()) * 1000000000       # 去除空格，转换为浮点数
                print("a:", a.strip())
                print("b:", b)
                print("c:", c)
            else:
                print("分割后的项数量不是三个")

            xls.setCell(entry.get(), 173, 3, b)
            xls.setCell(entry.get(), 173, 4, c)
            time.sleep(1)   # 确保数据写入单元格
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

        if vin >= 10:       # 根据输入电压的值来设置 CH4的通道配置和触发电平
            osc.chanset('CH4', -2, 0, '500.0000E+06', 3)
            # CH4 通道  -2 垂直便宜  0触发电平  500MHz 带宽限制 3 垂直刻度
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 4)
            osc.trigger('NORMAL', 'CH4', 'RISE', 6)
            # NORMAL 模式  CH4 通道 RISE 边缘类型 5 触发电平
        else:
            osc.chanset('CH4', -2, 0, '500.0000E+06', 2)
            # osc.math('MATH1', 'CH3-CH4', 0, -3, 2)
            osc.trigger('NORMAL', 'CH4', 'RISE', 3)

        # osc.label('CH3', "VIN_FET", 1.5, 2.5)
        osc.label('CH4', "HS_Vds", 1, 7)
        # CH4 通道 HS_Vds 标签文本  1 Y轴位置  7 X轴位置
        # osc.label('MATH:MATH1', "VDS_HIGH", 2, 7.5)
        # osc.write('HORIZONTAL:MODE:SAMPLERATE 1e10')  # 设置波形采样频率和scale

        if infinite_off_6.get() == 'True':
            osc.persistence('OFF')  # 关闭累积
        else:
            osc.persistence('INFPersist')  # 开启累积
        osc.write('DISplay:GLObal:CH4:STATE OFF')
        time.sleep(1)
        osc.write('DISplay:GLObal:CH4:STATE ON')    # 刷新配置
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

        if pg == 0: # 根据变量 pg 的值，决定哪些通道被激活
            osc.channel('ON', 'OFF', 'OFF', 'ON', 'OFF', 'OFF')
        else:
            osc.channel('ON', 'OFF', 'ON', 'ON', 'OFF', 'OFF')

        osc.horpos(25)  # 水平位置
        scale_v = float(dy / 3)
        scale_v = round(scale_v)        # 计算水平位置浮点数
        osc.chanset('CH1', 1, 0, '20.0000E+06', scale_v)
        # CH1:通道  1:电压偏移量  0:垂直位置  20MV:电压范围  scale_v:垂直比例尺度
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
    ld_ocp = ld_max                             # 初始负载电流设为最大值
    ldmax_step = 0.005 * ocp_spec               # 每次增加的负载电流步长
    while ld_ocp <= ocp_spec * 1.5:             # 直到负载电流超过标准的1.5倍
        ld_ocp = ld_max + ldmax_step * i
        i = i + 1
        el.static(1, 'MAX', ld_ocp)
        el.state('ON')
        time.sleep(1)
        tri_state = osc.query('ACQUIRE:STATE?') # 查询示波器的采集状态
        if MSO5 == 1:
            tri_state = tri_state[15:]
        tri_state = int(tri_state)              # 采集MS05的返回值
        # el.state('OFF')
        # time.sleep(2)

        if tri_state != 1:      # 状态不是1表示测试成功
            ocpwindow = messagebox.askquestion(title='OCP测试确认', message='OCP测试成功，请点击是确认存图，失败请点击否')
            if ocpwindow == 'yes':      # 如果用户确认则保存图片
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
        elif ld_ocp >= ocp_spec * 1.5:      # 电流超载
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

    powerup = messagebox.askquestion(title='电路上电确认',message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
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

    powerup = messagebox.askquestion(title='电路上电确认',message='请进行电路上电，上电成功请点击是确认存图，失败请点击否')
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
    global countmode        # 计数变量
    tdc = xls.getCell(entry.get(), 324, 3)      # 负载电流值
    el.static(1, 'MAX', tdc)
    el.state('ON')
    countmode = 'ON'
    count()

def t10_1():
    global countmode
    countmode = 'OFF'
    el.state('OFF')

# —————————————————————————————分割线———————————————————————————————
# 界面设置

def select_excel_path():    # 选择文件路径
    global file_path        # 全局变量声明
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    # 弹出一个文件对话框，让用户选择一个文件，类型是xlsx或者xls
    print(file_path)
    if file_path:           # 如果文件路径正确，更新文件路径的值到EnValues3控件
        EnValue3.set(file_path)


def select_pic_path():      # 选择图片路径
    global pic_path         # 全局变量声明
    pic_path = filedialog.askdirectory()    # 弹出对话框选择目录，存储于pic_path变量中
    print(pic_path)
    if pic_path:            # 如果文件路径正确，更新文件路径的值到EnValues4控件
        EnValue4.set(pic_path)


global dy, ld_max, freq, ocp_spec, temp, jitter, osc, el, xls, vin, display, counter, \
    infinite_off_6, infinite_off, ocpmode, osc_addr, el_addr, MSO5, EnValue3, rm, DPO7000, countmode, file_path, pic_path
# 全局变量声明

root = Tk()  # 创建一个Tkinter的主窗口实例，root是这个主窗口的引用
root.title('Suma Power Test')   # 标题：Suma Power Test
root.resizable(False, False)    # 禁止调整窗口大小
root.geometry('730x495')        # 设置初始框的大小为 730×495
# root.focus_force() # 强制聚焦窗口
visa_dll = 'c:/windows/system32/visa32.dll'     # 指定VISA库的DLL文件路径
time.sleep(1)

image_file = PhotoImage(file='suma1.png')   # 加载图片（png、gif）
image = Label(root, image=image_file)       # 将图片设置为标签的内容
image.grid(row=0, column=0, columnspan=3, padx=30, pady=20)     # 标签布局
Label(root, text='SheetName:').grid(row=3, column=0, sticky=E)  # 显示SheetName 右对齐
EnValue1 = StringVar()
EnValue2 = StringVar()
EnValue3 = StringVar()
EnValue4 = StringVar()      # 创建字符串变量，跟踪和更新控件中的文本值
entry = Entry(root, show=None, width=20, textvariable=EnValue1)
# 创建单行文本输入框，和 EnValue1相连，使其能够动态更新
entry.grid(row=3, column=1, columnspan=2)                       # 输入框布局
Label(root, text='输出电压：').grid(row=4, column=0, sticky=E)   # 标签组件，显示输出电压，右对齐

Entry(root, show=None, width=10, textvariable=EnValue2, state='readonly').grid(row=4, column=1)
# 创建一个只读的文本输入框，显示电压值，不允许修改
Label(root, text='V').grid(row=4, column=1, sticky=E)   # 标签组件，V，右对齐
Entry(root, show=None, textvariable=EnValue3, state='readonly').place(x=130, y=120, width=200, height=30)
Entry(root, show=None, textvariable=EnValue4, state='readonly').place(x=130, y=160, width=200, height=30)
# 创建两个只读的文本输入框，分别用于显示 EnValue3和 EnValue4的内容
v = IntVar()    # 创建一个整数变量 V
v.set(1)        # 初始值设置为 1

# Label(root, text='选择DPI：').grid(row=3, column=0, sticky=E)
# Radiobutton(root, text='100%', variable=v, value=0).grid(row=3, column=1)
# Radiobutton(root, text='125%', variable=v, value=1).grid(row=3, column=2)
# DPI组件，让用户选择，变量v, 100DPI—————0, 125DPI——————1

Label(root, text='工作模式：').grid(row=6, column=0, sticky=E)
# 标签控件，选择工作模式
x = IntVar()
x.set(0)        # 创建整形变量，用于存储单选按钮的选中值，初始值为0
Radiobutton(root, text='COT模式', variable=x, value=0).grid(row=6, column=1)
Radiobutton(root, text='PWM模式', variable=x, value=1).grid(row=6, column=2)
# 单选按钮两个, COT模式————0, PWM模式——————1
Label(root, text='PG信号：').grid(row=7, column=0, sticky=E)
# 标签控件，显示 PG 信号
pg = IntVar()
pg.set(1)   # 创建一个整型变量存储 PG 信号值，初始化为 1

Radiobutton(root, text='不输出', variable=pg, value=0).grid(row=7, column=2)
Radiobutton(root, text='输出', variable=pg, value=1).grid(row=7, column=1)
# 单选按钮两个, 同步更新PG的值, 不输出————0, 输出——————1
theButton = Button(root, text="选择测试报告路径", command=select_excel_path)
theButton.place(x=10, y=120, width=115, height=30)
# 创建一个按钮控件，用于选择测试报告路径, 调用 select_excel_path函数
theButton = Button(root, text="选择保存图片路径", command=select_pic_path)
theButton.place(x=10, y=160, width=115, height=30)
# 创建一个按钮控件，用于选择保存图片路径, 调用 select_pic_path函数
theButton = Button(root, text="连接仪器", command=instrument)
theButton.place(x=10, y=430, width=80, height=30)
# 创建一个按钮，用于连接仪器，点击时调用 instrument 函数
theButton11 = Button(root, text="读取表格", command=go)
theButton11.place(x=120, y=430, width=80, height=30)
# 创建一个按钮，用于读取表格数据，点击时调用 go 函数。
theButton12 = Button(root, text='保存并退出表格', command=tl11, activeforeground='white', activebackground='red')
theButton12.place(x=230, y=430, width=100, height=30)
# 创建一个按钮，用于保存并退出表格，点击时调用 tl11 函数。按钮的活动前景色和背景色设置为白色和红色。

group = LabelFrame(root, text='POL测试项', padx=5, pady=5)
group.grid(row=0, rowspan=12, column=3, padx=30, pady=15)
# 创建一个分组框（LabelFrame），用于将测试项按钮分组, 标题 POL测试项
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
root.mainloop()     # 启动 Tkinter 的事件循环，使窗口保持显示并响应用户操作