
ScriptVersion = '1.9.0'


import time

import pyvisa

# 在文件开头初始化变量
Data_Acquisition_DAQ973A_id = None
Data_Acquisition_34970A_id = None
Electronic_Load_6312A_id = None
Electronic_Load_63212E_id = None




class Data_Acquisition:
    global date, Data_Acquisition_id


    def __init__(self,Data_Acquisition_id):
        self.rm = pyvisa.ResourceManager()
        self.Data_Acquisition = self.rm.open_resource(Data_Acquisition_id)
        self.Data_Acquisition.timeout = 10000  # 设置超时为10秒

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

class Electronic_Load:
    global Electronic_Load_id
    def __init__(self,Electronic_Load_id):
        rm = pyvisa.ResourceManager()
        self.el = rm.open_resource(Electronic_Load_id)

    def reset(self):
        self.el.write('*RST')

    def channel_choose(self, date):
        self.el.write('CHAN %d' % date)  # 对应的通道选择为1 3 5 7 9 因为每个负载有两个通道

    def mode(self, date):
        self.el.write('MODE %s' % date)

    def state(self, date):
        self.el.write('LOAD %s' % date)  # 对应的通道状态 ON or OFF

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
        if load <= 0.6:
            self.el.write('MODE CCL')
        elif load <= 6:
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



def connection_init():
    global rm, Data_Acquisition_DAQ973A_id, Data_Acquisition_34970A_id, Electronic_Load_6312A_id, Electronic_Load_63212E_id, data_Acquisition, elstate
    data_Acquisition = 0
    elstate = 0
    rm = pyvisa.ResourceManager()
    instrument_list = rm.list_resources()
    print(instrument_list)

    for addr in instrument_list:
        if 'GPIB' in addr:
            ins = rm.open_resource(addr)
            device_id = ins.query('*IDN?').upper()

            # 识别数据采集仪和电子负载设备
            if 'DAQ973A' in device_id:
                Data_Acquisition_DAQ973A_id = addr
                data_Acquisition = 1
            elif '34970A' in device_id:
                Data_Acquisition_34970A_id = addr
                data_Acquisition = 1
            elif 'CHROMA,6312' in device_id:
                Electronic_Load_6312A_id = addr
                elstate = 1
            elif 'CHROMA,63212E' in device_id:
                Electronic_Load_63212E_id = addr
                elstate = 1

    # 打印连接状态
    if data_Acquisition:
        print('数据采集仪已正确连接')
    else:
        print('数据采集仪连接错误，请检查')

    if elstate:
        print('电子负载已正确连接')
    else:
        print('电子负载连接错误，请检查')

connection_init()


# 连接数据采集仪
if Data_Acquisition_34970A_id or Data_Acquisition_DAQ973A_id:
    if Data_Acquisition_34970A_id:
        t = Data_Acquisition(Data_Acquisition_34970A_id)
        t.Channel_Set()
    else:
        t = Data_Acquisition(Data_Acquisition_DAQ973A_id)
        t.Channel_Set()
else:
    print("数据采集仪未连接，请检查连接状态。")

# 连接电子负载
if Electronic_Load_6312A_id or Electronic_Load_63212E_id:
    if Electronic_Load_6312A_id:
        d = Electronic_Load(Electronic_Load_6312A_id)
    else:
        d = Electronic_Load(Electronic_Load_63212E_id)
else:
    print("电子负载未连接，请检查连接状态。")



# 在执行测试前获取输出模式


# End Of imports
print("##############################################################################")
print("------------------------------------------------------------------------------")
print("                              DC LOAD LINE                                    ")
print("------------------------------------------------------------------------------")
print("##############################################################################")
print("")

def generate_sequence(max_value):
    # 计算步长
    step = max_value / 10

    # 生成从0到最大值的数列，每步增加step
    sequence = [round(i * step, 2) for i in range(11)]  # 11表示包括0到最大值共11个数

    return sequence

max_value =6
output_mode = 'near'

test_currents = generate_sequence(max_value)


print("Begining Test")

print("")
print('.............................................................................................................')
# print('  Rail Voltage \t Rail Current\t Rail Ripple \t VR Current In\t VR Voltage In \t IMON')
if output_mode == 'near':
    print('Measured Vout1 (V) \t Measured Vin (V) \t Measured Iin (A)')
elif output_mode == 'far':
    print('Measured Vout1 (V) \t Measured Vin (V) \t Measured Vout2 (V)')
print('.............................................................................................................')

for test_current in test_currents:
    # 设置电流
    d.static(9, 'MAX', test_current)
    d.state('ON')
    time.sleep(3)
    t.Scan_Channel()
    date = t.Read_Date()

    # 更新测量数据的顺序
    Measured_Vout1 = float(date[0])  # 从通道 105
    Measured_Vin = float(date[1])  # 从通道 102
    Measured_Iin = float(date[2]) * 1000   # 从通道 103
    Measured_Vout2 = float(date[3])  # 从通道 104

    # 根据选择的模式，输出不同的数据列
    if output_mode == 'near':
        # 近端模式，输出第一列、第二列、第四列
        print("{:.4f}\t{:.4f}\t{:.4f}".format(Measured_Vout1, Measured_Vin, Measured_Iin))
    elif output_mode == 'far':
        # 远端模式，输出第一列、第二列、第四列（保持与近端模式相同的格式）
        print("{:.4f}\t{:.4f}\t{:.4f}".format(Measured_Vout1, Measured_Vin, Measured_Vout2))

# Turn off the load
if Electronic_Load:
    d.state('OFF')
print("---------------------------------------------------------------------------------------------------------------")
print("")
print("Test complete")
