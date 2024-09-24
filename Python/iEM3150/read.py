import time
import serial
import modbus_tk.defines as cst
from modbus_tk import modbus_rtu
import struct
import threading
import pymysql

# 设定串口从站
master = modbus_rtu.RtuMaster(serial.Serial(port='/dev/ttyUSB0', baudrate=19200, bytesize=8, parity='E', stopbits=1))
master.set_timeout(5.0)  # timeout表示若超过5秒没有连接上slave就会自动断开
master.set_verbose(True)  # 关闭debug的log输出

def run():
    while True:
        threading.Thread(target=data).start()
        time.sleep(3600)

def read():
    red = []
    try:
        # 读取保持寄存器float swapped(modsim32),Most(modscan32),顺序1234(MCGS)
        red1 = readfloat(master.execute(1, cst.HOLDING_REGISTERS, 3059, 2, '>HH'))
        time.sleep(1)
        red2 = readfloat(master.execute(1, cst.HOLDING_REGISTERS, 45099, 2, '>HH'))
        return red1, red2
    except Exception as ext:
        print(str(ext))
        return 0, 0

def readfloat(*args, reverse=True):
    for n, m in args:
        n, m = '%04x' % n, '%04x' % m
    if reverse:
        v = n + m
    else:
        v = m + n
    y_bytes = bytes.fromhex(v)
    y = struct.unpack('!f', y_bytes)[0]
    y = round(y, 3)
    return y

def data():
    # 打开数据库连接
    db = pymysql.connect(host='localhost', user='root', password='1234', database='mysql')

    # 使用 cursor() 方法创建一个游标对象 cursor
    cursor = db.cursor()

    # 获取时间戳
    timestr = time.strftime('%Y-%m-%d %H:%M:%S')

    # 读取power, energy
    power, energy = read()
    print("%s Power: %.3fkW, Energy: %.1fkWh" % (timestr, power, energy))

    # SQL 插入语句
    sql = "INSERT INTO iem3150(date, power, energy) \
           VALUES ('%s', '%.3f',  %.1f)" % (timestr, power, energy)
    try:
        # 执行sql语句
        cursor.execute(sql)
        # 执行sql语句
        db.commit()
    except:
        # 发生错误时回滚
        db.rollback()

    # 关闭数据库连接
    db.close()


if __name__ == '__main__':
    run()
