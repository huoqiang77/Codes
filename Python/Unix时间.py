import time
import datetime

# 转成Unix时间
timeNow = datetime.datetime.now()
unix_time = time.mktime(timeNow.timetuple())
print(timeNow)
print(unix_time)
print(hex(int(unix_time)))

# 转自Unix时间
unix_time_set = 1650243919
timeNow = datetime.datetime.fromtimestamp(unix_time_set)
print(timeNow)
print(type(timeNow))