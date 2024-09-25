#!/bin/bash
# command content

# 延时启动
sleep 30

# 进入虚拟环境
source /home/aiko/python/myenv/bin/activate

# 后台运行read.py脚本，并将日志输出到out.log文件
# nohup python3 -u /home/aiko/python/read.py > /home/aiko/python/out.log 2>&1 &
# 后台运行read.py脚本，只记录程序的异常日志
# nohup python3 /home/aiko/python/read.py > /dev/null 2>/home/aiko/python/error.log 2>&1 &
# 后台运行read.py脚本，且不打印日志
nohup python3 /home/aiko/python/read.py > /dev/null 2>&1 &

exit 0
