#!/bin/bash

### BEGIN INIT INFO
# Provides:     startup
# Required-Start:  $remote_fs $syslog
# Required-Stop:   $remote_fs $syslog
# Default-Start:   2 3 4 5
# Default-Stop:   0 1 6
# Short-Description: start startup
# Description:    start startup
### END INIT INFO

bash /home/aiko/python/run.sh
exit 0