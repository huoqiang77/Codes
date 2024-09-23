# -*- coding:utf-8 -*-

class CSCheck(object):
    # CS校验
    def CSCal(self, data):
        data = data.replace(' ', '')
        if (len(data) % 2) != 0:
            return 0
        else:
            list_hex = [data[i:i + 2] for i in range(0, len(data), 2)]
            n = len(list_hex)
            list_int = [[] for _ in range(n)]
            sum = 0
            for i in range(0, n, 1):
                list_int[i] = int(list_hex[i], 16)
                sum += list_int[i]
            return str(hex(sum))[-2:]


def Result():
    data = "12 34 56 78"
    check_test = CSCheck()
    print(check_test.CSCal(data))

'''
if __name__ == "__main__":
    Result()
'''