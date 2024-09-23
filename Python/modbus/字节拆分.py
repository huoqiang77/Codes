n = 10240099  # 0x009c4063
b1 = (n & 0xff000000) >> 24
b2 = (n & 0xff0000) >> 16
b3 = (n & 0xff00) >> 8
b4 = n & 0xff
# bs = bytes([b1, b2, b3, b4])
bs = bytes([0x01, 0x02, 0x21, 0xA1])
print(hex(b1), hex(b2), hex(b3), hex(b4))
print(bs)
