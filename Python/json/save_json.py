import json

a = {
    "cbBeaAddr": [
        13498904581234,
        56756734981234,
        34456495671234,
        43656798581234
    ]
}
b = json.dumps(a, indent=2) # indent=2避免json一行显示
print(b)
#f2 = open('new.json', 'w')
#f2.write(b)
#f2.close()

