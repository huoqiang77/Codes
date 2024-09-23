import json

f = open('new.json', 'r')
content = f.read()
a = json.loads(content)
b = a['cbBeaAddr']
list = []
for i in b:
    list.append(i)
print(list)
f.close()