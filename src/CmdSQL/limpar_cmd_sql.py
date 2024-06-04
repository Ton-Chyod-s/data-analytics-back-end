import os
cmd_sql = []
dict_cmd = {}
newCMD = []

for i in open(os.path.abspath('LogComandos.txt'),'r'):
    cmd_sql.append(i.split(' '))

for key, lin in enumerate(cmd_sql):
    for i in range(2, len(lin)):
        if i == len(lin)-1:
           newCMD.append(f'{str(lin[i]).replace('\n',';\n')}')
        else:
            newCMD.append(f'{lin[i]}')
    dict_cmd[key] = newCMD.copy()
    newCMD.clear()

with open('ComandosLog.txt','a') as f:
    for i in dict_cmd:
        for j in dict_cmd[i]:
            f.write(f' {j}')

#     for lin in cmd_sql:
#         for i in range(2, len(lin)):
#             if i == len(lin)-1:
#                 f.write(f' {lin[i]};')
#                 break
#             else:
#                 f.write(f' {lin[i]}')

lol='lol'