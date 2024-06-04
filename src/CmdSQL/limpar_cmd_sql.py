import os
cmd_sql = []
dict_cmd = {}
newCMD = []

file_path = os.path.join(os.path.dirname(__file__), 'LogComandos.sql')
save_path = os.path.join(os.path.dirname(__file__), 'ComandosLog.txt')

for i in open(file_path,'r'):
    cmd_sql.append(i.split(' '))

for key, lin in enumerate(cmd_sql):
    for i in range(2, len(lin)):
        if i == len(lin)-1:
           newCMD.append(f'{str(lin[i]).replace('\n',';\n')}')
        else:
            newCMD.append(f'{lin[i]}')
    dict_cmd[key] = newCMD.copy()
    newCMD.clear()

with open(save_path,'a') as f:
    for i in dict_cmd:
        for j in dict_cmd[i]:
            f.write(f' {j}')
